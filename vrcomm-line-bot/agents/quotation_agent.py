"""
agents/quotation_agent.py — VRCOMM Quotation Orchestrator Agent

Architecture:
  Quotation Agent (orchestrator) ──────calls──────▶ Product Agent
                                                        │
                                                  VRCOMM_CostSheet.xlsx
                                                        │
                                              return {found, missing}
                                                        │
  Quotation Agent ◀─────────────────────────────────────┘
        │
        ├── all costs found  → generate Excel → notify Admin
        └── costs missing    → push Admin alert (update CostSheet)
                             → tell user which items couldn't be priced

Conversation collects: customer_name + items (brand, product, qty)
Cost lookup is automatic via product_agent.get_cost_sheet()
User NEVER needs to enter unit_cost_thb manually.
"""
import os, re, json, logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── System prompts ────────────────────────────────────────────────────────────

_EXTRACT_SYSTEM = """You are a data extractor for VRCOMM's quotation system.
Given a conversation history, extract all available information.

Return a JSON object with these fields (use null for missing ones):
{
  "customer_name": "company name of the customer" or null,
  "items": [
    {"brand": "...", "product": "...", "qty": <integer>}
  ] or [],
  "margin_pct": <number> or null,
  "validity_days": <number> or null,
  "notes": "special notes" or null
}

Rules:
- qty must be a positive integer (default 1 if not stated)
- brand and product should be as specific as possible from the message
- Do NOT include unit_cost_thb — that is looked up automatically
- margin_pct: if user says "30% margin" → 30; leave null if not mentioned
- Return ONLY raw JSON. No explanation, no markdown fences.
"""

_CONVERSATION_SYSTEM = """You are VRCOMM's internal Quotation Assistant — helping sales staff create customer quotations quickly.
VRCOMM is a Network and Cybersecurity solutions company in Thailand.

Your job: collect the minimum information needed, then trigger quotation generation.

Required information:
1. Customer company name (ชื่อบริษัทลูกค้า)
2. Line items: brand, product/model name, quantity (จำนวน)
   — DO NOT ask for price/cost — the system looks it up automatically
3. Margin % — OPTIONAL, suggest 30% if not stated
4. Validity days — OPTIONAL, suggest 30 days if not stated

Rules:
- Ask for ONE missing piece at a time
- Once you have customer_name + at least one item (brand + product + qty) → trigger generation
  Use exactly this marker on its own line: [READY_TO_GENERATE]
- Suggest margin=30% and validity=30 days as defaults (don't make user provide them)
- Tone: warm, efficient, internal colleague
- Reply in the SAME language as the user (Thai → Thai, English → English)
- Plain text only — no markdown, no bullet points

Example — when ready:
[READY_TO_GENERATE]
สรุป: ลูกค้า ABC Corp, Sangfor SSL VPN 100U x2, margin 30%, valid 30 วัน
"""

# ── Data extraction ────────────────────────────────────────────────────────────

def _extract_quote_data(history: list, message: str) -> dict:
    """Use Haiku to parse conversation → structured quote data. No cost field."""
    conv_text = ""
    for turn in history:
        role    = "Staff" if turn.get("role") == "user" else "Bot"
        conv_text += "%s: %s\n" % (role, turn.get("content", ""))
    if message.strip():
        conv_text += "Staff: %s\n" % message

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            system=_EXTRACT_SYSTEM,
            messages=[{"role": "user", "content": conv_text}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.MULTILINE)
        raw = re.sub(r'\s*```$',          '', raw, flags=re.MULTILINE)
        data = json.loads(raw)
        logger.info("[quotation_agent] extracted: %s", str(data)[:200])
        return data
    except Exception as e:
        logger.error("[quotation_agent] extraction error: %s", e)
        return {"customer_name": None, "items": [], "margin_pct": None,
                "validity_days": None, "notes": None}


def _has_enough_to_proceed(data: dict) -> bool:
    """
    Ready to hand off to Product Agent when:
      - customer_name is known
      - at least one item has brand + product + qty
    Cost is NOT needed from user — Product Agent fetches it.
    """
    if not data.get("customer_name"):
        return False
    for item in data.get("items", []):
        if item.get("brand") and item.get("product") and item.get("qty"):
            return True
    return False


# ── Cost fetch via Product Agent ──────────────────────────────────────────────

def _fetch_costs(items: list) -> dict:
    """
    Call product_agent.get_cost_sheet() to look up cost prices.
    Returns {"found": [...], "missing": [...]}
    """
    try:
        from agents.product_agent import get_cost_sheet
        result = get_cost_sheet(items)
        logger.info("[quotation_agent] cost fetch: %d found, %d missing",
                    len(result["found"]), len(result["missing"]))
        return result
    except Exception as e:
        logger.error("[quotation_agent] cost fetch error: %s", e)
        return {"found": [], "missing": items}


# ── Orchestration: cost check → generate or alert ─────────────────────────────

def _orchestrate(data: dict, user_name: str, user_id: str, source: str) -> str:
    """
    1. Call Product Agent for costs
    2a. All found → generate Excel → notify Admin → return confirm message
    2b. Some missing → notify Admin to update CostSheet → return partial message
    """
    margin_pct    = float(data.get("margin_pct") or 30)
    validity_days = int(data.get("validity_days") or 30)
    customer      = data["customer_name"]
    items         = data.get("items", [])

    # ── Step: fetch costs ─────────────────────────────────────────────────────
    cost_result = _fetch_costs(items)
    found       = cost_result["found"]
    missing     = cost_result["missing"]

    # ── Case A: all items priced → full generation ────────────────────────────
    if found and not missing:
        return _generate_and_notify(
            customer=customer, found=found,
            margin_pct=margin_pct, validity_days=validity_days,
            notes=data.get("notes") or "",
            user_name=user_name,
        )

    # ── Case B: partial (some found, some missing) ────────────────────────────
    if found and missing:
        # Generate quote with what we have, flag the missing ones
        gen_reply = _generate_and_notify(
            customer=customer, found=found,
            margin_pct=margin_pct, validity_days=validity_days,
            notes=data.get("notes") or "",
            user_name=user_name,
        )
        missing_list = ", ".join(
            "%s %s" % (m["brand"], m["product"]) for m in missing
        )
        _push_missing_cost_alert(customer, missing, user_name)
        return (
            gen_reply + "\n\n"
            "⚠️ หมายเหตุ: สินค้าต่อไปนี้ไม่พบราคาทุนใน CostSheet จึงไม่รวมในใบเสนอราคา:\n"
            "%s\n\n"
            "แจ้ง PM อัพเดทไฟล์ VRCOMM_CostSheet.xlsx แล้วครับ"
        ) % missing_list

    # ── Case C: nothing found ─────────────────────────────────────────────────
    missing_list = ", ".join("%s %s" % (m["brand"], m["product"]) for m in missing)
    _push_missing_cost_alert(customer, missing, user_name)
    return (
        "⚠️ ไม่พบราคาทุนสำหรับสินค้าที่ขอทั้งหมดใน VRCOMM_CostSheet.xlsx:\n%s\n\n"
        "แจ้ง PM ให้อัพเดทราคาทุนในไฟล์ product/VRCOMM_CostSheet.xlsx แล้วครับ\n"
        "เมื่อ PM อัพเดทแล้ว ส่งคำขอ quote ใหม่ได้เลยครับ"
    ) % missing_list


def _generate_and_notify(customer: str, found: list,
                         margin_pct: float, validity_days: int,
                         notes: str, user_name: str) -> str:
    """Generate Excel quotation and push LINE notification to admin."""
    from quotation_generator import generate_quotation

    try:
        result = generate_quotation(
            customer_name=customer,
            items=found,
            margin_pct=margin_pct,
            validity_days=validity_days,
            notes=notes,
            prepared_by=user_name,
        )
    except Exception as e:
        logger.error("[quotation_agent] generate error: %s", e)
        return "เกิดข้อผิดพลาดในการสร้างใบเสนอราคาครับ กรุณาติดต่อทีม IT"

    quote_no    = result["quote_no"]
    grand_total = result["grand_total"]
    filename    = result["filename"]

    base_url     = os.environ.get("APP_BASE_URL", "https://vrcomm-line.onrender.com")
    download_url = "%s/download/quotation/%s" % (base_url, filename)

    _push_admin_notification(
        quote_no=quote_no, customer=customer, items=found,
        margin_pct=margin_pct, grand_total=grand_total,
        prepared_by=user_name, download_url=download_url,
    )

    total_fmt = "{:,.0f}".format(grand_total)
    return (
        "✅ สร้างใบเสนอราคาเรียบร้อยแล้วครับ\n\n"
        "เลขที่  : %s\n"
        "ลูกค้า  : %s\n"
        "ยอดรวม : %s THB (รวม VAT 7%%)\n\n"
        "📎 ดาวน์โหลด:\n%s\n\n"
        "แจ้ง Admin ตรวจสอบและส่งให้ลูกค้าแล้วครับ 🎯"
    ) % (quote_no, customer, total_fmt, download_url)


# ── LINE push helpers ──────────────────────────────────────────────────────────

def _push_admin_notification(quote_no, customer, items, margin_pct,
                              grand_total, prepared_by, download_url):
    """Push new quotation draft to admin LINE."""
    try:
        from linebot import LineBotApi
        from linebot.models import TextSendMessage

        token    = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
        admin_id = os.environ.get("ADMIN_LINE_USER_ID", "")
        if not token or not admin_id:
            logger.warning("[quotation_agent] admin push skipped — env not set")
            return

        items_text = "\n".join(
            "  %d. %s %s x%d (cost: %s THB)" % (
                i + 1,
                it.get("brand", ""), it.get("matched_product") or it.get("product", ""),
                it.get("qty", 1),
                "{:,.0f}".format(it.get("unit_cost_thb", 0)),
            )
            for i, it in enumerate(items)
        )
        total_fmt = "{:,.0f}".format(grand_total)

        msg = (
            "📋 ใบเสนอราคาใหม่ — รอตรวจสอบ\n\n"
            "เลขที่    : %s\n"
            "ลูกค้า    : %s\n"
            "จัดทำโดย : %s\n"
            "Margin   : %.0f%%\n"
            "ยอดรวม   : %s THB (รวม VAT)\n\n"
            "สินค้า (จาก CostSheet):\n%s\n\n"
            "📎 ดาวน์โหลด:\n%s"
        ) % (quote_no, customer, prepared_by, margin_pct, total_fmt, items_text, download_url)

        LineBotApi(token).push_message(admin_id, TextSendMessage(text=msg))
        logger.info("[quotation_agent] admin notified for %s", quote_no)
    except Exception as e:
        logger.error("[quotation_agent] admin push error: %s", e)


def _push_missing_cost_alert(customer: str, missing: list, requested_by: str):
    """Push alert to admin when some products are not in CostSheet."""
    try:
        from linebot import LineBotApi
        from linebot.models import TextSendMessage

        token    = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
        admin_id = os.environ.get("ADMIN_LINE_USER_ID", "")
        if not token or not admin_id:
            return

        missing_text = "\n".join(
            "  - %s %s x%d" % (m["brand"], m["product"], m.get("qty", 1))
            for m in missing
        )
        msg = (
            "⚠️ ราคาทุนหายไปจาก CostSheet\n\n"
            "Quote สำหรับ: %s\n"
            "ขอโดย: %s\n\n"
            "สินค้าที่ไม่พบราคาทุน:\n%s\n\n"
            "กรุณาอัพเดทไฟล์:\nproduct/VRCOMM_CostSheet.xlsx\n"
            "แล้วให้ staff ส่งคำขอ quote ใหม่อีกครั้งครับ"
        ) % (customer, requested_by, missing_text)

        LineBotApi(token).push_message(admin_id, TextSendMessage(text=msg))
        logger.info("[quotation_agent] missing-cost alert sent for %d items", len(missing))
    except Exception as e:
        logger.error("[quotation_agent] missing-cost alert error: %s", e)


# ── Cost sheet parser (for PM-uploaded Excel files) ───────────────────────────

def parse_cost_sheet(filepath: str) -> list:
    """
    Parse a PM-provided cost sheet Excel file uploaded by user.
    Flexible column detection: Brand, Product/Model, Qty, Cost.
    Returns list of {brand, product, qty, unit_cost_thb}.
    """
    import openpyxl
    wb   = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # Detect header row
    header_row_idx = 0
    headers = []
    for i, row in enumerate(rows):
        non_empty = [str(c).strip().lower() for c in row if c is not None]
        if len(non_empty) >= 3:
            headers = non_empty
            header_row_idx = i
            break

    col_map = {}
    kw_map  = {
        "brand":   ["brand", "vendor", "ยี่ห้อ"],
        "product": ["product", "model", "description", "item", "รุ่น", "สินค้า"],
        "qty":     ["qty", "quantity", "จำนวน"],
        "cost":    ["cost", "unit cost", "ราคาทุน", "unit price", "price"],
    }
    for key, keywords in kw_map.items():
        for i, h in enumerate(headers):
            if any(kw in h for kw in keywords):
                col_map[key] = i
                break

    items = []
    for row in rows[header_row_idx + 1:]:
        if not any(row):
            continue
        try:
            brand   = str(row[col_map["brand"]]).strip()   if "brand"   in col_map else ""
            product = str(row[col_map["product"]]).strip() if "product" in col_map else ""
            qty     = int(float(str(row[col_map["qty"]]))) if "qty" in col_map and row[col_map["qty"]] else 1
            cost_v  = row[col_map["cost"]]                 if "cost"    in col_map else None
            cost    = float(str(cost_v).replace(",", ""))  if cost_v else None

            if not brand and not product:
                continue
            items.append({"brand": brand, "product": product, "qty": qty, "unit_cost_thb": cost})
        except Exception as e:
            logger.warning("[quotation_agent] parse_cost_sheet row error: %s", e)

    logger.info("[quotation_agent] parsed %d items from uploaded file", len(items))
    return items


# ── Main handler ───────────────────────────────────────────────────────────────

def _build_context_summary(data: dict) -> str:
    lines = []
    if data.get("customer_name"):
        lines.append("customer: %s" % data["customer_name"])
    for it in data.get("items", []):
        lines.append("  item: %s %s x%s" % (
            it.get("brand", "?"), it.get("product", "?"), it.get("qty", "?")))
    if not lines:
        return ""
    return "\nExtracted so far:\n" + "\n".join(lines)


def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "quotation",
           cost_sheet_data: list = None,
           **kwargs) -> str:
    """
    Quotation orchestrator.
    Collects customer + items via conversation, then calls Product Agent
    for cost lookup, then generates Excel and notifies admin.

    cost_sheet_data: pre-parsed items from an uploaded file (passed by app.py)
    """
    if history is None:
        history = []

    if cost_sheet_data:
        items_text = "\n".join(
            "%s %s x%d" % (it.get("brand", ""), it.get("product", ""), it.get("qty", 1))
            for it in cost_sheet_data
        )
        message = "[Cost Sheet Uploaded — %d items]\n%s\n\n%s" % (
            len(cost_sheet_data), items_text, message
        )
        logger.info("[quotation_agent] %d items from uploaded cost sheet", len(cost_sheet_data))

    data = _extract_quote_data(history, message)

    if _has_enough_to_proceed(data):
        return _orchestrate(data, user_name, user_id, source)

    ctx      = _build_context_summary(data)
    messages = history + [{"role": "user", "content": message}]

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=400,
            system=_CONVERSATION_SYSTEM + ctx,
            messages=messages,
        )
        reply = resp.content[0].text.strip()

        if "[READY_TO_GENERATE]" in reply:
            final_data = _extract_quote_data(
                history + [{"role": "user", "content": message}], ""
            )
            if _has_enough_to_proceed(final_data):
                preamble = reply.replace("[READY_TO_GENERATE]", "").strip()
                gen      = _orchestrate(final_data, user_name, user_id, source)
                return (preamble + "\n\n" + gen).strip() if preamble else gen
            reply = reply.replace("[READY_TO_GENERATE]", "").strip()

        return reply

    except Exception as e:
        logger.error("[quotation_agent] Claude error: %s", e)
        return "ขออภัยครับ ระบบขัดข้องชั่วคราว กรุณาลองใหม่อีกครั้งครับ"
