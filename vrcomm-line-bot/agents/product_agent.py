"""
agents/product_agent.py — VRCOMM Product Information Agent

Data source: vrcomm-line-bot/product/VRCOMM_ProductList.xlsx
  - Column A: Brand/Product name
  - Column B: Website URL

Two-Step architecture (prevents Claude from hallucinating non-listed brands):

  STEP 1 — SELECT:
    Ask Claude: "From ONLY these 21 brands, which are relevant to this query?"
    → Returns e.g. ["Hillstone Networks", "Sangfor", "Safetica", "Varonis"]
    → Claude physically cannot pick Fortinet/Cisco because they're not in the list

  STEP 2 — ANSWER:
    Build system prompt with ONLY the selected brands.
    Claude answers without ever seeing non-listed brands in context.
"""
import os, re, logging, time
import requests
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

_BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PRODUCT_LIST = os.path.join(_BASE_DIR, "product", "VRCOMM_ProductList.xlsx")

# ── Product list cache ────────────────────────────────────────────────────────

_product_cache      = None
_product_cache_time = 0.0
_CACHE_TTL          = 300


def _load_product_list() -> list:
    """Load VRCOMM_ProductList.xlsx → list of {brand, url}. Cached 5 min."""
    global _product_cache, _product_cache_time

    path = _PRODUCT_LIST
    if not os.path.isfile(path):
        logger.error("ProductList not found: %s", path)
        return []

    mtime = os.path.getmtime(path)
    now   = time.time()
    if (_product_cache is not None
            and mtime == _product_cache_time
            and (now - _product_cache_time) < _CACHE_TTL):
        return _product_cache

    try:
        import openpyxl
        wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws   = wb.active
        products = []
        for row in ws.iter_rows(values_only=True):
            brand = str(row[0]).strip() if row[0] else ""
            url   = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            if brand.lower() in ("brand", "product", "name", "") or not brand:
                continue
            products.append({"brand": brand, "url": url})

        _product_cache      = products
        _product_cache_time = mtime
        logger.info("ProductList loaded: %d brands", len(products))
        return products
    except Exception as e:
        logger.error("ProductList load error: %s", e)
        return []


# ── STEP 1: Brand selector ────────────────────────────────────────────────────

_SELECT_PROMPT = """You are a product selector for VRCOMM, a cybersecurity company in Thailand.

A customer sent this query:
\"{message}\"

Below is the COMPLETE list of brands VRCOMM sells:
{brand_list}

Task: Select which brands from the list above are relevant to the customer's query.
- Consider the product category (firewall, endpoint, DLP, switch, backup, PKI, etc.)
- Select ALL relevant brands, including alternatives and combined solutions
- Return ONLY the exact brand names from the list above, one per line
- If truly nothing is relevant, return exactly: NONE
- Do NOT add brands outside the list. Do NOT add explanations."""


def _select_relevant_brands(message: str, product_list: list) -> list:
    """
    Step 1: Ask Claude to pick relevant brands from our list only.
    Returns list of matched {brand, url} dicts.
    """
    brand_list_text = "\n".join(
        "%d. %s" % (i + 1, p["brand"]) for i, p in enumerate(product_list)
    )

    prompt = _SELECT_PROMPT.format(
        message=message,
        brand_list=brand_list_text,
    )

    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        logger.info("[product_agent] Step1 selection: %s", raw[:120])

        if raw.strip().upper() == "NONE" or not raw:
            return []

        # Parse returned lines, strip numbering/bullets
        selected_names = [
            re.sub(r'^[\d\.\-\)\s]+', '', line).strip()
            for line in raw.splitlines()
            if line.strip() and line.strip().upper() != "NONE"
        ]

        # Match back to actual product list (case-insensitive)
        matched = []
        for name in selected_names:
            for p in product_list:
                if (p["brand"].lower() == name.lower()
                        or name.lower() in p["brand"].lower()
                        or p["brand"].lower() in name.lower()):
                    if p not in matched:
                        matched.append(p)

        logger.info("[product_agent] Step1 matched brands: %s",
                    [m["brand"] for m in matched])
        return matched

    except Exception as e:
        logger.error("[product_agent] Step1 selection error: %s", e)
        return []


# ── URL content fetcher (cached 1 hr) ────────────────────────────────────────

_url_cache: dict = {}
_URL_CACHE_TTL   = 3600


def _fetch_url(url: str) -> str:
    """Fetch URL → clean plain text. Cached 1 hour."""
    url = url.strip()
    if not url:
        return ""

    now    = time.time()
    cached = _url_cache.get(url)
    if cached and (now - cached["fetched_at"]) < _URL_CACHE_TTL:
        return cached["content"]

    try:
        resp = requests.get(url, headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9,th;q=0.8",
        }, timeout=8, allow_redirects=True)
        resp.raise_for_status()

        html = resp.text
        html = re.sub(r'<(script|style|nav|footer|header)[^>]*>.*?</\1>',
                      ' ', html, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<[^>]+>', ' ', html)
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n{3,}', '\n\n', text).strip()[:3000]

        _url_cache[url] = {"content": text, "fetched_at": now}
        logger.info("[product_agent] Fetched %d chars: %s", len(text), url[:60])
        return text
    except Exception as e:
        logger.warning("[product_agent] URL fetch failed [%s]: %s", url[:60], e)
        return ""


# ── Forbidden brands (post-processing safety net) ────────────────────────────

_FORBIDDEN_BRANDS = [
    "fortinet", "fortigate", "cisco", "palo alto", "check point", "checkpoint",
    "juniper", "sonicwall", "sonic wall", "barracuda", "trend micro", "trendmicro",
    "crowdstrike", "crowd strike", "sentinelone", "sentinel one", "mcafee",
    "symantec", "broadcom", "hp ", "hewlett", "dell ", "aruba", "meraki",
    "watchguard", "watch guard", "zyxel", "netgear", "ubiquiti",
]


def _contains_forbidden_brand(text: str) -> str:
    """Return the first forbidden brand found in text, or empty string."""
    text_lower = text.lower()
    for brand in _FORBIDDEN_BRANDS:
        if brand in text_lower:
            return brand
    return ""


# ── STEP 2: Answer prompt (internal staff tone) ───────────────────────────────

_ANSWER_SYSTEM_WITH_MATCH = """You are a VRCOMM pre-sales engineer briefing an internal staff member.
This is an INTERNAL conversation — not customer-facing.

The relevant VRCOMM products for this query are listed below.
These are the ONLY products you may discuss. Period.

{selected_brands_section}

HARD RULES:
- You may ONLY mention the brands listed above in your answer
- FORBIDDEN brands (never mention these under any circumstances):
  Fortinet, FortiGate, Cisco, Palo Alto, Check Point, Juniper, SonicWall,
  Barracuda, Trend Micro, CrowdStrike, SentinelOne, or ANY brand not listed above
- If you are tempted to suggest a product not in the list above — don't. Pick the closest match from the list instead.
- NEVER quote specific prices — say: "ให้ Sales ดึง cost sheet แล้วทำ quote ได้เลยครับ"
- Tone: direct, technical, colleague-to-colleague — not customer service language
- Reply in the SAME LANGUAGE as the staff member (Thai → Thai, English → English)
- Plain text only — no markdown, no bullets, no headers
- Max 4-5 short paragraphs"""

_ANSWER_SYSTEM_NO_MATCH = """You are a VRCOMM pre-sales engineer briefing an internal staff member.
This is an INTERNAL conversation — not customer-facing.

The staff member is asking about a product category or brand that VRCOMM does not carry.

VRCOMM's complete product portfolio (ALL brands we sell — nothing else):
{full_list}

Your task:
1. Confirm that VRCOMM does not carry the requested brand/product
2. Recommend the best alternative(s) from the list above for the same use case
3. If relevant, propose a combined solution using multiple brands from the list

HARD RULES:
- Recommend ONLY brands from the list above — nothing outside it
- FORBIDDEN: Fortinet, Cisco, Palo Alto, Check Point, or any brand NOT in the list
- NEVER quote prices — say: "ให้ Sales ดึง cost sheet แล้วทำ quote ได้เลยครับ"
- Tone: direct, technical, internal briefing — not customer service
- Reply in the SAME LANGUAGE as the staff member
- Plain text only — max 4-5 short paragraphs"""


def _build_answer_system(selected: list, product_list: list) -> str:
    if selected:
        # Build section with name + fetched web content
        sections = []
        for p in selected:
            content = _fetch_url(p["url"])
            section = "Brand: %s\nWebsite: %s" % (p["brand"], p["url"])
            if content:
                section += "\nProduct info:\n%s" % content[:1500]
            sections.append(section)
        selected_brands_section = "\n\n---\n".join(sections)
        return _ANSWER_SYSTEM_WITH_MATCH.format(
            selected_brands_section=selected_brands_section
        )
    else:
        full_list = "\n".join(
            "%d. %s" % (i + 1, p["brand"]) for i, p in enumerate(product_list)
        )
        return _ANSWER_SYSTEM_NO_MATCH.format(full_list=full_list)


# ── Main handler ──────────────────────────────────────────────────────────────

def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "product_info", **kwargs) -> str:
    """
    Two-step product handler:
      Step 1 — SELECT: Claude picks relevant brands from our list
      Step 2 — ANSWER: Claude answers using only those brands
    """
    if history is None:
        history = []

    # Load product list
    product_list = _load_product_list()
    if not product_list:
        logger.warning("[product_agent] empty product list — falling back to general")
        from agents.general_agent import handle as general_handle
        return general_handle(message=message, user_name=user_name,
                              user_id=user_id, source=source,
                              history=history, intent=intent)

    # STEP 1 — select relevant brands from our list
    selected = _select_relevant_brands(message, product_list)

    # STEP 2 — answer using only selected brands (or suggest alternatives)
    system = _build_answer_system(selected, product_list)

    # Internal framing — staff, not customer
    src_note = " (via email)" if source == "email" else ""
    content  = message if history else (
        "Staff: %s%s\n\nQuery: %s" % (user_name, src_note, message)
    )
    messages = history + [{"role": "user", "content": content}]

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=768,
            system=system,
            messages=messages,
        )
        reply = response.content[0].text.strip()

        # Post-processing safety net — retry once if forbidden brand detected
        forbidden = _contains_forbidden_brand(reply)
        if forbidden:
            logger.warning("[product_agent] Forbidden brand '%s' in reply — retrying", forbidden)
            retry_system = system + (
                "\n\nCRITICAL: Your previous answer mentioned '%s' which is NOT in VRCOMM's product list. "
                "Rewrite the answer using ONLY the VRCOMM brands provided above." % forbidden
            )
            retry_resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=768,
                system=retry_system,
                messages=messages,
            )
            reply = retry_resp.content[0].text.strip()
            logger.info("[product_agent] Retry reply: %s", reply[:80])

        logger.info("[product_agent] replied to %s (selected=%s): %s",
                    user_name, [s["brand"] for s in selected], reply[:80])
        return reply

    except Exception as e:
        logger.error("[product_agent] Claude API error: %s", e)
        return ("ขออภัยครับ ระบบขัดข้องชั่วคราว "
                "กรุณาติดต่อทีม VRCOMM ที่ iddhi.t@vrcomm.net ครับ")
