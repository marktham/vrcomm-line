"""
agents/subscription_agent.py — VRCOMM Subscription Management Agent

Data source: VRCOMM_Subscriptions.xlsx (place in vrcomm-line-bot/ folder)
Sheet: ATR2027 (or first sheet)

Handles:
  - Check expiry status for an account
  - List subscriptions expiring soon (with 🔴/🟡/🟢 status)
  - Summary of all active subscriptions
  - Renewal information and LIV$ reference
"""
import os, logging, time
from datetime import datetime, date
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── Paths ─────────────────────────────────────────────────────────────────────

_BASE_DIR       = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SUB_FILE       = os.environ.get(
    "SUBSCRIPTION_PATH",
    os.path.join(_BASE_DIR, "VRCOMM_Subscriptions.xlsx")
)

# ── Subscription cache ────────────────────────────────────────────────────────

_sub_cache      = None   # list of subscription dicts
_sub_cache_time = 0.0
_CACHE_TTL      = 300    # 5 minutes


def _load_subscriptions() -> list:
    """
    Load all subscriptions from Excel.
    Returns list of dicts with keys matching column headers.
    Cached for 5 minutes.
    """
    global _sub_cache, _sub_cache_time

    path = _SUB_FILE
    if not os.path.isfile(path):
        logger.error("Subscriptions file not found: %s", path)
        return []

    mtime = os.path.getmtime(path)
    now   = time.time()

    if (_sub_cache is not None
            and mtime == _sub_cache_time
            and (now - _sub_cache_time) < _CACHE_TTL):
        return _sub_cache

    try:
        import openpyxl
        wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            return []

        header = [str(h).strip() if h else "" for h in rows[0]]
        subs   = []
        for row in rows[1:]:
            if not any(row):
                continue
            record = dict(zip(header, row))
            # Normalise expire date to a date object
            exp = record.get("Expire Date")
            if isinstance(exp, datetime):
                record["_expire_date"] = exp.date()
            elif isinstance(exp, date):
                record["_expire_date"] = exp
            else:
                record["_expire_date"] = None
            subs.append(record)

        _sub_cache      = subs
        _sub_cache_time = mtime
        logger.info("Subscriptions loaded: %d records", len(subs))
        return subs

    except Exception as e:
        logger.error("Subscription load error: %s", e)
        return []


# ── Days-to-expiry helpers ────────────────────────────────────────────────────

def _days_to_expiry(exp_date) -> int | None:
    """Return days until expiry from today. Negative = already expired."""
    if exp_date is None:
        return None
    today = date.today()
    return (exp_date - today).days


def _status_emoji(days: int | None) -> str:
    if days is None:
        return "⬜"
    if days < 0:
        return "🔴"   # expired
    if days <= 30:
        return "🔴"   # critical
    if days <= 60:
        return "🟡"   # warning
    return "🟢"        # ok


# ── Format subscriptions as structured text for Claude ───────────────────────

def _format_subscriptions(subs: list, filter_account: str = None) -> str:
    """
    Format subscription records as readable text for Claude context.
    Optionally filter by account name (case-insensitive partial match).
    """
    today = date.today()
    lines = ["=== VRCOMM Subscription Data (as of %s) ===" % today.isoformat()]

    # Filter by account if specified
    if filter_account:
        query  = filter_account.lower()
        subset = [s for s in subs if query in str(s.get("Account", "")).lower()]
        if not subset:
            # Try partial word matching
            subset = [s for s in subs
                      if any(w in str(s.get("Account", "")).lower()
                             for w in query.split())]
        display_subs = subset if subset else subs
        if filter_account and not subset:
            lines.append("(ไม่พบบัญชี '%s' — แสดงทั้งหมดแทน)" % filter_account)
    else:
        display_subs = subs

    # Group by account for readability
    accounts: dict = {}
    for s in display_subs:
        acct = s.get("Account", "Unknown")
        accounts.setdefault(acct, []).append(s)

    for acct, records in sorted(accounts.items()):
        lines.append("\nAccount: %s" % acct)
        for r in records:
            days    = _days_to_expiry(r["_expire_date"])
            emoji   = _status_emoji(days)
            exp_str = r["_expire_date"].isoformat() if r["_expire_date"] else "N/A"
            days_str = ("%d days" % days) if days is not None else "N/A"
            if days is not None and days < 0:
                days_str = "EXPIRED %d days ago" % abs(days)

            lines.append(
                "  %s Sub: %-15s | Product: %-35s | QTY: %-4s | "
                "Expire: %s (%s) | LIV$: %s | Status: %s" % (
                    emoji,
                    str(r.get("Subscription", ""))[:15],
                    str(r.get("Product Name", ""))[:35],
                    str(r.get("QTY", "")),
                    exp_str,
                    days_str,
                    str(r.get("LIV$ (expiring)", "")),
                    str(r.get("Sophos Status", "")),
                )
            )

    # Summary counts
    all_days = [_days_to_expiry(s["_expire_date"]) for s in display_subs]
    critical = sum(1 for d in all_days if d is not None and d <= 30)
    warning  = sum(1 for d in all_days if d is not None and 31 <= d <= 60)
    lines.append(
        "\nSummary: %d subscriptions | 🔴 Critical (≤30d): %d | 🟡 Warning (31-60d): %d"
        % (len(display_subs), critical, warning)
    )

    return "\n".join(lines)


# ── System prompt ─────────────────────────────────────────────────────────────

_SYSTEM = """You are VRCOMM's Subscription Administrator.
VRCOMM is a Network and Cybersecurity solutions provider in Thailand.
You manage Sophos subscription renewals for VRCOMM's customers.

Use the subscription data below to answer questions accurately.

Status legend:
  🔴 Critical — expires within 30 days or already expired
  🟡 Warning  — expires within 31-60 days
  🟢 OK       — expires in more than 60 days

Your capabilities:
1. Check subscription status for any account
2. List subscriptions expiring soon
3. Provide renewal information using LIV$ as a reference cost
4. Recommend renewal action based on urgency

RULES:
- NEVER quote a final customer price — LIV$ is internal cost reference only
  If asked for a renewal price, say: "ทีม Sales จะจัดทำใบเสนอราคา renewal ให้ครับ"
- Reply in the SAME LANGUAGE as the user (Thai → Thai, English → English)
- Be precise with dates and account names
- Plain text only — no markdown, no headers, no bullet symbols

{sub_data}"""


# ── Main handler ──────────────────────────────────────────────────────────────

def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "subscription", **kwargs) -> str:
    """
    Handle subscription queries.

    1. Load subscriptions from Excel
    2. Try to detect if a specific account is mentioned
    3. Format relevant subscription data
    4. Claude answers with accurate dates/status
    """
    if history is None:
        history = []

    subs = _load_subscriptions()
    if not subs:
        return ("ขออภัยครับ ไม่สามารถโหลดข้อมูล Subscription ได้ในขณะนี้ "
                "กรุณาติดต่อทีม VRCOMM โดยตรงครับ")

    # Try to detect account name from the message
    # (Claude will handle fuzzy interpretation from the full data)
    account_hint = _extract_account_hint(message, subs)

    sub_data = _format_subscriptions(subs, filter_account=account_hint)
    system   = _SYSTEM.format(sub_data=sub_data)

    if history:
        content = message
    else:
        src_note = " (via email)" if source == "email" else ""
        content  = "Staff: %s%s\n\n%s" % (user_name, src_note, message)

    messages = history + [{"role": "user", "content": content}]

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=768,
            system=system,
            messages=messages,
        )
        reply = response.content[0].text.strip()
        logger.info("[subscription_agent] replied to %s (account_hint=%s): %s",
                    user_name, account_hint, reply[:80])
        return reply

    except Exception as e:
        logger.error("[subscription_agent] Claude API error: %s", e)
        return ("ขออภัยครับ ระบบขัดข้องชั่วคราว "
                "กรุณาติดต่อทีม VRCOMM ที่ iddhi.t@vrcomm.net ครับ")


def _extract_account_hint(message: str, subs: list) -> str:
    """
    Check if any known account name appears in the message.
    Returns the best-matched account name or empty string.

    Scoring:
    - Full name match        → score 999 (instant win)
    - Each matching word     → score +1 per word
    - Longer words weighted  → score += len(word) * 0.1
    - Min score to qualify   → 1 (at least one word must match)
    """
    msg_lower = message.lower()
    accounts  = list({s.get("Account", "") for s in subs})

    # Skip generic words that appear in many account names
    SKIP_WORDS = {"co", "ltd", "co.", "ltd.", "company", "limited",
                  "the", "and", "for", "of", "in", "public", "corporation"}

    best_score = 0
    best_acct  = ""

    for acct in accounts:
        if not acct:
            continue

        # Full name match — instant return
        if acct.lower() in msg_lower:
            return acct

        # Word-level scoring
        words = [w.strip(".,()-").lower() for w in acct.split()
                 if len(w.strip(".,()-")) >= 3 and w.lower() not in SKIP_WORDS]

        score = sum(
            1 + len(w) * 0.1
            for w in words
            if w in msg_lower
        )

        if score > best_score:
            best_score = score
            best_acct  = acct

    # Require at least one meaningful word match (score >= 1)
    return best_acct if best_score >= 1.0 else ""
