"""
agents/general_agent.py — General VRCOMM Admin agent.
Handles greetings, general questions, and acts as fallback
for intents whose dedicated agent is not yet built.
"""
import os, logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

_BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PRODUCT_LIST = os.path.join(_BASE_DIR, "product", "VRCOMM_ProductList.xlsx")


def _get_product_list_text() -> str:
    """Load brand names from ProductList for injection into system prompt."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(_PRODUCT_LIST, read_only=True, data_only=True)
        ws = wb.active
        brands = []
        for row in ws.iter_rows(values_only=True):
            brand = str(row[0]).strip() if row[0] else ""
            if brand.lower() in ("brand", "product", "name", "") or not brand:
                continue
            brands.append(brand)
        return ", ".join(brands)
    except Exception as e:
        logger.warning("[general_agent] Could not load product list: %s", e)
        return "various network and cybersecurity products"

_SYSTEM_PROMPT_TEMPLATE = """You are the VRCOMM Admin, the official first-point-of-contact for VRCOMM Company
-- a Network and Cybersecurity solutions provider in Thailand.

Your role:
1. Greet customers warmly and professionally on first contact
2. Remember everything discussed earlier in this conversation
3. Understand their inquiry (product info, quotation request, technical support,
   subscription renewal, general question)
4. Give a helpful, concise response in the SAME LANGUAGE they write in
   (Thai reply Thai, English reply English, mixed reply mixed)
5. If they need a quotation, ask for their email or company name
6. If they report a technical issue, gather: company name, product affected,
   issue description, urgency level
7. For subscription renewals, ask: company name and which product
8. Build on previous messages -- never ask for information already provided

VRCOMM's product portfolio (the ONLY brands we sell):
{product_list}

IMPORTANT: If someone asks about any brand in the list above, acknowledge it as part of our portfolio.
Do NOT say a brand is "not in our portfolio" if it appears in the list above.
NEVER mention Fortinet, Cisco, Palo Alto, Check Point, or any brand NOT in the list above.

Tone: Warm, professional, concise. Max 3-4 short paragraphs.
Do NOT mention you are an AI unless asked directly.
Format as plain text only -- no markdown, no bullet symbols, no headers."""


def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "general", **kwargs) -> str:
    """
    Handle general / fallback messages.

    Args:
        message   : user message text
        user_name : display name
        user_id   : LINE user_id or email address
        source    : "line" or "email"
        history   : conversation history
        intent    : classified intent (used for logging)
    """
    if history is None:
        history = []

    # Build system prompt with current product list
    system = _SYSTEM_PROMPT_TEMPLATE.format(product_list=_get_product_list_text())

    # First message in conversation — include context header
    if history:
        content = message
    else:
        source_note = ""
        if source == "email":
            source_note = " (via email)"
        elif source == "line_group":
            source_note = " (sent in a group chat)"
        content = "Customer: %s\nID: %s%s\n\n%s" % (
            user_name, user_id, source_note, message)

    messages = history + [{"role": "user", "content": content}]

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            system=system,
            messages=messages,
        )
        reply = response.content[0].text.strip()
        logger.info("[general_agent] replied to %s (intent=%s): %s",
                    user_name, intent, reply[:80])
        return reply
    except Exception as e:
        logger.error("[general_agent] Claude API error: %s", e)
        return ("ขอบคุณสำหรับการติดต่อครับ ทีมงาน VRCOMM ได้รับข้อความของคุณแล้ว "
                "และจะติดต่อกลับโดยเร็วที่สุดครับ")
