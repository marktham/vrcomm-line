"""
intent_router.py — Classify intent and route to the correct agent.
Works for both LINE and Email inputs.

Intent types:
  product_info  — สินค้า ยี่ห้อ รุ่น features availability
  quotation     — ขอราคา ใบเสนอราคา price inquiry
  subscription  — subscription renewal expiry license
  technical     — spec TOR SOW compliance network design
  general       — ทั่วไป greetings admin other
"""
import os, logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

INTENTS = ["product_info", "quotation", "subscription", "technical", "general"]

CLASSIFIER_SYSTEM = """You are an intent classifier for VRCOMM, a Network and Cybersecurity solutions company in Thailand.
Products sold: Fortinet, Sophos, Cisco, and other network/cybersecurity brands.

Classify the message into exactly ONE category:
- product_info  : asking about products, brands, models, features, specs, availability, comparison
- quotation     : requesting a price quote, asking for pricing, wanting a quotation document, cost inquiry, ขอใบเสนอราคา
- subscription  : subscription status, renewal, expiry, license management, ต่ออายุ
- technical     : technical support, spec compliance, TOR/SOW analysis, equipment recommendation, network design
- general       : greetings, company info, complaints, general questions, anything else

Reply with ONLY the category name. One word. No punctuation. No explanation."""


def classify_intent(message: str, history: list = None) -> str:
    """
    Classify the intent of a message using a fast, cheap Claude call.
    Returns one of: product_info, quotation, subscription, technical, general
    """
    try:
        # Include last 2 exchanges for context (keeps it fast)
        ctx_messages = []
        if history:
            ctx_messages = history[-4:]
        ctx_messages.append({"role": "user", "content": message})

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=10,
            system=CLASSIFIER_SYSTEM,
            messages=ctx_messages,
        )
        intent = response.content[0].text.strip().lower().rstrip(".")
        if intent not in INTENTS:
            logger.warning("Unknown intent '%s' — defaulting to general", intent)
            intent = "general"
        logger.info("Intent: '%s' | msg: %s", intent, message[:60])
        return intent
    except Exception as e:
        logger.error("Intent classification error: %s", e)
        return "general"


def route(intent: str, message: str, user_name: str, user_id: str,
          source: str = "line", history: list = None, **kwargs) -> str:
    """
    Route message to the appropriate agent based on classified intent.

    Args:
        intent    : classified intent string
        message   : raw message text
        user_name : display name (LINE) or sender name (email)
        user_id   : LINE user_id or sender email address
        source    : "line" or "email"
        history   : conversation history list [{"role":..., "content":...}]
        **kwargs  : extra context passed through to agents

    Returns:
        Agent reply string
    """
    from agents.general_agent       import handle as general_handle
    from agents.product_agent       import handle as product_handle
    from agents.subscription_agent  import handle as subscription_handle
    from agents.engineer_agent      import handle as engineer_handle

    # Map each intent to its handler.
    # Set to None until that agent is built — will fall back to general.
    agent_map = {
        "general":      general_handle,
        "product_info": product_handle,
        "subscription": subscription_handle,
        "technical":    engineer_handle,
        "quotation":    None,   # agents/quotation_agent.py — coming soon
    }

    handler = agent_map.get(intent)

    if handler is None:
        logger.info("Agent for '%s' not yet built — falling back to general", intent)
        handler = general_handle

    return handler(
        message=message,
        user_name=user_name,
        user_id=user_id,
        source=source,
        history=history or [],
        intent=intent,
        **kwargs
    )
