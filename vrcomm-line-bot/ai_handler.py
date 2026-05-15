"""ai_handler.py - Claude AI reply generator with conversation memory"""
import os, logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

SYSTEM_PROMPT = """
You are the VRCOMM Admin, the official first-point-of-contact for VRCOMM Company
-- a Network and Cybersecurity solutions provider in Thailand.

Your role via LINE Official Account:
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

Tone: Warm, professional, concise. Max 3-4 short paragraphs.
Do NOT mention you are an AI unless asked directly.
Company: VRCOMM | Products: Fortinet, Sophos, Cisco and other network/cybersecurity brands.
Format as plain text only -- no markdown, no bullet symbols, no headers.
"""


def process_with_ai(user_name, user_id, message, source_type="user", history=None):
    """
    Send message + conversation history to Claude and return reply string.
    history: list of {"role": "user"/"assistant", "content": "..."} dicts
    """
    if history is None:
        history = []

    note = ""
    if source_type == "group":
        note = " (sent in a group chat)"
    elif source_type == "room":
        note = " (sent in a multi-person chat)"

    # Build messages: history + current message
    new_message = "Customer: %s\nLINE ID: %s%s\n\n%s" % (
        user_name, user_id, note, message)

    # For follow-up messages, use plain message without the header
    if history:
        new_message = message

    messages = history + [{"role": "user", "content": new_message}]

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            system=SYSTEM_PROMPT,
            messages=messages,
        )
        reply = response.content[0].text.strip()
        logger.info("AI reply for %s (%d history turns): %s",
                    user_name, len(history), reply[:80])
        return reply
    except Exception as e:
        logger.error("Claude API error: %s", e)
        return ("Thank you for your message.\n"
                "Our VRCOMM team has received it and will get back to you shortly.")
