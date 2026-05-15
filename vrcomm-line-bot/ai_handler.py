"""
ai_handler.py — Claude AI message processor for VRCOMM LINE Bot
----------------------------------------------------------------
Takes an incoming LINE message, sends it to Claude with VRCOMM Admin
context, and returns an appropriate Thai/English reply.
"""

import os
import logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)

client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── System prompt: VRCOMM Admin persona ──────────────────────────────────────
SYSTEM_PROMPT = """
You are the VRCOMM Admin, the official first-point-of-contact for VRCOMM Company
— a Network and Cybersecurity solutions provider in Thailand.

Your role via LINE Official Account:
1. Greet customers warmly and professionally
2. Understand their inquiry (product info, quotation request, technical support,
   subscription renewal, general question)
3. Give a helpful, concise response in the SAME LANGUAGE they write in
   (Thai → reply Thai, English → reply English, mixed → reply mixed)
4. If they need a quotation or technical proposal, let them know the team will
   prepare one and ask for their email or company name
5. If they report a technical issue, gather: company name, product/system affected,
   issue description, urgency level
6. For subscription renewals, ask: company name and which product/software
7. Always end with a friendly closing

Tone: Warm, professional, concise. Max 3-4 short paragraphs.
Do NOT mention you are an AI unless asked directly.
Company name: VRCOMM | Products: Fortinet, Sophos, Cisco, and other network/cybersecurity brands.

Format your reply as plain text (no markdown, no bullet symbols with *, no #headers).
Use line breaks naturally for readability.
"""


def process_with_ai(
    user_name: str,
    user_id: str,
    message: str,
    source_type: str = "user",
) -> str:
    """
    Send a LINE message to Claude and return the AI reply string.
    Falls back to a polite default message on any error.
    """
    context_note = ""
    if source_type == "group":
        context_note = " (This message was sent in a group chat.)"
    elif source_type == "room":
        context_note = " (This message was sent in a multi-person chat.)"

    user_prompt = (
        f"Customer name: {user_name}\n"
        f"Customer LINE ID: {user_id}\n"
        f"Message:{context_note}\n\n"
        f"{message}"
    )

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",   # fast + cost-effective for chat
            max_tokens=512,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_prompt}],
        )
        reply = response.content[0].text.strip()
        logger.info(f"AI reply generated for {user_name}: {reply[:80]}")
        return reply

    except Exception as e:
        logger.error(f"Claude API error: {e}")
        return (
            "ขอบคุณสำหรับข้อความครับ/ค่ะ 🙏\n"
            "ทีมงาน VRCOMM ได้รับแล้ว และจะติดต่อกลับโดยเร็วที่สุดนะครับ/ค่ะ\n\n"
            "Thank you for your message. Our team will get back to you shortly."
        )
