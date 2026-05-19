"""
email_processor.py — Claude AI processes incoming emails for VRCOMM
--------------------------------------------------------------------
- Generates VRCOMM Task ID
- Summarises the email
- Drafts a professional reply
- Price info always flagged — never auto-replied
"""

import os, json, logging
from datetime import datetime
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

SYSTEM_PROMPT = """
You are VRCOMM Admin processing an incoming email.
VRCOMM is a Network and Cybersecurity solutions provider in Thailand.
Products: Fortinet, Sophos, Cisco and other network/cybersecurity brands.

Analyse the email and return a JSON object with exactly these keys:
{
  "summary": "2-3 sentence summary of what the sender wants",
  "draft_reply": "professional reply text ready to send",
  "category": one of: "quotation", "technical_support", "subscription_renewal", "general_inquiry", "complaint", "other"
}

STRICT RULES:
- NEVER include specific prices or cost figures in draft_reply
- If pricing is asked: acknowledge the request, say our sales team will prepare a quotation
- Match the language of the incoming email (Thai email = Thai reply, English = English)
- Keep draft_reply professional and concise (3-5 sentences max)
- Sign off as: VRCOMM Team
- Do NOT include subject line in draft_reply, just the body text
- Return ONLY valid JSON, no extra text
"""


def generate_task_id(conn=None) -> str:
    """Generate next Task ID in format VRCOMM-YYYYMMDD-NNN."""
    from db import get_conn
    today = datetime.now().strftime("%Y%m%d")
    with get_conn() as c:
        row = c.execute(
            "SELECT COUNT(*) FROM pending_approvals WHERE task_id LIKE ?",
            ("VRCOMM-%s-%%" % today,)
        ).fetchone()
        count = row[0] + 1
    return "VRCOMM-%s-%03d" % (today, count)


def process_email(sender_name: str, sender_email: str,
                  subject: str, body_text: str) -> dict:
    """
    Process an incoming email with Claude AI.
    Returns: {task_id, summary, draft_reply, category}
    """
    task_id = generate_task_id()

    prompt = (
        "New email received:\n\n"
        "From: %s <%s>\n"
        "Subject: %s\n\n"
        "Body:\n%s"
    ) % (sender_name, sender_email, subject, body_text[:3000])

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1024,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()

        # Parse JSON response
        result = json.loads(raw)
        result["task_id"] = task_id
        logger.info("Email processed — Task: %s Category: %s",
                    task_id, result.get("category", "?"))
        return result

    except json.JSONDecodeError as e:
        logger.error("AI response not valid JSON: %s | raw: %s", e, raw[:200])
    except Exception as e:
        logger.error("Email processing error: %s", e)

    # Fallback
    return {
        "task_id":     task_id,
        "summary":     "New email from %s — Subject: %s" % (sender_email, subject),
        "draft_reply": (
            "Dear %s,\n\n"
            "Thank you for contacting VRCOMM. We have received your email "
            "and our team will get back to you shortly.\n\n"
            "Best regards,\nVRCOMM Team"
        ) % (sender_name or "Customer"),
        "category":    "general_inquiry",
    }


def format_line_notification(task_id: str, sender_name: str, sender_email: str,
                              subject: str, summary: str, draft_reply: str,
                              category: str) -> str:
    """Format the LINE push notification for admin approval."""
    cat_emoji = {
        "quotation":           "💰",
        "technical_support":   "🔧",
        "subscription_renewal":"🔄",
        "general_inquiry":     "📋",
        "complaint":           "⚠️",
        "other":               "📌",
    }.get(category, "📌")

    return (
        "📧 New Email Task\n"
        "━━━━━━━━━━━━━━━━\n"
        "Task ID: %s\n"
        "%s Category: %s\n"
        "From: %s\n"
        "Email: %s\n"
        "Subject: %s\n\n"
        "📋 Summary:\n%s\n\n"
        "💬 Draft Reply:\n%s\n\n"
        "━━━━━━━━━━━━━━━━\n"
        "Reply with:\n"
        "SEND %s\n"
        "EDIT %s [your text]\n"
        "CANCEL %s"
    ) % (
        task_id, cat_emoji, category.replace("_", " ").title(),
        sender_name, sender_email, subject,
        summary, draft_reply,
        task_id, task_id, task_id
    )
