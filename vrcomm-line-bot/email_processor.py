"""
email_processor.py — Claude AI processes incoming emails for VRCOMM.
Uses the same intent_router as LINE messages so both channels
share the same agent logic.
"""
import os, logging
from datetime import datetime
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))


# ── Task ID generator ─────────────────────────────────────────────────────────

def generate_task_id() -> str:
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


# ── Category classifier (for emoji + Sheets logging) ─────────────────────────

_CATEGORY_SYSTEM = """You are an email classifier for VRCOMM, a Network and Cybersecurity company in Thailand.
Classify this email into ONE category:
quotation | technical_support | subscription_renewal | general_inquiry | complaint | other
Reply with ONLY the category name. One word."""

def _classify_email_category(subject: str, body: str) -> str:
    """Quick category classification used for emoji and Sheets logging."""
    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=10,
            system=_CATEGORY_SYSTEM,
            messages=[{"role": "user", "content":
                "Subject: %s\n\n%s" % (subject, body[:500])}],
        )
        cat = response.content[0].text.strip().lower().rstrip(".")
        valid = {"quotation", "technical_support", "subscription_renewal",
                 "general_inquiry", "complaint", "other"}
        return cat if cat in valid else "other"
    except Exception:
        return "general_inquiry"


# ── Summary generator ─────────────────────────────────────────────────────────

def _summarise_email(sender_name: str, sender_email: str,
                     subject: str, body_text: str) -> str:
    """Generate a 2-3 sentence summary for the LINE notification."""
    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=150,
            system="Summarise this email in 2-3 sentences. Reply in the same language as the email body.",
            messages=[{"role": "user", "content":
                "From: %s <%s>\nSubject: %s\n\n%s" % (
                    sender_name, sender_email, subject, body_text[:1500])}],
        )
        return response.content[0].text.strip()
    except Exception:
        return "New email from %s — Subject: %s" % (sender_email, subject)


# ── Main email processor ──────────────────────────────────────────────────────

def process_email(sender_name: str, sender_email: str,
                  subject: str, body_text: str) -> dict:
    """
    Process an incoming email using the shared intent_router.
    Both LINE and Email go through the same agents.

    Returns: {task_id, summary, draft_reply, category, intent}
    """
    from intent_router import classify_intent, route

    task_id  = generate_task_id()

    # Build message in a format agents can understand
    message  = "Subject: %s\n\n%s" % (subject, body_text[:3000])

    # Classify intent (shared with LINE)
    intent   = classify_intent(message)
    category = _classify_email_category(subject, body_text)

    # Get draft reply from the appropriate agent
    try:
        draft_reply = route(
            intent=intent,
            message=message,
            user_name=sender_name,
            user_id=sender_email,
            source="email",
            history=[],
        )
    except Exception as e:
        logger.error("Agent routing error: %s", e)
        draft_reply = (
            "Dear %s,\n\nThank you for contacting VRCOMM. "
            "We have received your email and our team will get back to you shortly.\n\n"
            "Best regards,\nVRCOMM Team"
        ) % (sender_name or "Customer")

    summary = _summarise_email(sender_name, sender_email, subject, body_text)

    logger.info("Email processed — Task: %s | Intent: %s | Category: %s",
                task_id, intent, category)
    return {
        "task_id":     task_id,
        "summary":     summary,
        "draft_reply": draft_reply,
        "category":    category,
        "intent":      intent,
    }


# ── LINE notification formatter ───────────────────────────────────────────────

def format_line_notification(task_id: str, sender_name: str, sender_email: str,
                              subject: str, summary: str, draft_reply: str,
                              category: str) -> str:
    """Format the LINE push notification for admin approval."""
    cat_emoji = {
        "quotation":            "💰",
        "technical_support":    "🔧",
        "subscription_renewal": "🔄",
        "general_inquiry":      "📋",
        "complaint":            "⚠️",
        "other":                "📌",
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
        summary, draft_reply[:400],
        task_id, task_id, task_id
    )
