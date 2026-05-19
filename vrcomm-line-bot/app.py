"""VRCOMM LINE Bot Webhook Server"""
import os, logging
from datetime import datetime
from flask import Flask, request, abort, jsonify, send_file, Response
from linebot import LineBotApi, WebhookParser
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    StickerMessage, ImageMessage, AudioMessage,
    VideoMessage, LocationMessage, FileMessage,
    FollowEvent, UnfollowEvent, JoinEvent, LeaveEvent
)
from db import (init_db, log_message, get_all_messages,
                get_history, save_turn, clear_history,
                save_pending, get_pending, resolve_pending,
                save_subscription, get_subscription)
from ai_handler import process_with_ai
from excel_export import export_to_excel
from sheets_logger import log_to_sheet, save_history_turn, log_email, update_email_status

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
app = Flask(__name__)

LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
EXPORT_PASSWORD           = os.environ.get("EXPORT_PASSWORD", "")
ADMIN_LINE_USER_ID        = os.environ.get("ADMIN_LINE_USER_ID", "")
MS_GRAPH_CLIENT_STATE     = "vrcomm-graph-secret-2026"

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
parser       = WebhookParser(LINE_CHANNEL_SECRET)
init_db()


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_display_name(source):
    try:
        if source.type == "user":
            return line_bot_api.get_profile(source.user_id).display_name
        elif source.type == "group":
            return line_bot_api.get_group_member_profile(
                source.group_id, source.user_id).display_name
        elif source.type == "room":
            return line_bot_api.get_room_member_profile(
                source.room_id, source.user_id).display_name
    except Exception as e:
        logger.warning("Could not fetch profile: %s", e)
    return "Unknown"


def get_source_id(source):
    if source.type == "group":
        return source.group_id
    elif source.type == "room":
        return source.room_id
    return source.user_id


def push_to_admin(text: str):
    """Send a LINE push message to the admin user."""
    if not ADMIN_LINE_USER_ID:
        logger.warning("ADMIN_LINE_USER_ID not set — cannot push notification")
        return
    try:
        line_bot_api.push_message(ADMIN_LINE_USER_ID, TextSendMessage(text=text))
        logger.info("Admin push sent (%d chars)", len(text))
    except LineBotApiError as e:
        logger.error("Admin push failed: %s", e)


# ── LINE webhook ───────────────────────────────────────────────────────────────

@app.route("/webhook", methods=["POST"])
def webhook():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        events = parser.parse(body, signature)
    except InvalidSignatureError:
        abort(400)
    except Exception as e:
        logger.error("Parse error: %s", e)
        abort(500)
    for event in events:
        handle_event(event)
    return "OK", 200


def handle_event(event):
    try:
        if isinstance(event, MessageEvent):
            handle_message(event)
        elif isinstance(event, FollowEvent):
            handle_follow(event)
        elif isinstance(event, UnfollowEvent):
            logger.info("Unfollowed: %s", event.source.user_id)
        elif isinstance(event, JoinEvent):
            logger.info("Joined: %s", get_source_id(event.source))
    except Exception as e:
        logger.error("Error handling %s: %s", type(event).__name__, e)


def handle_message(event):
    user_id      = event.source.user_id
    source_type  = event.source.type
    source_id    = get_source_id(event.source)
    reply_token  = event.reply_token
    timestamp    = datetime.utcfromtimestamp(event.timestamp / 1000).strftime(
                       "%Y-%m-%d %H:%M:%S UTC")
    display_name = get_display_name(event.source)
    msg          = event.message
    msg_type     = msg.type
    msg_text = msg_detail = ""

    if isinstance(msg, TextMessage):
        msg_text = msg_detail = msg.text
    elif isinstance(msg, StickerMessage):
        msg_text = msg_detail = "[Sticker] package=%s id=%s" % (msg.package_id, msg.sticker_id)
    elif isinstance(msg, ImageMessage):
        msg_text = "[Image]"
        msg_detail = "Image id=%s" % msg.id
    elif isinstance(msg, AudioMessage):
        msg_text = "[Audio]"
        msg_detail = "Audio id=%s" % msg.id
    elif isinstance(msg, VideoMessage):
        msg_text = "[Video]"
        msg_detail = "Video id=%s" % msg.id
    elif isinstance(msg, LocationMessage):
        msg_text = msg_detail = "[Location] %s lat=%s lon=%s" % (
            msg.title or "", msg.latitude, msg.longitude)
    elif isinstance(msg, FileMessage):
        msg_text = msg_detail = "[File] %s (%s bytes)" % (msg.file_name, msg.file_size)
    else:
        msg_text = msg_detail = "[%s]" % msg_type

    logger.info("MSG from %s [%s]: %s", display_name, source_type, msg_text[:80])

    log_message(
        user_id=user_id, display_name=display_name,
        source_type=source_type, source_id=source_id,
        msg_type=msg_type, msg_text=msg_text, msg_detail=msg_detail,
        reply_token=reply_token, timestamp=timestamp, message_id=msg.id,
    )

    ai_reply_text = ""

    if isinstance(msg, TextMessage):
        text_stripped = msg_text.strip()

        # ── Email approval commands: SEND / EDIT / CANCEL ─────────────────────
        upper = text_stripped.upper()
        if upper.startswith("SEND ") or upper.startswith("EDIT ") or upper.startswith("CANCEL "):
            ai_reply_text = handle_approval_command(text_stripped)
            try:
                line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
            except Exception as e:
                logger.error("Approval reply error: %s", e)

        # ── Reset conversation ─────────────────────────────────────────────────
        elif text_stripped.lower() in ("reset", "เริ่มใหม่", "/reset", "clear"):
            clear_history(user_id)
            ai_reply_text = "Conversation reset. How can I help you today?"
            try:
                line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
            except Exception:
                pass

        # ── Normal AI reply ────────────────────────────────────────────────────
        else:
            try:
                history = get_history(user_id, max_turns=10)
                ai_reply_text = process_with_ai(
                    user_name=display_name, user_id=user_id,
                    message=msg_text, source_type=source_type,
                    history=history)
                line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
                logger.info("Replied to %s: %s", display_name, ai_reply_text[:80])

                save_turn(user_id, "user", msg_text)
                save_turn(user_id, "assistant", ai_reply_text)
                save_history_turn(user_id, display_name, "user", msg_text)
                save_history_turn(user_id, display_name, "assistant", ai_reply_text)

            except LineBotApiError as e:
                logger.error("LINE API error: %s", e)
            except Exception as e:
                logger.error("AI error: %s", e)

    else:
        try:
            ai_reply_text = ("Thank you for your %s. "
                             "Our team has received it and will respond shortly." % msg_type)
            line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
        except Exception:
            pass

    log_to_sheet(
        user_id=user_id, display_name=display_name,
        source_type=source_type, source_id=source_id,
        msg_type=msg_type, msg_text=msg_text, msg_detail=msg_detail,
        reply_token=reply_token, message_id=msg.id,
        timestamp=timestamp, ai_reply=ai_reply_text,
    )


def handle_approval_command(text: str) -> str:
    """
    Parse SEND/EDIT/CANCEL [task_id] commands from admin LINE message.
    Returns a reply string to send back.
    """
    from email_handler import send_reply, get_email
    parts = text.strip().split(None, 2)   # ["SEND", "VRCOMM-...", optional_new_text]
    if len(parts) < 2:
        return "⚠️ Format: SEND [task_id] | EDIT [task_id] [new text] | CANCEL [task_id]"

    command  = parts[0].upper()
    task_id  = parts[1].upper()
    extra    = parts[2] if len(parts) > 2 else ""

    task = get_pending(task_id)
    if not task:
        return "⚠️ Task %s not found or already resolved." % task_id

    if command == "CANCEL":
        resolve_pending(task_id, "cancelled")
        update_email_status(task_id, "cancelled")
        return "❌ Task %s cancelled. No reply sent." % task_id

    if command == "SEND":
        reply_body = task["draft_reply"]
    elif command == "EDIT":
        if not extra:
            return "⚠️ EDIT requires new reply text.\nFormat: EDIT %s [your reply text]" % task_id
        reply_body = extra
    else:
        return "⚠️ Unknown command. Use SEND, EDIT, or CANCEL."

    # Send the email reply via Graph API
    success = send_reply(task["message_id"], reply_body)
    if success:
        resolve_pending(task_id, "sent")
        update_email_status(task_id, "sent")
        return "✅ Reply sent for Task %s\n\nTo: %s\n\n%s" % (
            task_id, task["sender_email"], reply_body[:200])
    else:
        return "❌ Failed to send reply for Task %s. Please try again." % task_id


def handle_follow(event):
    display_name = get_display_name(event.source)
    try:
        welcome = ("Hello %s, welcome to VRCOMM Official!\n\n"
                   "Send us a message and our team will assist you shortly.") % display_name
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=welcome))
    except Exception as e:
        logger.error("Welcome message failed: %s", e)


# ── Email webhook (Microsoft Graph) ──────────────────────────────────────────

@app.route("/email-webhook", methods=["GET", "POST"])
def email_webhook():
    # ── Graph validation challenge (GET or POST with validationToken) ─────────
    validation_token = request.args.get("validationToken")
    if validation_token:
        logger.info("Graph validation challenge received")
        return Response(validation_token, status=200,
                        mimetype="text/plain; charset=utf-8")

    # ── Incoming notification (POST) ──────────────────────────────────────────
    if request.method != "POST":
        return "OK", 200

    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        data = {}

    notifications = data.get("value", [])
    for notif in notifications:
        # Verify client state to prevent spoofed requests
        if notif.get("clientState") != MS_GRAPH_CLIENT_STATE:
            logger.warning("Graph: clientState mismatch — ignoring notification")
            continue

        resource_data = notif.get("resourceData", {})
        message_id    = resource_data.get("id", "")
        if not message_id:
            # Try to extract from resource string e.g. "Users/.../messages/AAB..."
            resource = notif.get("resource", "")
            if "/messages/" in resource:
                message_id = resource.split("/messages/")[-1]

        if not message_id:
            logger.warning("Graph: no message_id in notification, skipping")
            continue

        logger.info("Graph: new email notification, message_id=%s", message_id[:40])
        _process_incoming_email(message_id)

    return "OK", 202


def _process_incoming_email(message_id: str):
    """Fetch, AI-process, and notify admin via LINE for a new incoming email."""
    try:
        from email_handler import get_email
        from email_processor import process_email, format_line_notification

        email = get_email(message_id)
        result = process_email(
            sender_name=email["sender_name"],
            sender_email=email["sender_email"],
            subject=email["subject"],
            body_text=email["body_text"],
        )

        task_id    = result["task_id"]
        summary    = result["summary"]
        draft      = result["draft_reply"]
        category   = result["category"]

        # Persist to SQLite for approval workflow
        save_pending(
            task_id=task_id,
            sender_name=email["sender_name"],
            sender_email=email["sender_email"],
            subject=email["subject"],
            body_preview=email["body_preview"],
            full_body=email["body_text"],
            message_id=message_id,
            category=category,
            summary=summary,
            draft_reply=draft,
        )

        # Log to Google Sheets "Email Messages" tab
        log_email(
            task_id=task_id,
            sender_name=email["sender_name"],
            sender_email=email["sender_email"],
            subject=email["subject"],
            category=category,
            body_preview=email["body_preview"],
            summary=summary,
            draft_reply=draft,
            received_at=email.get("received_at", ""),
            status="pending",
        )

        # Push LINE notification to admin for approval
        notification = format_line_notification(
            task_id=task_id,
            sender_name=email["sender_name"],
            sender_email=email["sender_email"],
            subject=email["subject"],
            summary=summary,
            draft_reply=draft,
            category=category,
        )
        push_to_admin(notification)
        logger.info("Email processed and admin notified: task=%s", task_id)

    except Exception as e:
        logger.error("Error processing incoming email (msg_id=%s): %s", message_id, e)


# ── Email subscription setup ──────────────────────────────────────────────────

@app.route("/setup-email", methods=["GET"])
def setup_email():
    """Create or renew the Microsoft Graph email webhook subscription."""
    from email_handler import create_subscription, renew_subscription, list_subscriptions

    try:
        existing_subs = list_subscriptions()
        current = get_subscription()

        if current and any(s["id"] == current["subscription_id"] for s in existing_subs):
            # Renew the existing one
            sub = renew_subscription(current["subscription_id"])
            save_subscription(sub["id"], sub["expirationDateTime"])
            return jsonify({
                "action":    "renewed",
                "id":        sub["id"],
                "expiry":    sub["expirationDateTime"],
            })
        else:
            # Create a fresh subscription
            sub = create_subscription()
            save_subscription(sub["id"], sub["expirationDateTime"])
            return jsonify({
                "action":    "created",
                "id":        sub["id"],
                "expiry":    sub["expirationDateTime"],
            })
    except Exception as e:
        logger.error("setup-email error: %s", e)
        return jsonify({"error": str(e)}), 500


# ── Export & health ───────────────────────────────────────────────────────────

@app.route("/export", methods=["GET"])
def export():
    if EXPORT_PASSWORD:
        if request.args.get("password", "") != EXPORT_PASSWORD:
            return Response("Unauthorized -- provide ?password=YOUR_PASSWORD", status=401)
    try:
        filepath = export_to_excel()
        filename = "LINE_Messages_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(filepath, as_attachment=True, download_name=filename,
                         mimetype=("application/vnd.openxmlformats-officedocument"
                                   ".spreadsheetml.sheet"))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/", methods=["GET"])
def health():
    messages = get_all_messages()
    sub = get_subscription()
    return jsonify({
        "status":                "running",
        "service":               "VRCOMM LINE Bot",
        "total_messages_logged": len(messages),
        "email_subscription":    sub["subscription_id"] if sub else "none",
        "email_sub_expiry":      sub["expiry"] if sub else "n/a",
        "timestamp":             datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
