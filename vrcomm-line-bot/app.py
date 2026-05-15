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
from db import init_db, log_message, get_all_messages, get_history, save_turn, clear_history
from ai_handler import process_with_ai
from excel_export import export_to_excel
from sheets_logger import log_to_sheet, save_history_turn

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
app = Flask(__name__)
LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
EXPORT_PASSWORD           = os.environ.get("EXPORT_PASSWORD", "")
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
parser       = WebhookParser(LINE_CHANNEL_SECRET)
init_db()


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
    msg = event.message
    msg_type = msg.type
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
        # Check for reset command
        if msg_text.strip().lower() in ("reset", "เริ่มใหม่", "/reset", "clear"):
            clear_history(user_id)
            ai_reply_text = "Conversation reset. How can I help you today?"
            try:
                line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
            except Exception:
                pass
        else:
            try:
                # Load this user's conversation history
                history = get_history(user_id, max_turns=10)

                ai_reply_text = process_with_ai(
                    user_name=display_name, user_id=user_id,
                    message=msg_text, source_type=source_type,
                    history=history)

                line_bot_api.reply_message(reply_token, TextSendMessage(text=ai_reply_text))
                logger.info("Replied to %s: %s", display_name, ai_reply_text[:80])

                # Save to SQLite session cache
                save_turn(user_id, "user", msg_text)
                save_turn(user_id, "assistant", ai_reply_text)
                # Save to Google Sheets for persistence across restarts
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


def handle_follow(event):
    display_name = get_display_name(event.source)
    try:
        welcome = ("Hello %s, welcome to VRCOMM Official!\n\n"
                   "Send us a message and our team will assist you shortly.") % display_name
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=welcome))
    except Exception as e:
        logger.error("Welcome message failed: %s", e)


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
    return jsonify({"status": "running", "service": "VRCOMM LINE Bot",
                    "total_messages_logged": len(messages),
                    "timestamp": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
