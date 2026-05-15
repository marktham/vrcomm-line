"""sheets_logger.py - Google Sheets logger + persistent conversation memory"""
import os, json, logging
from datetime import datetime

logger = logging.getLogger(__name__)

SHEET_HEADERS = [
    "No.", "Timestamp (LINE)", "Logged At (Server)", "Display Name",
    "User ID", "Source Type", "Source ID", "Message Type",
    "Message", "Detail", "Message ID", "Reply Token", "AI Reply",
]

HISTORY_HEADERS = [
    "user_id", "display_name", "role", "content", "logged_at"
]

_spreadsheet  = None   # shared gspread Spreadsheet object
_worksheet    = None   # "LINE Messages" sheet
_hist_sheet   = None   # "Chat History" sheet
_row_counter  = 1


def _get_spreadsheet():
    """Lazy-init and return the gspread Spreadsheet object. Returns None if not configured."""
    global _spreadsheet
    if _spreadsheet is not None:
        return _spreadsheet

    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip().lstrip("\xef\xbb\xbf")
    sheet_id   = os.environ.get("GOOGLE_SHEET_ID", "").strip()
    logger.info("Sheets init -- creds_json length: %d, sheet_id: '%s'", len(creds_json), sheet_id)

    if not creds_json or not sheet_id:
        logger.warning("Google Sheets disabled -- env vars not set")
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = json.loads(creds_json)
        logger.info("Sheets credentials parsed OK -- project: %s", creds_dict.get("project_id", "?"))
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        _spreadsheet = client.open_by_key(sheet_id)
        logger.info("Google Sheets connected: %s", _spreadsheet.title)
        return _spreadsheet
    except Exception as e:
        logger.error("Google Sheets connection failed: %s", e)
        return None


def _init_messages_sheet():
    """Return the LINE Messages worksheet, creating it if needed."""
    global _worksheet, _row_counter
    if _worksheet is not None:
        return _worksheet
    sp = _get_spreadsheet()
    if sp is None:
        return None
    try:
        import gspread
        try:
            _worksheet = sp.worksheet("LINE Messages")
        except gspread.WorksheetNotFound:
            _worksheet = sp.add_worksheet(title="LINE Messages", rows=5000, cols=len(SHEET_HEADERS))
            _worksheet.append_row(SHEET_HEADERS, value_input_option="RAW")
        existing = _worksheet.get_all_values()
        _row_counter = max(1, len(existing))
        logger.info("LINE Messages sheet ready -- next row: %d", _row_counter + 1)
        return _worksheet
    except Exception as e:
        logger.error("LINE Messages sheet init failed: %s", e)
        return None


def _init_history_sheet():
    """Return the Chat History worksheet, creating it if needed."""
    global _hist_sheet
    if _hist_sheet is not None:
        return _hist_sheet
    sp = _get_spreadsheet()
    if sp is None:
        return None
    try:
        import gspread
        try:
            _hist_sheet = sp.worksheet("Chat History")
        except gspread.WorksheetNotFound:
            _hist_sheet = sp.add_worksheet(title="Chat History", rows=10000, cols=len(HISTORY_HEADERS))
            _hist_sheet.append_row(HISTORY_HEADERS, value_input_option="RAW")
            logger.info("Created new Chat History sheet")
        logger.info("Chat History sheet ready")
        return _hist_sheet
    except Exception as e:
        logger.error("Chat History sheet init failed: %s", e)
        return None


# ── Message logging ───────────────────────────────────────────────────────────

def log_to_sheet(user_id, display_name, source_type, source_id,
                 msg_type, msg_text, msg_detail, reply_token,
                 message_id, timestamp, ai_reply=""):
    global _row_counter
    ws = _init_messages_sheet()
    if ws is None:
        return
    try:
        row = [
            _row_counter,
            timestamp,
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            display_name, user_id, source_type, source_id,
            msg_type, msg_text, msg_detail, message_id, reply_token, ai_reply,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        _row_counter += 1
        logger.info("Sheets: logged message row %d for %s", _row_counter - 1, display_name)
    except Exception as e:
        logger.error("Sheets message log failed: %s", e)


# ── Conversation history (persistent memory) ──────────────────────────────────

def save_history_turn(user_id, display_name, role, content):
    """Append one conversation turn to the Chat History sheet."""
    ws = _init_history_sheet()
    if ws is None:
        return
    try:
        row = [
            user_id,
            display_name,
            role,
            content,
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        logger.info("Sheets history: saved %s turn for %s", role, display_name)
    except Exception as e:
        logger.error("Sheets history save failed: %s", e)


def load_user_history(user_id, max_turns=10):
    """
    Load the last `max_turns` exchanges for a user from Chat History sheet.
    Returns a list of {"role": ..., "content": ...} dicts, oldest first.
    """
    ws = _init_history_sheet()
    if ws is None:
        return []
    try:
        all_rows = ws.get_all_values()
        # all_rows[0] = header; rest = data
        data = [r for r in all_rows[1:] if len(r) >= 4 and r[0] == user_id]
        # take last max_turns * 2 rows (each exchange = user + assistant)
        recent = data[-(max_turns * 2):]
        history = [{"role": r[2], "content": r[3]} for r in recent]
        logger.info("Sheets history: loaded %d turns for user %s", len(history), user_id)
        return history
    except Exception as e:
        logger.error("Sheets history load failed: %s", e)
        return []


def clear_user_history(user_id):
    """Delete all Chat History rows for a given user_id."""
    ws = _init_history_sheet()
    if ws is None:
        return
    try:
        all_rows = ws.get_all_values()
        # Find rows to delete (iterate in reverse to preserve row indices)
        rows_to_delete = [
            i + 1  # 1-based sheet row (row 1 = header)
            for i, r in enumerate(all_rows)
            if i > 0 and len(r) >= 1 and r[0] == user_id
        ]
        # Delete from bottom to top so row numbers stay valid
        for row_idx in reversed(rows_to_delete):
            ws.delete_rows(row_idx)
        logger.info("Sheets history: cleared %d rows for user %s",
                    len(rows_to_delete), user_id)
    except Exception as e:
        logger.error("Sheets history clear failed: %s", e)
