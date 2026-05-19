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

EMAIL_HEADERS = [
    "No.", "Received At", "Logged At (Server)", "Task ID",
    "Sender Name", "Sender Email", "Subject", "Category",
    "Body Preview", "Summary", "Draft Reply", "Status",
]

_spreadsheet   = None   # shared gspread Spreadsheet object
_worksheet     = None   # "LINE Messages" sheet
_hist_sheet    = None   # "Chat History" sheet
_email_sheet   = None   # "Email Messages" sheet
_row_counter   = 1
_email_counter = 1


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


def _init_email_sheet():
    """Return the Email Messages worksheet, creating it if needed."""
    global _email_sheet, _email_counter
    if _email_sheet is not None:
        return _email_sheet
    sp = _get_spreadsheet()
    if sp is None:
        return None
    try:
        import gspread
        try:
            _email_sheet = sp.worksheet("Email Messages")
        except gspread.WorksheetNotFound:
            _email_sheet = sp.add_worksheet(title="Email Messages", rows=5000, cols=len(EMAIL_HEADERS))
            _email_sheet.append_row(EMAIL_HEADERS, value_input_option="RAW")
        existing = _email_sheet.get_all_values()
        _email_counter = max(1, len(existing))
        logger.info("Email Messages sheet ready -- next row: %d", _email_counter + 1)
        return _email_sheet
    except Exception as e:
        logger.error("Email Messages sheet init failed: %s", e)
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


# ── Email logging ─────────────────────────────────────────────────────────────

def log_email(task_id: str, sender_name: str, sender_email: str,
              subject: str, category: str, body_preview: str,
              summary: str, draft_reply: str,
              received_at: str = "", status: str = "pending"):
    """Log an incoming email to the Email Messages sheet."""
    global _email_counter
    ws = _init_email_sheet()
    if ws is None:
        return
    try:
        row = [
            _email_counter,
            received_at or datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            task_id,
            sender_name, sender_email, subject, category,
            body_preview[:300],
            summary,
            draft_reply,
            status,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        _email_counter += 1
        logger.info("Sheets: logged email row %d task=%s", _email_counter - 1, task_id)
    except Exception as e:
        logger.error("Sheets email log failed: %s", e)


def update_email_status(task_id: str, new_status: str):
    """Update the Status column for a given Task ID in Email Messages sheet."""
    ws = _init_email_sheet()
    if ws is None:
        return
    try:
        all_rows = ws.get_all_values()
        # Column D (index 3) = Task ID; Column L (index 11) = Status
        for i, row in enumerate(all_rows):
            if i == 0:
                continue  # skip header
            if len(row) >= 4 and row[3] == task_id:
                sheet_row = i + 1           # 1-based
                ws.update_cell(sheet_row, 12, new_status)
                logger.info("Sheets: updated email status task=%s -> %s", task_id, new_status)
                return
        logger.warning("Sheets: task_id %s not found in Email Messages", task_id)
    except Exception as e:
        logger.error("Sheets email status update failed: %s", e)


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
