"""
sheets_logger.py — Google Sheets real-time logger for VRCOMM LINE Bot
----------------------------------------------------------------------
Appends every incoming LINE message to a Google Sheet instantly.
Data persists permanently — survives Render restarts and redeploys.

Required environment variables:
  GOOGLE_CREDENTIALS_JSON   — full JSON string of the service account key
  GOOGLE_SHEET_ID           — ID from the Google Sheet URL
                              e.g. https://docs.google.com/spreadsheets/d/SHEET_ID/edit
"""

import os
import json
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

# ── Column headers (must match order of values appended in log_to_sheet) ──────
SHEET_HEADERS = [
    "No.",
    "Timestamp (LINE)",
    "Logged At (Server)",
    "Display Name",
    "User ID",
    "Source Type",
    "Source ID",
    "Message Type",
    "Message",
    "Detail",
    "Message ID",
    "Reply Token",
    "AI Reply",
]

_sheet_client  = None   # gspread client (initialised once)
_worksheet     = None   # target worksheet (initialised once)
_row_counter   = 1      # tracks next row number for "No." column


def _init_sheet():
    """
    Lazily initialise the gspread client and worksheet.
    Called on first log attempt — not at import time so missing credentials
    don't crash the whole server.
    """
    global _sheet_client, _worksheet, _row_counter

    if _worksheet is not None:
        return True     # already initialised

    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    sheet_id   = os.environ.get("GOOGLE_SHEET_ID", "")

    if not creds_json or not sheet_id:
        logger.warning(
            "Google Sheets logging disabled — "
            "GOOGLE_CREDENTIALS_JSON or GOOGLE_SHEET_ID not set"
        )
        return False

    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = json.loads(creds_json)
        creds      = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        _sheet_client = gspread.authorize(creds)

        spreadsheet = _sheet_client.open_by_key(sheet_id)

        # Use existing "LINE Messages" sheet or create it
        try:
            _worksheet = spreadsheet.worksheet("LINE Messages")
            logger.info("Connected to existing 'LINE Messages' sheet")
        except gspread.WorksheetNotFound:
            _worksheet = spreadsheet.add_worksheet(
                title="LINE Messages", rows=5000, cols=len(SHEET_HEADERS)
            )
            logger.info("Created new 'LINE Messages' sheet")
            _ensure_headers()

        # Determine current row count so No. continues from where it left off
        existing = _worksheet.get_all_values()
        # existing[0] = header row (if present), rest = data
        data_rows    = len(existing) - 1 if len(existing) > 0 else 0
        _row_counter = max(1, data_rows + 1)

        logger.info(
            f"Google Sheets logger ready — sheet_id={sheet_id} "
            f"next_row_no={_row_counter}"
        )
        return True

    except Exception as e:
        logger.error(f"Google Sheets init failed: {e}")
        _worksheet = None
        return False


def _ensure_headers():
    """Write the header row if the sheet is empty."""
    global _worksheet
    if _worksheet is None:
        return
    try:
        first_row = _worksheet.row_values(1)
        if not first_row:
            _worksheet.append_row(
                SHEET_HEADERS,
                value_input_option="RAW"
            )
            # Bold the header row
            _worksheet.format("A1:M1", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.106, "green": 0.310, "blue": 0.447},
            })
    except Exception as e:
        logger.warning(f"Could not write sheet headers: {e}")


def log_to_sheet(
    user_id: str,
    display_name: str,
    source_type: str,
    source_id: str,
    msg_type: str,
    msg_text: str,
    msg_detail: str,
    reply_token: str,
    message_id: str,
    timestamp: str,
    ai_reply: str = "",
):
    """
    Append one row to the Google Sheet.
    Silently skips if Sheets is not configured or unavailable.
    """
    global _row_counter

    if not _init_sheet():
        return      # Sheets not configured — skip silently

    try:
        _ensure_headers()   # idempotent — only writes if row 1 is empty

        row = [
            _row_counter,
            timestamp,
            datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
            display_name,
            user_id,
            source_type,
            source_id,
            msg_type,
            msg_text,
            msg_detail,
            message_id,
            reply_token,
            ai_reply,
        ]

        _worksheet.append_row(row, value_input_option="USER_ENTERED")
        _row_counter += 1
        logger.info(f"Google Sheets: logged row {_row_counter - 1} for {display_name}")

    except Exception as e:
        logger.error(f"Google Sheets append failed: {e}")


def update_ai_reply(row_no: int, ai_reply: str):
    """
    Update the AI Reply column (column M = 13) for a specific row number.
    Called after the AI reply is generated so the sheet shows both
    the incoming message and the outgoing reply in the same row.
    """
    if _worksheet is None:
        return
    try:
        # Find the row that has this row_no in column A
        cell = _worksheet.find(str(row_no), in_column=1)
        if cell:
            _worksheet.update_cell(cell.row, 13, ai_reply)
            logger.info(f"Google Sheets: updated AI reply for row {row_no}")
    except Exception as e:
        logger.warning(f"Could not update AI reply in Sheets: {e}")
