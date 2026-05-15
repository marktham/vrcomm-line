"""
excel_export.py — Export logged LINE messages to Excel (.xlsx)
--------------------------------------------------------------
Generates a formatted Excel file from the SQLite database.
Called by the /export Flask endpoint.
"""

import os
import logging
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

from db import get_all_messages

logger = logging.getLogger(__name__)

# Export file path (temp directory)
EXPORT_PATH = os.environ.get("EXPORT_PATH", "/tmp/line_messages_export.xlsx")

# Column definitions: (header_label, db_field, column_width)
COLUMNS = [
    ("No.",          "id",           6),
    ("Timestamp",    "timestamp",    22),
    ("Logged At",    "logged_at",    20),
    ("Display Name", "display_name", 20),
    ("User ID",      "user_id",      25),
    ("Source Type",  "source_type",  12),
    ("Source ID",    "source_id",    25),
    ("Msg Type",     "msg_type",     12),
    ("Message",      "msg_text",     50),
    ("Detail",       "msg_detail",   40),
    ("Message ID",   "message_id",   20),
    ("Reply Token",  "reply_token",  40),
]

# Styling colours
HEADER_BG   = "1B4F72"   # dark blue
HEADER_FONT = "FFFFFF"   # white
ALT_ROW_BG  = "EBF5FB"   # light blue
BORDER_CLR  = "AED6F1"


def _header_style() -> tuple:
    font  = Font(bold=True, color=HEADER_FONT, size=11)
    fill  = PatternFill("solid", fgColor=HEADER_BG)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin  = Side(border_style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return font, fill, align, border


def _cell_style(row_idx: int) -> tuple:
    fill = PatternFill("solid", fgColor=ALT_ROW_BG) if row_idx % 2 == 0 else None
    align = Alignment(vertical="top", wrap_text=True)
    thin  = Side(border_style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return fill, align, border


def export_to_excel(output_path: str = EXPORT_PATH) -> str:
    """
    Generate the Excel file and return its file path.
    Raises on failure.
    """
    messages = get_all_messages()
    logger.info(f"Exporting {len(messages)} messages to Excel")

    wb = Workbook()
    ws = wb.active
    ws.title = "LINE Messages"

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    title_cell = ws["A1"]
    title_cell.value = (
        f"VRCOMM LINE Bot — Message Log   |   "
        f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    title_cell.font      = Font(bold=True, size=13, color=HEADER_FONT)
    title_cell.fill      = PatternFill("solid", fgColor="154360")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Header row ────────────────────────────────────────────────────────────
    h_font, h_fill, h_align, h_border = _header_style()
    for col_idx, (label, _, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=label)
        cell.font      = h_font
        cell.fill      = h_fill
        cell.alignment = h_align
        cell.border    = h_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 20

    # ── Data rows ─────────────────────────────────────────────────────────────
    for row_offset, msg in enumerate(messages):
        excel_row = row_offset + 3          # row 1=title, 2=header, 3+=data
        c_fill, c_align, c_border = _cell_style(row_offset)

        for col_idx, (_, field, _) in enumerate(COLUMNS, start=1):
            value = msg.get(field, "")
            cell  = ws.cell(row=excel_row, column=col_idx, value=value)
            if c_fill:
                cell.fill  = c_fill
            cell.alignment = c_align
            cell.border    = c_border

        ws.row_dimensions[excel_row].height = 18

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "VRCOMM LINE Bot — Summary"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2.merge_cells("A1:D1")

    summary_data = [
        ("Total Messages",    len(messages)),
        ("Export Date",       datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]

    # Count by source type
    source_counts: dict = {}
    msg_type_counts: dict = {}
    user_counts: dict = {}
    for m in messages:
        source_counts[m["source_type"]] = source_counts.get(m["source_type"], 0) + 1
        msg_type_counts[m["msg_type"]]  = msg_type_counts.get(m["msg_type"], 0) + 1
        user_counts[m["user_id"]]       = user_counts.get(m["user_id"], 0) + 1

    summary_data.append(("", ""))
    summary_data.append(("Source Type", "Count"))
    for k, v in sorted(source_counts.items()):
        summary_data.append((f"  {k}", v))

    summary_data.append(("", ""))
    summary_data.append(("Message Type", "Count"))
    for k, v in sorted(msg_type_counts.items()):
        summary_data.append((f"  {k}", v))

    summary_data.append(("", ""))
    summary_data.append(("Unique Users", len(user_counts)))

    for r, (label, value) in enumerate(summary_data, start=3):
        ws2.cell(row=r, column=1, value=label)
        ws2.cell(row=r, column=2, value=value)

    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 20

    # ── Save ──────────────────────────────────────────────────────────────────
    wb.save(output_path)
    logger.info(f"Excel exported → {output_path} ({len(messages)} rows)")
    return output_path
