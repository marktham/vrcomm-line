"""
db.py — SQLite message logger for VRCOMM LINE Bot
---------------------------------------------------
Stores every incoming LINE message with full metadata.
The /export endpoint reads from here to generate Excel files.
"""

import sqlite3
import os
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

# Database file path (writable in Render's ephemeral filesystem)
DB_PATH = os.environ.get("DB_PATH", "vrcomm_line_messages.db")


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create tables if they don't exist."""
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id      TEXT    NOT NULL,
                display_name TEXT,
                source_type  TEXT,
                source_id    TEXT,
                msg_type     TEXT,
                msg_text     TEXT,
                msg_detail   TEXT,
                reply_token  TEXT,
                message_id   TEXT,
                timestamp    TEXT,
                logged_at    TEXT    DEFAULT (datetime('now'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ai_replies (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                message_id TEXT,
                user_id    TEXT,
                prompt     TEXT,
                reply      TEXT,
                logged_at  TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.commit()
    logger.info(f"Database initialised at: {DB_PATH}")


def log_message(
    user_id: str,
    display_name: str,
    source_type: str,
    source_id: str,
    msg_type: str,
    msg_text: str,
    msg_detail: str,
    reply_token: str,
    timestamp: str,
    message_id: str,
) -> int:
    """Insert a new message record. Returns the inserted row ID."""
    with get_conn() as conn:
        cursor = conn.execute(
            """
            INSERT INTO messages
                (user_id, display_name, source_type, source_id,
                 msg_type, msg_text, msg_detail, reply_token,
                 message_id, timestamp)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (user_id, display_name, source_type, source_id,
             msg_type, msg_text, msg_detail, reply_token,
             message_id, timestamp),
        )
        conn.commit()
        row_id = cursor.lastrowid
        logger.info(f"Logged message row_id={row_id} from {display_name}")
        return row_id


def log_ai_reply(message_id: str, user_id: str, prompt: str, reply: str):
    """Store the AI reply alongside the original message_id."""
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO ai_replies (message_id, user_id, prompt, reply)
            VALUES (?, ?, ?, ?)
            """,
            (message_id, user_id, prompt, reply),
        )
        conn.commit()


def get_all_messages() -> list:
    """Return all messages as a list of dicts, newest first."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM messages ORDER BY id DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def get_messages_by_user(user_id: str) -> list:
    """Return message history for a specific LINE user."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM messages WHERE user_id = ? ORDER BY id DESC LIMIT 50",
            (user_id,)
        ).fetchall()
    return [dict(r) for r in rows]
