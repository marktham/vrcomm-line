"""db.py - SQLite message logger + session cache for VRCOMM LINE Bot"""
import sqlite3, os, logging
from datetime import datetime

logger = logging.getLogger(__name__)
DB_PATH = os.environ.get("DB_PATH", "vrcomm_line_messages.db")

# Track which users have been seeded from Sheets in this Render session
_seeded_users = set()


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id      TEXT,
                display_name TEXT,
                source_type  TEXT,
                source_id    TEXT,
                msg_type     TEXT,
                msg_text     TEXT,
                msg_detail   TEXT,
                reply_token  TEXT,
                message_id   TEXT,
                timestamp    TEXT,
                logged_at    TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS conversation_history (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    TEXT NOT NULL,
                role       TEXT NOT NULL,
                content    TEXT NOT NULL,
                logged_at  TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_conv_user
            ON conversation_history (user_id, id)
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pending_approvals (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id       TEXT UNIQUE NOT NULL,
                sender_name   TEXT,
                sender_email  TEXT,
                subject       TEXT,
                body_preview  TEXT,
                full_body     TEXT,
                message_id    TEXT,
                category      TEXT,
                summary       TEXT,
                draft_reply   TEXT,
                status        TEXT DEFAULT 'pending',
                created_at    TEXT DEFAULT (datetime('now')),
                resolved_at   TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS graph_subscriptions (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                subscription_id TEXT UNIQUE,
                expiry          TEXT,
                updated_at      TEXT DEFAULT (datetime('now'))
            )
        """)
        conn.commit()
    logger.info("Database initialised at: %s", DB_PATH)


# ── Pending approvals ─────────────────────────────────────────────────────────

def save_pending(task_id, sender_name, sender_email, subject,
                 body_preview, full_body, message_id, category, summary, draft_reply):
    with get_conn() as conn:
        conn.execute("""
            INSERT OR REPLACE INTO pending_approvals
            (task_id, sender_name, sender_email, subject, body_preview,
             full_body, message_id, category, summary, draft_reply)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (task_id, sender_name, sender_email, subject, body_preview,
              full_body, message_id, category, summary, draft_reply))
        conn.commit()
    logger.info("Pending approval saved: %s", task_id)


def get_pending(task_id: str) -> dict:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM pending_approvals WHERE task_id = ? AND status = 'pending'",
            (task_id,)
        ).fetchone()
    return dict(row) if row else None


def resolve_pending(task_id: str, status: str):
    """Mark a task as sent/edited/cancelled."""
    with get_conn() as conn:
        conn.execute(
            """UPDATE pending_approvals SET status = ?, resolved_at = datetime('now')
               WHERE task_id = ?""",
            (status, task_id)
        )
        conn.commit()
    logger.info("Task %s resolved as: %s", task_id, status)


def get_all_pending() -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM pending_approvals WHERE status = 'pending' ORDER BY created_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]


# ── Graph subscription tracking ───────────────────────────────────────────────

def save_subscription(subscription_id: str, expiry: str):
    with get_conn() as conn:
        conn.execute("DELETE FROM graph_subscriptions")
        conn.execute(
            "INSERT INTO graph_subscriptions (subscription_id, expiry) VALUES (?, ?)",
            (subscription_id, expiry)
        )
        conn.commit()


def get_subscription() -> dict:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM graph_subscriptions ORDER BY id DESC LIMIT 1"
        ).fetchone()
    return dict(row) if row else None


# ── Conversation history ──────────────────────────────────────────────────────

def save_turn(user_id: str, role: str, content: str):
    """Save one conversation turn (role = 'user' or 'assistant')."""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO conversation_history (user_id, role, content) VALUES (?, ?, ?)",
            (user_id, role, content)
        )
        conn.commit()


def get_history(user_id: str, max_turns: int = 10) -> list:
    """
    Return the last `max_turns` exchanges oldest-first for Claude.
    On first call per user per session, seeds SQLite from Google Sheets
    so history survives Render restarts.
    """
    global _seeded_users
    if user_id not in _seeded_users:
        _seed_from_sheets(user_id, max_turns)
        _seeded_users.add(user_id)

    with get_conn() as conn:
        rows = conn.execute(
            """SELECT role, content FROM conversation_history
               WHERE user_id = ?
               ORDER BY id DESC LIMIT ?""",
            (user_id, max_turns * 2)
        ).fetchall()
    return [{"role": r["role"], "content": r["content"]} for r in reversed(rows)]


def _seed_from_sheets(user_id: str, max_turns: int):
    """Load history from Google Sheets into SQLite on first access per session."""
    try:
        from sheets_logger import load_user_history
        history = load_user_history(user_id, max_turns)
        if not history:
            return
        with get_conn() as conn:
            # Only insert if SQLite is empty for this user (avoid duplicates)
            existing = conn.execute(
                "SELECT COUNT(*) FROM conversation_history WHERE user_id = ?",
                (user_id,)
            ).fetchone()[0]
            if existing == 0:
                for turn in history:
                    conn.execute(
                        "INSERT INTO conversation_history (user_id, role, content) VALUES (?, ?, ?)",
                        (user_id, turn["role"], turn["content"])
                    )
                conn.commit()
                logger.info("Seeded %d history turns from Sheets for user %s",
                            len(history), user_id)
    except Exception as e:
        logger.warning("Could not seed history from Sheets: %s", e)


def clear_history(user_id: str):
    """Delete all conversation history for a user — both SQLite and Sheets."""
    global _seeded_users
    with get_conn() as conn:
        conn.execute(
            "DELETE FROM conversation_history WHERE user_id = ?", (user_id,)
        )
        conn.commit()
    _seeded_users.discard(user_id)
    # Also clear from Sheets
    try:
        from sheets_logger import clear_user_history
        clear_user_history(user_id)
    except Exception as e:
        logger.warning("Could not clear Sheets history: %s", e)
    logger.info("Conversation history cleared for user: %s", user_id)


def log_message(user_id, display_name, source_type, source_id,
                msg_type, msg_text, msg_detail, reply_token, timestamp, message_id):
    with get_conn() as conn:
        cursor = conn.execute(
            """INSERT INTO messages
               (user_id, display_name, source_type, source_id,
                msg_type, msg_text, msg_detail, reply_token, message_id, timestamp)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (user_id, display_name, source_type, source_id,
             msg_type, msg_text, msg_detail, reply_token, message_id, timestamp)
        )
        conn.commit()
        return cursor.lastrowid


def get_all_messages():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM messages ORDER BY id DESC").fetchall()
    return [dict(r) for r in rows]


def get_messages_by_user(user_id):
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM messages WHERE user_id = ? ORDER BY id DESC LIMIT 50",
            (user_id,)
        ).fetchall()
    return [dict(r) for r in rows]
