"""SQLite-backed chat history storage with JSON compatibility for existing data."""

import datetime
import json
import os
import sqlite3

from backend import core

CHAT_DIR = core.CHAT_DIR
DB_PATH = os.path.join(CHAT_DIR, "chat_history.db")


def _ensure_chat_dir():
    os.makedirs(CHAT_DIR, exist_ok=True)


def get_connection():
    _ensure_chat_dir()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    _ensure_chat_dir()
    connection = get_connection()
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL UNIQUE,
            title TEXT NOT NULL,
            created_at TEXT NOT NULL,
            messages TEXT NOT NULL,
            conversation_id TEXT
        )
        """
    )

    columns = [
        row[1] for row in connection.execute("PRAGMA table_info(chat_history)").fetchall()
    ]
    if "conversation_id" not in columns:
        connection.execute(
            "ALTER TABLE chat_history ADD COLUMN conversation_id TEXT"
        )

    connection.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_chat_history_conversation_id ON chat_history(conversation_id)"
    )
    connection.commit()
    connection.close()
    migrate_json_history()


def sanitize_title(title):
    cleaned = str(title or "untitled").strip()
    for character in ["\\", "/", ":", "*", "?", '"', "'", "<", ">", "|"]:
        cleaned = cleaned.replace(character, "_")
    cleaned = cleaned.strip() or "untitled"
    return cleaned


def build_filename(title):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return f"{timestamp}_{sanitize_title(title)}.json"


def save_chat(title, messages, conversation_id=None):
    init_db()
    payload = json.dumps(messages, ensure_ascii=False)
    created_at = datetime.datetime.now().isoformat()

    connection = get_connection()

    if conversation_id:
        existing = connection.execute(
            "SELECT filename FROM chat_history WHERE conversation_id = ?",
            (conversation_id,),
        ).fetchone()
        if existing:
            connection.execute(
                """
                UPDATE chat_history
                SET title = ?, created_at = ?, messages = ?
                WHERE conversation_id = ?
                """,
                (title, created_at, payload, conversation_id),
            )
            connection.commit()
            connection.close()
            return existing["filename"]

    filename = build_filename(title)
    connection.execute(
        """
        INSERT INTO chat_history (filename, title, created_at, messages, conversation_id)
        VALUES (?, ?, ?, ?, ?)
        """,
        (filename, title, created_at, payload, conversation_id),
    )
    connection.commit()
    connection.close()
    return filename


def list_history():
    init_db()
    connection = get_connection()
    rows = connection.execute(
        "SELECT filename FROM chat_history ORDER BY created_at DESC, filename DESC"
    ).fetchall()
    connection.close()
    return [row["filename"] for row in rows]


def load_chat(filename):
    init_db()
    connection = get_connection()
    row = connection.execute(
        "SELECT messages FROM chat_history WHERE filename = ?",
        (filename,),
    ).fetchone()
    connection.close()
    if row is None:
        return []

    try:
        return json.loads(row["messages"])
    except (TypeError, ValueError):
        return []


def delete_chat(filename):
    init_db()
    connection = get_connection()
    cursor = connection.execute(
        "DELETE FROM chat_history WHERE filename = ?",
        (filename,),
    )
    connection.commit()
    connection.close()
    return cursor.rowcount > 0


def migrate_json_history():
    _ensure_chat_dir()
    if not os.path.isdir(CHAT_DIR):
        return

    connection = get_connection()
    existing_rows = connection.execute(
        "SELECT filename FROM chat_history"
    ).fetchall()
    existing_filenames = {row["filename"] for row in existing_rows}
    migrated = False

    for entry in sorted(os.listdir(CHAT_DIR)):
        if not entry.endswith(".json") or entry == "chat_history.db":
            continue
        filepath = os.path.join(CHAT_DIR, entry)

        if entry in existing_filenames:
            try:
                os.remove(filepath)
                migrated = True
            except OSError:
                pass
            continue

        try:
            with open(filepath, "r", encoding="utf-8") as file:
                payload = json.load(file)
        except (OSError, ValueError):
            continue

        title = os.path.splitext(entry)[0].split("_", 1)[1] if "_" in os.path.splitext(entry)[0] else entry
        created_at = os.path.getmtime(filepath)
        created_at_iso = datetime.datetime.fromtimestamp(created_at).isoformat()

        connection.execute(
            """
            INSERT INTO chat_history (filename, title, created_at, messages)
            VALUES (?, ?, ?, ?)
            """,
            (entry, title, created_at_iso, json.dumps(payload, ensure_ascii=False)),
        )
        try:
            os.remove(filepath)
            migrated = True
        except OSError:
            pass

    connection.commit()
    connection.close()
    return migrated
