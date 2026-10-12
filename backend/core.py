"""Store shared chat state, upload paths, and the active knowledge context."""

import os
import uuid


CHAT_DIR = "chats"
UPLOAD_DIR = "uploads"
MAX_RECENT_MESSAGES = 6

SESSION_FILES = []
ACTIVE_CONTEXT_ID = str(uuid.uuid4())
CHAT_SUMMARY = ""
RECENT_MESSAGES = []

os.makedirs(CHAT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


def reset_chat_state(clear_knowledge=True):
    """Reset conversation state and optionally remove indexed file knowledge."""

    global ACTIVE_CONTEXT_ID, CHAT_SUMMARY

    SESSION_FILES.clear()
    ACTIVE_CONTEXT_ID = str(uuid.uuid4())
    RECENT_MESSAGES.clear()
    CHAT_SUMMARY = ""

    if clear_knowledge:
        try:
            from backend.services.rag_service import clear_collection
            clear_collection()
        except Exception as error:
            print("RESET RAG ERROR:", error)
