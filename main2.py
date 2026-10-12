import os
import json
import shutil
import datetime

import gradio as gr
from llama_cpp import Llama

# =====================================
# CONFIG
# =====================================

MODEL_PATH = r"models\qwen2.5-coder-7b-instruct-q4_k_m.gguf"

TEMPERATURE = 0.1
TOP_P = 0.9
MAX_TOKENS = 1024

# =====================================
# FOLDERS
# =====================================

UPLOAD_DIR = "uploads"
CHAT_DIR = "chats"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(CHAT_DIR, exist_ok=True)

# =====================================
# GLOBALS
# =====================================

CHAT_FILE = None
SESSION_FILES = []

# =====================================
# MODEL
# =====================================

llm = Llama(
    model_path=MODEL_PATH,
    n_ctx=8192,
    n_gpu_layers=-1,
    verbose=False,
    chat_format="chatml"
)

# =====================================
# SYSTEM PROMPT
# =====================================

SYSTEM_PROMPT = """
You are a Senior Enterprise Software Engineer.

Expertise:
- Java
- Python
- SQL
- IBM Maximo

Rules:
- Preserve business logic during migrations
- Use Python 3.12
- Generate pytest tests
- Follow PEP8
- Explain migration decisions

When returning code:
- Always wrap code inside triple backticks
- Include the language name
- Preserve indentation
- Never return HTML tags
"""

# =====================================
# CHAT FILE MANAGEMENT
# =====================================

def get_chat_file(first_message):

    global CHAT_FILE

    if CHAT_FILE is None:

        title = "".join(
            c if c.isalnum() or c in "_- "
            else "_"
            for c in first_message[:40]
        )

        title = title.replace(" ", "_")

        timestamp = datetime.datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )

        CHAT_FILE = os.path.join(
            CHAT_DIR,
            f"{timestamp}_{title}.json"
        )

    return CHAT_FILE


def save_chat(messages, first_message):

    filename = get_chat_file(
        first_message
    )

    with open(
        filename,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            messages,
            f,
            indent=4,
            ensure_ascii=False
        )

# =====================================
# FILE HELPERS
# =====================================

def load_file(filepath):

    try:

        with open(
            filepath,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as f:

            content = f.read()

        return f"""
====================================
FILE: {os.path.basename(filepath)}
====================================

{content}

"""

    except Exception as e:

        return f"""
====================================
FILE: {os.path.basename(filepath)}
====================================

ERROR:
{e}

"""


def find_relevant_files(question):

    matches = []

    question = question.lower()

    for filename in os.listdir(
        UPLOAD_DIR
    ):

        basename = os.path.splitext(
            filename
        )[0].lower()

        if basename in question:

            matches.append(
                os.path.join(
                    UPLOAD_DIR,
                    filename
                )
            )

    return matches

# =====================================
# CHAT FUNCTION
# =====================================

def chat(
    message,
    history,
    uploaded_files
):

    messages = [
        {
            "role": "system",
            "content": SYSTEM_PROMPT
        }
    ]

    for msg in history:

        if isinstance(msg, dict):

            messages.append(
                {
                    "role": msg.get("role"),
                    "content": str(
                        msg.get(
                            "content",
                            ""
                        )
                    )
                }
            )

    file_context = ""

    # =====================================
    # NEW FILES
    # =====================================

    if uploaded_files:

        for uploaded_file in uploaded_files:

            destination = os.path.join(
                UPLOAD_DIR,
                os.path.basename(
                    uploaded_file
                )
            )

            shutil.copy(
                uploaded_file,
                destination
            )

            if destination not in SESSION_FILES:

                SESSION_FILES.append(
                    destination
                )

            file_context += load_file(
                destination
            )

    # =====================================
    # EXISTING FILES
    # =====================================

    else:

        relevant_files = (
            find_relevant_files(
                message
            )
        )

        if relevant_files:

            for file_path in relevant_files:

                file_context += (
                    load_file(file_path)
                )

        elif SESSION_FILES:

            # Use the most recently uploaded file

            recent_file = (
                SESSION_FILES[-1]
            )

            file_context += load_file(
                recent_file
            )

    # =====================================
    # BUILD PROMPT
    # =====================================

    if file_context:

        final_message = f"""
Question:

{message}

File Context:

{file_context}
"""

    else:

        final_message = message

    messages.append(
        {
            "role": "user",
            "content": final_message
        }
    )

    # =====================================
    # STREAMING RESPONSE
    # =====================================

    stream = llm.create_chat_completion(
        messages=messages,
        temperature=TEMPERATURE,
        top_p=TOP_P,
        max_tokens=MAX_TOKENS,
        stream=True
    )

    answer = ""

    for chunk in stream:

        try:

            delta = (
                chunk["choices"][0]
                .get("delta", {})
            )

            content = delta.get(
                "content",
                ""
            )

            if content:

                answer += content

                yield answer

        except Exception as e:

            print(
                "Streaming Error:",
                e
            )

    messages.append(
        {
            "role": "assistant",
            "content": answer,
            "timestamp": str(
                datetime.datetime.now()
            )
        }
    )

    save_chat(
        messages,
        message
    )

# =====================================
# FILE UPLOAD
# =====================================

file_upload = gr.File(
    label="Upload Files",
    type="filepath",
    file_count="multiple"
)
# =====================================
# CHAT HISTORY
# =====================================

def get_chat_list():

    chats = []

    for filename in os.listdir(CHAT_DIR):

        if filename.endswith(".json"):

            chats.append(filename)

    chats.sort(reverse=True)

    return chats


def preview_chat(chat_file):

    if not chat_file:
        return ""

    filepath = os.path.join(
        CHAT_DIR,
        chat_file
    )

    try:

        with open(
            filepath,
            "r",
            encoding="utf-8"
        ) as f:

            content = json.load(f)

        output = ""

        for msg in content:

            role = msg.get("role")

            if role == "user":

                output += f"\n\nUSER:\n{msg.get('content','')}"

            elif role == "assistant":

                output += f"\n\nASSISTANT:\n{msg.get('content','')}"

        return output

    except Exception as e:

        return str(e)
# =====================================
# UI
# =====================================

with gr.Blocks(
    title="COMM_IT AI"
) as demo:

    gr.Markdown("# COMM_IT AI")

    chat_selector = gr.Dropdown(
        choices=get_chat_list(),
        label="Previous Chats"
    )

    preview_box = gr.Textbox(
        label="Chat Preview",
        lines=15
    )

    chat_selector.change(
        fn=preview_chat,
        inputs=chat_selector,
        outputs=preview_box
    )

    gr.ChatInterface(
        fn=chat,
        additional_inputs=[
            file_upload
        ]
    )

# =====================================
# START
# =====================================

demo.launch(
    server_name="127.0.0.1",
    server_port=7860,
    share=True
)


