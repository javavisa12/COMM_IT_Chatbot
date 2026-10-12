"""Build final LLM messages and generate non-streaming chat responses."""

from backend.services.llm_service import llm

SYSTEM_PROMPT = """
You are a Senior Enterprise Software Engineer.
"""


def generate_response(messages):

    final_messages = [
        {
            "role": "system",
            "content": SYSTEM_PROMPT
        }
    ] + messages

    response = llm.create_chat_completion(
        messages=final_messages
    )

    return response["choices"][0]["message"]["content"]