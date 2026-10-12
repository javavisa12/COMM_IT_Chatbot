"""Provide the client used to call the local language-model service."""

import requests


class LLMClient:
    """HTTP client for the local Transformers model server on port 8001."""

    def create_chat_completion_stream(
        self,
        messages
    ):
        """Yield response text chunks from the model's streaming endpoint."""

        prompt = messages[-1]["content"]

        with requests.post(
            "http://localhost:8001/chat/stream",
            json={
                "message": prompt
            },
            timeout=300,
            stream=True
        ) as response:

            response.raise_for_status()

            for chunk in response.iter_content(
                chunk_size=1,
                decode_unicode=True
            ):

                if chunk:
                    yield chunk

    def create_chat_completion(
        self,
        messages
    ):
        """Return a non-streaming response in the app's chat-completion format."""

        prompt = messages[-1]["content"]

        response = requests.post(
            "http://localhost:8001/chat",
            json={
                "message": prompt
            },
            timeout=300
        )

        response.raise_for_status()

        data = response.json()

        return {
            "choices": [
                {
                    "message": {
                        "content":
                            data["response"]
                    }
                }
            ]
        }


llm = LLMClient()