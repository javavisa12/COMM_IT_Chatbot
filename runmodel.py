print("starting server...")

from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from transformers import AutoTokenizer, AutoModelForCausalLM
from transformers import TextIteratorStreamer
import threading
import torch
import uvicorn

MODEL_PATH = "/home/aiadmin/Chatbot/models/Qwen2.5-Coder-7B-Instruct"

print("Loading tokenizer...")

tokenizer = AutoTokenizer.from_pretrained(
    MODEL_PATH,
    trust_remote_code=True
)

print("Loading model...")

model = AutoModelForCausalLM.from_pretrained(
    MODEL_PATH,
    torch_dtype=torch.bfloat16,
    device_map="auto",
    trust_remote_code=True
)

print("Model loaded successfully")

app = FastAPI()


class ChatRequest(BaseModel):
    message: str


@app.get("/")
def root():

    return {
        "status": "running",
        "model": "Qwen2.5-Coder-7B-Instruct"
    }


@app.post("/chat")
def chat(request: ChatRequest):

    messages = [
        {
            "role": "system",
            "content": (
                "Answer the user's question directly. Use uploaded document "
                "context when it is relevant, but answer general knowledge "
                "questions normally. Do not refuse a question merely because "
                "it is unrelated to the uploaded documents."
            )
        },
        {
            "role": "user",
            "content": request.message
        }
    ]

    text = tokenizer.apply_chat_template(
        messages,
        tokenize=False,
        add_generation_prompt=True
    )

    inputs = tokenizer(
        text,
        return_tensors="pt"
    ).to(model.device)

    outputs = model.generate(
        **inputs,
        max_new_tokens=1024,
        temperature=0.1,
        do_sample=True
    )

    generated = outputs[0][inputs.input_ids.shape[-1]:]

    answer = tokenizer.decode(
        generated,
        skip_special_tokens=True
    )

    return {
        "response": answer
    }


@app.post("/chat/stream")
def chat_stream(request: ChatRequest):

    messages = [
        {
            "role": "system",
            "content": (
                "Answer the user's question directly. Use uploaded document "
                "context when it is relevant, but answer general knowledge "
                "questions normally. Do not refuse a question merely because "
                "it is unrelated to the uploaded documents."
            )
        },
        {
            "role": "user",
            "content": request.message
        }
    ]

    text = tokenizer.apply_chat_template(
        messages,
        tokenize=False,
        add_generation_prompt=True
    )

    inputs = tokenizer(
        text,
        return_tensors="pt"
    ).to(model.device)

    streamer = TextIteratorStreamer(
        tokenizer,
        skip_prompt=True,
        skip_special_tokens=True
    )

    generation_args = {
        **inputs,
        "streamer": streamer,
        "max_new_tokens": 1024,
        "temperature": 0.2,
        "do_sample": True
    }

    generation_thread = threading.Thread(
        target=model.generate,
        kwargs=generation_args,
        daemon=True
    )
    generation_thread.start()

    def generate_chunks():

        for chunk in streamer:
            yield chunk

    return StreamingResponse(
        generate_chunks(),
        media_type="text/plain"
    )


if __name__ == "__main__":

    print("Starting server on port 8001")

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8001
    )