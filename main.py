from llama_cpp import Llama

llm = Llama(
    model_path=r"models\qwen2.5-coder-7b-instruct-q4_k_m.gguf",
    n_ctx=8192,
    n_gpu_layers=-1,
    verbose=False
)

messages = [
    {
        "role": "system",
        "content": """
        You are a senior software engineer.
        Expertise:
        - Java
        - Python
        - SQL
        - IBM Maximo
        """
    }
]

print("Chat started. Type 'exit' to quit.\n")

while True:
    user_input = input("You: ")

    if user_input.lower() == "exit":
        break

    messages.append(
        {
            "role": "user",
            "content": user_input
        }
    )

    response = llm.create_chat_completion(
        messages=messages
    )

    answer = response["choices"][0]["message"]["content"]

    print(f"\nAI: {answer}\n")

    messages.append(
        {
            "role": "assistant",
            "content": answer
        }
    )