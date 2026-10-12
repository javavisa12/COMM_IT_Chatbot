"""Start the FastAPI application and serve the chatbot frontend."""

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import uvicorn

from maximo_service import get_object
from backend.routers import router


app = FastAPI()

app.mount(
    "/css",
    StaticFiles(directory="frontend/css"),
    name="css"
)

app.mount(
    "/js",
    StaticFiles(directory="frontend/js"),
    name="js"
)

app.include_router(router)


if __name__ == "__main__":

    HOST = "0.0.0.0"
    PORT = 8000

    print("\n" + "=" * 60)
    print("COMM_IT AI STARTED")
    print(f"Frontend : http://localhost:{PORT}")
    print(f"Swagger  : http://localhost:{PORT}/docs")
    print(f"Health   : http://localhost:{PORT}/health")
    print("=" * 60 + "\n")

    uvicorn.run(
        app,
        host=HOST,
        port=PORT
    )
