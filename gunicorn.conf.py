# Read automatically by gunicorn from the app folder, so Azure App Service's default start command
# ("gunicorn --bind=0.0.0.0 --timeout 600 app:app") serves this FastAPI (ASGI) app correctly.
import os

worker_class = "uvicorn_worker.UvicornWorker"
# One worker: a run executes in a background thread of the worker that received it.
workers = int(os.environ.get("UAR_GUNICORN_WORKERS", "1"))
timeout = int(os.environ.get("UAR_GUNICORN_TIMEOUT", "600"))
graceful_timeout = 60
keepalive = 5
accesslog = "-"
errorlog = "-"
