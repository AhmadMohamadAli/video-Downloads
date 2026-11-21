import time
import socket
import threading
import uvicorn
import webview
from app.main import app

HOST = "127.0.0.1"
PORT = 8000


def wait_for_server(host: str, port: int, timeout: float = 10.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.1)
    return False


def start_server():
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


if __name__ == "__main__":
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    if wait_for_server(HOST, PORT):
        webview.create_window(
            title="استوديو التحميل والتحويل الشامل Pro",
            url=f"http://{HOST}:{PORT}",
            width=1180,
            height=820,
            min_size=(900, 650),
        )
        webview.start()
    else:
        print("تعذر تشغيل الخادم المحلي.")