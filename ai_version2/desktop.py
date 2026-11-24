"""Desktop launcher for ExcelSQL Studio 2.0.

Starts the local Flask server in a background thread and opens a native
pywebview window. Run with:  python desktop.py
"""

from __future__ import annotations

import threading
import time
import urllib.request

import webview

from app import app, _port_is_free


HOST = "127.0.0.1"
PORT_START = 8765
PORT_SEARCH_SPAN = 30

_server_error: list[Exception] = []


def pick_port() -> int:
    for port in range(PORT_START, PORT_START + PORT_SEARCH_SPAN):
        if _port_is_free(HOST, port):
            return port
    raise RuntimeError(
        f"No free port found in {PORT_START}-{PORT_START + PORT_SEARCH_SPAN - 1}."
    )


def serve(port: int) -> None:
    try:
        app.run(
            host=HOST,
            port=port,
            threaded=True,
            use_reloader=False,
        )
    except Exception as exc:  # surfaced by wait_until_ready()
        _server_error.append(exc)


def wait_until_ready(url: str) -> None:
    for _ in range(100):
        if _server_error:
            raise RuntimeError(f"Local server failed to start: {_server_error[0]}")
        try:
            with urllib.request.urlopen(url, timeout=0.2) as response:
                if response.status == 200:
                    return
        except Exception:
            time.sleep(0.1)
    raise RuntimeError("Timed out waiting for the local server to start.")


def main() -> None:
    port = pick_port()
    url = f"http://{HOST}:{port}/"

    threading.Thread(target=serve, args=(port,), daemon=True).start()
    wait_until_ready(url)

    webview.create_window(
        "ExcelSQL Studio 2.0",
        url,
        width=1440,
        height=900,
        min_size=(980, 620),
        background_color="#0f141c",
    )
    webview.start()


if __name__ == "__main__":
    main()
