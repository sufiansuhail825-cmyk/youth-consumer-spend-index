"""
Launcher for the packaged Youth Consumer Spend Index app.
Finds a free port, starts Streamlit, opens the browser.
Not used by Streamlit Cloud (which runs app.py directly).
"""

import sys
import os
import time
import socket
import threading
import webbrowser
from pathlib import Path

if getattr(sys, 'frozen', False):
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    BUNDLE_DIR = Path(__file__).parent

APP_FILE = BUNDLE_DIR / "app.py"


def find_free_port(preferred=8501, max_tries=20):
    for candidate in range(preferred, preferred + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", candidate))
                return candidate
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def open_browser(port):
    time.sleep(2)
    webbrowser.open(f"http://localhost:{port}")


def main():
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")

    port = find_free_port(8501)
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit",
        "run",
        str(APP_FILE),
        "--global.developmentMode=false",
        "--server.headless=true",
        f"--server.port={port}",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
