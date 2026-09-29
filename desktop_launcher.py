"""Start the bundled Streamlit app on this computer only."""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

from streamlit.web.cli import main as streamlit_main


def _free_local_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _open_when_ready(url: str) -> None:
    health_url = f"{url}/_stcore/health"
    for _ in range(120):
        try:
            with urllib.request.urlopen(health_url, timeout=1) as response:
                if response.status == 200:
                    if os.environ.get("OPERATING_PROFILE_NO_BROWSER") != "1":
                        webbrowser.open(url)
                    return
        except (OSError, urllib.error.URLError):
            time.sleep(1)
    print(f"The app did not start. Check the error above, then try again: {url}")


def main() -> None:
    app = Path(__file__).resolve().with_name("app.py")
    if not app.is_file():
        raise SystemExit(f"Missing bundled application: {app}")

    port = int(os.environ.get("OPERATING_PROFILE_PORT") or _free_local_port())
    url = f"http://127.0.0.1:{port}"
    print("Starting Vessel Operating Profile...", flush=True)
    print(f"If the browser does not open, visit {url}", flush=True)
    print("Close this window to stop the app.", flush=True)
    threading.Thread(target=_open_when_ready, args=(url,), daemon=True).start()

    sys.argv = [
        "streamlit",
        "run",
        str(app),
        "--global.developmentMode=false",
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        "--browser.gatherUsageStats=false",
    ]
    streamlit_main()


if __name__ == "__main__":
    main()
