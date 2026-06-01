"""
Excel Arabic Dashboard — desktop app (Flask server + browser).

Development:
    python excel_arabic_desktop.py

Build .exe (from project root):
    pip install -r requirements-build.txt
    pyinstaller excel_arabic_dashboard.spec
    → dist/excel_arabic_dashboard.exe
"""

from __future__ import annotations

import argparse
import os
import sys
import threading
import time
import webbrowser


def _default_host() -> str:
    return os.environ.get("EXCEL_ARABIC_HOST", "127.0.0.1")


def main() -> None:
    from app import DEFAULT_DASHBOARD_PORT, app as flask_app

    parser = argparse.ArgumentParser(description="Excel Arabic compliance dashboard")
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Start server only; do not open a browser window",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=None,
        help=f"HTTP port (default: {DEFAULT_DASHBOARD_PORT} or EXCEL_ARABIC_PORT)",
    )
    args = parser.parse_args()

    port = args.port or DEFAULT_DASHBOARD_PORT
    host = _default_host()
    url = f"http://{host}:{port}/"

    frozen = getattr(sys, "frozen", False)
    if frozen:
        print("Excel Arabic Dashboard (desktop)")
        print(f"Running from: {sys.executable}")
    else:
        print("Excel Arabic Dashboard (development)")
    print(f"Open: {url}")
    print("Press Ctrl+C to stop.")

    if not args.no_browser:

        def _open_browser() -> None:
            time.sleep(0.9)
            webbrowser.open(url)

        threading.Thread(target=_open_browser, daemon=True).start()

    flask_app.run(host=host, port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
