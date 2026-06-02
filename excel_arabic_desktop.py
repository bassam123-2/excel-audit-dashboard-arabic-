"""
Excel Arabic Dashboard — desktop app (Django server + browser).

Development:
    python excel_arabic_desktop.py
"""

from __future__ import annotations

import argparse
import os
import sys
import webbrowser


def _default_host() -> str:
    return os.environ.get("EXCEL_ARABIC_HOST", "127.0.0.1")


def main() -> None:
    default_port = int(os.environ.get("EXCEL_ARABIC_PORT", "8765"))

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
        help=f"HTTP port (default: {default_port} or EXCEL_ARABIC_PORT)",
    )
    args = parser.parse_args()

    port = args.port or default_port
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
        webbrowser.open(url)

    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "django_project.settings.local")
    os.environ["EXCEL_ARABIC_PORT"] = str(port)

    from web_app import main as run_web_app

    run_web_app(host=host, port=port)


if __name__ == "__main__":
    main()
