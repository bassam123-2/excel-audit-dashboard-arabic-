"""
تشغيل لوحة التحليل محلياً وفتح المتصفح — نقطة الدخول لبناء PyInstaller (ملف .exe واحد).

تشغيل أثناء التطوير:
    python excel_arabic_desktop.py
"""

from __future__ import annotations

import os
import threading
import time
import webbrowser


def main() -> None:
    from app import DEFAULT_DASHBOARD_PORT, app as flask_app

    port = DEFAULT_DASHBOARD_PORT
    host = os.environ.get("EXCEL_ARABIC_HOST", "127.0.0.1")
    url = f"http://{host}:{port}/"

    print(f"Excel Arabic Dashboard  http://{host}:{port}/")

    def _open_browser() -> None:
        time.sleep(0.9)
        webbrowser.open(url)

    threading.Thread(target=_open_browser, daemon=True).start()
    flask_app.run(host=host, port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
