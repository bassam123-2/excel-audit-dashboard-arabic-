"""Compatibility entry for Django runtime."""

import os
import sys


def main(host: str | None = None, port: int | None = None) -> None:
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "django_project.settings.local")
    from django.core.management import execute_from_command_line

    run_host = host or os.environ.get("EXCEL_ARABIC_HOST", "0.0.0.0")
    run_port = port or int(os.environ.get("EXCEL_ARABIC_PORT", "8765"))
    execute_from_command_line([sys.argv[0], "runserver", f"{run_host}:{run_port}", "--noreload"])


if __name__ == "__main__":
    main()
