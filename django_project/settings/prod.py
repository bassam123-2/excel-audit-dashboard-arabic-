from __future__ import annotations

import os

from .base import *  # noqa: F403,F401

DEBUG = False
ALLOWED_HOSTS = [h.strip() for h in os.environ.get("DJANGO_ALLOWED_HOSTS", "").split(",") if h.strip()]
