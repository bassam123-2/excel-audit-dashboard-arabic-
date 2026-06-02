from __future__ import annotations

import os

from .base import *  # noqa: F403,F401

DEBUG = os.environ.get("DJANGO_DEBUG", "1").strip() in {"1", "true", "True"}
