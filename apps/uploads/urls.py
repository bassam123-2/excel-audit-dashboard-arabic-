from __future__ import annotations

from django.urls import path

from .views import analyze_view, upload_view

urlpatterns = [
    path("", upload_view, name="upload"),
    path("analyze", analyze_view, name="analyze"),
]
