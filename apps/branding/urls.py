from __future__ import annotations

from django.urls import path

from .views import brand_logo_view, logo_view

urlpatterns = [
    path("api/brand-logo", brand_logo_view, name="api-brand-logo"),
    path("logo", logo_view, name="logo"),
]
