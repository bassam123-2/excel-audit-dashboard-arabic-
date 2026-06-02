from __future__ import annotations

from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("apps.uploads.urls")),
    path("", include("apps.dashboard.urls")),
    path("", include("apps.legal.urls")),
    path("", include("apps.branding.urls")),
]
