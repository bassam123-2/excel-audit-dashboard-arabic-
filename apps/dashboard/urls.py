from __future__ import annotations

from django.urls import path

from .views import aging_summary_view, audit_plan_panel_view, summary_view

urlpatterns = [
    path("api/summary", summary_view, name="api-summary"),
    path("api/aging-summary", aging_summary_view, name="api-aging-summary"),
    path("api/audit-plan-panel", audit_plan_panel_view, name="api-audit-plan-panel"),
]
