from __future__ import annotations

from django.urls import path

from .views import (
    export_dashboard_html_view,
    export_legal_text_pptx_view,
    extracted_image_view,
    legal_text_details_view,
    legal_text_row_images_view,
    send_legal_text_email_view,
)

urlpatterns = [
    path("extracted/<path:filename>", extracted_image_view, name="extracted-image"),
    path("api/legal-text-details", legal_text_details_view, name="api-legal-text-details"),
    path("api/legal-text-row-images", legal_text_row_images_view, name="api-legal-text-row-images"),
    path("api/export-legal-text-pptx", export_legal_text_pptx_view, name="api-export-legal-text-pptx"),
    path("api/send-legal-text-email", send_legal_text_email_view, name="api-send-legal-text-email"),
    path("api/export-dashboard-html", export_dashboard_html_view, name="api-export-dashboard-html"),
]
