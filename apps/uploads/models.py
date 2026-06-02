from __future__ import annotations

import uuid

from django.conf import settings
from django.db import models


class UploadSession(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    original_filename = models.CharField(max_length=255)
    stored_path = models.CharField(max_length=1024)
    file_hash = models.CharField(max_length=64, db_index=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="excel_upload_sessions",
    )
    is_active = models.BooleanField(default=True)

    class Meta:
        ordering = ["-uploaded_at"]


class ExportLog(models.Model):
    EXPORT_TYPES = (
        ("pptx", "PowerPoint"),
        ("html", "HTML Snapshot"),
    )
    STATUS_TYPES = (
        ("ok", "Success"),
        ("error", "Error"),
    )

    upload_session = models.ForeignKey(
        UploadSession,
        on_delete=models.CASCADE,
        related_name="export_logs",
    )
    export_type = models.CharField(max_length=16, choices=EXPORT_TYPES)
    created_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=16, choices=STATUS_TYPES, default="ok")
    error_message = models.TextField(blank=True)

    class Meta:
        ordering = ["-created_at"]


class EmailLog(models.Model):
    STATUS_TYPES = (
        ("ok", "Success"),
        ("error", "Error"),
    )

    upload_session = models.ForeignKey(
        UploadSession,
        on_delete=models.CASCADE,
        related_name="email_logs",
    )
    recipient = models.EmailField()
    subject = models.CharField(max_length=255, blank=True)
    sent_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=16, choices=STATUS_TYPES, default="ok")
    error_message = models.TextField(blank=True)

    class Meta:
        ordering = ["-sent_at"]
