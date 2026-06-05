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
    row_count = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["-uploaded_at"]


class ComplianceRecord(models.Model):
    """One row from sheet سجل الالتزام الموحد (Excel row 5+)."""

    upload_session = models.ForeignKey(
        UploadSession,
        on_delete=models.CASCADE,
        related_name="records",
    )
    excel_row = models.PositiveIntegerField(help_text="1-based row number in the Excel file")
    row_index = models.PositiveIntegerField(help_text="0-based pandas dataframe index")

    inherent_risk = models.CharField(max_length=255, blank=True, default="")
    residual_risk = models.CharField(max_length=255, blank=True, default="")
    status = models.CharField(max_length=255, blank=True, default="", db_index=True)
    target_date = models.DateField(null=True, blank=True)
    year_value = models.CharField(max_length=32, blank=True, default="", db_index=True)
    department = models.CharField(max_length=255, blank=True, default="")
    legislator = models.CharField(max_length=255, blank=True, default="")
    system_name = models.CharField(max_length=255, blank=True, default="")
    authority = models.CharField(max_length=255, blank=True, default="")
    regulation = models.CharField(max_length=255, blank=True, default="")
    legal_text = models.TextField(blank=True, default="")
    compliance_status = models.CharField(max_length=255, blank=True, default="", db_index=True)
    control_category = models.CharField(max_length=255, blank=True, default="")

    task_owner = models.CharField(max_length=255, blank=True, default="")
    responsible_person = models.CharField(max_length=255, blank=True, default="")
    corrective_plan = models.TextField(blank=True, default="")
    modified_date = models.DateField(null=True, blank=True)
    management_notes = models.TextField(blank=True, default="")
    compliance_notes = models.TextField(blank=True, default="")
    holding_company = models.CharField(max_length=255, blank=True, default="")
    subsidiary_company = models.CharField(max_length=255, blank=True, default="")
    email = models.CharField(max_length=255, blank=True, default="")

    imported_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["excel_row"]
        constraints = [
            models.UniqueConstraint(
                fields=["upload_session", "excel_row"],
                name="uniq_record_per_upload_excel_row",
            )
        ]
        indexes = [
            models.Index(fields=["upload_session", "status"]),
            models.Index(fields=["upload_session", "residual_risk"]),
        ]

    def __str__(self) -> str:
        preview = (self.legal_text or self.status or "")[:60]
        return f"Row {self.excel_row}: {preview}"


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
