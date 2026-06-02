from __future__ import annotations

import uuid

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):
    initial = True

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name="UploadSession",
            fields=[
                ("id", models.UUIDField(default=uuid.uuid4, editable=False, primary_key=True, serialize=False)),
                ("original_filename", models.CharField(max_length=255)),
                ("stored_path", models.CharField(max_length=1024)),
                ("file_hash", models.CharField(db_index=True, max_length=64)),
                ("uploaded_at", models.DateTimeField(auto_now_add=True)),
                ("is_active", models.BooleanField(default=True)),
                (
                    "uploaded_by",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="excel_upload_sessions",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
            ],
            options={"ordering": ["-uploaded_at"]},
        ),
        migrations.CreateModel(
            name="ExportLog",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("export_type", models.CharField(choices=[("pptx", "PowerPoint"), ("html", "HTML Snapshot")], max_length=16)),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("status", models.CharField(choices=[("ok", "Success"), ("error", "Error")], default="ok", max_length=16)),
                ("error_message", models.TextField(blank=True)),
                (
                    "upload_session",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="export_logs",
                        to="uploads.uploadsession",
                    ),
                ),
            ],
            options={"ordering": ["-created_at"]},
        ),
        migrations.CreateModel(
            name="EmailLog",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("recipient", models.EmailField(max_length=254)),
                ("subject", models.CharField(blank=True, max_length=255)),
                ("sent_at", models.DateTimeField(auto_now_add=True)),
                ("status", models.CharField(choices=[("ok", "Success"), ("error", "Error")], default="ok", max_length=16)),
                ("error_message", models.TextField(blank=True)),
                (
                    "upload_session",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="email_logs",
                        to="uploads.uploadsession",
                    ),
                ),
            ],
            options={"ordering": ["-sent_at"]},
        ),
    ]
