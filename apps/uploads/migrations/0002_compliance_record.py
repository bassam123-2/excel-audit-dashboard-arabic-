from __future__ import annotations

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("uploads", "0001_initial"),
    ]

    operations = [
        migrations.AddField(
            model_name="uploadsession",
            name="row_count",
            field=models.PositiveIntegerField(default=0),
        ),
        migrations.CreateModel(
            name="ComplianceRecord",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("excel_row", models.PositiveIntegerField(help_text="1-based row number in the Excel file")),
                ("row_index", models.PositiveIntegerField(help_text="0-based pandas dataframe index")),
                ("inherent_risk", models.CharField(blank=True, default="", max_length=255)),
                ("residual_risk", models.CharField(blank=True, default="", max_length=255)),
                ("status", models.CharField(blank=True, db_index=True, default="", max_length=255)),
                ("target_date", models.DateField(blank=True, null=True)),
                ("year_value", models.CharField(blank=True, db_index=True, default="", max_length=32)),
                ("department", models.CharField(blank=True, default="", max_length=255)),
                ("legislator", models.CharField(blank=True, default="", max_length=255)),
                ("system_name", models.CharField(blank=True, default="", max_length=255)),
                ("authority", models.CharField(blank=True, default="", max_length=255)),
                ("regulation", models.CharField(blank=True, default="", max_length=255)),
                ("legal_text", models.TextField(blank=True, default="")),
                ("compliance_status", models.CharField(blank=True, db_index=True, default="", max_length=255)),
                ("control_category", models.CharField(blank=True, default="", max_length=255)),
                ("task_owner", models.CharField(blank=True, default="", max_length=255)),
                ("responsible_person", models.CharField(blank=True, default="", max_length=255)),
                ("corrective_plan", models.TextField(blank=True, default="")),
                ("modified_date", models.DateField(blank=True, null=True)),
                ("management_notes", models.TextField(blank=True, default="")),
                ("compliance_notes", models.TextField(blank=True, default="")),
                ("holding_company", models.CharField(blank=True, default="", max_length=255)),
                ("subsidiary_company", models.CharField(blank=True, default="", max_length=255)),
                ("email", models.CharField(blank=True, default="", max_length=255)),
                ("imported_at", models.DateTimeField(auto_now_add=True)),
                (
                    "upload_session",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="records",
                        to="uploads.uploadsession",
                    ),
                ),
            ],
            options={
                "ordering": ["excel_row"],
            },
        ),
        migrations.AddIndex(
            model_name="compliancerecord",
            index=models.Index(fields=["upload_session", "status"], name="uploads_com_upload__a8e2c1_idx"),
        ),
        migrations.AddIndex(
            model_name="compliancerecord",
            index=models.Index(fields=["upload_session", "residual_risk"], name="uploads_com_upload__b3f4d2_idx"),
        ),
        migrations.AddConstraint(
            model_name="compliancerecord",
            constraint=models.UniqueConstraint(
                fields=("upload_session", "excel_row"),
                name="uniq_record_per_upload_excel_row",
            ),
        ),
    ]
