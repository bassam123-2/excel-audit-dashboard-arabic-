from __future__ import annotations

from django.conf import settings
from django.core.management.base import BaseCommand
from django.db import connection

from apps.uploads.models import ComplianceRecord, UploadSession


class Command(BaseCommand):
    help = "Show database connection and upload/import row counts."

    def handle(self, *args, **options):
        db = settings.DATABASES["default"]
        self.stdout.write(f"Database engine: {db['ENGINE']}")
        self.stdout.write(f"Database name:   {db.get('NAME')}")
        self.stdout.write(f"Database host:   {db.get('HOST') or '(local)'}")

        with connection.cursor() as cursor:
            cursor.execute("SHOW TABLES")
            tables = [row[0] for row in cursor.fetchall() if row[0].startswith("uploads_")]
        self.stdout.write(f"Upload tables:   {', '.join(tables) or '(none)'}")

        sessions = UploadSession.objects.count()
        records = ComplianceRecord.objects.count()
        self.stdout.write(f"Upload sessions: {sessions}")
        self.stdout.write(f"Compliance rows: {records}")

        latest = UploadSession.objects.order_by("-uploaded_at").first()
        if latest:
            self.stdout.write(
                f"Latest upload:   {latest.original_filename!r} "
                f"(row_count={latest.row_count}, db_rows={latest.records.count()})"
            )
        else:
            self.stdout.write("Latest upload:   (none — upload a file in the app)")
