from __future__ import annotations

from django.core.management.base import BaseCommand

from apps.uploads.models import ComplianceRecord, UploadSession


class Command(BaseCommand):
    help = "Remove upload sessions that have no imported rows (test/bad uploads)."

    def handle(self, *args, **options):
        ids = list(UploadSession.objects.filter(row_count=0).values_list("id", flat=True))
        count = len(ids)
        UploadSession.objects.filter(id__in=ids).delete()
        self.stdout.write(self.style.SUCCESS(f"Removed {count} empty upload session(s)."))
