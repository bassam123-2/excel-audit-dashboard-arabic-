from __future__ import annotations

from django.test import TestCase

from apps.uploads.models import UploadSession


class UploadSessionModelTests(TestCase):
    def test_upload_session_defaults(self):
        item = UploadSession.objects.create(
            original_filename="sample.xlsx",
            stored_path="C:/tmp/sample.xlsx",
            file_hash="a" * 64,
        )
        self.assertTrue(item.is_active)
        self.assertEqual(item.original_filename, "sample.xlsx")
