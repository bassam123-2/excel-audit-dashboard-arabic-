from __future__ import annotations

from django.test import Client, TestCase


class BasicRoutesTests(TestCase):
    def setUp(self):
        self.client = Client()

    def test_upload_page_loads(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)

    def test_summary_without_upload_returns_error(self):
        response = self.client.get("/api/summary")
        self.assertEqual(response.status_code, 400)
