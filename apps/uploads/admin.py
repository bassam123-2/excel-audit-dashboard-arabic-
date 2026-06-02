from django.contrib import admin

from .models import EmailLog, ExportLog, UploadSession


@admin.register(UploadSession)
class UploadSessionAdmin(admin.ModelAdmin):
    list_display = ("id", "original_filename", "uploaded_at", "uploaded_by", "is_active")
    search_fields = ("original_filename", "file_hash")
    list_filter = ("is_active", "uploaded_at")


@admin.register(ExportLog)
class ExportLogAdmin(admin.ModelAdmin):
    list_display = ("id", "upload_session", "export_type", "status", "created_at")
    list_filter = ("export_type", "status", "created_at")


@admin.register(EmailLog)
class EmailLogAdmin(admin.ModelAdmin):
    list_display = ("id", "upload_session", "recipient", "status", "sent_at")
    list_filter = ("status", "sent_at")
