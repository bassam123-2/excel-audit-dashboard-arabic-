from __future__ import annotations

import logging
from pathlib import Path

from django.conf import settings
from django.http import HttpRequest, HttpResponse
from django.shortcuts import redirect, render

logger = logging.getLogger(__name__)

from apps.dashboard.services.legacy_bridge import legacy
from apps.dashboard.services.query_service import build_company_columns
from apps.uploads.services.import_service import import_records_from_excel
from apps.uploads.services.upload_service import get_active_upload_from_request, save_upload


def upload_view(request: HttpRequest) -> HttpResponse:
    if request.method == "POST":
        uploaded_file = request.FILES.get("excel_file")
        if not uploaded_file or not uploaded_file.name:
            return render(request, "upload.html", {"error": "الرجاء اختيار ملف Excel."})

        suffix = Path(uploaded_file.name).suffix.lower()
        if suffix not in {".xlsx", ".xls"}:
            return render(request, "upload.html", {"error": "الملف يجب أن يكون بصيغة Excel."})

        upload_session = save_upload(uploaded_file, user=request.user)
        saved_path = Path(upload_session.stored_path)

        try:
            legacy._load_dashboard_df(saved_path)
            row_count = import_records_from_excel(upload_session)
            if row_count <= 0:
                raise ValueError(
                    "No data rows found. Use the compliance register sheet: سجل الالتزام الموحد"
                )
            logger.info("Imported %s rows for upload %s", row_count, upload_session.id)
        except Exception as exc:
            logger.exception("Upload/import failed for %s", uploaded_file.name)
            saved_path.unlink(missing_ok=True)
            upload_session.delete()
            request.session.pop("active_upload_session_id", None)
            error = "تعذر قراءة الملف. تأكد من أن التنسيق يطابق نموذج سجل الالتزام."
            if settings.DEBUG:
                error = f"{error} ({exc})"
            return render(request, "upload.html", {"error": error})

        request.session["active_upload_session_id"] = str(upload_session.id)
        return redirect("analyze")

    return render(request, "upload.html")


def analyze_view(request: HttpRequest) -> HttpResponse:
    upload_session = get_active_upload_from_request(request)
    if not upload_session:
        return redirect("upload")

    company_columns = build_company_columns(Path(upload_session.stored_path))
    response = render(
        request,
        "analyze.html",
        {
            "company_columns": company_columns,
            "snapshot_pack_json": None,
            "snapshot_logo_url": "",
            "upload_row_count": upload_session.row_count,
            "upload_filename": upload_session.original_filename,
        },
    )
    response["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response["Pragma"] = "no-cache"
    return response
