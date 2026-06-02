from __future__ import annotations

from io import BytesIO
from pathlib import Path

import os

from django.http import FileResponse, HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import api_view

from apps.dashboard.services.legacy_bridge import legacy
from apps.legal.services.legal_service import (
    build_dashboard_snapshot_html,
    collect_legal_text_record,
    export_legal_text_pptx,
    log_email,
    log_export,
    send_legal_text_email,
)
from apps.uploads.services.upload_service import get_active_upload_from_request


def _apply_email_api_cors(request, response):
    origin = (request.headers.get("Origin") or "").strip()
    port = os.environ.get("EXCEL_ARABIC_PORT", "8765")
    allowed = {"null", f"http://127.0.0.1:{port}", f"http://localhost:{port}"}
    if origin in allowed:
        response["Access-Control-Allow-Origin"] = origin
        response["Vary"] = "Origin"
        response["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response["Access-Control-Allow-Headers"] = "Content-Type"
    return response


def _active_path_or_error(request):
    upload_session = get_active_upload_from_request(request)
    if not upload_session:
        return None, None, JsonResponse({"error": "No uploaded file"}, status=400)
    path = Path(upload_session.stored_path)
    if not path.exists():
        return None, None, JsonResponse({"error": "Uploaded file not found"}, status=400)
    return upload_session, path, None


def extracted_image_view(request, filename: str):
    safe_name = (filename or "").strip().replace("\\", "/")
    if ".." in safe_name:
        return HttpResponse(status=404)
    file_path = legacy.EXTRACTED_IMAGES_DIR / safe_name
    if not file_path.exists() or not file_path.is_file():
        return HttpResponse(status=404)
    return FileResponse(open(file_path, "rb"))


@api_view(["GET", "POST"])
def legal_text_details_view(request):
    upload_session, path, err = _active_path_or_error(request)
    if err:
        return err

    if request.method == "POST":
        payload = request.data or {}
        legal_text = str(payload.get("text", "")).strip()
        include_images = bool(payload.get("include_images", False))
    else:
        legal_text = (request.GET.get("text") or "").strip()
        include_images = (request.GET.get("include_images", "1").strip().lower() not in {"0", "false", "no"})

    if not legal_text:
        return JsonResponse({"error": "Missing text"}, status=400)

    try:
        record = collect_legal_text_record(path, legal_text, include_images=include_images)
    except LookupError:
        return JsonResponse({"error": "Not found"}, status=404)
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)
    return JsonResponse(record)


@api_view(["GET"])
def legal_text_row_images_view(request):
    upload_session, path, err = _active_path_or_error(request)
    if err:
        return err
    excel_row = request.GET.get("excel_row")
    try:
        excel_row_int = int(excel_row)
        if excel_row_int < 1:
            raise ValueError
    except Exception:
        return JsonResponse({"error": "Invalid excel_row"}, status=400)

    images = legacy._extract_images_for_excel_row(path, excel_row_int)
    return JsonResponse({"images": images})


@api_view(["POST"])
def export_legal_text_pptx_view(request):
    payload = request.data or {}
    legal_text = str(payload.get("text", "")).strip()
    if not legal_text:
        return JsonResponse({"error": "Missing text"}, status=400)

    upload_session = get_active_upload_from_request(request)
    path = Path(upload_session.stored_path) if upload_session else None
    if path is not None and not path.exists():
        path = None

    try:
        pptx_buf, fname = export_legal_text_pptx(path, payload)
        log_export(upload_session, "pptx", ok=True)
    except LookupError:
        log_export(upload_session, "pptx", ok=False, error_message="Not found")
        return JsonResponse({"error": "Not found"}, status=404)
    except ValueError as exc:
        log_export(upload_session, "pptx", ok=False, error_message=str(exc))
        return JsonResponse({"error": str(exc)}, status=400)
    except Exception:
        log_export(upload_session, "pptx", ok=False, error_message="PPTX export failed")
        return JsonResponse({"error": "فشل إنشاء ملف PowerPoint."}, status=500)

    response = FileResponse(
        pptx_buf,
        as_attachment=True,
        filename=fname,
        content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    return response


@csrf_exempt
@api_view(["POST", "OPTIONS"])
def send_legal_text_email_view(request):
    if request.method == "OPTIONS":
        return _apply_email_api_cors(request, HttpResponse(status=204))

    payload = request.data or {}
    upload_session = get_active_upload_from_request(request)
    path = Path(upload_session.stored_path) if upload_session else None
    if path is not None and not path.exists():
        path = None

    try:
        to_addr = send_legal_text_email(path, payload)
        log_email(upload_session, to_addr, ok=True)
        return _apply_email_api_cors(request, JsonResponse({"ok": True, "to": to_addr}))
    except LookupError:
        log_email(upload_session, "", ok=False, error_message="Not found")
        return _apply_email_api_cors(request, JsonResponse({"error": "Not found"}, status=404))
    except ValueError as exc:
        log_email(upload_session, "", ok=False, error_message=str(exc))
        return _apply_email_api_cors(request, JsonResponse({"error": str(exc)}, status=400))
    except Exception as exc:
        log_email(upload_session, "", ok=False, error_message=str(exc))
        return _apply_email_api_cors(request, JsonResponse({"error": str(exc)}, status=500))


@api_view(["GET"])
def export_dashboard_html_view(request):
    upload_session, path, err = _active_path_or_error(request)
    if err:
        return err
    try:
        html_bytes, fname = build_dashboard_snapshot_html(path, request.GET)
        log_export(upload_session, "html", ok=True)
    except Exception:
        log_export(upload_session, "html", ok=False, error_message="HTML snapshot export failed")
        return JsonResponse({"error": "فشل بناء ملف التصدير."}, status=500)

    return FileResponse(BytesIO(html_bytes), as_attachment=True, filename=fname, content_type="application/octet-stream")
