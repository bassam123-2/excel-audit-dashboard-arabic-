from __future__ import annotations

from pathlib import Path

from django.http import JsonResponse
from rest_framework.decorators import api_view

from apps.dashboard.services.query_service import (
    build_aging_payload,
    build_audit_plan_payload,
    build_summary_payload,
    parse_multi_filter_from_request,
)
from apps.uploads.services.upload_service import get_active_upload_from_request


def _active_path_or_error(request):
    upload_session = get_active_upload_from_request(request)
    if not upload_session:
        return None, JsonResponse({"error": "No uploaded file"}, status=400)
    path = Path(upload_session.stored_path)
    if not path.exists():
        return None, JsonResponse({"error": "Uploaded file not found"}, status=400)
    return path, None


@api_view(["GET"])
def summary_view(request):
    path, err = _active_path_or_error(request)
    if err:
        return err
    selected = parse_multi_filter_from_request(request)
    payload = build_summary_payload(path, selected)
    return JsonResponse(payload)


@api_view(["GET"])
def aging_summary_view(request):
    path, err = _active_path_or_error(request)
    if err:
        return err
    reference_raw = (request.GET.get("reference") or "").strip()
    if not reference_raw:
        return JsonResponse({"error": "Missing reference date"}, status=400)

    selected = parse_multi_filter_from_request(request)
    date_source = request.GET.get("aging_date_source", "target")
    try:
        payload = build_aging_payload(path, selected, reference_raw=reference_raw, date_source=date_source)
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)
    except KeyError as exc:
        return JsonResponse({"error": str(exc)}, status=400)
    return JsonResponse(payload)


@api_view(["GET"])
def audit_plan_panel_view(request):
    path, err = _active_path_or_error(request)
    if err:
        return err
    selected = parse_multi_filter_from_request(request)
    payload = build_audit_plan_payload(path, selected)
    return JsonResponse(payload)
