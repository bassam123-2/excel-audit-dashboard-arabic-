from __future__ import annotations

from django.http import FileResponse, HttpResponse
from rest_framework.decorators import api_view

from apps.dashboard.services.legacy_bridge import legacy


@api_view(["GET"])
def brand_logo_view(request):
    code = (request.GET.get("code") or "").strip()
    if code not in legacy.BRAND_LOGO_CODES:
        code = (
            legacy._pick_brand_logo_code(
                request.GET.getlist("subsidiary_company"),
                request.GET.getlist("holding_company"),
            )
            or ""
        )
    if not code:
        return HttpResponse(status=204)
    path = legacy._brand_logo_path(code)
    if path is None:
        return HttpResponse(status=204)
    return FileResponse(open(path, "rb"), content_type=legacy._brand_logo_mime(path))


@api_view(["GET"])
def logo_view(request):
    if legacy.LOGO_PATH is not None and legacy.LOGO_PATH.is_file():
        return FileResponse(open(legacy.LOGO_PATH, "rb"))
    return HttpResponse(legacy._TINY_PNG, content_type="image/png")
