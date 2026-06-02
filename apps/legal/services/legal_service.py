from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path

from django.template.loader import render_to_string

from apps.dashboard.services.legacy_bridge import legacy
from apps.dashboard.services.query_service import build_company_columns
from apps.uploads.models import EmailLog, ExportLog, UploadSession


def collect_legal_text_record(path: Path, legal_text: str, *, include_images: bool) -> dict[str, object]:
    return legacy._collect_legal_text_record(path, legal_text, include_images=include_images)


def export_legal_text_pptx(path: Path | None, payload: dict[str, object]) -> tuple[BytesIO, str]:
    legal_text = str(payload.get("text", "")).strip()
    if not legal_text:
        raise ValueError("Missing text")

    record: dict[str, object] | None = None
    if path is not None and path.exists():
        record = collect_legal_text_record(path, legal_text, include_images=True)
    if record is None:
        fields = payload.get("fields")
        if isinstance(fields, list) and fields:
            record = {
                "legal_text": legal_text,
                "recipient_email": str(payload.get("recipient_email", "")).strip(),
                "fields": fields,
                "images": payload.get("images") if isinstance(payload.get("images"), list) else [],
            }
        else:
            raise ValueError("No uploaded file")

    pptx_buf = legacy._build_legal_text_pptx(record)
    fname = legacy._safe_pptx_filename(str(record.get("legal_text") or legal_text))
    return pptx_buf, fname


def send_legal_text_email(path: Path | None, payload: dict[str, object]) -> str:
    legal_text = str(payload.get("text", "")).strip()
    if not legal_text:
        raise ValueError("Missing text")

    direct_to = legacy._normalize_recipient_email(payload.get("to"))
    if direct_to:
        legacy._send_smtp_message(direct_to, legacy.LEGAL_EMAIL_SUBJECT_PLACEHOLDER, legal_text)
        return direct_to

    if path is None or not path.exists():
        raise ValueError("No uploaded file")

    df = legacy._load_full_df(path)
    if legacy.COL_LEGAL_TEXT not in df.columns:
        raise ValueError("Legal text column missing")

    normalized = df[legacy.COL_LEGAL_TEXT].map(legacy._normalize_value)
    lookup_key = legacy._normalize_value(legal_text)
    matches = df[normalized == lookup_key]
    if matches.empty:
        raise LookupError("Not found")

    row = legacy._pick_best_legal_match(matches)
    idx = int(row.name)
    excel_row = legacy._excel_row_for_df_index(idx)
    to_addr = legacy._normalize_recipient_email(legacy._read_email_at_excel_row(path, excel_row))
    if not to_addr and legacy.COL_EMAIL in df.columns:
        raw_mail = row[legacy.COL_EMAIL] if legacy.COL_EMAIL in row.index else None
        to_addr = legacy._normalize_recipient_email(raw_mail)

    if not to_addr:
        if legacy._find_email_excel_column_1based(path) is None:
            raise ValueError("لم يُعثر على عمود email في أول صفوف الملف.")
        raise ValueError("خلية البريد فارغة أو غير صالحة في صف الإكسل لهذا السجل.")

    legacy._send_smtp_message(to_addr, legacy.LEGAL_EMAIL_SUBJECT_PLACEHOLDER, lookup_key)
    return to_addr


def build_dashboard_snapshot_html(path: Path, request_get) -> tuple[bytes, str]:
    payload = legacy._build_snapshot_payload(path)
    snapshot_json = json.dumps(payload, ensure_ascii=False).replace("<", "\\u003c")
    brand_code = legacy._pick_brand_logo_code(
        request_get.getlist("subsidiary_company"),
        request_get.getlist("holding_company"),
    )
    html = render_to_string(
        "analyze.html",
        {
            "snapshot_pack_json": snapshot_json,
            "snapshot_logo_url": legacy._brand_logo_data_url(brand_code),
            "company_columns": build_company_columns(path),
        },
    )
    fname = f"dashboard-snapshot-{legacy.pd.Timestamp.now().strftime('%Y-%m-%d')}.html"
    return html.encode("utf-8"), fname


def log_export(upload_session: UploadSession | None, export_type: str, *, ok: bool, error_message: str = "") -> None:
    if upload_session is None:
        return
    ExportLog.objects.create(
        upload_session=upload_session,
        export_type=export_type,
        status="ok" if ok else "error",
        error_message=error_message,
    )


def log_email(upload_session: UploadSession | None, recipient: str, *, ok: bool, error_message: str = "") -> None:
    if upload_session is None:
        return
    EmailLog.objects.create(
        upload_session=upload_session,
        recipient=recipient or "unknown@example.com",
        subject=legacy.LEGAL_EMAIL_SUBJECT_PLACEHOLDER,
        status="ok" if ok else "error",
        error_message=error_message,
    )
