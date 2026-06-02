from __future__ import annotations

import hashlib
from pathlib import Path
from uuid import uuid4

from django.conf import settings

from apps.uploads.models import UploadSession


def _uploads_dir() -> Path:
    root = Path(settings.MEDIA_ROOT)
    root.mkdir(parents=True, exist_ok=True)
    return root


def save_upload(uploaded_file, user=None) -> UploadSession:
    suffix = Path(uploaded_file.name).suffix.lower()
    stored_name = f"{uuid4().hex}{suffix}"
    stored_path = _uploads_dir() / stored_name

    sha = hashlib.sha256()
    with stored_path.open("wb") as target:
        for chunk in uploaded_file.chunks():
            target.write(chunk)
            sha.update(chunk)

    return UploadSession.objects.create(
        original_filename=uploaded_file.name,
        stored_path=str(stored_path.resolve()),
        file_hash=sha.hexdigest(),
        uploaded_by=user if getattr(user, "is_authenticated", False) else None,
        is_active=True,
    )


def get_active_upload_from_request(request) -> UploadSession | None:
    upload_id = request.session.get("active_upload_session_id")
    if not upload_id:
        return None
    try:
        return UploadSession.objects.get(id=upload_id, is_active=True)
    except UploadSession.DoesNotExist:
        return None
