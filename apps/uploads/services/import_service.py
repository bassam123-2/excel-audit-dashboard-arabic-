from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

from apps.dashboard.services.legacy_bridge import legacy
from apps.uploads.models import ComplianceRecord, UploadSession


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _cell_date(value: object) -> date | None:
    parsed = legacy._parse_excel_date_value(value)
    if parsed is None or pd.isna(parsed):
        return None
    return parsed.date()


def _row_value(row: pd.Series, column: str) -> object:
    if column not in row.index:
        return None
    return row[column]


def import_records_from_excel(upload_session: UploadSession) -> int:
    """Parse uploaded workbook and persist all rows into ComplianceRecord."""
    path = Path(upload_session.stored_path)
    df = legacy._load_full_df(path)
    df = legacy._add_year_column_to_df(df)

    ComplianceRecord.objects.filter(upload_session=upload_session).delete()

    records: list[ComplianceRecord] = []
    for idx, row in df.iterrows():
        excel_row = legacy._excel_row_for_df_index(int(idx))
        year_raw = _row_value(row, legacy.YEAR_COL)
        if year_raw is None or (isinstance(year_raw, float) and pd.isna(year_raw)):
            year_raw = _row_value(row, legacy.COL_YEAR)
        year_value = legacy._extract_valid_year(year_raw) or _cell_text(year_raw)

        records.append(
            ComplianceRecord(
                upload_session=upload_session,
                excel_row=excel_row,
                row_index=int(idx),
                inherent_risk=_cell_text(_row_value(row, legacy.COL_INHERENT_RISK)),
                residual_risk=_cell_text(_row_value(row, legacy.COL_RESIDUAL_RISK)),
                status=_cell_text(_row_value(row, legacy.COL_STATUS)),
                target_date=_cell_date(_row_value(row, legacy.COL_TARGET_DATE)),
                year_value=year_value or "",
                department=_cell_text(_row_value(row, legacy.COL_DEPARTMENT)),
                legislator=_cell_text(_row_value(row, legacy.COL_LEGISLATOR)),
                system_name=_cell_text(_row_value(row, legacy.COL_SYSTEM_NAME)),
                authority=_cell_text(_row_value(row, legacy.COL_AUTHORITY)),
                regulation=_cell_text(_row_value(row, legacy.COL_REGULATION)),
                legal_text=_cell_text(_row_value(row, legacy.COL_LEGAL_TEXT)),
                compliance_status=_cell_text(_row_value(row, legacy.COL_COMPLIANCE_STATUS)),
                control_category=_cell_text(_row_value(row, legacy.COL_CONTROL_CATEGORY)),
                task_owner=_cell_text(_row_value(row, legacy.COL_TASK_OWNER)),
                responsible_person=_cell_text(_row_value(row, legacy.COL_RESPONSIBLE_PERSON)),
                corrective_plan=_cell_text(_row_value(row, legacy.COL_CORRECTIVE_PLAN)),
                modified_date=_cell_date(_row_value(row, legacy.COL_MODIFIED_DATE)),
                management_notes=_cell_text(_row_value(row, legacy.COL_MANAGEMENT_NOTES)),
                compliance_notes=_cell_text(_row_value(row, legacy.COL_COMPLIANCE_NOTES)),
                holding_company=_cell_text(_row_value(row, legacy.COL_HOLDING_COMPANY)),
                subsidiary_company=_cell_text(_row_value(row, legacy.COL_SUBSIDIARY_COMPANY)),
                email=_cell_text(_row_value(row, legacy.COL_EMAIL)),
            )
        )

    ComplianceRecord.objects.bulk_create(records, batch_size=500)
    upload_session.row_count = len(records)
    upload_session.save(update_fields=["row_count"])
    return len(records)
