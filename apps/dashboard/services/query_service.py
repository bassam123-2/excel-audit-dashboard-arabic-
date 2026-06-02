from __future__ import annotations

from pathlib import Path

from apps.dashboard.services.legacy_bridge import legacy


def _multi_filter_param_map() -> dict[str, str]:
    return legacy._multi_filter_param_map()


def parse_multi_filter_from_request(request) -> dict[str, list[str]]:
    mapping = _multi_filter_param_map()
    return {col: [v.strip() for v in request.GET.getlist(param) if v.strip()] for col, param in mapping.items()}


def _normalize_for_filters(df):
    out = legacy._add_year_column_to_df(df)
    out = legacy._normalize_filter_dimensions(out)
    return out


def build_company_columns(path: Path) -> dict[str, bool]:
    try:
        df = legacy._load_full_df(path)
        return {
            "holding": legacy.COL_HOLDING_COMPANY in df.columns,
            "subsidiary": legacy.COL_SUBSIDIARY_COMPANY in df.columns,
        }
    except Exception:
        return {"holding": False, "subsidiary": False}


def build_summary_payload(path: Path, selected: dict[str, list[str]]) -> dict[str, object]:
    df = legacy._load_full_df(path)
    df = _normalize_for_filters(df)
    fully_filtered = legacy._apply_filters(df, selected)

    groups = {}
    for key in legacy._filter_dimension_cols_in_df(df) + [legacy.YEAR_COL]:
        group_df = legacy._apply_filters(df, selected, skip_key=key)
        groups[key] = legacy._build_group_data(group_df, key)

    return {
        "total": int(len(fully_filtered)),
        "selected": selected,
        "groups": groups,
        "company_columns": {
            "holding": legacy.COL_HOLDING_COMPANY in df.columns,
            "subsidiary": legacy.COL_SUBSIDIARY_COMPANY in df.columns,
        },
    }


def build_aging_payload(path: Path, selected: dict[str, list[str]], *, reference_raw: str, date_source: str) -> dict[str, object]:
    reference = legacy.pd.to_datetime(reference_raw, errors="coerce")
    if legacy.pd.isna(reference):
        raise ValueError("Invalid reference date")

    date_source = date_source.strip().lower()
    if date_source not in {"target", "modified"}:
        raise ValueError("Invalid aging_date_source")
    date_col = legacy.COL_MODIFIED_DATE if date_source == "modified" else legacy.COL_TARGET_DATE

    work = legacy._prepare_filtered_full_df(path, selected)
    if date_col not in work.columns:
        raise KeyError(f"Missing column: {date_col}")
    if legacy.COL_STATUS not in work.columns or legacy.COL_RESIDUAL_RISK not in work.columns:
        raise KeyError("Missing status or risk column")

    before_open = len(work)
    work = work[work[legacy.COL_STATUS].map(legacy._is_open_status_for_aging)]
    skipped_other_status = int(before_open - len(work))

    risk_keys = [k for k, _, _ in legacy.AGING_RISK_COLS]
    matrix = {tid: {rk: 0 for rk in risk_keys} for tid, _ in legacy.AGING_TIME_ROWS}
    unknown_time = 0

    compare_series = legacy._parse_excel_date_series(work[date_col])
    for idx in work.index:
        rkey = legacy._aging_risk_key(legacy._normalize_value(work.loc[idx, legacy.COL_RESIDUAL_RISK])) or "other"
        cdate = compare_series.loc[idx]
        if legacy.pd.isna(cdate):
            unknown_time += 1
            continue
        tkey = legacy._aging_time_bucket(cdate, reference)
        if tkey is None:
            unknown_time += 1
            continue
        matrix[tkey][rkey] += 1

    rows_out = []
    for tid, tlabel in legacy.AGING_TIME_ROWS:
        cells = matrix[tid]
        row_total = sum(int(cells[k]) for k in risk_keys)
        rows_out.append({"id": tid, "label": tlabel, "cells": cells, "total": row_total})

    col_totals = {k: sum(matrix[tid][k] for tid, _ in legacy.AGING_TIME_ROWS) for k in risk_keys}
    grand_total = sum(col_totals.values())

    return {
        "reference": reference.strftime("%Y-%m-%d"),
        "date_field": date_col,
        "date_source": date_source,
        "risk_columns": [{"id": k, "label": lab, "color": col} for k, lab, col in legacy.AGING_RISK_COLS],
        "time_rows": rows_out,
        "column_totals": col_totals,
        "grand_total": int(grand_total),
        "status_filter": "open_only",
        "skipped_other_status": skipped_other_status,
        "skipped_unknown_time": int(unknown_time),
    }


def build_audit_plan_payload(path: Path, selected: dict[str, list[str]]) -> dict[str, object]:
    work = legacy._prepare_filtered_full_df(path, selected)
    cols = legacy._audit_plan_column_list(work)

    columns_out = []
    for col in cols:
        labels = legacy._series_labels_for_audit_column(col, work[col])
        vc = labels.value_counts(dropna=False)
        items = [{"label": str(k), "count": int(v)} for k, v in vc.head(legacy.AUDIT_PLAN_VALUECAP).items()]
        non_null = int((labels != legacy.BLANK_LABEL).sum())
        columns_out.append(
            {
                "name": col,
                "entries": items,
                "truncated": bool(len(vc) > legacy.AUDIT_PLAN_VALUECAP),
                "distinct": int(len(vc)),
                "non_null": non_null,
            }
        )
    return {"total_rows": int(len(work)), "columns": columns_out}
