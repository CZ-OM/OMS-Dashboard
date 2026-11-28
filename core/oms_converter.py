import math
import re
import warnings
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, BinaryIO, Optional, Union

import numpy as np
import pandas as pd

from .loader import (
    HEADER_KEYWORDS,
    _ensure_seekable_source,
    _get_bytes_or_buffer,
    _read_csv_with_fallback,
    _read_raw_bytes,
    _read_xls_dataframe,
    _safe_seek_zero,
    detect_header_row,
    load_dataset,
)

OMS_DATETIME_COLUMNS: tuple[str, ...] = (
    "Updated At",
    "Created At",
    "Assigned At",
    "Acknowledgement SLA Deadline",
    "Acknowledged On",
    "In Progress Status Time",
    "Hold Status Time",
    "Resolved At Time",
    "Restored At Time",
    "Closed At",
    "Resolution SLA Deadline",
    "Approval Requested At Time",
    "Approval Accepted At Time",
    "CET Approval Requested At Time",
    "CET Approval Accepted At Time",
)

RESOLVED_TIME_RAW_COL: str = "Resolved  total time(Days:Hours)"
RESOLVED_TIME_DATE_ONLY_COL: str = "Resolved  total time(Days)"

TARGET_OMS_COLUMNS: tuple[str, ...] = (
    "Ticket ID",
    "Ticket Source",
    "Creator's Department",
    "Status",
    "Issue Category",
    "Issue Subcategory",
    "Department",
    "Asset",
    "Station",
    "Assigned To",
    "Created By",
    "Location",
    "Priority",
    "Ticket Type",
    "Asset ID",
    "Charger",
    "Updated At",
    "Created At",
    "Assigned At",
    "Acknowledgement SLA Deadline",
    "Acknowledged On",
    "Acknowledgement time taken(Days)",
    "Acknowledgement SLA Status",
    "In Progress Status Time",
    "In Progress total time(Days)",
    "Hold Status Time",
    "Hold total time(Days)",
    "Resolved At Time",
    "Resolved  total time(Days)",
    "Restored At Time",
    "Restore total time(Days)",
    "Closed At",
    "Total time taken (Days)",
    "Resolution SLA Deadline",
    "SLA Breached As per Actual (Days)",
    "SLA Breached As per OMS",
    "Approval Requested At Time",
    "Approval Accepted At Time",
    "Approval Time taken (Days)",
    "CET Approval Requested At Time",
    "CET Approval Accepted At Time",
    "CET Approval Time taken (Days)",
    "Resolution Note",
)

EXCLUDED_OMS_RAW_COLUMNS: tuple[str, ...] = (
    "Call Type",
    "Customer Phone",
    "Call SID",
    "Is Reverted",
    "Reverted At",
    "Issue Title",
)

def standardize_to_target_42(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    col_map: dict[str, str] = {}
    for c in out.columns:
        c_str = str(c).strip()
        c_clean = c_str.replace("’", "'").replace("  ", " ").lower()
        if c_clean in ("creator's department", "creator department", "creators department", "creator dept", "creator's dept"):
            col_map[c] = "Creator's Department"
        elif c_clean in ("resolution note", "remarks", "resolution remarks", "closing remarks", "closure remarks"):
            col_map[c] = "Resolution Note"
        elif c_clean in (
            "in progress total time(days)",
            "in progress total time (days)",
            "in progress time taken(days)",
            "in progress time taken (days)",
            "in progress days",
            "in progress total time",
            "in progress time taken",
        ):
            col_map[c] = "In Progress total time(Days)"
        elif c_clean in ("sla breached as per oms", "sla breached", "sla breach as per oms"):
            if "SLA Breached As per OMS" not in out.columns or c == "SLA Breached As per OMS":
                col_map[c] = "SLA Breached As per OMS"

    if col_map:
        out = out.rename(columns=col_map)

    if "SLA Breached" in out.columns:
        if "SLA Breached As per OMS" in out.columns:
            out["SLA Breached As per OMS"] = out["SLA Breached As per OMS"].fillna(out["SLA Breached"])
            out = out.drop(columns=["SLA Breached"])
        else:
            out = out.rename(columns={"SLA Breached": "SLA Breached As per OMS"})

    to_drop = [c for c in EXCLUDED_OMS_RAW_COLUMNS if c in out.columns]
    if to_drop:
        out = out.drop(columns=to_drop)

    for c in TARGET_OMS_COLUMNS:
        if c not in out.columns:
            out[c] = pd.NA

    res = out[list(TARGET_OMS_COLUMNS)].copy()
    if "today_date" in df.attrs:
        res.attrs["today_date"] = df.attrs["today_date"]
    return res


def is_full_raw_oms_dump(df_or_cols: Union[pd.DataFrame, list[str], set[str], tuple[str, ...]]) -> bool:
    if hasattr(df_or_cols, "columns"):
        cols = set(df_or_cols.columns)
    else:
        cols = set(df_or_cols)
    col_str_lower = {str(c).replace("’", "'").strip().lower() for c in cols}
    has_call_tracking = bool(col_str_lower.intersection({
        "call sid", "call type", "customer phone", "is reverted", "reverted at", "issue title"
    }))
    has_creators_dept = bool(col_str_lower.intersection({
        "creator's department", "creator department", "creators department", "creator dept", "creator's dept"
    }))
    return (has_call_tracking and has_creators_dept) or (len(cols) >= 35 and has_creators_dept) or ("call sid" in col_str_lower and "ticket id" in col_str_lower)

OMS_DELTA_SPECS: tuple[tuple[str, str, Optional[str]], ...] = (
    ("Acknowledgement time taken(Days)", "Assigned At", "Acknowledged On"),
    ("Hold total time(Days)", "Hold Status Time", None),
    ("Resolved  total time(Days)", "Assigned At", "Resolved At Time"),
    ("Restore total time(Days)", "Assigned At", "Restored At Time"),
    ("Total time taken (Days)", "Assigned At", "Closed At"),
    ("SLA Breached As per Actual (Days)", "Assigned At", "Resolution SLA Deadline"),
    ("Approval Time taken (Days)", "Approval Requested At Time", "Approval Accepted At Time"),
    ("CET Approval Time taken (Days)", "CET Approval Requested At Time", "CET Approval Accepted At Time"),
)

OMS_DELTA_COLUMNS: tuple[str, ...] = tuple(spec[0] for spec in OMS_DELTA_SPECS)
IN_PROGRESS_TIME_COL: str = "In Progress total time(Days)"

_SENTINEL_STRINGS = {"", "-", "--", "NAN", "NONE", "NAT", "NULL", "N/A", "NA"}

_ISO_FORMATS = (
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d %I:%M:%S %p",
    "%Y-%m-%d %I:%M %p",
    "%Y-%m-%d",
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
    "%Y/%m/%d %H:%M:%S.%f",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y/%m/%d",
)

_EXPLICIT_NON_ISO_FORMATS = (
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y %I:%M:%S %p",
    "%d-%m-%Y %I:%M %p",
    "%d-%m-%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %I:%M:%S %p",
    "%d/%m/%Y %I:%M %p",
    "%d/%m/%Y",
    "%d-%m-%y %H:%M:%S",
    "%d-%m-%y %H:%M",
    "%d-%m-%y %I:%M:%S %p",
    "%d-%m-%y %I:%M %p",
    "%d-%m-%y",
    "%d/%m/%y %H:%M:%S",
    "%d/%m/%y %H:%M",
    "%d/%m/%y %I:%M:%S %p",
    "%d/%m/%y %I:%M %p",
    "%d/%m/%y",
    "%d.%m.%Y %H:%M:%S",
    "%d.%m.%Y %H:%M",
    "%d.%m.%Y %I:%M:%S %p",
    "%d.%m.%Y %I:%M %p",
    "%d.%m.%Y",
    "%d.%m.%y %H:%M:%S",
    "%d.%m.%y %H:%M",
    "%d.%m.%y %I:%M:%S %p",
    "%d.%m.%y %I:%M %p",
    "%d.%m.%y",
    "%d %b %Y, %I:%M:%S %p",
    "%d %b %Y, %I:%M %p",
    "%d %b %Y, %H:%M:%S",
    "%d %b %Y, %H:%M",
    "%d %b %Y %I:%M:%S %p",
    "%d %b %Y %I:%M %p",
    "%d %b %Y %H:%M:%S",
    "%d %b %Y %H:%M",
    "%d %b %Y",
    "%d %B %Y, %I:%M:%S %p",
    "%d %B %Y, %I:%M %p",
    "%d %B %Y, %H:%M:%S",
    "%d %B %Y, %H:%M",
    "%d %B %Y %I:%M:%S %p",
    "%d %B %Y %H:%M:%S",
    "%d %B %Y",
    "%d-%b-%Y %H:%M:%S",
    "%d-%b-%Y %H:%M",
    "%d-%b-%Y",
    "%d-%b-%y %H:%M:%S",
    "%d-%b-%y %H:%M",
    "%d-%b-%y %I:%M:%S %p",
    "%d-%b-%y %I:%M %p",
    "%d-%b-%y",
    "%b %d, %Y %I:%M:%S %p",
    "%b %d, %Y %H:%M:%S",
    "%b %d, %Y",
)

_OMS_KNOWN_HEADERS = {
    "ticket id",
    RESOLVED_TIME_RAW_COL.lower(),
    RESOLVED_TIME_DATE_ONLY_COL.lower(),
    *(col.lower() for col in OMS_DATETIME_COLUMNS),
    *(col.lower() for col in OMS_DELTA_COLUMNS),
}


def _is_missing_scalar(value: Any) -> bool:
    if value is None or isinstance(value, bool):
        return value is None
    try:
        res = pd.isna(value)
        if isinstance(res, (bool, np.bool_)):
            return bool(res)
    except (TypeError, ValueError):
        pass
    return False


def _excel_serial_to_date(serial: float) -> Optional[date]:
    if not math.isfinite(serial) or not (20000.0 <= serial <= 80000.0):
        return None
    try:
        d = (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        return d if _is_valid_oms_year(d) else None
    except (OverflowError, ValueError):
        return None


_TZ_ABBREV_PATTERN = (
    r"(?:UTC|GMT|UT|IST|EST|EDT|CST|CDT|MST|MDT|PST|PDT|"
    r"AKST|AKDT|HST|AST|ADT|CET|CEST|EET|EEST|WET|WEST|BST|"
    r"JST|KST|SGT|HKT|MYT|PHT|WIB|WITA|WIT|ICT|PKT|NPT|BDT|"
    r"GST|IRST|IRDT|AEST|AEDT|ACST|ACDT|AWST|NZST|NZDT|"
    r"BRT|BRST|ART|CLT|CLST|CAT|EAT|WAT|SAST)"
)
_TZ_OFFSET_PATTERN = r"(?:[+-](?:0?\d|1[0-4]):?[0-5]\d)"
_TIME_TRAILING_TZ_RE = re.compile(
    rf"(?:"
    rf"\s*\(\s*{_TZ_ABBREV_PATTERN}(?:{_TZ_OFFSET_PATTERN})?\s*\)"
    rf"|(?<=[\s\d]){_TZ_ABBREV_PATTERN}(?:{_TZ_OFFSET_PATTERN})?"
    rf"|\s*{_TZ_OFFSET_PATTERN}"
    rf"|(?<=\d)Z"
    rf")\s*$",
    re.IGNORECASE,
)
_DATE_TRAILING_TZ_RE = re.compile(
    rf"\s+(?:\(\s*)?{_TZ_ABBREV_PATTERN}(?:{_TZ_OFFSET_PATTERN})?(?:\s*\))?\s*$",
    re.IGNORECASE,
)


def _strip_trailing_timezone(s: str) -> str:
    if re.search(r"(?<!\d)\d{1,2}:\d{2}(?!\d)", s):
        cleaned = _TIME_TRAILING_TZ_RE.sub("", s).strip()
    else:
        cleaned = _DATE_TRAILING_TZ_RE.sub("", s).strip()
    return re.sub(r"\s+", " ", cleaned)


def _is_valid_oms_year(d: date) -> bool:
    return 1970 <= d.year <= 2100


def parse_date_only(value: Any) -> Optional[date]:
    if value is None or isinstance(value, (bool, np.bool_)):
        return None

    if _is_missing_scalar(value):
        return None

    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        d = value.date()
        return d if _is_valid_oms_year(d) else None

    if isinstance(value, datetime):
        d = value.date()
        return d if _is_valid_oms_year(d) else None

    if isinstance(value, date):
        return value if _is_valid_oms_year(value) else None

    if isinstance(value, (int, float, np.integer, np.floating)):
        return _excel_serial_to_date(float(value))

    s = str(value).strip()
    if not s or s.upper() in _SENTINEL_STRINGS:
        return None

    s_lower = s.lower()
    if (
        s_lower.startswith("=")
        or "colomn" in s_lower
        or "column" in s_lower
        or s_lower.startswith("today")
        or not re.search(r"\d", s)
    ):
        return None

    if re.fullmatch(r"\d{5}(?:\.\d+)?", s):
        return _excel_serial_to_date(float(s))

    if re.fullmatch(r"\d{8}", s):
        try:
            d = datetime.strptime(s, "%Y%m%d").date()
            return d if _is_valid_oms_year(d) else None
        except ValueError:
            return None

    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", s):
        return None

    s_clean = _strip_trailing_timezone(s)
    if not s_clean or re.fullmatch(
        r"\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*(?:AM|PM))?",
        s_clean,
        re.IGNORECASE,
    ):
        return None

    # ISO-first (YYYY-MM-DD...) so YYYY-MM-DD is never swapped to YYYY-DD-MM
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}(?:[T\s].*)?$", s_clean):
        for candidate in (s, s_clean):
            try:
                d = datetime.fromisoformat(candidate).date()
                return d if _is_valid_oms_year(d) else None
            except ValueError:
                pass
        for fmt in _ISO_FORMATS:
            try:
                d = datetime.strptime(s_clean, fmt).date()
                return d if _is_valid_oms_year(d) else None
            except ValueError:
                continue

    if re.match(r"^\d{4}/\d{1,2}/\d{1,2}(?:\s.*)?$", s_clean):
        for fmt in (
            "%Y/%m/%d %H:%M:%S.%f",
            "%Y/%m/%d %H:%M:%S",
            "%Y/%m/%d %H:%M",
            "%Y/%m/%d %I:%M:%S %p",
            "%Y/%m/%d %I:%M %p",
            "%Y/%m/%d",
        ):
            try:
                d = datetime.strptime(s_clean, fmt).date()
                return d if _is_valid_oms_year(d) else None
            except ValueError:
                continue

    if re.match(r"^\d{4}\.\d{1,2}\.\d{1,2}(?:\s.*)?$", s_clean):
        for fmt in (
            "%Y.%m.%d %H:%M:%S.%f",
            "%Y.%m.%d %H:%M:%S",
            "%Y.%m.%d %H:%M",
            "%Y.%m.%d %I:%M:%S %p",
            "%Y.%m.%d %I:%M %p",
            "%Y.%m.%d",
        ):
            try:
                d = datetime.strptime(s_clean, fmt).date()
                return d if _is_valid_oms_year(d) else None
            except ValueError:
                continue

    for fmt in _EXPLICIT_NON_ISO_FORMATS:
        try:
            d = datetime.strptime(s_clean, fmt).date()
            return d if _is_valid_oms_year(d) else None
        except ValueError:
            continue

    # Structural pre-validation guard before pd.to_datetime fallback
    if re.match(r"^\d{1,4}[-/.]\d{1,4}(?:$|\s)", s_clean):
        return None
    if not re.search(r"(?<!\d)(?:19[7-9]\d|20\d\d|2100)(?!\d)", s_clean):
        return None
    has_3part_num = bool(re.search(r"(?<!\d)\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?!\d)", s_clean))
    no_time = re.sub(r"\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?", "", s_clean)
    has_month_day = bool(
        re.search(r"[A-Za-z]{3,}", no_time)
        and re.search(r"(?<!\d)\d{1,2}(?!\d)", no_time)
    )
    if not (has_3part_num or has_month_day):
        return None

    dayfirst = not bool(re.match(r"^\d{4}[-/.]", s_clean))
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        ts = pd.to_datetime(s_clean, dayfirst=dayfirst, errors="coerce")
    if not pd.isna(ts):
        d = ts.date()
        if _is_valid_oms_year(d):
            return d

    return None


def strip_to_date_str(value: Any) -> str:
    parsed = parse_date_only(value)
    return parsed.strftime("%Y-%m-%d") if parsed is not None else ""


def _normalize_today_date(today_date: Optional[date]) -> date:
    if today_date is None:
        return date.today()
    if isinstance(today_date, datetime):
        return today_date.date()
    return today_date


def convert_oms_to_date_only(
    df: pd.DataFrame,
    today_date: Optional[date] = None,
    standardize_schema: Optional[bool] = None,
) -> pd.DataFrame:
    ref_today = _normalize_today_date(today_date)
    out = df.copy()

    should_standardize = (
        standardize_schema
        if standardize_schema is not None
        else is_full_raw_oms_dump(out)
    )

    if RESOLVED_TIME_RAW_COL in out.columns:
        if RESOLVED_TIME_DATE_ONLY_COL not in out.columns:
            out = out.rename(columns={RESOLVED_TIME_RAW_COL: RESOLVED_TIME_DATE_ONLY_COL})

    if should_standardize:
        if "SLA Breached" in out.columns:
            if "SLA Breached As per OMS" not in out.columns:
                out = out.rename(columns={"SLA Breached": "SLA Breached As per OMS"})
            else:
                out["SLA Breached As per OMS"] = out["SLA Breached As per OMS"].fillna(out["SLA Breached"])
                out = out.drop(columns=["SLA Breached"])

        if "Remarks" in out.columns and "Resolution Note" not in out.columns:
            out = out.rename(columns={"Remarks": "Resolution Note"})

    parsed_dates: dict[str, list[Optional[date]]] = {}
    for col in OMS_DATETIME_COLUMNS:
        if col in out.columns:
            dates_list = [parse_date_only(val) for val in out[col].tolist()]
            parsed_dates[col] = dates_list
            out[col] = pd.Series(
                [d.strftime("%Y-%m-%d") if d is not None else "" for d in dates_list],
                index=out.index,
                dtype=object,
            )

    n_rows = len(out)
    for target_col, start_col, end_col in OMS_DELTA_SPECS:
        has_start = start_col in out.columns
        has_end = (end_col is None) or (end_col in out.columns)
        has_sources = has_start and has_end

        if not has_sources and target_col not in out.columns:
            if should_standardize and target_col in TARGET_OMS_COLUMNS:
                out[target_col] = pd.Series([pd.NA] * n_rows, index=out.index, dtype="Int64")
            continue

        if not has_sources:
            out[target_col] = pd.Series([pd.NA] * n_rows, index=out.index, dtype="Int64")
            continue

        start_list = parsed_dates.get(start_col)
        if start_list is None:
            start_list = [parse_date_only(val) for val in out[start_col].tolist()]
            parsed_dates[start_col] = start_list

        if end_col is None:
            end_list: list[Optional[date]] = [ref_today] * n_rows
        else:
            end_list_cached = parsed_dates.get(end_col)
            if end_list_cached is None:
                end_list_cached = [parse_date_only(val) for val in out[end_col].tolist()]
                parsed_dates[end_col] = end_list_cached
            end_list = end_list_cached

        delta_vals = [
            (e_d - s_d).days if (s_d is not None and e_d is not None) else pd.NA
            for s_d, e_d in zip(start_list, end_list)
        ]
        out[target_col] = pd.Series(delta_vals, index=out.index, dtype="Int64")

    if should_standardize or IN_PROGRESS_TIME_COL in out.columns:
        has_prog = "In Progress Status Time" in out.columns
        has_ack = "Acknowledged On" in out.columns
        if has_prog and has_ack:
            prog_list = parsed_dates.get("In Progress Status Time")
            if prog_list is None:
                prog_list = [parse_date_only(val) for val in out["In Progress Status Time"].tolist()]
                parsed_dates["In Progress Status Time"] = prog_list
            ack_list = parsed_dates.get("Acknowledged On")
            if ack_list is None:
                ack_list = [parse_date_only(val) for val in out["Acknowledged On"].tolist()]
                parsed_dates["Acknowledged On"] = ack_list
            prog_delta_vals = [
                (p - a).days if (p is not None and a is not None) else pd.NA
                for a, p in zip(ack_list, prog_list)
            ]
            out[IN_PROGRESS_TIME_COL] = pd.Series(prog_delta_vals, index=out.index, dtype="Int64")
        elif should_standardize:
            out[IN_PROGRESS_TIME_COL] = pd.Series([pd.NA] * n_rows, index=out.index, dtype="Int64")

    if should_standardize:
        out = standardize_to_target_42(out)

    out.attrs["today_date"] = ref_today.strftime("%Y-%m-%d")
    return out


def recalculate_hold_days(
    df: pd.DataFrame,
    today_date: Optional[date] = None,
) -> pd.DataFrame:
    ref_today = _normalize_today_date(today_date)
    out = df.copy()

    hold_col = "Hold total time(Days)"
    if "Hold Status Time" in out.columns:
        hold_dates = [parse_date_only(val) for val in out["Hold Status Time"].tolist()]
        hold_vals = [
            (ref_today - d).days if d is not None else pd.NA
            for d in hold_dates
        ]
        out[hold_col] = pd.Series(hold_vals, index=out.index, dtype="Int64")
    elif hold_col in out.columns:
        out[hold_col] = pd.Series([pd.NA] * len(out), index=out.index, dtype="Int64")

    out.attrs["today_date"] = ref_today.strftime("%Y-%m-%d")
    return out


def _peek_top_rows(
    source: Union[str, bytes, BinaryIO, Path],
    filename: str = "",
    max_rows: int = 15,
) -> list[list[str]]:
    name = (filename or (str(source) if isinstance(source, (str, Path)) else "")).lower()
    buffer, is_stream = _get_bytes_or_buffer(source)
    try:
        if name.endswith(".csv"):
            df_head = _read_csv_with_fallback(buffer, is_stream, header=None, nrows=max_rows)
        elif name.endswith(".xls"):
            raw_bytes = _read_raw_bytes(buffer if is_stream else source)
            df_head = _read_xls_dataframe(raw_bytes, header=None, nrows=max_rows)
        else:
            if is_stream:
                _safe_seek_zero(buffer)
            df_head = pd.read_excel(buffer, header=None, nrows=max_rows)
        if df_head.empty:
            return []
        return [
            [str(v).strip() for v in df_head.iloc[idx].dropna().tolist() if str(v).strip() != ""]
            for idx in range(len(df_head))
        ]
    except Exception:
        return []
    finally:
        if is_stream:
            _safe_seek_zero(buffer)


def _peek_first_row(
    source: Union[str, bytes, BinaryIO, Path],
    filename: str = "",
) -> list[str]:
    rows = _peek_top_rows(source, filename=filename, max_rows=1)
    return rows[0] if rows else []


def load_oms_raw_file(
    source: Union[str, bytes, BinaryIO, Path],
    filename: str = "",
) -> pd.DataFrame:
    source = _ensure_seekable_source(source)
    top_rows = _peek_top_rows(source, filename=filename, max_rows=15)
    row1 = top_rows[0] if top_rows else []
    row1_oms_matches = {cell.lower() for cell in row1 if cell.lower() in _OMS_KNOWN_HEADERS}

    if len(row1_oms_matches) >= 2:
        header_row: Optional[int] = 1
    elif len(row1_oms_matches) == 1:
        candidate_row = detect_header_row(source, filename=filename)
        header_row = 1
        if candidate_row > 1 and (candidate_row - 1) < len(top_rows):
            cand_cells = top_rows[candidate_row - 1]
            cand_oms_matches = {
                cell.lower() for cell in cand_cells if cell.lower() in _OMS_KNOWN_HEADERS
            }
            cand_exact_kw = {
                cell.lower() for cell in cand_cells if cell.lower() in HEADER_KEYWORDS
            }
            row1_has_date = any(parse_date_only(cell) is not None for cell in row1)
            row1_is_banner = (len(row1) <= 2 and row1_has_date) or (
                len(row1) == 1 and len(cand_cells) > 1
            )
            if len(cand_oms_matches) >= 2 or (
                row1_is_banner
                and (len(cand_oms_matches) >= 1 or len(cand_exact_kw) >= 2)
                and len(cand_cells) > len(row1)
            ):
                header_row = candidate_row
    else:
        header_row = None

    return load_dataset(source, filename=filename, header_row=header_row)
