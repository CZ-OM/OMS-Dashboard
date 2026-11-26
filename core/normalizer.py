import re
from datetime import datetime, timedelta
from typing import Any, Union, Optional
import pandas as pd


def normalize_ocpp_id(
    series_or_val: Union[pd.Series, str, Any],
    strict: bool = False
) -> Union[pd.Series, str]:
    if isinstance(series_or_val, pd.Series):
        s = series_or_val.fillna("").astype(str).str.strip().str.upper()
        s = s.replace({"NAN": "", "NONE": "", "NAT": "", "NULL": "", "-": ""})
        if not strict:
            s = s.str.replace(r"[\s\-_]+", "", regex=True)
        return s

    if series_or_val is None or pd.isna(series_or_val):
        return ""
    val_str = str(series_or_val).strip().upper()
    if val_str in {"NAN", "NONE", "NAT", "NULL", "-"}:
        return ""
    if not strict:
        val_str = re.sub(r"[\s\-_]+", "", val_str)
    return val_str


def parse_issue_age(age_val: Any) -> float:
    if age_val is None or pd.isna(age_val):
        return 0.0

    if isinstance(age_val, (int, float)):
        return float(age_val)

    s = str(age_val).strip()
    if not s or s.upper() in {"NAN", "NONE", "NULL", "-"}:
        return 0.0

    match_dh = re.search(r"(\d+)\s*Days?\s*(\d+)\s*Hours?", s, re.IGNORECASE)
    if match_dh:
        days = float(match_dh.group(1))
        hours = float(match_dh.group(2))
        return round(days + (hours / 24.0), 3)

    match_d = re.search(r"(\d+(?:\.\d+)?)\s*Days?", s, re.IGNORECASE)
    if match_d:
        return round(float(match_d.group(1)), 3)

    match_h = re.search(r"(\d+(?:\.\d+)?)\s*Hours?", s, re.IGNORECASE)
    if match_h:
        return round(float(match_h.group(1)) / 24.0, 3)

    try:
        return round(float(s), 3)
    except (ValueError, TypeError):
        return 0.0


def excel_serial_to_datetime(serial: float) -> datetime:
    base_date = datetime(1899, 12, 30)
    return base_date + timedelta(days=serial)


def parse_datetime(val: Any) -> Optional[datetime]:
    if val is None or pd.isna(val):
        return None

    if isinstance(val, (datetime, pd.Timestamp)):
        return val.to_pydatetime() if isinstance(val, pd.Timestamp) else val

    if isinstance(val, (int, float)):
        if 20000 <= val <= 80000:
            try:
                return excel_serial_to_datetime(float(val))
            except Exception:
                return None

    val_str = str(val).strip()
    if not val_str or val_str.upper() in {"NAN", "NONE", "NAT", "NULL", "-"}:
        return None

    try:
        dt = pd.to_datetime(val_str)
        return dt.to_pydatetime()
    except Exception:
        return None


def format_date_str(val: Any) -> str:
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        if val.hour or val.minute or val.second:
            return val.strftime("%Y-%m-%d %H:%M:%S")
        return val.strftime("%Y-%m-%d")

    s = str(val).strip()
    if not s or s.upper() in {"NAT", "NAN", "NONE", "NULL", "-"}:
        return ""

    dt = parse_datetime(s)
    if dt:
        if dt.hour or dt.minute or dt.second:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return dt.strftime("%Y-%m-%d")
    return s
