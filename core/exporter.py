import io
import math
import re
from datetime import date, datetime
from typing import Dict, Any
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from .models import ReconciliationResult
from .oms_converter import OMS_DELTA_COLUMNS

_ILLEGAL_EXCEL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _sanitize_excel_str(val: str) -> str:
    return _ILLEGAL_EXCEL_CHARS_RE.sub("", val)


def _assign_excel_cell(sheet, row: int, column: int, value: Any):
    if isinstance(value, str):
        clean_val = _sanitize_excel_str(value)
        cell = sheet.cell(row=row, column=column, value=clean_val)
        cell.data_type = "s"
        return cell
    return sheet.cell(row=row, column=column, value=value)


def generate_reconciliation_excel(result: ReconciliationResult) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    accent_fill = PatternFill(start_color="C8102E", end_color="C8102E", fill_type="solid")
    accent_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws_summary = wb.active
    ws_summary.title = "Executive Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:D1")
    title_cell = ws_summary["A1"]
    title_cell.value = "EV CHARGING INFRASTRUCTURE RECONCILIATION AUDIT"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_summary.row_dimensions[1].height = 28

    ws_summary["A3"] = "Generated At:"
    _assign_excel_cell(ws_summary, 3, 2, str(result.summary.timestamp))
    ws_summary["A4"] = "Execution Time:"
    _assign_excel_cell(ws_summary, 4, 2, f"{result.summary.execution_time_ms} ms")
    for r in [3, 4]:
        ws_summary[f"A{r}"].font = Font(bold=True)

    headers_summary = ["Metric Category", "Operational Value", "Benchmark / Status", "Prescribed Directive"]
    for col_idx, h in enumerate(headers_summary, start=1):
        cell = ws_summary.cell(row=6, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_summary.row_dimensions[6].height = 22

    summary_rows = [
        ("Total FIR Chargers Audited", result.summary.total_fir, "Physical Field Roster", "Baseline operational footprint"),
        ("Chargers in Maintenance", result.summary.in_maintenance, "Faulted / Unavailable", "Active outage roster needing coverage"),
        ("Healthy / Available Chargers", result.summary.available_cleared, "Cleared", "Field operating normally"),
        ("Ticket Generation Gaps", result.summary.ticket_gaps, "CRITICAL DISCREPANCY", "Immediately generate OMS tickets and dispatch ZME"),
        ("Phantom Resolutions", result.summary.phantom_resolutions, "PREMATURE CLOSURE", "Reopen OMS tickets; hardware remains down in field"),
        ("Active Verified Outages", result.summary.active_verified, "TRACKED IN OMS", "Monitor technician resolution progress vs SLA"),
        ("Coverage Gap Rate", f"{result.summary.coverage_gap_rate}%", "Threshold: < 5.0%", "Operational discrepancy index = (Gaps + Phantoms) / Maintenance"),
        ("OMS Untracked Open Tickets", result.summary.untracked_oms_open, "BLIND SPOT RISK", "Active tickets for chargers outside morning FIR snapshot"),
    ]

    for row_idx, data in enumerate(summary_rows, start=7):
        for col_idx, val in enumerate(data, start=1):
            cell = _assign_excel_cell(ws_summary, row_idx, col_idx, val)
            cell.border = thin_border
            if col_idx == 2:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 3 and "CRITICAL" in str(val):
                cell.font = Font(bold=True, color="C8102E")
            elif col_idx == 3 and "PREMATURE" in str(val):
                cell.font = Font(bold=True, color="D9534F")
        ws_summary.row_dimensions[row_idx].height = 20

    def _append_df_to_sheet(sheet, df: pd.DataFrame, sheet_title: str):
        sheet.title = sheet_title
        sheet.views.sheetView[0].showGridLines = True

        clean_df = df.drop(columns=[c for c in ["_norm_ocpp", "_sort_created", "Tentative_ETA", "Tentative ETA"] if c in df.columns])

        for col_idx, col_name in enumerate(clean_df.columns, start=1):
            cell = _assign_excel_cell(sheet, 1, col_idx, str(col_name))
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 24

        for row_idx, row_values in enumerate(clean_df.itertuples(index=False, name=None), start=2):
            for col_idx, value in enumerate(row_values, start=1):
                if isinstance(value, (pd.Timestamp, datetime)):
                    value = "" if pd.isna(value) else value.strftime("%Y-%m-%d %H:%M:%S")
                elif isinstance(value, date):
                    value = value.strftime("%Y-%m-%d")
                else:
                    try:
                        if pd.isna(value):
                            value = ""
                    except (TypeError, ValueError):
                        pass
                cell = _assign_excel_cell(sheet, row_idx, col_idx, value)
                cell.border = thin_border
            sheet.row_dimensions[row_idx].height = 19

        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    ws_gaps = wb.create_sheet()
    _append_df_to_sheet(ws_gaps, result.gaps_df, "Ticket Generation Gaps")

    ws_phantoms = wb.create_sheet()
    _append_df_to_sheet(ws_phantoms, result.phantoms_df, "Phantom Resolutions")

    ws_active = wb.create_sheet()
    _append_df_to_sheet(ws_active, result.active_df, "Active Verified")

    ws_master = wb.create_sheet()
    _append_df_to_sheet(ws_master, result.reconciled_df, "Master Reconciled")

    ws_untracked = wb.create_sheet()
    _append_df_to_sheet(ws_untracked, result.untracked_oms_df, "OMS Untracked Open")

    for col in ws_summary.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 55)

    wb.save(output)
    return output.getvalue()


def export_reconciliation_csv(df: pd.DataFrame) -> str:
    clean_df = df.drop(columns=[c for c in ["_norm_ocpp", "_sort_created", "Tentative_ETA", "Tentative ETA"] if c in df.columns])
    return clean_df.to_csv(index=False)


def _coerce_delta_series_for_csv(series: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(series.dtype):
        return series.astype(object)
    if pd.api.types.is_integer_dtype(series.dtype):
        return series.astype("Int64")

    num = pd.to_numeric(series, errors="coerce")
    if pd.api.types.is_bool_dtype(num.dtype):
        return num.astype(object)
    if pd.api.types.is_integer_dtype(num.dtype):
        return num.astype("Int64")

    arr = num.to_numpy(dtype=float, na_value=np.nan)
    finite_mask = np.isfinite(arr)
    finite_vals = arr[finite_mask]
    if (
        np.all(np.floor(finite_vals) == finite_vals)
        and np.all((finite_vals >= -9223372036854775808.0) & (finite_vals < 9223372036854775808.0))
    ):
        cleaned = pd.Series(
            np.where(finite_mask, arr, np.nan),
            index=series.index,
            dtype="float64",
        )
        return cleaned.astype("Int64")

    out_vals = []
    for val, is_fin in zip(arr, finite_mask):
        if not is_fin:
            out_vals.append("")
        elif float(val).is_integer():
            out_vals.append(int(val))
        else:
            out_vals.append(float(val))
    return pd.Series(out_vals, index=series.index, dtype=object)


def export_oms_date_only_csv(df: pd.DataFrame) -> str:
    export_df = df.copy()
    delta_set = set(OMS_DELTA_COLUMNS)

    for col in export_df.columns:
        series = export_df[col]
        if col in delta_set or pd.api.types.is_integer_dtype(series.dtype):
            export_df[col] = _coerce_delta_series_for_csv(series)
        elif pd.api.types.is_datetime64_any_dtype(series.dtype):
            export_df[col] = series.dt.strftime("%Y-%m-%d").fillna("")
        else:
            obj_series = series.astype(object)
            export_df[col] = obj_series.where(~pd.isna(obj_series), "")

    return export_df.to_csv(index=False)


def export_oms_date_only_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws = wb.active
    ws.title = "OMS Date Only"
    ws.views.sheetView[0].showGridLines = True

    columns = list(df.columns)
    delta_set = set(OMS_DELTA_COLUMNS)

    for col_idx, col_name in enumerate(columns, start=1):
        cell = _assign_excel_cell(ws, 1, col_idx, str(col_name))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for row_idx, row_values in enumerate(df.itertuples(index=False, name=None), start=2):
        for col_idx, (col_name, value) in enumerate(zip(columns, row_values), start=1):
            is_na_val = False
            if value is None:
                is_na_val = True
            else:
                try:
                    na_res = pd.isna(value)
                    if isinstance(na_res, (bool, np.bool_)) and bool(na_res):
                        is_na_val = True
                except (TypeError, ValueError):
                    pass

            if is_na_val or (isinstance(value, str) and value == ""):
                cell_val: Any = None
            elif isinstance(value, (bool, np.bool_)):
                cell_val = bool(value)
            elif isinstance(value, (float, np.floating)) and not math.isfinite(float(value)):
                cell_val = None
            elif col_name in delta_set:
                if isinstance(value, (int, np.integer)):
                    cell_val = int(value)
                elif isinstance(value, str) and re.fullmatch(r"[+-]?\d+", value.strip()):
                    cell_val = int(value.strip())
                else:
                    try:
                        fval = float(value)
                    except (TypeError, ValueError):
                        cell_val = str(value)
                    else:
                        if not math.isfinite(fval):
                            cell_val = None
                        elif fval.is_integer():
                            cell_val = int(fval)
                        else:
                            cell_val = fval
            elif isinstance(value, (int, np.integer)):
                cell_val = int(value)
            elif isinstance(value, np.floating):
                cell_val = float(value)
            elif isinstance(value, (pd.Timestamp, datetime, date)):
                cell_val = value.strftime("%Y-%m-%d")
            else:
                cell_val = value

            cell = _assign_excel_cell(ws, row_idx, col_idx, cell_val)
            cell.border = thin_border
            if isinstance(cell_val, int) and not isinstance(cell_val, bool):
                cell.number_format = "0"
        ws.row_dimensions[row_idx].height = 19

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 42)

    wb.save(output)
    return output.getvalue()
