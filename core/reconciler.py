import time
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd

from .models import (
    OPEN_OMS_STATUSES,
    CLOSED_OMS_STATUSES,
    DOWN_FIR_STATUSES,
    CATEGORY_GAP,
    CATEGORY_PHANTOM,
    CATEGORY_ACTIVE,
    CATEGORY_HEALTHY,
    CATEGORY_UNTRACKED,
    RISK_CRITICAL,
    RISK_HIGH,
    RISK_MEDIUM,
    RISK_LOW,
    ReconciliationSummary,
    ReconciliationResult,
)
from .normalizer import normalize_ocpp_id, parse_issue_age, parse_datetime, format_date_str


def _find_column(df: pd.DataFrame, patterns: List[str], default: Optional[str] = None) -> Optional[str]:
    col_map = {str(c).strip().lower(): c for c in df.columns}

    for pat in patterns:
        p_clean = pat.strip().lower()
        if p_clean in col_map:
            return col_map[p_clean]

    for pat in patterns:
        p_clean = pat.strip().lower()
        for c_lower, original_col in col_map.items():
            if p_clean in c_lower:
                return original_col

    return default


def reconcile_incident_records(fir_df: pd.DataFrame, oms_df: pd.DataFrame) -> ReconciliationResult:
    start_time = time.perf_counter()

    fir = fir_df.copy()
    oms = oms_df.copy()

    default_report_date = fir.attrs.get("report_date", datetime.now().strftime("%Y-%m-%d"))

    fir_ocpp_col = _find_column(fir, ["ocpp id", "ocpp", "charger id", "charger_id", "charger"], default="OCPP ID")
    fir_status_col = _find_column(fir, ["status", "fir_status", "state"], default="Status")
    fir_station_col = _find_column(fir, ["station name", "station", "site"], default="Station Name")
    fir_age_col = _find_column(fir, ["issue_age", "issue age", "age", "duration"], default="Issue_Age")
    fir_zone_col = _find_column(fir, ["zone", "region"], default="Zone")
    fir_make_col = _find_column(fir, ["charger make", "make", "oem", "brand"], default="Charger Make")
    fir_tenant_col = _find_column(fir, ["tenant name", "tenant", "client"], default="Tenant Name")
    fir_error_col = _find_column(fir, ["error info", "error", "fault", "alarm"], default="Error Info")
    fir_priority_col = _find_column(fir, ["priority_label", "priority", "severity"], default="Priority_Label")
    fir_zme_col = _find_column(fir, ["zme name", "zme", "manager"], default="ZME Name")
    fir_contact_col = _find_column(fir, ["zme contact number", "contact", "phone", "mobile"], default="ZME Contact Number")
    fir_eta_col = _find_column(fir, ["tentative eta", "eta", "expected completion", "target date"], default="Tentative ETA")
    fir_date_col = _find_column(fir, ["report date", "date", "created at", "snapshot date"], default="Report Date")

    oms_charger_col = _find_column(oms, ["charger", "ocpp", "charger_id", "device"], default="Charger")
    oms_ticket_col = _find_column(oms, ["ticket id", "ticket_id", "ticket", "incident"], default="Ticket ID")
    oms_status_col = _find_column(oms, ["status", "ticket_status", "state"], default="Status")
    oms_category_col = _find_column(oms, ["issue category", "category"], default="Issue Category")
    oms_subcat_col = _find_column(oms, ["issue subcategory", "subcategory"], default="Issue Subcategory")
    oms_assigned_col = _find_column(oms, ["assigned to", "assigned", "technician", "owner"], default="Assigned To")
    oms_created_col = _find_column(oms, ["created at", "created_at", "created date", "created"], default="Created At")
    oms_updated_col = _find_column(oms, ["updated at", "updated_at", "updated"], default="Updated At")
    oms_closed_col = _find_column(oms, ["closed at", "closed_at", "closed date", "resolved at", "resolved_at"], default="Closed At")
    oms_sla_deadline_col = _find_column(oms, ["resolution sla deadline", "resolution deadline", "sla deadline"], default="Resolution SLA Deadline")
    oms_ack_deadline_col = _find_column(oms, ["acknowledgement sla deadline", "ack deadline"], default="Acknowledgement SLA Deadline")
    oms_priority_col = _find_column(oms, ["priority", "severity"], default="Priority")
    oms_sla_col = _find_column(oms, ["sla breached", "sla_breached", "sla"], default="SLA Breached")
    oms_station_col = _find_column(oms, ["station", "site"], default="Station")
    oms_zone_col = _find_column(oms, ["zone", "region"], default="Zone")

    fir["_norm_ocpp"] = normalize_ocpp_id(fir[fir_ocpp_col] if fir_ocpp_col in fir.columns else pd.Series([""] * len(fir)))
    oms["_norm_ocpp"] = normalize_ocpp_id(oms[oms_charger_col] if oms_charger_col in oms.columns else pd.Series([""] * len(oms)))

    sort_cols = []
    temp_sort_cols = []
    if oms_created_col in oms.columns:
        oms["_sort_created"] = pd.to_datetime(oms[oms_created_col], errors="coerce", utc=True)
        sort_cols.append("_sort_created")
    if "Sno" in oms.columns:
        sno_series = oms["Sno"]
        if pd.api.types.is_numeric_dtype(sno_series.dtype) and not pd.api.types.is_bool_dtype(sno_series.dtype):
            sort_cols.append("Sno")
        else:
            oms["_sort_sno_num"] = pd.to_numeric(sno_series, errors="coerce")
            oms["_sort_sno_str"] = sno_series.fillna("").astype(str).str.strip()
            sort_cols.extend(["_sort_sno_num", "_sort_sno_str"])
            temp_sort_cols.extend(["_sort_sno_num", "_sort_sno_str"])

    if sort_cols:
        oms_sorted = oms.sort_values(by=sort_cols, ascending=True, na_position="last")
        if temp_sort_cols:
            oms_sorted = oms_sorted.drop(columns=temp_sort_cols)
            oms = oms.drop(columns=temp_sort_cols)
    else:
        oms_sorted = oms

    oms_valid = oms_sorted[oms_sorted["_norm_ocpp"] != ""]
    oms_lookup: Dict[str, List[Dict[str, Any]]] = {}
    for norm_id, group in oms_valid.groupby("_norm_ocpp", sort=False):
        oms_lookup[str(norm_id)] = group.to_dict(orient="records")

    audit_categories: List[str] = []
    risk_levels: List[str] = []
    prescribed_actions: List[str] = []

    latest_ticket_ids: List[str] = []
    latest_ticket_statuses: List[str] = []
    latest_ticket_categories: List[str] = []
    latest_ticket_subcats: List[str] = []
    latest_ticket_assigned: List[str] = []
    latest_ticket_created: List[str] = []
    latest_ticket_updated: List[str] = []
    latest_ticket_closed: List[str] = []
    latest_ticket_sla_deadline: List[str] = []
    latest_ticket_ack_deadline: List[str] = []
    latest_ticket_sla: List[str] = []

    report_dates: List[str] = []
    tentative_etas: List[str] = []

    total_tickets_count: List[int] = []
    open_tickets_count: List[int] = []
    parsed_age_days: List[float] = []

    for _, row in fir.iterrows():
        norm_id = str(row["_norm_ocpp"])
        fir_st = str(row.get(fir_status_col, "")).strip().upper()
        raw_age = row.get(fir_age_col, 0.0)
        age_in_days = parse_issue_age(raw_age)
        parsed_age_days.append(age_in_days)

        row_date = str(row.get(fir_date_col, default_report_date)).strip()
        report_dates.append(row_date if row_date else default_report_date)

        raw_eta = row.get(fir_eta_col, "")
        tentative_etas.append(format_date_str(raw_eta))

        is_down = fir_st in DOWN_FIR_STATUSES
        charger_tickets = oms_lookup.get(norm_id, [])
        total_tickets_count.append(len(charger_tickets))

        open_tickets = [
            t for t in charger_tickets
            if str(t.get(oms_status_col, "")).strip().upper() in OPEN_OMS_STATUSES
        ]
        closed_tickets = [
            t for t in charger_tickets
            if str(t.get(oms_status_col, "")).strip().upper() not in OPEN_OMS_STATUSES
        ]
        open_tickets_count.append(len(open_tickets))

        latest_t: Dict[str, Any] = {}
        if open_tickets:
            latest_t = open_tickets[-1]
        elif closed_tickets:
            latest_t = closed_tickets[-1]

        t_id = str(latest_t.get(oms_ticket_col, "")) if latest_t else ""
        t_st = str(latest_t.get(oms_status_col, "")) if latest_t else ""
        t_cat = str(latest_t.get(oms_category_col, "")) if latest_t else ""
        t_subcat = str(latest_t.get(oms_subcat_col, "")) if latest_t else ""
        t_assigned = str(latest_t.get(oms_assigned_col, "")) if latest_t else ""
        t_created = format_date_str(latest_t.get(oms_created_col, "")) if latest_t else ""
        t_updated = format_date_str(latest_t.get(oms_updated_col, "")) if latest_t else ""
        t_closed = format_date_str(latest_t.get(oms_closed_col, "")) if latest_t else ""
        t_sla_deadline = format_date_str(latest_t.get(oms_sla_deadline_col, "")) if latest_t else ""
        t_ack_deadline = format_date_str(latest_t.get(oms_ack_deadline_col, "")) if latest_t else ""
        t_sla = str(latest_t.get(oms_sla_col, "")).strip() if latest_t else ""

        latest_ticket_ids.append(t_id)
        latest_ticket_statuses.append(t_st)
        latest_ticket_categories.append(t_cat)
        latest_ticket_subcats.append(t_subcat)
        latest_ticket_assigned.append(t_assigned)
        latest_ticket_created.append(t_created)
        latest_ticket_updated.append(t_updated)
        latest_ticket_closed.append(t_closed)
        latest_ticket_sla_deadline.append(t_sla_deadline)
        latest_ticket_ack_deadline.append(t_ack_deadline)
        latest_ticket_sla.append(t_sla)

        if not is_down:
            audit_categories.append(CATEGORY_HEALTHY)
            risk_levels.append(RISK_LOW)
            prescribed_actions.append("Operational: Charger is operating normally in the field.")
        else:
            if len(charger_tickets) == 0:
                audit_categories.append(CATEGORY_GAP)
                risk_levels.append(RISK_CRITICAL)
                prescribed_actions.append("Critical: Immediately generate an OMS ticket and dispatch ZME.")
            elif len(open_tickets) == 0:
                audit_categories.append(CATEGORY_PHANTOM)
                is_sla_breach = t_sla.upper() == "YES"
                risk_levels.append(RISK_CRITICAL if is_sla_breach else RISK_HIGH)
                prescribed_actions.append("Reopen ticket in OMS or create a follow-up ticket. Field unit is still down.")
            else:
                audit_categories.append(CATEGORY_ACTIVE)
                is_sla_breach = t_sla.upper() == "YES"
                if is_sla_breach or age_in_days > 5.0:
                    risk_levels.append(RISK_HIGH)
                else:
                    risk_levels.append(RISK_MEDIUM)
                prescribed_actions.append("Monitor active ticket progress and adhere to SLA resolution timeline.")

    reconciled = fir.copy()
    reconciled["Report_Date"] = report_dates
    reconciled["Tentative_ETA"] = tentative_etas
    reconciled["Audit_Category"] = audit_categories
    reconciled["Risk_Level"] = risk_levels
    reconciled["Prescribed_Action"] = prescribed_actions
    reconciled["Latest_Ticket_ID"] = latest_ticket_ids
    reconciled["Latest_Ticket_Status"] = latest_ticket_statuses
    reconciled["Latest_Ticket_Created_At"] = latest_ticket_created
    reconciled["Latest_Ticket_Updated_At"] = latest_ticket_updated
    reconciled["Latest_Ticket_Closed_At"] = latest_ticket_closed
    reconciled["Latest_Ticket_SLA_Deadline"] = latest_ticket_sla_deadline
    reconciled["Latest_Ticket_Ack_Deadline"] = latest_ticket_ack_deadline
    reconciled["Latest_Ticket_Category"] = latest_ticket_categories
    reconciled["Latest_Ticket_Subcategory"] = latest_ticket_subcats
    reconciled["Latest_Ticket_Assigned_To"] = latest_ticket_assigned
    reconciled["Latest_Ticket_SLA_Breached"] = latest_ticket_sla
    reconciled["Total_OMS_Tickets_Count"] = total_tickets_count
    reconciled["Open_OMS_Tickets_Count"] = open_tickets_count
    reconciled["Parsed_Issue_Age_Days"] = parsed_age_days

    fir_norm_set = set(reconciled["_norm_ocpp"])
    if oms_status_col in oms.columns and "_norm_ocpp" in oms.columns:
        untracked_mask = (
            (~oms["_norm_ocpp"].isin(fir_norm_set))
            & (oms[oms_status_col].astype(str).str.strip().str.upper().isin(OPEN_OMS_STATUSES))
        )
        untracked_oms_df = oms[untracked_mask].copy()
    else:
        untracked_oms_df = pd.DataFrame(columns=list(oms.columns))
    untracked_oms_df["Audit_Category"] = CATEGORY_UNTRACKED

    # Clean date formatting on untracked OMS records
    for col_to_fmt in [oms_created_col, oms_updated_col, oms_closed_col, oms_sla_deadline_col, oms_ack_deadline_col]:
        if col_to_fmt in untracked_oms_df.columns:
            untracked_oms_df[col_to_fmt] = untracked_oms_df[col_to_fmt].apply(format_date_str)

    gaps_df = reconciled[reconciled["Audit_Category"] == CATEGORY_GAP].copy()
    phantoms_df = reconciled[reconciled["Audit_Category"] == CATEGORY_PHANTOM].copy()
    active_df = reconciled[reconciled["Audit_Category"] == CATEGORY_ACTIVE].copy()
    healthy_df = reconciled[reconciled["Audit_Category"] == CATEGORY_HEALTHY].copy()

    total_fir_count = len(reconciled)
    down_fir_count = len(gaps_df) + len(phantoms_df) + len(active_df)
    healthy_count = len(healthy_df)
    gap_count = len(gaps_df)
    phantom_count = len(phantoms_df)
    active_count = len(active_df)
    untracked_count = len(untracked_oms_df)

    coverage_gap_rate = round(((gap_count + phantom_count) / down_fir_count * 100), 2) if down_fir_count > 0 else 0.0
    execution_time_ms = round((time.perf_counter() - start_time) * 1000, 2)

    summary = ReconciliationSummary(
        total_fir=total_fir_count,
        in_maintenance=down_fir_count,
        available_cleared=healthy_count,
        ticket_gaps=gap_count,
        phantom_resolutions=phantom_count,
        active_verified=active_count,
        coverage_gap_rate=coverage_gap_rate,
        untracked_oms_open=untracked_count,
        execution_time_ms=execution_time_ms,
    )

    return ReconciliationResult(
        summary=summary,
        reconciled_df=reconciled,
        gaps_df=gaps_df,
        phantoms_df=phantoms_df,
        active_df=active_df,
        healthy_df=healthy_df,
        untracked_oms_df=untracked_oms_df,
    )


def create_filtered_result(base_result: ReconciliationResult, filtered_df: pd.DataFrame) -> ReconciliationResult:
    gaps = filtered_df[filtered_df["Audit_Category"] == CATEGORY_GAP].copy()
    phantoms = filtered_df[filtered_df["Audit_Category"] == CATEGORY_PHANTOM].copy()
    active = filtered_df[filtered_df["Audit_Category"] == CATEGORY_ACTIVE].copy()
    healthy = filtered_df[filtered_df["Audit_Category"] == CATEGORY_HEALTHY].copy()

    down_count = len(gaps) + len(phantoms) + len(active)
    gap_rate = round(((len(gaps) + len(phantoms)) / down_count * 100), 2) if down_count > 0 else 0.0

    filtered_summary = ReconciliationSummary(
        total_fir=len(filtered_df),
        in_maintenance=down_count,
        available_cleared=len(healthy),
        ticket_gaps=len(gaps),
        phantom_resolutions=len(phantoms),
        active_verified=len(active),
        coverage_gap_rate=gap_rate,
        untracked_oms_open=base_result.summary.untracked_oms_open,
        execution_time_ms=base_result.summary.execution_time_ms,
        timestamp=base_result.summary.timestamp,
    )
    return ReconciliationResult(
        summary=filtered_summary,
        reconciled_df=filtered_df,
        gaps_df=gaps,
        phantoms_df=phantoms,
        active_df=active,
        healthy_df=healthy,
        untracked_oms_df=base_result.untracked_oms_df,
    )

