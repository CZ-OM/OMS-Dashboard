import html
import io
from datetime import date, datetime
from pathlib import Path
from typing import Any

import altair as alt
import pandas as pd
import streamlit as st

from core.exporter import export_oms_date_only_csv, export_oms_date_only_excel
from core.models import CLOSED_OMS_STATUSES, OPEN_OMS_STATUSES
from core.oms_converter import (
    RESOLVED_TIME_DATE_ONLY_COL,
    convert_oms_to_date_only,
    load_oms_raw_file,
    recalculate_hold_days,
)
from views.styles import (
    _coerce_display_cell,
    calc_table_height,
    sanitize_for_dataframe as _sanitize_for_dataframe,
)

_OMS_DASH_FILTER_KEYS: tuple[str, ...] = (
    "oms_dash_filter_pill",
    "oms_dash_filter_search",
    "oms_dash_filter_zones",
    "oms_dash_filter_statuses",
    "oms_dash_filter_priorities",
    "oms_dash_filter_categories",
    "oms_dash_filter_sla",
)

_SEARCHABLE_COLUMNS: tuple[str, ...] = (
    "Ticket ID",
    "Charger",
    "Station",
    "Location",
    "Assigned To",
    "Issue Category",
    "Issue Subcategory",
    "Status",
    "Zone",
    "Hold Reason",
    "Priority",
)

_TRACKER_PRIORITY_COLUMNS: tuple[str, ...] = (
    "Ticket ID",
    "Status",
    "Priority",
    "Zone",
    "Station",
    "Charger",
    "Issue Category",
    "Issue Subcategory",
    "Assigned To",
    "SLA Breached",
    "Created At",
    "Assigned At",
    "Hold Status Time",
    "Hold total time(Days)",
    "Resolved At Time",
    RESOLVED_TIME_DATE_ONLY_COL,
    "Closed At",
    "Total time taken (Days)",
    "SLA Breached As per Actual (Days)",
)


def init_oms_dashboard_state() -> None:
    if "oms_dash_raw_df" not in st.session_state:
        st.session_state["oms_dash_raw_df"] = None
    if "oms_dash_df" not in st.session_state:
        st.session_state["oms_dash_df"] = None
    if "oms_dash_filename" not in st.session_state:
        st.session_state["oms_dash_filename"] = None
    if "oms_dash_file_sig" not in st.session_state:
        st.session_state["oms_dash_file_sig"] = None
    if "oms_dash_active_today" not in st.session_state:
        st.session_state["oms_dash_active_today"] = date.today()


def _reset_oms_dashboard_filters() -> None:
    for key in _OMS_DASH_FILTER_KEYS:
        if key in st.session_state:
            del st.session_state[key]


def _clear_oms_dashboard_data() -> None:
    st.session_state["oms_dash_raw_df"] = None
    st.session_state["oms_dash_df"] = None
    st.session_state["oms_dash_filename"] = None
    st.session_state["oms_dash_file_sig"] = None
    _reset_oms_dashboard_filters()
    if "oms_dash_uploader" in st.session_state:
        del st.session_state["oms_dash_uploader"]


def _insert_row_number_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "No." in out.columns:
        alt_name = "Original No."
        suffix = 1
        while alt_name in out.columns:
            suffix += 1
            alt_name = f"Original No. ({suffix})"
        out = out.rename(columns={"No.": alt_name})
    out.insert(0, "No.", range(1, len(out) + 1))
    return out


def render_oms_dashboard_sidebar() -> None:
    init_oms_dashboard_state()

    st.title("OMS Issue Tracker")
    st.caption("Convert raw OMS dumps to date_only & monitor ticket SLA health.")

    st.markdown("---")
    st.subheader("Step 1: Upload OMS Dump")

    uploaded = st.file_uploader(
        "Raw OMS Dump (CSV / XLSX / XLS)",
        type=["xlsx", "xls", "csv"],
        key="oms_dash_uploader",
        help="Upload a raw OMS ticket export (e.g., Nov1-25-OMS format).",
    )

    clear_clicked = st.button(
        "Clear OMS Data",
        key="oms_dash_clear_btn",
        width="stretch",
        disabled=st.session_state["oms_dash_df"] is None,
    )
    if clear_clicked:
        _clear_oms_dashboard_data()
        st.rerun()

    st.markdown("---")
    st.subheader("Step 2: Reference Date ('Today')")

    if "oms_dash_today" not in st.session_state:
        st.session_state["oms_dash_today"] = st.session_state["oms_dash_active_today"]

    selected_today = st.date_input(
        '"Today" Reference Date',
        key="oms_dash_today",
        help='Reference date used to calculate "Hold total time(Days)" from Hold Status Time.',
    )

    active_today = (
        selected_today
        if isinstance(selected_today, date) and not isinstance(selected_today, datetime)
        else date.today()
    )
    st.session_state["oms_dash_active_today"] = active_today

    if uploaded is None:
        st.session_state["oms_dash_file_sig"] = None
    else:
        sig = (getattr(uploaded, "file_id", None), uploaded.name, uploaded.size)
        if sig != st.session_state["oms_dash_file_sig"]:
            try:
                raw_bytes = uploaded.getvalue()
                raw_df = load_oms_raw_file(io.BytesIO(raw_bytes), filename=uploaded.name)
                conv_df = convert_oms_to_date_only(raw_df, today_date=active_today)
                st.session_state["oms_dash_raw_df"] = raw_df
                st.session_state["oms_dash_df"] = conv_df
                st.session_state["oms_dash_filename"] = uploaded.name
                st.session_state["oms_dash_file_sig"] = sig
                _reset_oms_dashboard_filters()
            except Exception as exc:
                st.error(f"Could not read OMS dump '{uploaded.name}': {exc}")

    dash_df = st.session_state["oms_dash_df"]
    if dash_df is not None and dash_df.attrs.get("today_date") != active_today.strftime("%Y-%m-%d"):
        st.session_state["oms_dash_df"] = recalculate_hold_days(dash_df, today_date=active_today)
        dash_df = st.session_state["oms_dash_df"]

    if dash_df is not None:
        st.markdown("---")
        filename = st.session_state["oms_dash_filename"] or "OMS Dump"
        st.success(f"Loaded: {filename} ({len(dash_df):,} tickets)")
        st.caption(
            f"Active 'today' date: {active_today.strftime('%Y-%m-%d')} | {len(dash_df.columns)} columns"
        )


def render_oms_landing_view() -> None:
    st.markdown(
        """
        <div class="header-box">
            <h1>OMS Issue Tracker Dashboard</h1>
            <p>Convert raw OMS ticket dumps into the clean <code>date_only</code> schema, adjust the "today" hold-age reference date, monitor SLA performance, and export ready-to-share sheets.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            """
            <div class="step-box">
                <div class="step-number">Step 1</div>
                <div class="step-title">Upload Raw OMS Dump</div>
                <div class="step-desc">
                    Upload a <b>.csv</b>, <b>.xlsx</b>, or <b>.xls</b> OMS ticket export in the left sidebar. Timestamps are stripped to <code>YYYY-MM-DD</code> and all 8 day-delta columns are computed automatically.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            """
            <div class="step-box">
                <div class="step-number">Step 2</div>
                <div class="step-title">Set "Today" Reference Date</div>
                <div class="step-desc">
                    Defaults to today's system date. Change the <b>"Today" Reference Date</b> in the sidebar anytime to recalculate <b>Hold total time(Days)</b> on the fly.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            """
            <div class="step-box">
                <div class="step-number">Step 3</div>
                <div class="step-title">Analyze &amp; Export date_only</div>
                <div class="step-desc">
                    Explore SLA breaches, ticket status breakdowns, hold aging charts, searchable ticket tables, and download the converted <b>date_only</b> CSV or Excel workbook.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.info(
        "Upload an OMS dump (.csv, .xlsx, or .xls) in the left sidebar to launch the dashboard and generate the converted date_only sheet."
    )


def _normalized_upper_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series("", index=df.index, dtype="string")
    return df[col].fillna("").astype(str).str.strip().str.upper()


def _format_mean_days(series: pd.Series) -> str | None:
    vals = pd.to_numeric(series, errors="coerce").dropna()
    if len(vals) == 0:
        return None
    return f"{vals.mean():.1f} d"


def _format_unique_items(series: pd.Series, max_items: int = 5) -> str:
    if series is None or series.empty:
        return "-"
    valid = [
        str(x).strip()
        for x in series
        if x is not None and pd.notna(x) and str(x).strip() not in {"", "-", "--", "None", "nan", "NAN"}
    ]
    seen: list[str] = []
    for x in valid:
        if x not in seen:
            seen.append(x)
    if not seen:
        return "-"
    if len(seen) <= max_items:
        return ", ".join(seen)
    return ", ".join(seen[:max_items]) + f" (+{len(seen) - max_items} more)"


def compute_top_categories_and_subtypes(
    df: pd.DataFrame, top_cats: int = 10, top_subs: int = 5
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    if df.empty or "Issue Category" not in df.columns:
        return pd.DataFrame(), []

    tot_df = len(df)
    clean_cat = df["Issue Category"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
    cat_counts = clean_cat.value_counts().head(top_cats)

    summary_rows: list[dict[str, Any]] = []
    cat_details: list[dict[str, Any]] = []

    for cat_name, cat_count in cat_counts.items():
        share_pct = round(cat_count / tot_df * 100, 1) if tot_df else 0.0
        summary_rows.append(
            {
                "Issue Category": cat_name,
                "Total Tickets": int(cat_count),
                "Share %": share_pct,
            }
        )

        c_slice = df[clean_cat == cat_name].copy()
        clean_sub = (
            c_slice["Issue Subcategory"]
            .fillna("Unspecified")
            .astype(str)
            .str.strip()
            .replace("", "Unspecified")
            if "Issue Subcategory" in c_slice.columns
            else pd.Series(["Unspecified"] * len(c_slice), index=c_slice.index)
        )
        c_slice["_sub_clean"] = clean_sub

        sub_counts = clean_sub.value_counts().head(top_subs)
        sub_rows: list[dict[str, Any]] = []
        for sub_name, sub_count in sub_counts.items():
            sub_slice = c_slice[c_slice["_sub_clean"] == sub_name]
            uniq_chargers = (
                sub_slice["Charger"].dropna().astype(str).str.strip().nunique()
                if "Charger" in sub_slice.columns
                else 0
            )
            sample_chargers = (
                _format_unique_items(sub_slice["Charger"])
                if "Charger" in sub_slice.columns
                else "-"
            )
            uniq_stations = (
                sub_slice["Station"].dropna().astype(str).str.strip().nunique()
                if "Station" in sub_slice.columns
                else 0
            )
            sub_share_pct = round(sub_count / cat_count * 100, 1) if cat_count else 0.0

            sub_rows.append(
                {
                    "Issue Subcategory": sub_name,
                    "Occurrences": int(sub_count),
                    "Sub-Type Share %": sub_share_pct,
                    "Unique Chargers": uniq_chargers,
                    "Affected Chargers": sample_chargers,
                    "Unique Stations": uniq_stations,
                }
            )

        cat_details.append(
            {
                "category": cat_name,
                "total_tickets": int(cat_count),
                "share_pct": share_pct,
                "sub_df": pd.DataFrame(sub_rows),
                "raw_slice": c_slice,
            }
        )

    return pd.DataFrame(summary_rows), cat_details


def compute_station_repetitive_issues(
    df: pd.DataFrame, top_stations: int = 10, top_subs_per_station: int = 5
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    if df.empty or "Station" not in df.columns:
        return pd.DataFrame(), []

    clean_st = df["Station"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
    st_counts = clean_st.value_counts().head(top_stations)

    summary_rows: list[dict[str, Any]] = []
    station_details: list[dict[str, Any]] = []

    for st_name, st_count in st_counts.items():
        st_slice = df[clean_st == st_name].copy()
        clean_sub = (
            st_slice["Issue Subcategory"]
            .fillna("Unspecified")
            .astype(str)
            .str.strip()
            .replace("", "Unspecified")
            if "Issue Subcategory" in st_slice.columns
            else pd.Series(["Unspecified"] * len(st_slice), index=st_slice.index)
        )
        st_slice["_sub_clean"] = clean_sub

        top_sub_name = clean_sub.value_counts().index[0] if not clean_sub.empty else "N/A"
        top_sub_cnt = int(clean_sub.value_counts().iloc[0]) if not clean_sub.empty else 0
        uniq_chargers = (
            st_slice["Charger"].dropna().astype(str).str.strip().nunique()
            if "Charger" in st_slice.columns
            else 0
        )

        summary_rows.append(
            {
                "Station": st_name,
                "Total Tickets": int(st_count),
                "Unique Chargers": uniq_chargers,
                "Top Failure Mode": top_sub_name,
                "Top Mode Count": top_sub_cnt,
            }
        )

        sub_counts = clean_sub.value_counts().head(top_subs_per_station)
        sub_rows: list[dict[str, Any]] = []
        for sub_name, sub_count in sub_counts.items():
            sub_slice = st_slice[st_slice["_sub_clean"] == sub_name]
            sample_chargers = (
                _format_unique_items(sub_slice["Charger"])
                if "Charger" in sub_slice.columns
                else "-"
            )
            engineers = (
                _format_unique_items(sub_slice["Assigned To"])
                if "Assigned To" in sub_slice.columns
                else "-"
            )
            sub_share_pct = round(sub_count / st_count * 100, 1) if st_count else 0.0

            sub_rows.append(
                {
                    "Issue Subcategory": sub_name,
                    "Occurrences": int(sub_count),
                    "Sub-Type Share %": sub_share_pct,
                    "Affected Chargers": sample_chargers,
                    "Assigned Engineers": engineers,
                }
            )

        station_details.append(
            {
                "station": st_name,
                "total_tickets": int(st_count),
                "unique_chargers": uniq_chargers,
                "sub_df": pd.DataFrame(sub_rows),
                "raw_slice": st_slice,
            }
        )

    return pd.DataFrame(summary_rows), station_details


def compute_repetitive_fault_summary(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty or "Issue Category" not in df.columns or "Issue Subcategory" not in df.columns:
        return {
            "total_repetitive_tickets": 0,
            "repetitive_share_pct": 0.0,
            "top_failure_mode": "N/A",
            "top_failure_count": 0,
            "top_station_hotspot": "N/A",
            "top_station_count": 0,
            "distinct_hotspots": 0,
            "modes_df": pd.DataFrame(),
        }

    clean_df = df.copy()
    clean_df["_cat"] = clean_df["Issue Category"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
    clean_df["_sub"] = clean_df["Issue Subcategory"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")

    grouped = clean_df.groupby(["_cat", "_sub"]).agg(
        Occurrences=("Ticket ID", "count") if "Ticket ID" in clean_df.columns else ("_cat", "count"),
        Unique_Stations=("Station", lambda s: s.dropna().astype(str).str.strip().nunique()) if "Station" in clean_df.columns else ("_cat", lambda _: 0),
        Unique_Chargers=("Charger", lambda s: s.dropna().astype(str).str.strip().nunique()) if "Charger" in clean_df.columns else ("_cat", lambda _: 0),
    ).reset_index()

    grouped.rename(columns={"_cat": "Issue Category", "_sub": "Issue Subcategory"}, inplace=True)
    rep_modes = grouped[grouped["Occurrences"] > 1].sort_values(by="Occurrences", ascending=False)
    total_rep = int(rep_modes["Occurrences"].sum()) if not rep_modes.empty else 0
    rep_share = round(total_rep / len(df) * 100, 1) if len(df) else 0.0

    top_mode_str = "N/A"
    top_mode_cnt = 0
    if not rep_modes.empty:
        top_row = rep_modes.iloc[0]
        top_mode_str = f"{top_row['Issue Category']} — {top_row['Issue Subcategory']}"
        top_mode_cnt = int(top_row["Occurrences"])

    top_st_str = "N/A"
    top_st_cnt = 0
    distinct_hotspots = 0
    if "Station" in df.columns:
        st_s = df["Station"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified").value_counts()
        if not st_s.empty:
            top_st_str = st_s.index[0]
            top_st_cnt = int(st_s.iloc[0])
            distinct_hotspots = int((st_s > 1).sum())

    return {
        "total_repetitive_tickets": total_rep,
        "repetitive_share_pct": rep_share,
        "top_failure_mode": top_mode_str,
        "top_failure_count": top_mode_cnt,
        "top_station_hotspot": top_st_str,
        "top_station_count": top_st_cnt,
        "distinct_hotspots": distinct_hotspots,
        "modes_df": rep_modes,
    }


def compute_engineer_performance(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Assigned To" not in df.columns:
        return pd.DataFrame()

    eng_series = df["Assigned To"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
    status_clean = df["Status"].fillna("Unspecified").astype(str).str.strip() if "Status" in df.columns else pd.Series(["Unspecified"] * len(df), index=df.index)

    if "SLA Breached As per Actual (Days)" in df.columns:
        sla_actual_breach = pd.to_numeric(df["SLA Breached As per Actual (Days)"], errors="coerce") > 0
    else:
        sla_actual_breach = pd.Series(False, index=df.index)

    if "SLA Breached" in df.columns:
        sla_col_breach = df["SLA Breached"].fillna("").astype(str).str.strip().str.upper() == "YES"
    else:
        sla_col_breach = pd.Series(False, index=df.index)

    sla_breached_series = sla_actual_breach | sla_col_breach

    resolved_days = (
        pd.to_numeric(df[RESOLVED_TIME_DATE_ONLY_COL], errors="coerce")
        if RESOLVED_TIME_DATE_ONLY_COL in df.columns
        else pd.Series(dtype="float64", index=df.index)
    )
    hold_days = (
        pd.to_numeric(df["Hold total time(Days)"], errors="coerce")
        if "Hold total time(Days)" in df.columns
        else pd.Series(dtype="float64", index=df.index)
    )

    work_df = pd.DataFrame(
        {
            "Engineer": eng_series,
            "Status": status_clean,
            "SLA_Breached": sla_breached_series,
            "Resolved_Days": resolved_days,
            "Hold_Days": hold_days,
        }
    )

    perf = (
        work_df.groupby("Engineer")
        .agg(
            Total=("Status", "count"),
            Resolved=("Status", lambda s: s.isin(["Resolved", "Closed", "Complete", "Completed"]).sum()),
            In_Progress=("Status", lambda s: (s == "In Progress").sum()),
            Pending_Approval=("Status", lambda s: s.str.contains("Pending", case=False, na=False).sum()),
            On_Hold=("Status", lambda s: (s == "On Hold").sum()),
            Assigned=("Status", lambda s: (s == "Assigned").sum()),
            SLA_Breached=("SLA_Breached", "sum"),
            Avg_Res_Days=("Resolved_Days", lambda s: s[s.notna()].mean() if s.notna().any() else None),
            Avg_Hold_Days=("Hold_Days", lambda s: s[s.notna()].mean() if s.notna().any() else None),
        )
        .reset_index()
        .sort_values(by="Total", ascending=False)
    )

    perf["Resolution Rate %"] = (perf["Resolved"] / perf["Total"] * 100).round(1)
    perf["SLA Adherence %"] = ((perf["Total"] - perf["SLA_Breached"]).clip(lower=0) / perf["Total"] * 100).round(1)

    return perf


def render_oms_dashboard_main() -> None:
    init_oms_dashboard_state()
    base_df: pd.DataFrame | None = st.session_state["oms_dash_df"]
    if base_df is None:
        render_oms_landing_view()
        return

    filename: str = st.session_state["oms_dash_filename"] or "oms_dump"
    file_stem = Path(filename).stem or "oms_dump"
    active_today: date = st.session_state["oms_dash_active_today"]
    today_str = active_today.strftime("%Y-%m-%d")

    csv_payload = export_oms_date_only_csv(base_df)
    xlsx_payload = export_oms_date_only_excel(base_df)

    col_title, col_csv, col_xlsx = st.columns([2.6, 1.2, 1.2])
    with col_title:
        st.markdown(
            f"""
            <div style="margin-bottom: 8px;">
                <h1 style="font-size: 1.75rem; font-weight: 800; color: #0F172A; margin: 0 0 4px 0;">OMS Issue Tracker Dashboard</h1>
                <p style="margin: 0; color: #64748B; font-size: 0.92rem;">
                    File: {html.escape(filename)} &nbsp;|&nbsp; Tickets: {len(base_df):,} &nbsp;|&nbsp; "Today" Ref: {today_str} &nbsp;|&nbsp; Schema: {len(base_df.columns)} columns
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_csv:
        st.download_button(
            label="Download date_only CSV",
            data=csv_payload,
            file_name=f"{file_stem}_date_only.csv",
            mime="text/csv",
            key="oms_dash_dl_csv_top",
            width="stretch",
        )
    with col_xlsx:
        st.download_button(
            label="Download date_only Excel",
            data=xlsx_payload,
            file_name=f"{file_stem}_date_only.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="oms_dash_dl_xlsx_top",
            width="stretch",
        )

    base_status_upper = _normalized_upper_series(base_df, "Status")
    base_sla_upper = _normalized_upper_series(base_df, "SLA Breached")

    total_count = len(base_df)
    open_count = int(base_status_upper.isin(OPEN_OMS_STATUSES).sum())
    closed_count = int(base_status_upper.isin(CLOSED_OMS_STATUSES).sum())
    hold_count = int((base_status_upper == "ON HOLD").sum())
    sla_yes_count = int((base_sla_upper == "YES").sum())

    st.markdown(
        """
        <div class="filter-card">
            <div class="filter-title">Dynamic Filters &amp; Live Search</div>
            <div style="font-size: 0.85rem; color: #64748B; margin-bottom: 8px;">
                Select a quick preset pill, search by Ticket ID, charger, station, or assignee, or refine using the dropdown filters below.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pill_options = [
        f"All Tickets ({total_count})",
        f"Open / Active ({open_count})",
        f"Closed / Resolved ({closed_count})",
        f"On Hold ({hold_count})",
        f"SLA Breached ({sla_yes_count})",
    ]
    selected_pill = st.pills(
        "OMS Quick Presets",
        options=pill_options,
        default=pill_options[0],
        key="oms_dash_filter_pill",
        label_visibility="collapsed",
    )

    col_s1, col_s2 = st.columns([5, 1])
    with col_s1:
        search_val = st.text_input(
            "OMS Live Search",
            placeholder="Search Ticket ID, Charger, Station, Location, Assignee, Category, or Zone...",
            key="oms_dash_filter_search",
            label_visibility="collapsed",
        ).strip()
    with col_s2:
        if st.button("Reset Filters", key="oms_dash_filter_reset", width="stretch"):
            _reset_oms_dashboard_filters()
            st.rerun()

    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)

    def _clean_options(col_name: str) -> list[str]:
        if col_name not in base_df.columns:
            return []
        vals = base_df[col_name].dropna().astype(str).str.strip()
        return sorted({v for v in vals if v != ""})

    with col_f1:
        zone_opts = _clean_options("Zone")
        selected_zones = st.multiselect(
            "Zone",
            zone_opts,
            key="oms_dash_filter_zones",
            placeholder="All Zones" if zone_opts else "No Zone column",
            disabled=len(zone_opts) == 0,
        )

    with col_f2:
        status_opts = _clean_options("Status")
        selected_statuses = st.multiselect(
            "Status",
            status_opts,
            key="oms_dash_filter_statuses",
            placeholder="All Statuses" if status_opts else "No Status column",
            disabled=len(status_opts) == 0,
        )

    with col_f3:
        priority_opts = _clean_options("Priority")
        selected_priorities = st.multiselect(
            "Priority",
            priority_opts,
            key="oms_dash_filter_priorities",
            placeholder="All Priorities" if priority_opts else "No Priority column",
            disabled=len(priority_opts) == 0,
        )

    with col_f4:
        cat_opts = _clean_options("Issue Category")
        selected_categories = st.multiselect(
            "Issue Category",
            cat_opts,
            key="oms_dash_filter_categories",
            placeholder="All Categories" if cat_opts else "No Category column",
            disabled=len(cat_opts) == 0,
        )

    with col_f5:
        sla_opts = _clean_options("SLA Breached")
        selected_sla = st.multiselect(
            "SLA Breached",
            sla_opts,
            key="oms_dash_filter_sla",
            placeholder="All SLA" if sla_opts else "No SLA column",
            disabled=len(sla_opts) == 0,
        )

    filtered_df = base_df.copy()

    if selected_pill and selected_pill.startswith("Open / Active"):
        f_status = _normalized_upper_series(filtered_df, "Status")
        filtered_df = filtered_df[f_status.isin(OPEN_OMS_STATUSES)]
    elif selected_pill and selected_pill.startswith("Closed / Resolved"):
        f_status = _normalized_upper_series(filtered_df, "Status")
        filtered_df = filtered_df[f_status.isin(CLOSED_OMS_STATUSES)]
    elif selected_pill and selected_pill.startswith("On Hold"):
        f_status = _normalized_upper_series(filtered_df, "Status")
        filtered_df = filtered_df[f_status == "ON HOLD"]
    elif selected_pill and selected_pill.startswith("SLA Breached"):
        f_sla = _normalized_upper_series(filtered_df, "SLA Breached")
        filtered_df = filtered_df[f_sla == "YES"]

    for col_name, selected_vals in (
        ("Zone", selected_zones),
        ("Status", selected_statuses),
        ("Priority", selected_priorities),
        ("Issue Category", selected_categories),
        ("SLA Breached", selected_sla),
    ):
        if selected_vals and col_name in filtered_df.columns:
            col_clean = filtered_df[col_name].fillna("").astype(str).str.strip()
            filtered_df = filtered_df[col_clean.isin(selected_vals)]

    if search_val:
        q = search_val.lower()
        search_mask = pd.Series(False, index=filtered_df.index)
        for col in _SEARCHABLE_COLUMNS:
            if col in filtered_df.columns:
                search_mask |= (
                    filtered_df[col].fillna("").astype(str).str.lower().str.contains(q, regex=False)
                )
        filtered_df = filtered_df[search_mask]

    active_filters: list[str] = []
    if selected_pill and not selected_pill.startswith("All Tickets"):
        active_filters.append(f"Preset: {selected_pill.split('(')[0].strip()}")
    if selected_zones:
        active_filters.append(f"Zone: {', '.join(selected_zones)}")
    if selected_statuses:
        active_filters.append(f"Status: {', '.join(selected_statuses)}")
    if selected_priorities:
        active_filters.append(f"Priority: {', '.join(selected_priorities)}")
    if selected_categories:
        active_filters.append(f"Category: {', '.join(selected_categories)}")
    if selected_sla:
        active_filters.append(f"SLA: {', '.join(selected_sla)}")
    if search_val:
        active_filters.append(f"Search: '{search_val}'")

    is_filtered = len(active_filters) > 0 or len(filtered_df) < total_count
    if is_filtered:
        tags_html = " ".join(
            f'<span class="filter-tag">{html.escape(f)}</span>' for f in active_filters
        )
        pct = round(len(filtered_df) / total_count * 100, 1) if total_count else 0.0
        st.markdown(
            f"""
            <div class="filter-summary-strip">
                <b>Active Filters:</b> {tags_html} &nbsp;&bull;&nbsp;
                <b>Showing {len(filtered_df):,} of {total_count:,} tickets</b> ({pct}% of total).
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div class="filter-summary-strip">
                Showing all <b>{total_count:,} tickets</b>. Use the quick presets or dropdowns above to filter dynamically.
            </div>
            """,
            unsafe_allow_html=True,
        )

    cur_status = _normalized_upper_series(filtered_df, "Status")
    cur_sla = _normalized_upper_series(filtered_df, "SLA Breached")

    cur_total = len(filtered_df)
    cur_open = int(cur_status.isin(OPEN_OMS_STATUSES).sum())
    cur_closed = int(cur_status.isin(CLOSED_OMS_STATUSES).sum())
    cur_on_hold = int((cur_status == "ON HOLD").sum())
    cur_sla_yes = int((cur_sla == "YES").sum())

    open_pct = round(cur_open / cur_total * 100, 1) if cur_total > 0 else 0.0
    closed_pct = round(cur_closed / cur_total * 100, 1) if cur_total > 0 else 0.0
    sla_pct = round(cur_sla_yes / cur_total * 100, 1) if cur_total > 0 else 0.0

    if "Hold total time(Days)" in filtered_df.columns:
        on_hold_mask = cur_status == "ON HOLD"
        on_hold_mean_str = (
            _format_mean_days(filtered_df.loc[on_hold_mask, "Hold total time(Days)"])
            if bool(on_hold_mask.any())
            else None
        )
        if on_hold_mean_str is not None:
            hold_sub = f"Avg On Hold age: {on_hold_mean_str}"
        else:
            all_hold_mean_str = _format_mean_days(filtered_df["Hold total time(Days)"])
            if all_hold_mean_str is not None:
                hold_sub = f"Avg hold-date age (all): {all_hold_mean_str}"
            else:
                hold_sub = "Avg On Hold age: N/A"
    else:
        hold_sub = "Hold age: N/A (no Hold Status Time)"

    ack_mean_str = (
        _format_mean_days(filtered_df["Acknowledgement time taken(Days)"])
        if "Acknowledgement time taken(Days)" in filtered_df.columns
        else None
    )

    res_value = "N/A"
    res_sub = "No turnaround dates"
    res_mean_str = (
        _format_mean_days(filtered_df[RESOLVED_TIME_DATE_ONLY_COL])
        if RESOLVED_TIME_DATE_ONLY_COL in filtered_df.columns
        else None
    )
    if res_mean_str is not None:
        res_value = res_mean_str
        res_sub = f"Avg Ack: {ack_mean_str}" if ack_mean_str is not None else "Avg Ack: N/A"
    elif "Total time taken (Days)" in filtered_df.columns:
        tot_mean_str = _format_mean_days(filtered_df["Total time taken (Days)"])
        if tot_mean_str is not None:
            res_value = tot_mean_str
            res_sub = "Total Closed - Assigned"
    if res_value == "N/A" and "Created At" in filtered_df.columns and "Closed At" in filtered_df.columns:
        created_dt = pd.to_datetime(filtered_df["Created At"], errors="coerce")
        closed_dt = pd.to_datetime(filtered_df["Closed At"], errors="coerce")
        turn_days = (closed_dt - created_dt).dt.days.dropna()
        if len(turn_days) > 0:
            res_value = f"{turn_days.mean():.1f} d"
            res_sub = "Closed - Created (fallback)"
    if res_value == "N/A" and ack_mean_str is not None:
        res_value = ack_mean_str
        res_sub = "Avg Ack only (no resolution)"

    st.markdown("### Operational Summary")
    k1, k2, k3, k4, k5, k6 = st.columns(6)

    with k1:
        sub_txt = f"of {total_count:,} total (filtered)" if is_filtered else "Tickets in OMS dump"
        st.markdown(
            f"""
            <div class="kpi-card card-neutral">
                <div class="kpi-title">Total Tickets</div>
                <div class="kpi-value" style="color: #1E293B;">{cur_total:,}</div>
                <div class="kpi-subtitle">{sub_txt}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k2:
        st.markdown(
            f"""
            <div class="kpi-card card-warning">
                <div class="kpi-title">Open / Active</div>
                <div class="kpi-value" style="color: #D97706;">{cur_open:,}</div>
                <div class="kpi-subtitle">{open_pct}% of shown</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k3:
        st.markdown(
            f"""
            <div class="kpi-card card-success">
                <div class="kpi-title">Closed / Resolved</div>
                <div class="kpi-value" style="color: #059669;">{cur_closed:,}</div>
                <div class="kpi-subtitle">Resolution rate: {closed_pct}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k4:
        sla_sub = f"Breach rate: {sla_pct}%" if "SLA Breached" in filtered_df.columns else "No SLA Breached column"
        st.markdown(
            f"""
            <div class="kpi-card card-danger">
                <div class="kpi-title">SLA Breached</div>
                <div class="kpi-value" style="color: #DC2626;">{cur_sla_yes:,}</div>
                <div class="kpi-subtitle">{sla_sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k5:
        st.markdown(
            f"""
            <div class="kpi-card card-gray">
                <div class="kpi-title">On Hold &amp; Avg Age</div>
                <div class="kpi-value" style="color: #475569;">{cur_on_hold:,}</div>
                <div class="kpi-subtitle">{hold_sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k6:
        st.markdown(
            f"""
            <div class="kpi-card card-neutral">
                <div class="kpi-title">Avg Resolution / Ack</div>
                <div class="kpi-value" style="color: #1D4ED8;">{res_value}</div>
                <div class="kpi-subtitle">{res_sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "📊 Summary & SLA Health",
            "🔁 Repetitive Issues & Top Types",
            "👷 Engineer Performance",
            f"📋 Ticket Tracker ({len(filtered_df):,})",
            f"📁 Converted date_only Schema ({len(filtered_df):,})",
        ]
    )

    with tab1:
        st.subheader("Executive Overview & SLA Health")
        st.caption(f"Live analytics for the filtered OMS dataset (Active 'today' reference: {today_str}).")

        col_c1, col_c2 = st.columns(2)

        with col_c1:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Tickets by Status</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Ticket distribution across operational workflow states.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Status" in filtered_df.columns and len(filtered_df) > 0:
                status_df = filtered_df.copy()
                status_df["Status_Clean"] = status_df["Status"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
                if "Zone" in status_df.columns:
                    status_df["Zone_Clean"] = status_df["Zone"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
                    chart_status_df = status_df.groupby(["Status_Clean", "Zone_Clean"]).size().reset_index(name="Tickets")
                    chart_status = (
                        alt.Chart(chart_status_df)
                        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                        .encode(
                            x=alt.X("Status_Clean:N", title="Status", axis=alt.Axis(labelAngle=-20)),
                            y=alt.Y("Tickets:Q", title="Tickets"),
                            color=alt.Color(
                                "Zone_Clean:N",
                                title="Zone",
                                scale=alt.Scale(range=["#F87171", "#FBBF24", "#34D399", "#60A5FA", "#818CF8", "#94A3B8"]),
                            ),
                            tooltip=[
                                alt.Tooltip("Status_Clean:N", title="Status"),
                                alt.Tooltip("Zone_Clean:N", title="Zone"),
                                alt.Tooltip("Tickets:Q", title="Tickets"),
                            ],
                        )
                        .properties(height=280)
                    )
                else:
                    chart_status_df = status_df.groupby("Status_Clean").size().reset_index(name="Tickets")
                    chart_status = (
                        alt.Chart(chart_status_df)
                        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3, color="#60A5FA")
                        .encode(
                            x=alt.X("Status_Clean:N", title="Status", axis=alt.Axis(labelAngle=-20)),
                            y=alt.Y("Tickets:Q", title="Tickets"),
                            tooltip=[
                                alt.Tooltip("Status_Clean:N", title="Status"),
                                alt.Tooltip("Tickets:Q", title="Tickets"),
                            ],
                        )
                        .properties(height=280)
                    )
                st.altair_chart(chart_status, width="stretch")
            else:
                st.info("No ticket status data matching current filters.")

        with col_c2:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Tickets by Priority &amp; SLA Status</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Priority breakdown highlighting SLA breaches.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Priority" in filtered_df.columns and len(filtered_df) > 0:
                prio_df = filtered_df.copy()
                prio_df["Priority_Clean"] = prio_df["Priority"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
                if "SLA Breached" in prio_df.columns:
                    prio_df["SLA_Clean"] = prio_df["SLA Breached"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
                    chart_prio_df = prio_df.groupby(["Priority_Clean", "SLA_Clean"]).size().reset_index(name="Tickets")
                    chart_prio = (
                        alt.Chart(chart_prio_df)
                        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                        .encode(
                            x=alt.X("Priority_Clean:N", title="Priority", axis=alt.Axis(labelAngle=0)),
                            y=alt.Y("Tickets:Q", title="Tickets"),
                            color=alt.Color(
                                "SLA_Clean:N",
                                title="SLA Breached",
                                scale=alt.Scale(
                                    domain=["Yes", "No", "Unknown"],
                                    range=["#F87171", "#34D399", "#94A3B8"],
                                ),
                            ),
                            tooltip=[
                                alt.Tooltip("Priority_Clean:N", title="Priority"),
                                alt.Tooltip("SLA_Clean:N", title="SLA Breached"),
                                alt.Tooltip("Tickets:Q", title="Tickets"),
                            ],
                        )
                        .properties(height=280)
                    )
                else:
                    chart_prio_df = prio_df.groupby("Priority_Clean").size().reset_index(name="Tickets")
                    chart_prio = (
                        alt.Chart(chart_prio_df)
                        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3, color="#F87171")
                        .encode(
                            x=alt.X("Priority_Clean:N", title="Priority", axis=alt.Axis(labelAngle=0)),
                            y=alt.Y("Tickets:Q", title="Tickets"),
                            tooltip=[
                                alt.Tooltip("Priority_Clean:N", title="Priority"),
                                alt.Tooltip("Tickets:Q", title="Tickets"),
                            ],
                        )
                        .properties(height=280)
                    )
                st.altair_chart(chart_prio, width="stretch")
            else:
                st.info("No priority data matching current filters.")

        col_c3, col_c4 = st.columns(2)

        with col_c3:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Top 10 Issue Categories</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Most frequent issue categories across matching OMS tickets.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Issue Category" in filtered_df.columns and len(filtered_df) > 0:
                cat_df = filtered_df.copy()
                cat_df["Category_Clean"] = cat_df["Issue Category"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
                cat_chart_df = (
                    cat_df.groupby("Category_Clean")
                    .size()
                    .reset_index(name="Count")
                    .sort_values(by="Count", ascending=False)
                    .head(10)
                )
                chart_cat = (
                    alt.Chart(cat_chart_df)
                    .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#818CF8")
                    .encode(
                        x=alt.X("Count:Q", title="Tickets"),
                        y=alt.Y("Category_Clean:N", sort="-x", title="Issue Category"),
                        tooltip=[
                            alt.Tooltip("Category_Clean:N", title="Category"),
                            alt.Tooltip("Count:Q", title="Tickets"),
                        ],
                    )
                    .properties(height=280)
                )
                st.altair_chart(chart_cat, width="stretch")
            else:
                st.info("No issue category data matching current filters.")

        with col_c4:
            has_hold_vals = False
            if "Hold total time(Days)" in filtered_df.columns and "Ticket ID" in filtered_df.columns and len(filtered_df) > 0:
                hold_chart_df = pd.DataFrame(
                    {
                        "Ticket ID": filtered_df["Ticket ID"].fillna("").astype(str).str.strip(),
                        "Hold_Days": pd.to_numeric(filtered_df["Hold total time(Days)"], errors="coerce"),
                    }
                ).dropna(subset=["Hold_Days"])
                hold_chart_df = hold_chart_df[hold_chart_df["Ticket ID"] != ""]
                if len(hold_chart_df) > 0:
                    has_hold_vals = True
                    hold_chart_df["Hold_Days"] = hold_chart_df["Hold_Days"].astype(int)
                    hold_chart_df = (
                        hold_chart_df.sort_values(by="Hold_Days", ascending=False)
                        .head(15)
                    )
                    st.markdown(
                        f"""
                        <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Hold Aging by Ticket (Days as of {today_str})</div>
                        <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Dynamically recalculated from Hold Status Time using the selected "Today" reference date.</div>
                        """,
                        unsafe_allow_html=True,
                    )
                    chart_hold = (
                        alt.Chart(hold_chart_df)
                        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3, color="#FBBF24")
                        .encode(
                            x=alt.X("Ticket ID:N", sort="-y", title="Ticket ID", axis=alt.Axis(labelAngle=-25)),
                            y=alt.Y("Hold_Days:Q", title="Hold Days"),
                            tooltip=[
                                alt.Tooltip("Ticket ID:N", title="Ticket ID"),
                                alt.Tooltip("Hold_Days:Q", title="Hold total time (Days)"),
                            ],
                        )
                        .properties(height=280)
                    )
                    st.altair_chart(chart_hold, width="stretch")

            if not has_hold_vals:
                st.markdown(
                    """
                    <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Top 10 Stations by Ticket Volume</div>
                    <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Stations generating the highest volume of OMS tickets.</div>
                    """,
                    unsafe_allow_html=True,
                )
                if "Station" in filtered_df.columns and len(filtered_df) > 0:
                    st_df = filtered_df.copy()
                    st_df["Station_Clean"] = st_df["Station"].fillna("Unspecified").astype(str).str.strip().replace("", "Unspecified")
                    station_chart_df = (
                        st_df.groupby("Station_Clean")
                        .size()
                        .reset_index(name="Tickets")
                        .sort_values(by="Tickets", ascending=False)
                        .head(10)
                    )
                    chart_station = (
                        alt.Chart(station_chart_df)
                        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#60A5FA")
                        .encode(
                            x=alt.X("Tickets:Q", title="Tickets"),
                            y=alt.Y("Station_Clean:N", sort="-x", title="Station"),
                            tooltip=[
                                alt.Tooltip("Station_Clean:N", title="Station"),
                                alt.Tooltip("Tickets:Q", title="Tickets"),
                            ],
                        )
                        .properties(height=280)
                    )
                    st.altair_chart(chart_station, width="stretch")
                else:
                    st.info("No hold aging or station data matching current filters.")

    with tab2:
        st.subheader("Repetitive Issues & Root Cause Hotspots")
        st.caption("Network-wide recurring failure modes, Top 10 issue categories & Top 5 subcategories, and site-level failure hotspots.")

        rep_summary = compute_repetitive_fault_summary(filtered_df)
        rk1, rk2, rk3, rk4 = st.columns(4)

        with rk1:
            st.markdown(
                f"""
                <div class="kpi-card card-danger">
                    <div class="kpi-title">Repetitive Incidents</div>
                    <div class="kpi-value" style="color: #DC2626;">{rep_summary['total_repetitive_tickets']:,}</div>
                    <div class="kpi-subtitle">{rep_summary['repetitive_share_pct']}% of Total Network Faults</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with rk2:
            st.markdown(
                f"""
                <div class="kpi-card card-warning">
                    <div class="kpi-title">#1 Failure Mode</div>
                    <div class="kpi-value" style="color: #D97706; font-size: 1.25rem; line-height: 1.3;">{html.escape(rep_summary['top_failure_mode'][:32])}</div>
                    <div class="kpi-subtitle">{rep_summary['top_failure_count']:,} Occurrences</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with rk3:
            st.markdown(
                f"""
                <div class="kpi-card card-neutral">
                    <div class="kpi-title">#1 Station Hotspot</div>
                    <div class="kpi-value" style="color: #1D4ED8; font-size: 1.25rem; line-height: 1.3;">{html.escape(rep_summary['top_station_hotspot'][:30])}</div>
                    <div class="kpi-subtitle">{rep_summary['top_station_count']:,} Recurring Tickets</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with rk4:
            st.markdown(
                f"""
                <div class="kpi-card card-gray">
                    <div class="kpi-title">Hotspot Stations</div>
                    <div class="kpi-value" style="color: #475569;">{rep_summary['distinct_hotspots']:,}</div>
                    <div class="kpi-subtitle">Sites with &gt; 1 logged fault</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        rep_subtab1, rep_subtab2, rep_subtab3 = st.tabs(
            [
                "⚡ Top 10 Issue Types & Top 5 Sub-Types",
                "📍 Top 10 Stations (Station Repetitive Issues)",
                "🔁 Network-Wide Repetitive Fault Modes",
            ]
        )

        with rep_subtab1:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1.05rem; color: #0F172A; margin-bottom: 2px;">Top 10 Issue Categories &amp; Top 5 Subcategories Breakdown</div>
                <div style="font-size: 0.84rem; color: #64748B; margin-bottom: 12px;">High-frequency equipment and infrastructure categories. Expand any category below to view its top 5 sub-types, affected chargers, and matching tickets.</div>
                """,
                unsafe_allow_html=True,
            )

            top_cats_df, cat_details = compute_top_categories_and_subtypes(filtered_df, top_cats=10, top_subs=5)

            if not top_cats_df.empty:
                col_tc1, col_tc2 = st.columns([1.1, 1.0])
                with col_tc1:
                    chart_top_cats = (
                        alt.Chart(top_cats_df)
                        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#EF4444")
                        .encode(
                            x=alt.X("Total Tickets:Q", title="Total Tickets"),
                            y=alt.Y("Issue Category:N", sort="-x", title="Issue Category"),
                            tooltip=[
                                alt.Tooltip("Issue Category:N", title="Category"),
                                alt.Tooltip("Total Tickets:Q", title="Total Tickets"),
                                alt.Tooltip("Share %:Q", title="Share of Total (%)", format=".1f"),
                            ],
                        )
                        .properties(height=300)
                    )
                    st.altair_chart(chart_top_cats, width="stretch")

                with col_tc2:
                    st.write("##### Summary Table: Top 10 Categories")
                    disp_top_cats = _insert_row_number_column(top_cats_df)
                    st.dataframe(
                        disp_top_cats,
                        width="stretch",
                        hide_index=True,
                        height=calc_table_height(len(disp_top_cats), max_h=300),
                    )

                st.markdown("---")
                st.write("##### 📂 Deep-Dive Expanders for Top 10 Issue Categories")

                for idx, item in enumerate(cat_details):
                    cat_name = item["category"]
                    cat_tot = item["total_tickets"]
                    share_pct = item["share_pct"]
                    sub_df = item["sub_df"]
                    raw_slice = item["raw_slice"]

                    with st.expander(
                        f"⚡ #{idx + 1}: {cat_name} — {cat_tot:,} Tickets ({share_pct}% of total)",
                        expanded=False,
                    ):
                        col_sub_c1, col_sub_c2 = st.columns([1.1, 1.4])

                        with col_sub_c1:
                            if not sub_df.empty:
                                chart_sub = (
                                    alt.Chart(sub_df)
                                    .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#F87171")
                                    .encode(
                                        x=alt.X("Occurrences:Q", title="Occurrences"),
                                        y=alt.Y("Issue Subcategory:N", sort="-x", title="Subcategory"),
                                        tooltip=[
                                            alt.Tooltip("Issue Subcategory:N", title="Subcategory"),
                                            alt.Tooltip("Occurrences:Q", title="Occurrences"),
                                            alt.Tooltip("Sub-Type Share %:Q", title="Category Share (%)", format=".1f"),
                                            alt.Tooltip("Unique Chargers:Q", title="Unique Chargers"),
                                        ],
                                    )
                                    .properties(height=max(180, len(sub_df) * 38))
                                )
                                st.altair_chart(chart_sub, width="stretch")

                        with col_sub_c2:
                            st.write(f"**Top 5 Sub-Types Table for '{cat_name}':**")
                            disp_sub = _insert_row_number_column(sub_df)
                            st.dataframe(
                                disp_sub,
                                width="stretch",
                                hide_index=True,
                                height=calc_table_height(len(disp_sub), max_h=240),
                            )
            else:
                st.info("No issue category data available matching current filters.")

        with rep_subtab2:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1.05rem; color: #0F172A; margin-bottom: 2px;">Top 10 Station Hotspots &amp; Site-Wise Repetitive Issues</div>
                <div style="font-size: 0.84rem; color: #64748B; margin-bottom: 12px;">Charging station locations experiencing the highest volume of recurring incidents. Expand any station to inspect its recurring subcategories, chargers, and field engineers.</div>
                """,
                unsafe_allow_html=True,
            )

            top_st_df, st_details = compute_station_repetitive_issues(filtered_df, top_stations=10, top_subs_per_station=5)

            if not top_st_df.empty:
                col_st1, col_st2 = st.columns([1.1, 1.0])
                with col_st1:
                    chart_st = (
                        alt.Chart(top_st_df)
                        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#3B82F6")
                        .encode(
                            x=alt.X("Total Tickets:Q", title="Total Incidents"),
                            y=alt.Y("Station:N", sort="-x", title="Station Location"),
                            tooltip=[
                                alt.Tooltip("Station:N", title="Station"),
                                alt.Tooltip("Total Tickets:Q", title="Total Incidents"),
                                alt.Tooltip("Unique Chargers:Q", title="Unique Chargers"),
                                alt.Tooltip("Top Failure Mode:N", title="Top Mode"),
                            ],
                        )
                        .properties(height=300)
                    )
                    st.altair_chart(chart_st, width="stretch")

                with col_st2:
                    st.write("##### Summary Table: Top 10 Station Hotspots")
                    disp_top_st = _insert_row_number_column(top_st_df)
                    st.dataframe(
                        disp_top_st,
                        width="stretch",
                        hide_index=True,
                        height=calc_table_height(len(disp_top_st), max_h=300),
                    )

                st.markdown("---")
                st.write("##### 📂 Station Hotspots & Recurring Issues Expanders")

                for idx, item in enumerate(st_details):
                    st_name = item["station"]
                    st_tot = item["total_tickets"]
                    uniq_chargers = item["unique_chargers"]
                    sub_df = item["sub_df"]
                    raw_slice = item["raw_slice"]

                    with st.expander(
                        f"📍 #{idx + 1}: {st_name} — {st_tot:,} Tickets ({uniq_chargers} Chargers)",
                        expanded=False,
                    ):
                        st.write(f"**Recurring Failure Subcategories at '{st_name}':**")
                        disp_st_sub = _insert_row_number_column(sub_df)
                        st.dataframe(
                            disp_st_sub,
                            width="stretch",
                            hide_index=True,
                            height=calc_table_height(len(disp_st_sub), max_h=220),
                        )
            else:
                st.info("No station location data available matching current filters.")

        with rep_subtab3:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1.05rem; color: #0F172A; margin-bottom: 2px;">Network-Wide Recurring Failure Modes (Occurrences &gt; 1)</div>
                <div style="font-size: 0.84rem; color: #64748B; margin-bottom: 12px;">All paired Issue Category &amp; Subcategory combinations occurring repeatedly across chargers and stations.</div>
                """,
                unsafe_allow_html=True,
            )

            modes_df = rep_summary.get("modes_df", pd.DataFrame())
            if not modes_df.empty:
                disp_modes = _insert_row_number_column(modes_df)
                st.dataframe(
                    disp_modes,
                    width="stretch",
                    hide_index=True,
                    height=calc_table_height(len(disp_modes), max_h=460),
                )
            else:
                st.info("No recurring fault modes detected (all failure modes occurred only once in the filtered dataset).")

    with tab3:
        st.subheader("Field Engineer Performance & Workload Governance")
        st.caption("Operational scorecard, turnaround time, and SLA adherence breakdown by assigned maintenance engineer (Assigned To).")

        eng_perf_df = compute_engineer_performance(filtered_df)

        if not eng_perf_df.empty:
            total_engineers = len(eng_perf_df)
            top_eng_row = eng_perf_df.iloc[0]
            top_eng_name = str(top_eng_row["Engineer"])
            top_eng_tickets = int(top_eng_row["Total"])

            net_assigned = int(eng_perf_df["Total"].sum())
            net_resolved = int(eng_perf_df["Resolved"].sum())
            net_breached = int(eng_perf_df["SLA_Breached"].sum())

            net_res_rate = round(net_resolved / net_assigned * 100, 1) if net_assigned else 0.0
            net_sla_rate = round(max(0, net_assigned - net_breached) / net_assigned * 100, 1) if net_assigned else 0.0

            ek1, ek2, ek3, ek4 = st.columns(4)

            with ek1:
                st.markdown(
                    f"""
                    <div class="kpi-card card-neutral">
                        <div class="kpi-title">Active Engineers</div>
                        <div class="kpi-value" style="color: #1D4ED8;">{total_engineers:,}</div>
                        <div class="kpi-subtitle">Assigned field engineers</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with ek2:
                st.markdown(
                    f"""
                    <div class="kpi-card card-danger">
                        <div class="kpi-title">Top Workload Engineer</div>
                        <div class="kpi-value" style="color: #DC2626; font-size: 1.35rem; line-height: 1.25;">{html.escape(top_eng_name[:28])}</div>
                        <div class="kpi-subtitle">{top_eng_tickets:,} Assigned Tickets</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with ek3:
                st.markdown(
                    f"""
                    <div class="kpi-card card-success">
                        <div class="kpi-title">Network Resolution %</div>
                        <div class="kpi-value" style="color: #059669;">{net_res_rate:.1f}%</div>
                        <div class="kpi-subtitle">{net_resolved:,} of {net_assigned:,} resolved</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with ek4:
                st.markdown(
                    f"""
                    <div class="kpi-card card-warning">
                        <div class="kpi-title">Network SLA Adherence %</div>
                        <div class="kpi-value" style="color: #D97706;">{net_sla_rate:.1f}%</div>
                        <div class="kpi-subtitle">{net_breached:,} SLA breaches recorded</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            top10_eng = eng_perf_df.head(10).copy()
            status_melt = pd.melt(
                top10_eng,
                id_vars=["Engineer"],
                value_vars=[c for c in ["Resolved", "In_Progress", "Pending_Approval", "On_Hold", "Assigned"] if c in top10_eng.columns],
                var_name="Status Category",
                value_name="Tickets",
            )
            status_melt["Status Category"] = status_melt["Status Category"].replace(
                {
                    "In_Progress": "In Progress",
                    "Pending_Approval": "Pending Approval",
                    "On_Hold": "On Hold",
                }
            )

            st.write("##### Engineer Workload & Status Breakdown (Top 10 by Tickets)")
            chart_eng = (
                alt.Chart(status_melt)
                .mark_bar(cornerRadiusTopRight=2, cornerRadiusBottomRight=2)
                .encode(
                    x=alt.X("Tickets:Q", title="Tickets"),
                    y=alt.Y("Engineer:N", sort="-x", title="Field Engineer (Assigned To)"),
                    color=alt.Color(
                        "Status Category:N",
                        title="Status",
                        scale=alt.Scale(
                            domain=["Resolved", "In Progress", "Pending Approval", "On Hold", "Assigned"],
                            range=["#10B981", "#3B82F6", "#F59E0B", "#64748B", "#EF4444"],
                        ),
                    ),
                    tooltip=[
                        alt.Tooltip("Engineer:N", title="Engineer"),
                        alt.Tooltip("Status Category:N", title="Status"),
                        alt.Tooltip("Tickets:Q", title="Tickets"),
                    ],
                )
                .properties(height=320)
            )
            st.altair_chart(chart_eng, width="stretch")

            st.write("##### Field Engineering Scorecard Table")
            scorecard_cols = [
                "Engineer",
                "Total",
                "Resolved",
                "In_Progress",
                "Pending_Approval",
                "On_Hold",
                "Assigned",
                "SLA_Breached",
                "Resolution Rate %",
                "SLA Adherence %",
                "Avg_Res_Days",
                "Avg_Hold_Days",
            ]
            clean_scorecard = eng_perf_df[[c for c in scorecard_cols if c in eng_perf_df.columns]].copy()
            rename_map = {
                "Total": "Total Assigned",
                "In_Progress": "In Progress",
                "Pending_Approval": "Pending Approval",
                "On_Hold": "On Hold",
                "SLA_Breached": "SLA Breached",
                "Avg_Res_Days": "Avg Res (Days)",
                "Avg_Hold_Days": "Avg Hold (Days)",
            }
            clean_scorecard.rename(columns=rename_map, inplace=True)
            disp_scorecard = _insert_row_number_column(clean_scorecard)

            st.dataframe(
                disp_scorecard,
                width="stretch",
                hide_index=True,
                height=calc_table_height(len(disp_scorecard), max_h=420),
            )

            st.markdown("---")
            st.write("##### 🔍 Field Engineer Ticket Inspector")
            eng_options = sorted(list(eng_perf_df["Engineer"].unique()))
            if eng_options:
                col_eng_sel, col_eng_chk = st.columns([1.5, 1.0])
                with col_eng_sel:
                    selected_eng = st.selectbox(
                        "Select Field Engineer:",
                        options=eng_options,
                        index=0,
                        key="oms_dash_eng_picker",
                    )
                with col_eng_chk:
                    inspect_checked = st.checkbox(
                        "Show Assigned Tickets Table",
                        value=False,
                        key="oms_dash_eng_inspect_chk",
                    )

                if inspect_checked:
                    eng_slice = filtered_df[
                        filtered_df["Assigned To"].fillna("Unassigned").astype(str).str.strip().replace("", "Unassigned")
                        == selected_eng
                    ]
                    st.caption(f"Showing **{len(eng_slice):,}** tickets assigned to **{selected_eng}**:")
                    eng_disp_cols = [
                        c
                        for c in [
                            "Ticket ID",
                            "Station",
                            "Charger",
                            "Issue Category",
                            "Issue Subcategory",
                            "Status",
                            "Priority",
                            "SLA Breached",
                            "Created At",
                            "Assigned At",
                            "Resolved At Time",
                            RESOLVED_TIME_DATE_ONLY_COL,
                            "Hold total time(Days)",
                        ]
                        if c in eng_slice.columns
                    ]
                    disp_eng_tickets = _insert_row_number_column(_sanitize_for_dataframe(eng_slice[eng_disp_cols]))
                    st.dataframe(
                        disp_eng_tickets,
                        width="stretch",
                        hide_index=True,
                        height=calc_table_height(len(disp_eng_tickets), max_h=320),
                    )
        else:
            st.info("No engineer assignment data ('Assigned To' column) available matching current filters.")

    with tab4:
        st.subheader("Operational Ticket Tracker")
        st.caption("Key operational fields and calculated turnaround/hold-day deltas for matching OMS tickets.")

        tracker_cols = [c for c in _TRACKER_PRIORITY_COLUMNS if c in filtered_df.columns]
        if not tracker_cols:
            tracker_cols = list(filtered_df.columns)
        tracker_df = _insert_row_number_column(_sanitize_for_dataframe(filtered_df[tracker_cols]))

        st.dataframe(
            tracker_df,
            width="stretch",
            hide_index=True,
            height=calc_table_height(len(tracker_df)),
        )

    with tab5:
        st.subheader("Converted date_only Schema")
        col_p1, col_p2, col_p3 = st.columns([2.6, 1.2, 1.2])
        with col_p1:
            st.caption(
                f"Full {len(filtered_df.columns)}-column converted date_only sheet "
                f"(showing {len(filtered_df):,} of {len(base_df):,} rows; 'today' = {today_str})."
            )
        tab5_csv_payload = export_oms_date_only_csv(filtered_df) if is_filtered else csv_payload
        tab5_xlsx_payload = export_oms_date_only_excel(filtered_df) if is_filtered else xlsx_payload
        with col_p2:
            st.download_button(
                label="Download Filtered CSV" if is_filtered else "Download CSV",
                data=tab5_csv_payload,
                file_name=f"{file_stem}_date_only{'_filtered' if is_filtered else ''}.csv",
                mime="text/csv",
                key="oms_dash_dl_csv_preview",
                width="stretch",
            )
        with col_p3:
            st.download_button(
                label="Download Filtered Excel" if is_filtered else "Download Excel",
                data=tab5_xlsx_payload,
                file_name=f"{file_stem}_date_only{'_filtered' if is_filtered else ''}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="oms_dash_dl_xlsx_preview",
                width="stretch",
            )

        schema_df = _insert_row_number_column(_sanitize_for_dataframe(filtered_df))
        st.dataframe(
            schema_df,
            width="stretch",
            hide_index=True,
            height=calc_table_height(len(schema_df)),
        )
