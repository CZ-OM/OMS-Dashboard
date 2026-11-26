import html
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

from core.models import (
    CATEGORY_GAP,
    CATEGORY_PHANTOM,
    CATEGORY_ACTIVE,
    CATEGORY_HEALTHY,
    ReconciliationResult,
    ReconciliationSummary,
)
from core.loader import load_dataset
from core.reconciler import reconcile_incident_records, create_filtered_result
from core.exporter import generate_reconciliation_excel, export_reconciliation_csv
from views.styles import calc_table_height, sanitize_for_dataframe

SAMPLE_FIR_PATH = None
SAMPLE_OMS_PATH = None


def init_compare_session_state() -> None:
    if "reconciliation_result" not in st.session_state:
        st.session_state["reconciliation_result"] = None
    if "fir_df" not in st.session_state:
        st.session_state["fir_df"] = None
    if "oms_df" not in st.session_state:
        st.session_state["oms_df"] = None
    if "dataset_source" not in st.session_state:
        st.session_state["dataset_source"] = None
    if "fir_uploader_sig" not in st.session_state:
        st.session_state["fir_uploader_sig"] = None
    if "oms_uploader_sig" not in st.session_state:
        st.session_state["oms_uploader_sig"] = None


def _upload_signature(uploaded):
    if uploaded is None:
        return None
    return (getattr(uploaded, "file_id", None), uploaded.name, uploaded.size)


def _clean_compare_options(df: pd.DataFrame, col_name: str) -> list[str]:
    if col_name not in df.columns:
        return []
    vals = df[col_name].dropna().astype(str).str.strip()
    return sorted({v for v in vals if v != ""})


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


def render_compare_sidebar() -> None:
    st.title("Charger Checker")
    st.caption("Check broken chargers against system tickets.")

    st.markdown("---")
    st.subheader("Step 1: Upload Files")

    uploaded_fir = st.file_uploader(
        "1. Field Report (FIR)",
        type=["xlsx", "xls", "csv"],
        key="fir_uploader",
        help="Upload the morning Excel or CSV file that lists broken or down chargers.",
    )

    uploaded_oms = st.file_uploader(
        "2. Helpdesk Tickets (OMS)",
        type=["xlsx", "xls", "csv"],
        key="oms_uploader",
        help="Upload the ticket dump file from your complaint or work order system.",
    )

    clear_btn = st.button("Clear All", width="stretch")

    if clear_btn:
        st.session_state["reconciliation_result"] = None
        st.session_state["fir_df"] = None
        st.session_state["oms_df"] = None
        st.session_state["dataset_source"] = None
        st.session_state["fir_uploader_sig"] = None
        st.session_state["oms_uploader_sig"] = None
        for k in ["fir_uploader", "oms_uploader"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    if uploaded_fir is None:
        st.session_state["fir_uploader_sig"] = None
    if uploaded_oms is None:
        st.session_state["oms_uploader_sig"] = None

    if uploaded_fir is not None:
        fir_sig = _upload_signature(uploaded_fir)
        if fir_sig != st.session_state["fir_uploader_sig"]:
            try:
                st.session_state["fir_df"] = load_dataset(uploaded_fir, filename=uploaded_fir.name)
                st.session_state["dataset_source"] = "Your Uploaded Files"
                st.session_state["fir_uploader_sig"] = fir_sig
            except Exception as e:
                st.error(f"Could not read Field Report: {e}")

    if uploaded_oms is not None:
        oms_sig = _upload_signature(uploaded_oms)
        if oms_sig != st.session_state["oms_uploader_sig"]:
            try:
                st.session_state["oms_df"] = load_dataset(uploaded_oms, filename=uploaded_oms.name)
                st.session_state["dataset_source"] = "Your Uploaded Files"
                st.session_state["oms_uploader_sig"] = oms_sig
            except Exception as e:
                st.error(f"Could not read Helpdesk Tickets file: {e}")

    st.markdown("---")
    st.subheader("Step 2: Check & Compare")

    can_run = st.session_state["fir_df"] is not None and st.session_state["oms_df"] is not None
    if not can_run:
        st.info("Upload both files above to start.")

    run_audit = st.button(
        "Compare Both Files",
        type="primary",
        disabled=not can_run,
        width="stretch",
    )

    if run_audit and can_run:
        with st.spinner("Comparing chargers with tickets..."):
            audit_result = reconcile_incident_records(
                st.session_state["fir_df"],
                st.session_state["oms_df"],
            )
            st.session_state["reconciliation_result"] = audit_result
            st.success("Comparison complete.")

    if st.session_state["reconciliation_result"] is not None:
        st.markdown("---")
        st.subheader("Comparison Ready")
        st.caption("Live dynamic filters and search are active above the dashboard tables.")


def render_landing_view() -> None:
    st.markdown(
        """
        <div class="header-box">
            <h1>EV Charger Complaint & Status Checker</h1>
            <p>Easily check which broken chargers on the ground are missing a ticket, and which tickets were closed too early.</p>
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
                <div class="step-title">Upload Today's Files</div>
                <div class="step-desc">
                    In the left sidebar, add your <b>Field Report (FIR)</b> and your <b>Helpdesk Tickets (OMS)</b>.
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
                <div class="step-title">Click Compare</div>
                <div class="step-desc">
                    Click the red <b>Compare Both Files</b> button. The system will match every charger with its latest ticket status in seconds.
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
                <div class="step-title">See What Needs Fixing</div>
                <div class="step-desc">
                    View clear lists of missing tickets, wrongly closed tickets, and download an Excel report to share with your team.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("File Upload Status")
    col_f1, col_f2 = st.columns(2)

    with col_f1:
        if st.session_state["fir_df"] is not None:
            st.success(f"Field Report ready: {len(st.session_state['fir_df'])} chargers found.")
        else:
            st.info("Field Report (FIR): Not uploaded yet. Add it on the left.")

    with col_f2:
        if st.session_state["oms_df"] is not None:
            st.success(f"Helpdesk Tickets ready: {len(st.session_state['oms_df'])} tickets found.")
        else:
            st.info("Helpdesk Tickets (OMS): Not uploaded yet. Add it on the left.")


def render_audit_view(result: ReconciliationResult) -> None:
    summary: ReconciliationSummary = result.summary

    friendly_column_config = {
        "No.": st.column_config.NumberColumn("No.", width="small"),
        "Report_Date": st.column_config.TextColumn("Report Date"),
        "Issue_Age": st.column_config.TextColumn("Duration Down"),
        "Latest_Ticket_ID": st.column_config.TextColumn("Latest Ticket ID"),
        "Latest_Ticket_Status": st.column_config.TextColumn("Ticket Status"),
        "Latest_Ticket_Created_At": st.column_config.TextColumn("Ticket Created Date"),
        "Latest_Ticket_Closed_At": st.column_config.TextColumn("Ticket Closed Date"),
        "Latest_Ticket_Updated_At": st.column_config.TextColumn("Ticket Last Updated"),
        "Latest_Ticket_SLA_Deadline": st.column_config.TextColumn("SLA Target Date"),
        "Latest_Ticket_Ack_Deadline": st.column_config.TextColumn("Acknowledgement Deadline"),
        "Latest_Ticket_Category": st.column_config.TextColumn("Ticket Category"),
        "Latest_Ticket_Subcategory": st.column_config.TextColumn("Ticket Subcategory"),
        "Latest_Ticket_Assigned_To": st.column_config.TextColumn("Assigned Technician"),
        "Latest_Ticket_SLA_Breached": st.column_config.TextColumn("SLA Overdue"),
        "Audit_Category": st.column_config.TextColumn("Audit Status"),
        "Risk_Level": st.column_config.TextColumn("Priority"),
        "Prescribed_Action": st.column_config.TextColumn("Action Needed"),
        "OEM Ticket Required": st.column_config.TextColumn("Vendor Ticket Needed"),
        "Priority_Label": st.column_config.TextColumn("Priority"),
        "Error Info": st.column_config.TextColumn("Error Details"),
        "Created At": st.column_config.TextColumn("Ticket Created Date"),
        "Updated At": st.column_config.TextColumn("Ticket Last Updated"),
        "Closed At": st.column_config.TextColumn("Ticket Closed Date"),
        "Resolution SLA Deadline": st.column_config.TextColumn("Resolution Target Date"),
        "Acknowledgement SLA Deadline": st.column_config.TextColumn("Acknowledgement Deadline"),
        "SLA Breached": st.column_config.TextColumn("SLA Overdue"),
    }

    col_title, col_btn = st.columns([3, 1])
    with col_title:
        st.markdown(
            f"""
            <div style="margin-bottom: 8px;">
                <h1 style="font-size: 1.75rem; font-weight: 800; color: #0F172A; margin: 0 0 4px 0;">Charger & Ticket Check Results</h1>
                <p style="margin: 0; color: #64748B; font-size: 0.92rem;">
                    Source: {html.escape(str(st.session_state.get('dataset_source') or ''))} &nbsp;|&nbsp; Checked at: {html.escape(str(summary.timestamp))}
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    base_df = result.reconciled_df.copy()
    total_count = len(base_df)
    gap_count = len(base_df[base_df["Audit_Category"] == CATEGORY_GAP])
    phantom_count = len(base_df[base_df["Audit_Category"] == CATEGORY_PHANTOM])
    active_count = len(base_df[base_df["Audit_Category"] == CATEGORY_ACTIVE])
    healthy_count = len(base_df[base_df["Audit_Category"] == CATEGORY_HEALTHY])
    critical_count = len(base_df[base_df["Risk_Level"] == "Critical"])

    st.markdown(
        """
        <div class="filter-card">
            <div class="filter-title">Dynamic Filters & Live Search</div>
            <div style="font-size: 0.85rem; color: #64748B; margin-bottom: 8px;">
                Select any quick preset pill, type a search keyword, or use the dropdowns to instantly filter results in real time.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pill_options = [
        f"All Chargers ({total_count})",
        f"Missing Tickets ({gap_count})",
        f"Wrongly Closed ({phantom_count})",
        f"Being Worked On ({active_count})",
        f"Working Fine ({healthy_count})",
        f"Critical Priority ({critical_count})",
    ]
    selected_pill = st.pills(
        "Quick Presets",
        options=pill_options,
        default=pill_options[0],
        key="dyn_quick_pill",
        label_visibility="collapsed",
    )

    col_s1, col_s2 = st.columns([5, 1])
    with col_s1:
        search_val = st.text_input(
            "Live Search",
            placeholder="Search across Charger ID, Station Name, City, Technician, Error, or Ticket ID...",
            key="dyn_search_input",
            label_visibility="collapsed",
        ).strip()
    with col_s2:
        reset_clicked = st.button("Reset Filters", width="stretch")
        if reset_clicked:
            for k in ["dyn_quick_pill", "dyn_search_input", "dyn_zones", "dyn_tenants", "dyn_makes", "dyn_hw_status"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        all_zones = _clean_compare_options(base_df, "Zone")
        selected_zones = st.multiselect("Region / City", all_zones, key="dyn_zones", placeholder="All Regions")

    if selected_zones and "Zone" in base_df.columns:
        df_zone_sub = base_df[base_df["Zone"].fillna("").astype(str).str.strip().isin(selected_zones)]
    else:
        df_zone_sub = base_df
    avail_tenants = _clean_compare_options(df_zone_sub, "Tenant Name")
    if "dyn_tenants" in st.session_state:
        st.session_state["dyn_tenants"] = [t for t in st.session_state["dyn_tenants"] if t in avail_tenants]

    with col_f2:
        selected_tenants = st.multiselect("Client / Fleet", avail_tenants, key="dyn_tenants", placeholder="All Clients")

    if selected_tenants and "Tenant Name" in df_zone_sub.columns:
        df_tenant_sub = df_zone_sub[df_zone_sub["Tenant Name"].fillna("").astype(str).str.strip().isin(selected_tenants)]
    else:
        df_tenant_sub = df_zone_sub
    avail_makes = _clean_compare_options(df_tenant_sub, "Charger Make")
    if "dyn_makes" in st.session_state:
        st.session_state["dyn_makes"] = [m for m in st.session_state["dyn_makes"] if m in avail_makes]

    with col_f3:
        selected_makes = st.multiselect("Charger Brand", avail_makes, key="dyn_makes", placeholder="All Brands")

    avail_statuses = _clean_compare_options(df_tenant_sub, "Status")
    if "dyn_hw_status" in st.session_state:
        st.session_state["dyn_hw_status"] = [s for s in st.session_state["dyn_hw_status"] if s in avail_statuses]

    with col_f4:
        selected_statuses = st.multiselect("Hardware Status", avail_statuses, key="dyn_hw_status", placeholder="All Statuses")

    df_all = base_df.copy()

    if selected_pill and "Missing Tickets" in selected_pill:
        df_all = df_all[df_all["Audit_Category"] == CATEGORY_GAP]
    elif selected_pill and "Wrongly Closed" in selected_pill:
        df_all = df_all[df_all["Audit_Category"] == CATEGORY_PHANTOM]
    elif selected_pill and "Being Worked On" in selected_pill:
        df_all = df_all[df_all["Audit_Category"] == CATEGORY_ACTIVE]
    elif selected_pill and "Working Fine" in selected_pill:
        df_all = df_all[df_all["Audit_Category"] == CATEGORY_HEALTHY]
    elif selected_pill and "Critical Priority" in selected_pill:
        df_all = df_all[df_all["Risk_Level"] == "Critical"]

    for col_name, selected_vals in (
        ("Zone", selected_zones),
        ("Tenant Name", selected_tenants),
        ("Charger Make", selected_makes),
        ("Status", selected_statuses),
    ):
        if selected_vals and col_name in df_all.columns:
            col_clean = df_all[col_name].fillna("").astype(str).str.strip()
            df_all = df_all[col_clean.isin(selected_vals)]

    if search_val:
        q = search_val.lower()
        search_mask = pd.Series(False, index=df_all.index)
        cols_to_search = [
            "OCPP ID", "Station Name", "Zone", "Tenant Name",
            "Charger Make", "Error Info", "ZME Name", "Latest_Ticket_ID", "Status"
        ]
        for col in cols_to_search:
            if col in df_all.columns:
                search_mask |= df_all[col].astype(str).str.lower().str.contains(q, regex=False, na=False)
        df_all = df_all[search_mask]

    active_filters = []
    if selected_pill and not selected_pill.startswith("All"):
        active_filters.append(f"Preset: {selected_pill.split('(')[0].strip()}")
    if selected_zones:
        active_filters.append(f"Region: {', '.join(selected_zones)}")
    if selected_tenants:
        active_filters.append(f"Client: {', '.join(selected_tenants)}")
    if selected_makes:
        active_filters.append(f"Brand: {', '.join(selected_makes)}")
    if selected_statuses:
        active_filters.append(f"Status: {', '.join(selected_statuses)}")
    if search_val:
        active_filters.append(f"Search: '{search_val}'")

    is_filtered = len(active_filters) > 0 or len(df_all) < total_count

    if is_filtered:
        tags_html = " ".join([f'<span class="filter-tag">{html.escape(f)}</span>' for f in active_filters])
        st.markdown(
            f"""
            <div class="filter-summary-strip">
                <b>Active Filters:</b> {tags_html} &nbsp;&bull;&nbsp;
                <b>Showing {len(df_all)} of {total_count} chargers</b> ({round(len(df_all)/total_count*100, 1) if total_count else 0}% of total).
                {total_count - len(df_all)} chargers filtered out.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div class="filter-summary-strip">
                Showing all <b>{total_count} chargers</b>. Use the preset pills or dropdowns above to filter dynamically.
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_btn:
        active_excel_result = create_filtered_result(result, df_all)
        excel_bytes = generate_reconciliation_excel(active_excel_result)
        btn_label = "Download Filtered Excel" if is_filtered else "Download Excel Report"
        st.download_button(
            label=btn_label,
            data=excel_bytes,
            file_name="Charger_Ticket_Audit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    filtered_gaps = df_all[df_all["Audit_Category"] == CATEGORY_GAP]
    filtered_phantoms = df_all[df_all["Audit_Category"] == CATEGORY_PHANTOM]
    filtered_active = df_all[df_all["Audit_Category"] == CATEGORY_ACTIVE]
    filtered_healthy = df_all[df_all["Audit_Category"] == CATEGORY_HEALTHY]

    cur_total = len(df_all)
    cur_gaps = len(filtered_gaps)
    cur_phantoms = len(filtered_phantoms)
    cur_active = len(filtered_active)
    cur_down = cur_gaps + cur_phantoms + cur_active
    cur_healthy = len(filtered_healthy)
    cur_gap_rate = round(((cur_gaps + cur_phantoms) / cur_down * 100), 1) if cur_down > 0 else 0.0

    st.markdown("### Quick Numbers")
    k1, k2, k3, k4, k5, k6 = st.columns(6)

    with k1:
        sub_txt = f"of {summary.total_fir} total (filtered)" if is_filtered else "Chargers in report"
        st.markdown(
            f"""
            <div class="kpi-card card-neutral">
                <div class="kpi-title">Total Checked</div>
                <div class="kpi-value" style="color: #1E293B;">{cur_total}</div>
                <div class="kpi-subtitle">{sub_txt}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k2:
        st.markdown(
            f"""
            <div class="kpi-card card-gray">
                <div class="kpi-title">Broken / Down</div>
                <div class="kpi-value" style="color: #475569;">{cur_down}</div>
                <div class="kpi-subtitle">{cur_healthy} working fine</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k3:
        sub_txt = f"of {summary.ticket_gaps} total (filtered)" if is_filtered else "Need ticket right away!"
        st.markdown(
            f"""
            <div class="kpi-card card-danger">
                <div class="kpi-title">Missing Tickets</div>
                <div class="kpi-value" style="color: #DC2626;">{cur_gaps}</div>
                <div class="kpi-subtitle">{sub_txt}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k4:
        sub_txt = f"of {summary.phantom_resolutions} total (filtered)" if is_filtered else "Ticket closed, still down!"
        st.markdown(
            f"""
            <div class="kpi-card card-warning">
                <div class="kpi-title">Wrongly Closed</div>
                <div class="kpi-value" style="color: #D97706;">{cur_phantoms}</div>
                <div class="kpi-subtitle">{sub_txt}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k5:
        sub_txt = f"of {summary.active_verified} total (filtered)" if is_filtered else "Open ticket in system"
        st.markdown(
            f"""
            <div class="kpi-card card-success">
                <div class="kpi-title">Being Worked On</div>
                <div class="kpi-value" style="color: #059669;">{cur_active}</div>
                <div class="kpi-subtitle">{sub_txt}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with k6:
        gap_color = "#DC2626" if cur_gap_rate > 15 else "#059669"
        st.markdown(
            f"""
            <div class="kpi-card card-danger">
                <div class="kpi-title">Missing Ticket Rate</div>
                <div class="kpi-value" style="color: {gap_color};">{cur_gap_rate}%</div>
                <div class="kpi-subtitle">Broken chargers missing tickets</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    untracked_df = result.untracked_oms_df
    untracked_display = untracked_df.copy()
    if "Zone" in untracked_display.columns and selected_zones:
        untracked_display = untracked_display[
            untracked_display["Zone"].fillna("").astype(str).str.strip().isin(selected_zones)
        ]
    if search_val:
        q = search_val.lower()
        u_mask = pd.Series(False, index=untracked_display.index)
        for c in ["Ticket ID", "Station", "Charger", "Status", "Issue Category", "Assigned To", "Zone"]:
            if c in untracked_display.columns:
                u_mask |= untracked_display[c].astype(str).str.lower().str.contains(q, regex=False, na=False)
        untracked_display = untracked_display[u_mask]

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            f"Missing Tickets ({len(filtered_gaps)})",
            f"Wrongly Closed Tickets ({len(filtered_phantoms)})",
            f"Being Worked On ({len(filtered_active)})",
            "Charts & Trends",
            f"Extra Tickets in OMS ({len(untracked_display)})",
            f"All Chargers List ({len(df_all)})",
        ]
    )

    with tab1:
        st.subheader("Missing Tickets: Chargers Down with No Ticket")
        st.markdown(
            f"""
            <div class="notice-box notice-danger">
                <b>Action Required:</b> There are <b>{len(filtered_gaps)} chargers</b> reported broken or faulted in the field,
                but <b>nobody opened a ticket yet</b> in the helpdesk system. Please create tickets immediately and send a technician.
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_t1, col_t2 = st.columns([3, 1])
        with col_t1:
            if is_filtered:
                st.caption(f"Showing {len(filtered_gaps)} of {summary.ticket_gaps} chargers with missing tickets (filtered):")
            else:
                st.caption(f"Showing all {len(filtered_gaps)} chargers with missing tickets:")
        with col_t2:
            st.download_button(
                "Download List as CSV",
                data=export_reconciliation_csv(filtered_gaps),
                file_name="missing_tickets_chargers_down.csv",
                mime="text/csv",
                width="stretch",
            )

        cols_to_display = [
            "Station Name",
            "OCPP ID",
            "Status",
            "Report_Date",
            "Issue_Age",
            "Tenant Name",
            "Priority_Label",
            "Error Info",
            "ZME Name",
            "ZME Contact Number",
            "Zone",
            "Charger Make",
            "OEM Ticket Required",
            "Prescribed_Action",
        ]
        display_cols = [c for c in cols_to_display if c in filtered_gaps.columns]
        display_df = filtered_gaps[display_cols].copy()
        display_df = _insert_row_number_column(sanitize_for_dataframe(display_df))
        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config=friendly_column_config,
            height=calc_table_height(len(display_df)),
        )

    with tab2:
        st.subheader("Wrongly Closed Tickets: Ticket Closed but Charger Still Down")
        st.markdown(
            f"""
            <div class="notice-box notice-warning">
                <b>Check with Technician:</b> There are <b>{len(filtered_phantoms)} chargers</b> where the ticket was marked
                <b>Closed or Resolved</b>, but the charger is <b>still broken on the ground</b>. Please reopen the ticket or check if repair is finished.
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            if is_filtered:
                st.caption(f"Showing {len(filtered_phantoms)} of {summary.phantom_resolutions} wrongly closed tickets (filtered)")
            else:
                st.caption(f"Showing all {len(filtered_phantoms)} wrongly closed tickets:")
        with col_p2:
            st.download_button(
                "Download List as CSV",
                data=export_reconciliation_csv(filtered_phantoms),
                file_name="wrongly_closed_tickets.csv",
                mime="text/csv",
                width="stretch",
            )

        cols_to_display = [
            "Station Name",
            "OCPP ID",
            "Status",
            "Latest_Ticket_ID",
            "Latest_Ticket_Status",
            "Latest_Ticket_Created_At",
            "Latest_Ticket_Closed_At",
            "Latest_Ticket_Updated_At",
            "Latest_Ticket_SLA_Deadline",
            "Report_Date",
            "Latest_Ticket_Category",
            "Latest_Ticket_Subcategory",
            "Latest_Ticket_Assigned_To",
            "Latest_Ticket_SLA_Breached",
            "Issue_Age",
            "Zone",
            "Prescribed_Action",
        ]
        display_cols = [c for c in cols_to_display if c in filtered_phantoms.columns]
        display_df = filtered_phantoms[display_cols].copy()
        display_df = _insert_row_number_column(sanitize_for_dataframe(display_df))
        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config=friendly_column_config,
            height=calc_table_height(len(display_df)),
        )

    with tab3:
        st.subheader("Being Worked On: Broken Chargers with Open Tickets")
        st.markdown(
            f"""
            <div class="notice-box notice-success">
                <b>Normal:</b> There are <b>{len(filtered_active)} chargers</b> down with an open ticket active in the helpdesk system.
                The team is already tracking these repairs.
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_a1, col_a2 = st.columns([3, 1])
        with col_a1:
            if is_filtered:
                st.caption(f"Showing {len(filtered_active)} of {summary.active_verified} open tickets (filtered)")
            else:
                st.caption(f"Showing all {len(filtered_active)} chargers being worked on:")
        with col_a2:
            st.download_button(
                "Download List as CSV",
                data=export_reconciliation_csv(filtered_active),
                file_name="chargers_being_worked_on.csv",
                mime="text/csv",
                width="stretch",
            )

        cols_to_display = [
            "Station Name",
            "OCPP ID",
            "Status",
            "Latest_Ticket_ID",
            "Latest_Ticket_Status",
            "Latest_Ticket_Created_At",
            "Latest_Ticket_Updated_At",
            "Latest_Ticket_SLA_Deadline",
            "Report_Date",
            "Latest_Ticket_Assigned_To",
            "Latest_Ticket_SLA_Breached",
            "Issue_Age",
            "Risk_Level",
            "Zone",
            "Prescribed_Action",
        ]
        display_cols = [c for c in cols_to_display if c in filtered_active.columns]
        display_df = filtered_active[display_cols].copy()
        display_df = _insert_row_number_column(sanitize_for_dataframe(display_df))
        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config=friendly_column_config,
            height=calc_table_height(len(display_df)),
        )

    with tab4:
        st.subheader("Charts & Trends")
        st.caption("Visual overview of charger health, regional impact, client outages, and top hardware faults.")

        chart_status_map = {
            CATEGORY_GAP: "Missing Ticket",
            CATEGORY_PHANTOM: "Wrongly Closed",
            CATEGORY_ACTIVE: "Being Worked On",
            CATEGORY_HEALTHY: "Working Fine",
        }

        col_ch1, col_ch2 = st.columns(2)

        with col_ch1:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Charger Status by Region</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Breakdown of missing tickets, wrongly closed, and active issues across regions.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Zone" in df_all.columns and "Audit_Category" in df_all.columns and len(df_all) > 0:
                zone_df = df_all.dropna(subset=["Zone"]).copy()
                zone_df["Zone"] = zone_df["Zone"].astype(str).str.strip()
                zone_df["Status_Label"] = zone_df["Audit_Category"].map(chart_status_map).fillna(zone_df["Audit_Category"])
                zone_chart_df = zone_df.groupby(["Zone", "Status_Label"]).size().reset_index(name="Chargers")
                chart_zone = (
                    alt.Chart(zone_chart_df)
                    .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                    .encode(
                        x=alt.X("Zone:N", title="Region", axis=alt.Axis(labelAngle=0)),
                        y=alt.Y("Chargers:Q", title="Number of Chargers"),
                        color=alt.Color(
                            "Status_Label:N",
                            title="Ticket Status",
                            scale=alt.Scale(
                                domain=["Missing Ticket", "Wrongly Closed", "Being Worked On", "Working Fine"],
                                range=["#F87171", "#FBBF24", "#34D399", "#60A5FA"],
                            ),
                        ),
                        tooltip=[
                            alt.Tooltip("Zone:N", title="Region"),
                            alt.Tooltip("Status_Label:N", title="Status"),
                            alt.Tooltip("Chargers:Q", title="Chargers"),
                        ],
                    )
                    .properties(height=280)
                )
                st.altair_chart(chart_zone, width="stretch")
            else:
                st.info("No charger data matching current filters.")

        with col_ch2:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Broken Chargers by Client</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Clients with the highest number of chargers currently offline.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Tenant Name" in df_all.columns:
                broken_df = df_all[df_all["Audit_Category"] != CATEGORY_HEALTHY].dropna(subset=["Tenant Name"]).copy()
                broken_df["Tenant Name"] = broken_df["Tenant Name"].astype(str).str.strip()
                tenant_chart_df = (
                    broken_df
                    .groupby("Tenant Name")
                    .size()
                    .reset_index(name="Broken_Chargers")
                    .sort_values(by="Broken_Chargers", ascending=False)
                    .head(10)
                )
                if len(tenant_chart_df) > 0:
                    chart_tenant = (
                        alt.Chart(tenant_chart_df)
                        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#F87171")
                        .encode(
                            x=alt.X("Broken_Chargers:Q", title="Number of Broken Chargers"),
                            y=alt.Y("Tenant Name:N", sort="-x", title="Client Name"),
                            tooltip=[
                                alt.Tooltip("Tenant Name:N", title="Client"),
                                alt.Tooltip("Broken_Chargers:Q", title="Broken Chargers"),
                            ],
                        )
                        .properties(height=280)
                    )
                    st.altair_chart(chart_tenant, width="stretch")
                else:
                    st.info("No broken chargers found under current filters.")

        col_ch3, col_ch4 = st.columns(2)

        with col_ch3:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Top Reported Faults</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Most frequent hardware and power failure reasons reported from the field.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Error Info" in df_all.columns:
                fault_df = df_all[df_all["Audit_Category"] != CATEGORY_HEALTHY].dropna(subset=["Error Info"]).copy()
                fault_df["Error Info"] = fault_df["Error Info"].astype(str).str.strip()
                error_chart_df = (
                    fault_df
                    .groupby("Error Info")
                    .size()
                    .reset_index(name="Count")
                    .sort_values(by="Count", ascending=False)
                    .head(10)
                )
                if len(error_chart_df) > 0:
                    chart_error = (
                        alt.Chart(error_chart_df)
                        .mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3, color="#818CF8")
                        .encode(
                            x=alt.X("Count:Q", title="Number of Chargers"),
                            y=alt.Y("Error Info:N", sort="-x", title="Reported Fault"),
                            tooltip=[
                                alt.Tooltip("Error Info:N", title="Reported Fault"),
                                alt.Tooltip("Count:Q", title="Affected Chargers"),
                            ],
                        )
                        .properties(height=280)
                    )
                    st.altair_chart(chart_error, width="stretch")
                else:
                    st.info("No faults reported under current filters.")

        with col_ch4:
            st.markdown(
                """
                <div style="font-weight: 700; font-size: 1rem; color: #0F172A; margin-bottom: 2px;">Charger Health by Brand</div>
                <div style="font-size: 0.82rem; color: #64748B; margin-bottom: 8px;">Field status (Faulted, Unavailable, Available) for each charger brand.</div>
                """,
                unsafe_allow_html=True,
            )
            if "Charger Make" in df_all.columns and "Status" in df_all.columns and len(df_all) > 0:
                make_df = df_all.dropna(subset=["Charger Make", "Status"]).copy()
                make_df["Charger Make"] = make_df["Charger Make"].astype(str).str.strip()
                make_df["Status"] = make_df["Status"].astype(str).str.strip()
                make_chart_df = (
                    make_df.groupby(["Charger Make", "Status"])
                    .size()
                    .reset_index(name="Count")
                )
                chart_make = (
                    alt.Chart(make_chart_df)
                    .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
                    .encode(
                        x=alt.X("Charger Make:N", title="Charger Brand", axis=alt.Axis(labelAngle=0)),
                        y=alt.Y("Count:Q", title="Number of Chargers"),
                        color=alt.Color(
                            "Status:N",
                            title="Hardware Status",
                            scale=alt.Scale(
                                domain=["Faulted", "Unavailable", "Available"],
                                range=["#F87171", "#FBBF24", "#34D399"],
                            ),
                        ),
                        tooltip=[
                            alt.Tooltip("Charger Make:N", title="Brand"),
                            alt.Tooltip("Status:N", title="Status"),
                            alt.Tooltip("Count:Q", title="Chargers"),
                        ],
                    )
                    .properties(height=280)
                )
                st.altair_chart(chart_make, width="stretch")
            else:
                st.info("No brand data matching current filters.")

    with tab5:
        st.subheader("Extra Tickets in OMS (Not in Today's Field Report)")

        raw_oms = st.session_state["oms_df"]

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Total Tickets in File", len(raw_oms) if raw_oms is not None else 0)
        with col_m2:
            st.metric("Extra Open Tickets", len(untracked_display))
        with col_m3:
            sla_breach_count = len(untracked_display[untracked_display["SLA Breached"].astype(str).str.upper() == "YES"]) if "SLA Breached" in untracked_display.columns else 0
            st.metric("Overdue Tickets", sla_breach_count)
        with col_m4:
            st.metric("Different Chargers", untracked_display["Charger"].dropna().nunique() if "Charger" in untracked_display.columns else 0)

        st.markdown("---")
        st.markdown(
            f"<b>Why this matters:</b> These <b>{len(untracked_display)} open tickets</b> exist in your ticket system, "
            "but these chargers were not listed as broken in today's morning field report. Check if these are already fixed or from other stations.",
            unsafe_allow_html=True,
        )

        col_u1, col_u2 = st.columns([3, 1])
        with col_u1:
            if len(untracked_display) < len(untracked_df):
                st.caption(f"Showing {min(len(untracked_display), 100)} of {len(untracked_display)} matching extra open tickets (filtered):")
            else:
                st.caption(f"Showing up to 100 of {len(untracked_display)} extra open tickets in OMS:")
        with col_u2:
            st.download_button(
                "Download List as CSV",
                data=export_reconciliation_csv(untracked_display),
                file_name="extra_open_tickets.csv",
                mime="text/csv",
                width="stretch",
            )

        oms_display_cols = [c for c in ["Sno", "Ticket ID", "Station", "Charger", "Status", "Created At", "Updated At", "Closed At", "Resolution SLA Deadline", "Acknowledgement SLA Deadline", "Issue Category", "Issue Subcategory", "Assigned To", "SLA Breached", "Zone"] if c in untracked_display.columns]
        display_df = untracked_display[oms_display_cols].head(100).copy()
        display_df = _insert_row_number_column(sanitize_for_dataframe(display_df))
        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config=friendly_column_config,
            height=calc_table_height(len(display_df)),
        )

    with tab6:
        st.subheader("All Chargers Checked Today")
        st.markdown("Complete list of all chargers from today's field report with their latest ticket status.")

        col_e1, col_e2 = st.columns([3, 1])
        with col_e1:
            if is_filtered:
                st.caption(f"Showing {len(df_all)} of {summary.total_fir} total chargers (filtered)")
            else:
                st.caption(f"Showing all {len(df_all)} chargers:")
        with col_e2:
            st.download_button(
                "Download All as CSV",
                data=export_reconciliation_csv(df_all),
                file_name="all_chargers_checked.csv",
                mime="text/csv",
                width="stretch",
            )

        cols_explorer = [
            "Station Name",
            "OCPP ID",
            "Status",
            "Audit_Category",
            "Risk_Level",
            "Report_Date",
            "Latest_Ticket_ID",
            "Latest_Ticket_Status",
            "Latest_Ticket_Created_At",
            "Latest_Ticket_Closed_At",
            "Latest_Ticket_Updated_At",
            "Latest_Ticket_SLA_Deadline",
            "Latest_Ticket_Assigned_To",
            "Latest_Ticket_SLA_Breached",
            "Tenant Name",
            "Zone",
            "Charger Make",
            "Error Info",
            "Issue_Age",
            "Prescribed_Action",
        ]
        valid_explorer_cols = [c for c in cols_explorer if c in df_all.columns]
        display_df = df_all[valid_explorer_cols].copy()
        display_df = _insert_row_number_column(sanitize_for_dataframe(display_df))
        st.dataframe(
            display_df,
            width="stretch",
            hide_index=True,
            column_config=friendly_column_config,
            height=calc_table_height(len(display_df)),
        )


def render_compare_main() -> None:
    if st.session_state["reconciliation_result"] is None:
        render_landing_view()
    else:
        render_audit_view(st.session_state["reconciliation_result"])
