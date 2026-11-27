import streamlit as st

# Backward-compatible re-exports for external modules & test runners
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

from views.styles import SHARED_CSS, inject_shared_styles, calc_table_height
from views.compare_view import (
    SAMPLE_FIR_PATH,
    SAMPLE_OMS_PATH,
    init_compare_session_state,
    render_compare_sidebar,
    render_landing_view,
    render_audit_view,
    render_compare_main,
)
from views.oms_dashboard import (
    init_oms_dashboard_state,
    render_oms_dashboard_sidebar,
    render_oms_landing_view,
    render_oms_dashboard_main,
)

__all__ = [
    "CATEGORY_GAP",
    "CATEGORY_PHANTOM",
    "CATEGORY_ACTIVE",
    "CATEGORY_HEALTHY",
    "ReconciliationResult",
    "ReconciliationSummary",
    "load_dataset",
    "reconcile_incident_records",
    "create_filtered_result",
    "generate_reconciliation_excel",
    "export_reconciliation_csv",
    "SHARED_CSS",
    "inject_shared_styles",
    "calc_table_height",
    "SAMPLE_FIR_PATH",
    "SAMPLE_OMS_PATH",
    "init_compare_session_state",
    "render_compare_sidebar",
    "render_landing_view",
    "render_audit_view",
    "render_compare_main",
    "init_oms_dashboard_state",
    "render_oms_dashboard_sidebar",
    "render_oms_landing_view",
    "render_oms_dashboard_main",
    "init_session_state",
    "main",
]


def init_session_state() -> None:
    init_oms_dashboard_state()
    init_compare_session_state()


def main() -> None:
    st.set_page_config(
        page_title="EV Charger Ticket Checker",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_shared_styles()
    init_session_state()

    with st.sidebar:
        active_view = st.radio(
            "Navigation",
            options=["OMS Dashboard", "OMS-FIR Compare"],
            index=0,
            key="nav_view",
        )
        st.markdown("---")
        if active_view == "OMS Dashboard":
            render_oms_dashboard_sidebar()
        else:
            render_compare_sidebar()

    if active_view == "OMS Dashboard":
        render_oms_dashboard_main()
    else:
        render_compare_main()


if __name__ == "__main__":
    main()
