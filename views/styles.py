from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st

SHARED_CSS: str = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #1E293B;
    }

    .header-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 24px 28px;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
    .header-box h1 {
        margin: 0 0 6px 0;
        font-size: 1.75rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.02em;
    }
    .header-box p {
        margin: 0;
        font-size: 0.95rem;
        color: #64748B;
    }

    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px 18px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .kpi-title {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #64748B;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 2.1rem;
        font-weight: 800;
        line-height: 1.1;
        margin-bottom: 6px;
        font-family: 'Inter', sans-serif;
    }
    .kpi-subtitle {
        font-size: 0.8rem;
        color: #64748B;
        font-weight: 500;
    }

    .card-danger {
        border-top: 4px solid #EF4444;
    }
    .card-warning {
        border-top: 4px solid #F59E0B;
    }
    .card-success {
        border-top: 4px solid #10B981;
    }
    .card-neutral {
        border-top: 4px solid #3B82F6;
    }
    .card-gray {
        border-top: 4px solid #64748B;
    }

    .step-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 18px 20px;
        margin-bottom: 12px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .step-number {
        font-size: 0.85rem;
        font-weight: 800;
        color: #C8102E;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 4px;
    }
    .step-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 6px;
    }
    .step-desc {
        font-size: 0.88rem;
        color: #64748B;
        line-height: 1.45;
    }

    .notice-box {
        padding: 14px 18px;
        border-radius: 8px;
        margin-bottom: 16px;
        font-size: 0.92rem;
        line-height: 1.45;
    }
    .notice-danger {
        background: #FEF2F2;
        border: 1px solid #FECACA;
        color: #991B1B;
    }
    .notice-warning {
        background: #FFFBEB;
        border: 1px solid #FDE68A;
        color: #92400E;
    }
    .notice-success {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        color: #166534;
    }

    .filter-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px 14px 20px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
    .filter-title {
        font-size: 0.96rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .filter-tag {
        display: inline-block;
        background: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #BFDBFE;
        border-radius: 9999px;
        padding: 3px 10px;
        font-size: 0.78rem;
        font-weight: 600;
        margin-right: 6px;
        margin-bottom: 4px;
    }
    .filter-summary-strip {
        font-size: 0.88rem;
        color: #475569;
        margin-top: 10px;
        padding-top: 8px;
        border-top: 1px solid #F1F5F9;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        border-radius: 8px 8px 0px 0px;
        font-weight: 600;
        font-size: 0.92rem;
        padding: 8px 16px;
    }
    .stTabs [aria-selected="true"] {
        color: #C8102E !important;
        border-bottom: 3px solid #C8102E !important;
    }
    </style>
    """


def inject_shared_styles() -> None:
    st.markdown(SHARED_CSS, unsafe_allow_html=True)


def calc_table_height(row_count: int, max_h: int = 560) -> int:
    if row_count <= 0:
        return 140
    return min(max(38 + 20 + (row_count * 36) + 25, 180), max_h)


def _coerce_display_cell(val: Any) -> str:
    if val is None or isinstance(val, bool):
        return "" if val is None else str(val)
    try:
        if bool(pd.isna(val)):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(val, (pd.Timestamp, datetime, date)):
        return val.strftime("%Y-%m-%d")
    return str(val)


def sanitize_for_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        series = out[col]
        if (
            pd.api.types.is_integer_dtype(series.dtype)
            or pd.api.types.is_float_dtype(series.dtype)
            or pd.api.types.is_bool_dtype(series.dtype)
            or pd.api.types.is_datetime64_any_dtype(series.dtype)
        ):
            continue
        if isinstance(series.dtype, pd.StringDtype):
            out[col] = series.fillna("")
            continue
        if series.dtype == object:
            out[col] = pd.Series(
                [_coerce_display_cell(v) for v in series.tolist()],
                index=out.index,
                dtype="string",
            )
    return out
