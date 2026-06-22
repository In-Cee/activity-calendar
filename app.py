# ============================================================
#  Activity Calendar - Streamlit App - v0.6.1
#  Mastercard Foundation - Enterprise Planning
# ============================================================

import os
import calendar as cal
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date, timedelta, datetime
from lookups import (TYPES, FUNCTIONS, SUB_FUNCTIONS, ATTENDEE_CATEGORIES,
                     DELIVERY_MODE, STATUS, WEIGHTING,
                     THRESHOLD_ELEVATED, THRESHOLD_CRITICAL,
                     FOUNDATION_ORANGE, FOUNDATION_BG, FOUNDATION_TEXT,
                     FOUNDATION_AMBER, FOUNDATION_RED, FOUNDATION_GREEN,
                     TYPE_COLORS, FUNCTION_COLORS, COUNTRY_TO_REGION)
from explainers import KPI_EXPLAINERS, CHART_EXPLAINERS, chart_explainer_markdown
from countries import flag
from settings_store import init_settings, get_thresholds, log_change
from exports import build_pptx, build_excel_export, build_template_excel

# ============================================================
#  Page setup + Foundation theme
# ============================================================
st.set_page_config(page_title="Activity Calendar", page_icon="📅", layout="wide")
init_settings()

# ============================================================
#  Optional access gate (shared password)
# ============================================================
ENABLE_LOGIN = True
SHARED_PASSWORD = "foundation2026"

if ENABLE_LOGIN:
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    if not st.session_state.auth_ok:
        st.markdown(f"<h1 style='color:{FOUNDATION_ORANGE};'>Activity Calendar</h1>", unsafe_allow_html=True)
        st.markdown("Mastercard Foundation · Enterprise Planning")
        st.write("")
        pw = st.text_input("Enter access password", type="password")
        if st.button("Sign in", type="primary"):
            if pw == SHARED_PASSWORD:
                st.session_state.auth_ok = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        st.caption("Contact Radintshi Monyobo for access.")
        st.stop()

# ============================================================
#  Foundation theme
# ============================================================
st.markdown(f"""
<style>
    .stApp {{ background-color: {FOUNDATION_BG}; }}
    h1, h2, h3 {{ color: {FOUNDATION_TEXT}; }}
    .insight-strip {{
        background: #FFF4E6; border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 10px 16px; margin: 8px 0 20px 0; font-style: italic;
        color: #555; border-radius: 4px;
    }}
    div[data-baseweb="tag"] {{ background-color: {FOUNDATION_ORANGE} !important; }}
    .detail-card {{
        background: white; border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 16px 20px; border-radius: 6px; margin: 8px 0;
    }}
    .region-chip {{
        display: inline-block; background: white;
        border: 1px solid #E0E0E0; border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 10px 16px; margin: 4px 8px 4px 0;
        border-radius: 4px; font-size: 14px;
    }}
    .cal-cell {{
        background: white; border: 1px solid #E8E8E8;
        border-radius: 4px; padding: 6px; min-height: 90px; font-size: 11px;
    }}
    .cal-day-num {{ font-weight: bold; color: {FOUNDATION_TEXT}; font-size: 12px; }}
    .cal-today {{ background: #FFF4E6; border: 2px solid {FOUNDATION_ORANGE}; }}
    .cal-weekend {{ background: #F5F5F0; }}
    .cal-outside {{ background: #FAFAFA; opacity: 0.5; }}
    .cal-activity {{
        display: block; margin-top: 2px; padding: 2px 4px;
        border-radius: 3px; font-size: 10px; color: white;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }}
    .exec-hero {{
        background: white; border-top: 6px solid {FOUNDATION_ORANGE};
        padding: 24px 28px; border-radius: 6px; margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .exec-hero h2 {{ color: {FOUNDATION_ORANGE}; margin: 0 0 8px 0; }}
    .exec-hero p {{ font-size: 15px; color: #444; line-height: 1.55; margin: 0; }}
    @media print {{
        .stSidebar, [data-testid="stSidebar"], .stTabs {{ display: none !important; }}
    }}
</style>
""", unsafe_allow_html=True)

# ---- Header ----
col_title, col_toggle = st.columns([3, 1])
with col_title:
    st.markdown(
        f"<h1 style='color:{FOUNDATION_ORANGE}; margin-bottom:0;'>Activity Calendar</h1>"
        f"<p style='color:#666; margin-top:0;'>Mastercard Foundation - Enterprise Planning</p>",
        unsafe_allow_html=True
    )
with col_toggle:
    st.write("")
    view_mode = st.radio("View mode", ["Executive", "Analyst"],
                         horizontal=True, label_visibility="collapsed")

# ============================================================
#  Helpers
# ============================================================
def insight(text):
    st.markdown(f"<div class='insight-strip'>💡 {text}</div>", unsafe_allow_html=True)

def chart_title_with_explainer(key, level="###"):
    e = CHART_EXPLAINERS.get(key, {})
    title = e.get("title", key)
    col_t, col_i = st.columns([20, 1])
    with col_t:
        st.markdown(f"{level} {title}")
    with col_i:
        with st.popover("ⓘ", use_container_width=True):
            st.markdown(chart_explainer_markdown(key))

def kpi_with_tooltip(col, label, value, delta=None):
    help_text = KPI_EXPLAINERS.get(label, "")
    col.metric(label, value, delta, help=help_text)

def normalize_status(df):
    """Standardise Status column so case/whitespace doesn't hide rows."""
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip().str.title()
        df.loc[df["Status"].isin(["", "Nan", "None"]), "Status"] = "Approved"
    return df

# ============================================================
#  Data loading
# ============================================================
@st.cache_data
def load_data(file_mtime):
    df = pd.read_excel("activities.xlsx", sheet_name="Activities", header=1)
    df["StartDate"] = pd.to_datetime(df["StartDate"])
    df["EndDate"]   = pd.to_datetime(df["EndDate"])
    if "Status" not in df.columns:
        df["Status"] = "Approved"
    if "Weighting" not in df.columns:
        df["Weighting"] = "Medium"
    df = normalize_status(df)
    df["Country"] = df["Location"].astype(str).str.split(", ").str[-1].str.strip()
    df["Region"]  = df["Country"].map(COUNTRY_TO_REGION).fillna("Other")
    return df

file_mtime = os.path.getmtime("activities.xlsx")
if "activities" not in st.session_state:
    st.session_state.activities = load_data(file_mtime)

# Always normalize Status on every rerun (covers newly submitted rows)
st.session_state.activities = normalize_status(st.session_state.activities)
df = st.session_state.activities

# ============================================================
#  Sidebar filters
# ============================================================
with st.sidebar:
    st.markdown("### 🔍 Filters")
    f_type = st.multiselect("Type", TYPES, default=TYPES)
    f_hosting = st.multiselect("Hosting function", FUNCTIONS, default=FUNCTIONS)
    f_participating = st.multiselect(
        "Participating function", FUNCTIONS,
        help="Filter activities where any of the chosen functions are participating."
    )
    f_country = st.multiselect("Country", sorted(df["Country"].dropna().unique()))
    f_region = st.multiselect("Region", sorted(df["Region"].dropna().unique()))
    f_status = st.multiselect("Status", STATUS, default=STATUS)
    f_weighting = st.multiselect("Weighting", WEIGHTING, default=WEIGHTING)
    f_delivery = st.multiselect("Internal/External", DELIVERY_MODE, default=DELIVERY_MODE)

    min_d, max_d = df["StartDate"].min().date(), df["EndDate"].max().date()
    f_dates = st.date_input("Date range", value=(min_d, max_d),
                            min_value=min_d, max_value=max_d)

    st.markdown("---")
    st.caption(f"v0.6.1 - {len(df)} activities loaded - {df['Country'].nunique()} countries")
    if ENABLE_LOGIN:
        if st.button("Sign out"):
            st.session_state.auth_ok = False
            st.rerun()

# ============================================================
