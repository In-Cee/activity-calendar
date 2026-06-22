# ============================================================
#  Activity Calendar — Streamlit App · v0.2
#  Mastercard Foundation · Enterprise Planning
# ============================================================

import os
import streamlit as st
import pandas as pd
import plotly.express as px
from lookups import (TYPES, FUNCTIONS, SUB_FUNCTIONS, ATTENDEE_CATEGORIES,
                     DELIVERY_MODE, STATUS, WEIGHTING,
                     THRESHOLD_ELEVATED, THRESHOLD_CRITICAL,
                     FOUNDATION_ORANGE, FOUNDATION_BG, FOUNDATION_TEXT,
                     FOUNDATION_AMBER, FOUNDATION_RED, FOUNDATION_GREEN,
                     TYPE_COLORS, FUNCTION_COLORS, COUNTRY_TO_REGION)

# ============================================================
#  Page setup + Foundation theme
# ============================================================
st.set_page_config(page_title="Activity Calendar", page_icon="📅", layout="wide")

# Inject Foundation styling
st.markdown(f"""
<style>
    .stApp {{ background-color: {FOUNDATION_BG}; }}
    h1, h2, h3 {{ color: {FOUNDATION_TEXT}; }}
    .insight-strip {{
        background: #FFF4E6;
        border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 10px 16px;
        margin: 8px 0 20px 0;
        font-style: italic;
        color: #555;
        border-radius: 4px;
    }}
    .kpi-card {{
        background: white; padding: 16px; border-radius: 8px;
        border-top: 4px solid {FOUNDATION_ORANGE};
    }}
    div[data-baseweb="tag"] {{ background-color: {FOUNDATION_ORANGE} !important; }}
</style>
""", unsafe_allow_html=True)

# ---- Header ----
col_title, col_toggle = st.columns([3, 1])
with col_title:
    st.markdown(
        f"<h1 style='color:{FOUNDATION_ORANGE}; margin-bottom:0;'>Activity Calendar</h1>"
        f"<p style='color:#666; margin-top:0;'>Mastercard Foundation · Enterprise Planning</p>",
        unsafe_allow_html=True
    )
with col_toggle:
    st.write("")  # spacer
    view_mode = st.radio("View mode", ["Executive", "Analyst"],
                         horizontal=True, label_visibility="collapsed")

# ============================================================
#  Data loading (cache-aware)
# ============================================================
@st.cache_data
def load_data(file_mtime):
    df = pd.read_excel("activities.xlsx", sheet_name="Activities", header=1)
    df["StartDate"] = pd.to_datetime(df["StartDate"])
    df["EndDate"]   = pd.to_datetime(df["EndDate"])

    # Fill in workflow fields if not present
    if "Status" not in df.columns:
        df["Status"] = "Approved"
    if "Weighting" not in df.columns:
        df["Weighting"] = "Medium"

    # Derive country from Location
    df["Country"] = df["Location"].astype(str).str.split(", ").str[-1].str.strip()
    df["Region"]  = df["Country"].map(COUNTRY_TO_REGION).fillna("Other")

    return df

file_mtime = os.path.getmtime("activities.xlsx")
if "activities" not in st.session_state:
    st.session_state.activities = load_data(file_mtime)

df = st.session_state.activities

# ============================================================
#  Sidebar — full filter set
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
    st.caption(f"v0.2 · {len(df)} activities loaded · {df['Country'].nunique()} countries")

# ============================================================
#  Apply filters
# ============================================================
mask = (
    df["Type"].isin(f_type)
    & df["Initiating Function"].isin(f_hosting)
    & df["Status"].isin(f_status)
    & df["Weighting"].isin(f_weighting)
    & df["Internal/External"].isin(f_delivery)
)

if f_country:
    mask &= df["Country"].isin(f_country)
if f_region:
    mask &= df["Region"].isin(f_region)
if f_participating:
    # Match if any chosen participating function appears in the comma-separated cell
    pattern = "|".join([f"\\b{p}\\b" for p in f_participating])
    mask &= df["Participating Function"].astype(str).str.contains(pattern, regex=True, na=False)

if len(f_dates) == 2:
    start, end = f_dates
    mask &= (df["StartDate"].dt.date >= start) & (df["StartDate"].dt.date <= end)

view = df[mask].copy()

# ---- Helper: render insight strip ----
def insight(text):
    st.markdown(f"<div class='insight-strip'>💡 {text}</div>", unsafe_allow_html=True)

# ============================================================
#  Tabs
# ============================================================
if view_mode == "Executive":
    tab_dashboard, tab_calendar, tab_heatmap, tab_gantt, tab_location, tab_submit, tab_upload = st.tabs(
        ["📈 Dashboard", "📅 Calendar", "🔥 Heatmap", "📊 Gantt",
         "📍 Location", "➕ Submit", "📤 Mass Upload"]
    )
else:
    tab_calendar, tab_heatmap, tab_gantt, tab_location, tab_dashboard, tab_submit, tab_upload = st.tabs(
        ["📅 Calendar", "🔥 Heatmap", "📊 Gantt", "📍 Location",
         "📈 Dashboard", "➕ Submit", "📤 Mass Upload"]
    )

# ============================================================
#  Tab: Executive Dashboard
# ============================================================
with tab_dashboard:
    st.subheader("Executive Dashboard")

    # ---- KPIs ----
    c1, c2, c3, c4, c5 = st.columns(5)
    total = len(view)
    approved = int((view["Status"] == "Approved").sum())
    pending  = int((view["Status"] == "Pending").sum())
    high_w   = int((view["Weighting"] == "High").sum())
    countries = view["Country"].nunique()

    c1.metric("Total activities", total)
    c2.metric("Approved", approved, f"{approved/max(total,1)*100:.0f}%")
    c3.metric("Pending", pending, f"{pending/max(total,1)*100:.0f}%")
    c4.metric("High weighting", high_w)
    c5.metric("Countries", countries)

    if total > 0:
        top_func = view["Initiating Function"].value_counts().idxmax()
        top_func_n = view["Initiating Function"].value_counts().max()
        peak_month = view.assign(M=view["StartDate"].dt.to_period("M").astype(str))["M"].value_counts().idxmax()
        insight(
            f"As of today, {total} activities are planned across {countries} countries. "
            f"{approved} are approved ({approved/total*100:.0f}%), {pending} pending. "
            f"{top_func} is the most active function with {top_func_n} activities. "
            f"Peak month: {peak_month}."
        )
    else:
        insight("No activities match your filters. Adjust the sidebar to see results.")

    # ---- Chart: When is the Foundation busiest? ----
    st.markdown("#### When is the Foundation busiest?")
    if not view.empty:
        v = view.copy()
        v["Month"] = v["StartDate"].dt.to_period("M").astype(str)
        by_month = v.groupby(["Month", "Type"]).size().reset_index(name="Count")
        fig = px.bar(by_month, x="Month", y="Count", color="Type", barmode="stack",
                     color_discrete_map=TYPE_COLORS)
        fig.update_layout(plot_bgcolor="white", height=380,
                          margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    # ---- Chart: Are we focused on the right work? ----
    st.markdown("#### Are we focused on the right work?")
    if not view.empty:
        by_w = view["Weighting"].value_counts().reset_index()
        by_w.columns = ["Weighting", "Count"]
        fig = px.pie(by_w, values="Count", names="Weighting", hole=0.5,
                     color_discrete_sequence=[FOUNDATION_ORANGE, FOUNDATION_AMBER, FOUNDATION_GREEN])
        fig.update_layout(height=320, margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
        high_pct = (view["Weighting"] == "High").mean() * 100
        insight(f"{high_pct:.0f}% of activities are High weighting. "
                f"{'Healthy strategic focus.' if high_pct >= 30 else 'Consider lifting more activities to High to sharpen focus.'}")

# ============================================================
#  Tab: Calendar (list view)
# ============================================================
with tab_calendar:
    st.subheader("Calendar")
    st.caption(f"{len(view)} activities · {view['Country'].nunique()} countries")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        insight(f"Showing {len(view)} activities. Sort columns by clicking the headers. "
                f"Use the sidebar filters to narrow further.")
        st.dataframe(
            view[["StartDate", "EndDate", "Type", "Title", "Country",
                  "Initiating Function", "Initiating Sub-Function",
                  "Attendee Category", "Participating Function",
                  "Internal/External", "Status", "Weighting", "Note"]],
            use_container_width=True, hide_index=True
        )

# ============================================================
#  Tab: Heatmap
# ============================================================
with tab_heatmap:
    st.subheader("Where is workload pressure building?")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        v = view.copy()
        v["Week"] = v["StartDate"].dt.to_period("W").apply(lambda p: p.start_time)
        heat = v.groupby(["Week", "Initiating Function"]).size().reset_index(name="Count")
        pivot = heat.pivot(index="Week", columns="Initiating Function", values="Count").fillna(0)

        # Find pressure points
        max_count = int(pivot.max().max())
        if max_count >= THRESHOLD_CRITICAL:
            insight(f"⚠️ Critical pressure detected: at least one function exceeds {THRESHOLD_CRITICAL} activities in a single week. Action needed.")
        elif max_count >= THRESHOLD_ELEVATED:
            insight(f"Elevated pressure detected: at least one function reaches {THRESHOLD_ELEVATED}+ activities in a week. Monitor closely.")
        else:
            insight(f"Workload is balanced — no function exceeds the comfort threshold of {THRESHOLD_ELEVATED} activities/week.")

        fig = px.imshow(pivot.T, color_continuous_scale="Oranges",
                        labels=dict(x="Week", y="Function", color="Activities"),
                        aspect="auto")
        fig.update_layout(height=420, margin=dict(l=20, r=20, t=20, b=20),
                          plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
        st.caption(f"Comfort thresholds — Elevated ≥ {THRESHOLD_ELEVATED}, Critical ≥ {THRESHOLD_CRITICAL}")

# ============================================================
#  Tab: Gantt
# ============================================================
with tab_gantt:
    st.subheader("Activity timeline (Gantt)")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        insight(f"Showing {len(view)} activities across {(view['EndDate'].max() - view['StartDate'].min()).days} days. Hover for details.")
        fig = px.timeline(view, x_start="StartDate", x_end="EndDate", y="Title",
                          color="Initiating Function",
                          color_discrete_map=FUNCTION_COLORS,
                          hover_data=["Country", "Type", "Weighting", "Status", "Note"])
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=max(500, len(view) * 18), plot_bgcolor="white", margin=dict(l=20, r=20, t=20, b=20))
