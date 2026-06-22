# ============================================================
#  Activity Calendar - Streamlit App - v0.2.1
#  Mastercard Foundation - Enterprise Planning
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
    div[data-baseweb="tag"] {{ background-color: {FOUNDATION_ORANGE} !important; }}
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
#  Data loading (cache-aware)
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

    df["Country"] = df["Location"].astype(str).str.split(", ").str[-1].str.strip()
    df["Region"]  = df["Country"].map(COUNTRY_TO_REGION).fillna("Other")

    return df

file_mtime = os.path.getmtime("activities.xlsx")
if "activities" not in st.session_state:
    st.session_state.activities = load_data(file_mtime)

df = st.session_state.activities

# ============================================================
#  Sidebar - full filter set
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
    st.caption(f"v0.2.1 - {len(df)} activities loaded - {df['Country'].nunique()} countries")

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
    pattern = "|".join([f"\\b{p}\\b" for p in f_participating])
    mask &= df["Participating Function"].astype(str).str.contains(pattern, regex=True, na=False)

if isinstance(f_dates, tuple) and len(f_dates) == 2:
    start, end = f_dates
    mask &= (df["StartDate"].dt.date >= start) & (df["StartDate"].dt.date <= end)

view = df[mask].copy()

# ---- Helper: insight strip ----
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
    try:
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
            top_func_n = int(view["Initiating Function"].value_counts().max())
            peak_month = view.assign(M=view["StartDate"].dt.to_period("M").astype(str))["M"].value_counts().idxmax()
            insight(
                f"As of today, {total} activities are planned across {countries} countries. "
                f"{approved} are approved ({approved/total*100:.0f}%), {pending} pending. "
                f"{top_func} is the most active function with {top_func_n} activities. "
                f"Peak month: {peak_month}."
            )
        else:
            insight("No activities match your filters. Adjust the sidebar to see results.")

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
    except Exception as e:
        st.error(f"Dashboard could not render: {e}")

# ============================================================
#  Tab: Calendar (list view)
# ============================================================
with tab_calendar:
    st.subheader("Calendar")
    try:
        st.caption(f"{len(view)} activities - {view['Country'].nunique()} countries")
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
    except Exception as e:
        st.error(f"Calendar could not render: {e}")

# ============================================================
#  Tab: Heatmap
# ============================================================
with tab_heatmap:
    st.subheader("Where is workload pressure building?")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            v = view.copy()
            v["Week"] = v["StartDate"].dt.to_period("W").apply(lambda p: p.start_time)
            heat = v.groupby(["Week", "Initiating Function"]).size().reset_index(name="Count")
            pivot = heat.pivot(index="Week", columns="Initiating Function", values="Count").fillna(0)

            max_count = int(pivot.max().max()) if not pivot.empty else 0
            if max_count >= THRESHOLD_CRITICAL:
                insight(f"⚠️ Critical pressure detected: at least one function exceeds {THRESHOLD_CRITICAL} activities in a single week. Action needed.")
            elif max_count >= THRESHOLD_ELEVATED:
                insight(f"Elevated pressure detected: at least one function reaches {THRESHOLD_ELEVATED}+ activities in a week. Monitor closely.")
            else:
                insight(f"Workload is balanced - no function exceeds the comfort threshold of {THRESHOLD_ELEVATED} activities/week.")

            fig = px.imshow(pivot.T, color_continuous_scale="Oranges",
                            labels=dict(x="Week", y="Function", color="Activities"),
                            aspect="auto")
            fig.update_layout(height=420, margin=dict(l=20, r=20, t=20, b=20),
                              plot_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
            st.caption(f"Comfort thresholds - Elevated >= {THRESHOLD_ELEVATED}, Critical >= {THRESHOLD_CRITICAL}")
    except Exception as e:
        st.error(f"Heatmap could not render: {e}")

# ============================================================
#  Tab: Gantt
# ============================================================
with tab_gantt:
    st.subheader("Activity timeline (Gantt)")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            MAX_ROWS = 80
            g = view.sort_values("StartDate").head(MAX_ROWS).copy()
            g["Label"] = (g["Title"].astype(str) + " - " + g["Country"].astype(str)
                          + " - " + g["StartDate"].dt.strftime("%b %d"))

            if len(view) > MAX_ROWS:
                insight(f"Showing the first {MAX_ROWS} of {len(view)} activities (sorted by date). Use sidebar filters to narrow further.")
            else:
                insight(f"Showing {len(g)} activities across {(g['EndDate'].max() - g['StartDate'].min()).days} days. Hover for details.")

            fig = px.timeline(g, x_start="StartDate", x_end="EndDate", y="Label",
                              color="Initiating Function",
                              color_discrete_map=FUNCTION_COLORS,
                              hover_data=["Country", "Type", "Weighting", "Status", "Note"])
            fig.update_yaxes(autorange="reversed")
            fig.update_layout(height=max(500, len(g) * 22), plot_bgcolor="white",
                              margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Gantt could not render: {e}")

# ============================================================
#  Tab: Location
# ============================================================
with tab_location:
    st.subheader("Where in our markets is activity concentrated?")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            by_region = (view.groupby("Region").size().reset_index(name="Activities")
                            .sort_values("Activities", ascending=False))
            top_country = view["Country"].value_counts().idxmax()
            top_country_n = int(view["Country"].value_counts().max())

            if not by_region.empty:
                top_region = by_region.iloc[0]["Region"]
                top_region_n = int(by_region.iloc[0]["Activities"])
                insight(f"{top_country} leads activity volume ({top_country_n} activities). "
                        f"{top_region} region carries {top_region_n/len(view)*100:.0f}% of total load.")
            else:
                insight(f"{top_country} leads activity volume ({top_country_n} activities).")

            st.markdown("##### By region")
            st.dataframe(by_region, use_container_width=True, hide_index=True)

            st.markdown("##### By country")
            by_country = (view.groupby("Country").size().reset_index(name="Activities")
                             .sort_values("Activities", ascending=True))
            fig = px.bar(by_country, x="Activities", y="Country", orientation="h",
                         color_discrete_sequence=[FOUNDATION_ORANGE])
            fig.update_layout(height=max(400, len(by_country)*22), plot_bgcolor="white",
                              margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Location could not render: {e}")

# ============================================================
#  Tab: Submit
# ============================================================
with tab_submit:
    st.subheader("Submit a new activity")
    try:
        with st.form("submit", clear_on_submit=True):
            col1, col2 = st.columns(2)
            title    = col1.text_input("Activity name")
            a_type   = col2.selectbox("Type", TYPES)
            start    = col1.date_input("Start date")
            end      = col2.date_input("End date")
            location = col1.text_input("Location (City, Country)")
            func     = col2.selectbox("Initiating Function", FUNCTIONS)
            subfunc  = col1.selectbox("Initiating Sub-Function", SUB_FUNCTIONS[func])
            attendee = col2.selectbox("Attendee Category", ATTENDEE_CATEGORIES)
            delivery = col1.selectbox("Internal/External", DELIVERY_MODE)
            weighting = col2.selectbox("Weighting", WEIGHTING, index=1)
            note     = st.text_area("Note", max_chars=250)

            if st.form_submit_button("✅ Submit", type="primary"):
                country = location.split(", ")[-1].strip() if ", " in location else ""
                region = COUNTRY_TO_REGION.get(country, "Other")
                new = pd.DataFrame([{
                    "StartDate": pd.to_datetime(start), "EndDate": pd.to_datetime(end),
                    "Type": a_type, "Title": title, "Location": location,
                    "Country": country, "Region": region,
                    "Initiating Function": func, "Initiating Sub-Function": subfunc,
                    "Attendee Category": attendee, "Participating Function": "",
                    "Participating Sub-Function": "", "Internal/External": delivery,
                    "Month (Auto/Manual)": pd.to_datetime(start).strftime("%b"),
                    "Year (Auto/Manual)": pd.to_datetime(start).year,
                    "Note": note, "Status": "Pending", "Weighting": weighting
                }])
                st.session_state.activities = pd.concat(
                    [st.session_state.activities, new], ignore_index=True)
                st.success(f"✅ '{title}' submitted as Pending with {weighting} weighting.")
    except Exception as e:
        st.error(f"Submit form could not render: {e}")

# ============================================================
#  Tab: Mass Upload
# ============================================================
with tab_upload:
    st.subheader("Mass upload")
    try:
        st.write("Upload an Excel file matching the template. Rows are validated against project lookups.")
        upload = st.file_uploader("Choose Excel file", type=["xlsx"])
        if upload:
            new_df = pd.read_excel(upload, sheet_name="Activities", header=1)
            valid, errors = [], []
            for i, row in new_df.iterrows():
                if row["Type"] not in TYPES:
                    errors.append(f"Row {i+3}: invalid Type '{row['Type']}'")
                elif row["Initiating Function"] not in FUNCTIONS:
                    errors.append(f"Row {i+3}: invalid Function '{row['Initiating Function']}'")
                elif row["Initiating Sub-Function"] not in SUB_FUNCTIONS.get(row["Initiating Function"], []):
                    errors.append(f"Row {i+3}: '{row['Initiating Sub-Function']}' not allowed under {row['Initiating Function']}")
                else:
                    valid.append(row)

            if errors:
                st.error(f"Parsed {len(new_df)} rows: {len(valid)} valid, {len(errors)} with errors.")
                with st.expander("See errors"):
                    for e in errors:
                        st.write("• " + e)
            else:
                st.success(f"✅ All {len(new_df)} rows valid.")
                if st.button("Load into calendar"):
                    new_df["Country"] = new_df["Location"].astype(str).str.split(", ").str[-1].str.strip()
                    new_df["Region"] = new_df["Country"].map(COUNTRY_TO_REGION).fillna("Other")
                    if "Status" not in new_df.columns:
                        new_df["Status"] = "Pending"
                    if "Weighting" not in new_df.columns:
                        new_df["Weighting"] = "Medium"
                    st.session_state.activities = pd.concat(
                        [st.session_state.activities, new_df], ignore_index=True)
                    st.success(f"Loaded {len(new_df)} activities into the calendar.")
    except Exception as e:
        st.error(f"Mass Upload could not render: {e}")
