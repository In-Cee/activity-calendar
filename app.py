# ============================================================
#  Activity Calendar - Streamlit App - v0.3
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
from explainers import KPI_EXPLAINERS, CHART_EXPLAINERS, chart_explainer_markdown

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
    .detail-card {{
        background: white;
        border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 16px 20px;
        border-radius: 6px;
        margin: 8px 0;
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
#  Helper functions
# ============================================================
def insight(text):
    """Render an italic orange-bordered insight strip."""
    st.markdown(f"<div class='insight-strip'>💡 {text}</div>", unsafe_allow_html=True)

def chart_title_with_explainer(key, level="###"):
    """Render a chart title with an ⓘ explainer popover next to it."""
    e = CHART_EXPLAINERS.get(key, {})
    title = e.get("title", key)
    col_t, col_i = st.columns([20, 1])
    with col_t:
        st.markdown(f"{level} {title}")
    with col_i:
        with st.popover("ⓘ", use_container_width=True):
            st.markdown(chart_explainer_markdown(key))

def kpi_with_tooltip(col, label, value, delta=None):
    """Render a KPI metric with an info tooltip."""
    help_text = KPI_EXPLAINERS.get(label, "")
    col.metric(label, value, delta, help=help_text)

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

    df["Country"] = df["Location"].astype(str).str.split(", ").str[-1].str.strip()
    df["Region"]  = df["Country"].map(COUNTRY_TO_REGION).fillna("Other")

    return df

file_mtime = os.path.getmtime("activities.xlsx")
if "activities" not in st.session_state:
    st.session_state.activities = load_data(file_mtime)

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
    st.caption(f"v0.3 - {len(df)} activities loaded - {df['Country'].nunique()} countries")

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

        kpi_with_tooltip(c1, "Total activities", total)
        kpi_with_tooltip(c2, "Approved", approved, f"{approved/max(total,1)*100:.0f}%")
        kpi_with_tooltip(c3, "Pending", pending, f"{pending/max(total,1)*100:.0f}%")
        kpi_with_tooltip(c4, "High weighting", high_w)
        kpi_with_tooltip(c5, "Countries", countries)

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

        chart_title_with_explainer("busiest", "####")
        if not view.empty:
            v = view.copy()
            v["Month"] = v["StartDate"].dt.to_period("M").astype(str)
            by_month = v.groupby(["Month", "Type"]).size().reset_index(name="Count")
            fig = px.bar(by_month, x="Month", y="Count", color="Type", barmode="stack",
                         color_discrete_map=TYPE_COLORS)
            fig.update_layout(plot_bgcolor="white", height=380,
                              margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)

        chart_title_with_explainer("priority", "####")
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
#  Tab: Calendar (with click-row drill-down)
# ============================================================
with tab_calendar:
    chart_title_with_explainer("calendar", "###")
    try:
        st.caption(f"{len(view)} activities - {view['Country'].nunique()} countries")
        if view.empty:
            st.info("No activities match your filters.")
        else:
            insight(f"Showing {len(view)} activities. Sort columns by clicking the headers. "
                    f"Pick an activity below to see its full details.")

            # Searchable, sortable table
            display = view[["StartDate", "EndDate", "Type", "Title", "Country",
                            "Initiating Function", "Initiating Sub-Function",
                            "Attendee Category", "Participating Function",
                            "Internal/External", "Status", "Weighting", "Note"]].reset_index(drop=True)
            st.dataframe(display, use_container_width=True, hide_index=True)

            # ---- Click-row drill-down ----
            st.markdown("##### 🔍 Activity drill-down")
            options = view.assign(
                _label=view["Title"].astype(str) + " - " + view["Country"].astype(str)
                        + " - " + view["StartDate"].dt.strftime("%b %d, %Y")
            )["_label"].tolist()
            pick = st.selectbox("Pick an activity to see full details", options=[""] + options)

            if pick:
                row = view.assign(
                    _label=view["Title"].astype(str) + " - " + view["Country"].astype(str)
                            + " - " + view["StartDate"].dt.strftime("%b %d, %Y")
                ).query("_label == @pick").iloc[0]

                d1, d2 = st.columns([2, 1])
                with d1:
                    st.markdown(f"""
<div class='detail-card'>
<h4 style='color:{FOUNDATION_ORANGE}; margin-top:0;'>{row['Title']}</h4>
<p style='color:#666;'><strong>{row['Type']}</strong> - {row['Country']} - {row['StartDate'].strftime('%b %d')} to {row['EndDate'].strftime('%b %d, %Y')}</p>
<p><strong>Hosting:</strong> {row['Initiating Function']} ({row['Initiating Sub-Function']})<br>
<strong>Participating:</strong> {row['Participating Function'] if str(row['Participating Function']) != 'nan' else 'None'}<br>
<strong>Attendees:</strong> {row['Attendee Category']}<br>
<strong>Mode:</strong> {row['Internal/External']}<br>
<strong>Status:</strong> {row['Status']} - <strong>Weighting:</strong> {row['Weighting']}</p>
<p style='color:#444;'><em>{row['Note'] if str(row['Note']) != 'nan' else ''}</em></p>
</div>
""", unsafe_allow_html=True)

                with d2:
                    # Find nearby activities (within 7 days)
                    start = row["StartDate"]
                    end = row["EndDate"]
                    nearby = view[
                        (view["StartDate"] >= start - pd.Timedelta(days=7))
                        & (view["StartDate"] <= end + pd.Timedelta(days=7))
                        & (view["Title"] != row["Title"])
                    ]
                    st.markdown(f"##### Nearby activities (±7 days)")
                    if nearby.empty:
                        st.caption("No overlapping activities within ±7 days.")
                    else:
                        st.caption(f"{len(nearby)} other activities within ±7 days")
                        for _, n in nearby.head(8).iterrows():
                            st.markdown(
                                f"- **{n['Title']}** - {n['Country']} - {n['StartDate'].strftime('%b %d')} ({n['Initiating Function']})"
                            )
    except Exception as e:
        st.error(f"Calendar could not render: {e}")

# ============================================================
#  Tab: Heatmap (with cell drill-down)
# ============================================================
with tab_heatmap:
    chart_title_with_explainer("pressure", "###")
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

            # ---- Cell drill-down ----
            st.markdown("##### 🔍 Pressure point drill-down")
            d1, d2 = st.columns(2)
            with d1:
                pick_func = st.selectbox("Function", FUNCTIONS, key="hm_func")
            with d2:
                week_options = sorted(v["Week"].unique())
                week_labels = {w: pd.to_datetime(w).strftime("Week of %b %d, %Y") for w in week_options}
                pick_week = st.selectbox("Week", options=week_options,
                                          format_func=lambda w: week_labels[w], key="hm_week")

            cell_rows = v[(v["Initiating Function"] == pick_func) & (v["Week"] == pick_week)]
            if cell_rows.empty:
                st.caption(f"No {pick_func} activities in {week_labels[pick_week]}.")
            else:
                st.caption(f"{len(cell_rows)} {pick_func} activities in {week_labels[pick_week]}")
                st.dataframe(
                    cell_rows[["Title", "Country", "StartDate", "EndDate", "Type",
                               "Attendee Category", "Status", "Weighting"]],
                    use_container_width=True, hide_index=True
                )
    except Exception as e:
        st.error(f"Heatmap could not render: {e}")

# ============================================================
#  Tab: Gantt
# ============================================================
with tab_gantt:
    chart_title_with_explainer("gantt", "###")
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
    chart_title_with_explainer("location", "###")
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
