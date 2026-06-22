# ============================================================
#  Activity Calendar — Streamlit App
#  Mastercard Foundation · Enterprise Planning
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from lookups import (TYPES, FUNCTIONS, SUB_FUNCTIONS, ATTENDEE_CATEGORIES,
                     DELIVERY_MODE, THRESHOLD_ELEVATED, THRESHOLD_CRITICAL,
                     FOUNDATION_ORANGE)

# ---- Page setup ----
st.set_page_config(page_title="Activity Calendar", page_icon="📅", layout="wide")

st.markdown(
    f"<h1 style='color:{FOUNDATION_ORANGE}; margin-bottom:0;'>Activity Calendar</h1>"
    f"<p style='color:#666; margin-top:0;'>Mastercard Foundation · Enterprise Planning</p>",
    unsafe_allow_html=True
)

# ---- Load data ----
@st.cache_data
def load_data():
    df = pd.read_excel("activities.xlsx", sheet_name="Activities", header=1)
    df["StartDate"] = pd.to_datetime(df["StartDate"])
    df["EndDate"]   = pd.to_datetime(df["EndDate"])
    df["Status"]    = "Approved"
    df["Weighting"] = "Medium"
    return df

if "activities" not in st.session_state:
    st.session_state.activities = load_data()

df = st.session_state.activities

# ---- Sidebar filters ----
with st.sidebar:
    st.markdown("### 🔍 Filters")
    f_type     = st.multiselect("Type", TYPES, default=TYPES)
    f_function = st.multiselect("Hosting function", FUNCTIONS, default=FUNCTIONS)
    f_country  = st.multiselect("Country",
                                sorted(df["Location"].str.split(", ").str[-1].unique()))

mask = df["Type"].isin(f_type) & df["Initiating Function"].isin(f_function)
if f_country:
    mask &= df["Location"].str.split(", ").str[-1].isin(f_country)
view = df[mask].copy()

# ---- Tabs ----
tabs = st.tabs(["📅 Calendar", "🔥 Heatmap", "📊 Gantt", "📍 Location",
                "📈 Dashboard", "➕ Submit", "📤 Mass Upload"])

# Calendar
with tabsst.subheader("Calendar")
    st.caption(f"{len(view)} activities · {view['Location'].str.split(', ').str[-1].nunique()} countries")
    st.dataframe(
        view[["StartDate","EndDate","Type","Title","Location",
              "Initiating Function","Initiating Sub-Function",
              "Attendee Category","Participating Function","Internal/External","Note"]],
        use_container_width=True, hide_index=True
    )

# Heatmap
with tabsst.subheader("Heatmap · weekly pressure by function")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        view["Week"] = view["StartDate"].dt.to_period("W").apply(lambda p: p.start_time)
        heat = view.groupby(["Week","Initiating Function"]).size().reset_index(name="Count")
        pivot = heat.pivot(index="Week", columns="Initiating Function", values="Count").fillna(0)
        fig = px.imshow(pivot.T, color_continuous_scale="Oranges",
                        labels=dict(x="Week", y="Function", color="Activities"),
                        aspect="auto")
        fig.update_layout(height=400, margin=dict(l=20,r=20,t=30,b=20))
        st.plotly_chart(fig, use_container_width=True)
        st.caption(f"Comfort thresholds — Elevated ≥ {THRESHOLD_ELEVATED}, Critical ≥ {THRESHOLD_CRITICAL}")

# Gantt
with tabsst.subheader("Gantt")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        fig = px.timeline(view, x_start="StartDate", x_end="EndDate", y="Title",
                          color="Initiating Function", hover_data=["Location","Type","Note"])
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=600, margin=dict(l=20,r=20,t=30,b=20))
        st.plotly_chart(fig, use_container_width=True)

# Location
with tabsst.subheader("Activities by country")
    if view.empty:
        st.info("No activities match your filters.")
    else:
        view["Country"] = view["Location"].str.split(", ").str[-1]
        by_country = view.groupby("Country").size().reset_index(name="Activities") \
                        .sort_values("Activities", ascending=True)
        fig = px.bar(by_country, x="Activities", y="Country", orientation="h",
                     color_discrete_sequence=[FOUNDATION_ORANGE])
        fig.update_layout(height=500, margin=dict(l=20,r=20,t=30,b=20))
        st.plotly_chart(fig, use_container_width=True)

# Dashboard
with tabsst.subheader("Executive Dashboard")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total activities", len(view))
    c2.metric("Approved", int((view["Status"]=="Approved").sum()))
    c3.metric("Pending", int((view["Status"]=="Pending").sum()))
    c4.metric("Countries", view["Location"].str.split(", ").str[-1].nunique())

    st.markdown("#### When is the Foundation busiest?")
    if not view.empty:
        view["Month"] = view["StartDate"].dt.to_period("M").astype(str)
        by_month = view.groupby(["Month","Type"]).size().reset_index(name="Count")
        fig = px.bar(by_month, x="Month", y="Count", color="Type", barmode="stack")
        st.plotly_chart(fig, use_container_width=True)

# Submit
with tabsst.subheader("Submit a new activity")
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
        note     = st.text_area("Note", max_chars=250)

        if st.form_submit_button("✅ Submit", type="primary"):
            new = pd.DataFrame([{
                "StartDate": pd.to_datetime(start), "EndDate": pd.to_datetime(end),
                "Type": a_type, "Title": title, "Location": location,
                "Initiating Function": func, "Initiating Sub-Function": subfunc,
                "Attendee Category": attendee, "Participating Function": "",
                "Participating Sub-Function": "", "Internal/External": delivery,
                "Month (Auto/Manual)": pd.to_datetime(start).strftime("%b"),
                "Year (Auto/Manual)": pd.to_datetime(start).year,
                "Note": note, "Status": "Pending", "Weighting": "Medium"
            }])
            st.session_state.activities = pd.concat(
                [st.session_state.activities, new], ignore_index=True)
            st.success(f"✅ '{title}' submitted as Pending.")

# Mass Upload
with tabsst.subheader("Mass upload")
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
                st.session_state.activities = pd.concat(
                    [st.session_state.activities, pd.DataFrame(valid)], ignore_index=True)
                st.success("Loaded successfully.")