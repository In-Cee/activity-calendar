# ============================================================
#  Activity Calendar - Streamlit App - v0.6.4
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
from persistence import (load_submissions, save_submission,
                         load_status_overrides, save_status_override, make_key)

st.set_page_config(page_title="Activity Calendar", page_icon="📅", layout="wide")
init_settings()

ENABLE_LOGIN = True
SHARED_PASSWORD = "foundation2026"

if ENABLE_LOGIN:
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    if not st.session_state.auth_ok:
        st.markdown(f"<h1 style='color:{FOUNDATION_ORANGE};'>Activity Calendar</h1>", unsafe_allow_html=True)
        st.markdown("Mastercard Foundation - Enterprise Planning")
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
    .region-chip {{
        display: inline-block; background: white;
        border: 1px solid #E0E0E0; border-left: 4px solid {FOUNDATION_ORANGE};
        padding: 10px 16px; margin: 4px 8px 4px 0;
        border-radius: 4px; font-size: 14px;
    }}
    .cal-cell {{
        background: white; border: 1px solid #D8D8D8;
        border-radius: 4px; padding: 6px; min-height: 100px; font-size: 11px;
    }}
    .cal-day-num {{ font-weight: bold; color: {FOUNDATION_TEXT}; font-size: 12px; }}
    .cal-today {{ background: #FFF4E6; border: 2px solid {FOUNDATION_ORANGE}; }}
    .cal-weekend {{ background: #F5F5F0; }}
    .cal-outside {{ background: #FAFAFA; opacity: 0.5; }}
    .cal-activity {{
        display: block; margin-top: 2px; padding: 3px 5px;
        border-radius: 3px; font-size: 10px; color: white;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        cursor: help;
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
    if "Status" in df.columns:
        df["Status"] = df["Status"].astype(str).str.strip().str.title()
        df.loc[df["Status"].isin(["", "Nan", "None"]), "Status"] = "Approved"
    return df

def ensure_derived_columns(df):
    if "Country" not in df.columns or df["Country"].isna().any():
        df["Country"] = df["Location"].astype(str).str.split(", ").str[-1].str.strip()
    if "Region" not in df.columns or df["Region"].isna().any():
        df["Region"] = df["Country"].map(COUNTRY_TO_REGION).fillna("Other")
    return df

def build_full_dataset():
    df = pd.read_excel("activities.xlsx", sheet_name="Activities", header=1)
    df["StartDate"] = pd.to_datetime(df["StartDate"])
    df["EndDate"]   = pd.to_datetime(df["EndDate"])
    if "Status" not in df.columns:
        df["Status"] = "Approved"
    if "Weighting" not in df.columns:
        df["Weighting"] = "Medium"
    subs = load_submissions()
    if not subs.empty:
        df = pd.concat([df, subs], ignore_index=True)
    overrides = load_status_overrides()
    if overrides:
        for i, row in df.iterrows():
            key = make_key(row.get("Title", ""), row.get("StartDate", ""))
            if key in overrides:
                df.at[i, "Status"] = overrides[key]
    df = normalize_status(df)
    df = ensure_derived_columns(df)
    return df

def activity_tooltip(row):
    """Build a rich HTML tooltip for hover."""
    parts = [
        f"{row.get('Title', '')}",
        f"Type: {row.get('Type', '')}",
        f"Location: {row.get('Location', '')}",
        f"Dates: {pd.to_datetime(row['StartDate']).strftime('%b %d, %Y')} - {pd.to_datetime(row['EndDate']).strftime('%b %d, %Y')}",
        f"Host: {row.get('Initiating Function', '')} ({row.get('Initiating Sub-Function', '')})",
        f"Attendees: {row.get('Attendee Category', '')}",
        f"Mode: {row.get('Internal/External', '')}",
        f"Status: {row.get('Status', '')} | Weighting: {row.get('Weighting', '')}",
    ]
    part = str(row.get("Participating Function", ""))
    if part and part.lower() != "nan":
        parts.append(f"Participating: {part}")
    note = str(row.get("Note", ""))
    if note and note.lower() != "nan":
        parts.append(f"Note: {note}")
    return " | ".join(parts).replace("'", "").replace('"', "")

st.session_state.activities = build_full_dataset()
df = st.session_state.activities

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
    st.caption(f"v0.6.4 - {len(df)} activities loaded - {df['Country'].nunique()} countries")
    if ENABLE_LOGIN:
        if st.button("Sign out"):
            st.session_state.auth_ok = False
            st.rerun()

mask = (
    df["Type"].isin(f_type)
    & df["Initiating Function"].isin(f_hosting)
    & df["Status"].isin(f_status)
    & df["Weighting"].isin(f_weighting)
    & df["Internal/External"].isin(f_delivery)
)
if f_country: mask &= df["Country"].isin(f_country)
if f_region: mask &= df["Region"].isin(f_region)
if f_participating:
    pattern = "|".join([f"\\b{p}\\b" for p in f_participating])
    mask &= df["Participating Function"].astype(str).str.contains(pattern, regex=True, na=False)
if isinstance(f_dates, tuple) and len(f_dates) == 2:
    start, end = f_dates
    mask &= (df["StartDate"].dt.date >= start) & (df["StartDate"].dt.date <= end)
view = df[mask].copy()

# ============================================================
#  Tabs (Dashboard merged - no separate Exec Brief)
# ============================================================
if view_mode == "Executive":
    (tab_dashboard, tab_calendar, tab_heatmap, tab_gantt,
     tab_location, tab_approvals, tab_submit, tab_upload, tab_settings) = st.tabs(
        ["📈 Dashboard", "📅 Calendar", "🔥 Heatmap", "📊 Gantt",
         "📍 Location", "✅ Approvals", "➕ Submit", "📤 Mass Upload", "⚙️ Settings"]
    )
else:
    (tab_calendar, tab_heatmap, tab_gantt, tab_location,
     tab_dashboard, tab_approvals, tab_submit, tab_upload, tab_settings) = st.tabs(
        ["📅 Calendar", "🔥 Heatmap", "📊 Gantt", "📍 Location",
         "📈 Dashboard", "✅ Approvals", "➕ Submit", "📤 Mass Upload", "⚙️ Settings"]
    )

ENT_ELEV, ENT_CRIT = get_thresholds()

# ============================================================
#  Tab: Dashboard (now includes the Exec Brief hero + exports)
# ============================================================
with tab_dashboard:
    try:
        total = len(view)
        approved = int((view["Status"] == "Approved").sum())
        pending  = int((view["Status"] == "Pending").sum())
        high_w   = int((view["Weighting"] == "High").sum())
        countries = view["Country"].nunique()

        narrative_plain = ""; top_func = "-"; top_country = "-"
        if total > 0:
            top_func = view["Initiating Function"].value_counts().idxmax()
            top_func_n = int(view["Initiating Function"].value_counts().max())
            peak_month = view.assign(M=view["StartDate"].dt.to_period("M").astype(str))["M"].value_counts().idxmax()
            top_country = view["Country"].value_counts().idxmax()
            top_country_n = int(view["Country"].value_counts().max())
            high_pct = (view["Weighting"] == "High").mean() * 100
            narrative_plain = (
                f"As of {date.today().strftime('%d %B %Y')}, the Foundation has {total} activities "
                f"planned across {countries} countries. {approved} are approved "
                f"({approved/total*100:.0f}%), {pending} pending review. {top_func} is "
                f"hosting the most work ({top_func_n} activities), and {top_country} "
                f"leads activity volume ({top_country_n}). {high_pct:.0f}% of work is High weighting. "
                f"Peak month: {peak_month}."
            )
        else:
            narrative_plain = "No activities match the current filters."

        # Executive hero card
        st.markdown(f"""
<div class='exec-hero'>
<p style='color:#888; font-size:12px; margin:0 0 4px 0;'>EXECUTIVE BRIEF · {date.today().strftime('%d %B %Y')}</p>
<h2>What you need to know</h2>
<p>{narrative_plain}</p>
</div>
""", unsafe_allow_html=True)

        # KPI strip
        c1, c2, c3, c4, c5 = st.columns(5)
        kpi_with_tooltip(c1, "Total activities", total)
        kpi_with_tooltip(c2, "Approved", approved, f"{approved/max(total,1)*100:.0f}%")
        kpi_with_tooltip(c3, "Pending", pending, f"{pending/max(total,1)*100:.0f}%")
        kpi_with_tooltip(c4, "High weighting", high_w)
        kpi_with_tooltip(c5, "Countries", countries)

        if not view.empty:
            chart_title_with_explainer("busiest", "####")
            v = view.copy()
            v["Month"] = v["StartDate"].dt.to_period("M").astype(str)
            by_month = v.groupby(["Month", "Type"]).size().reset_index(name="Count")
            fig = px.bar(by_month, x="Month", y="Count", color="Type", barmode="stack",
                         color_discrete_map=TYPE_COLORS)
            fig.update_layout(plot_bgcolor="white", height=380, margin=dict(l=20, r=20, t=20, b=20),
                              xaxis=dict(showgrid=True, gridcolor="#EEE"),
                              yaxis=dict(showgrid=True, gridcolor="#EEE"))
            st.plotly_chart(fig, use_container_width=True)

            chart_title_with_explainer("priority", "####")
            by_w = view["Weighting"].value_counts().reset_index()
            by_w.columns = ["Weighting", "Count"]
            fig = px.pie(by_w, values="Count", names="Weighting", hole=0.5,
                         color_discrete_sequence=[FOUNDATION_ORANGE, FOUNDATION_AMBER, FOUNDATION_GREEN])
            fig.update_layout(height=320, margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
            insight(f"{high_pct:.0f}% of activities are High weighting. "
                    f"{'Healthy strategic focus.' if high_pct >= 30 else 'Consider lifting more activities to High to sharpen focus.'}")

            # Functions needing attention
            st.markdown("#### Functions needing attention")
            v2 = view.copy()
            v2["Week"] = v2["StartDate"].dt.to_period("W").apply(lambda p: p.start_time)
            heat = v2.groupby(["Week", "Initiating Function"]).size().reset_index(name="Count")
            pivot = heat.pivot(index="Week", columns="Initiating Function", values="Count").fillna(0)
            attention = []
            for func in pivot.columns:
                peak = int(pivot[func].max())
                elev, crit = get_thresholds(func)
                if peak >= crit:
                    attention.append([func, peak, "🔴 Critical", elev, crit])
                elif peak >= elev:
                    attention.append([func, peak, "🟠 Elevated", elev, crit])
            attention_df = pd.DataFrame(attention,
                                        columns=["Function", "Peak week", "Status", "Elevated", "Critical"])
            if attention_df.empty:
                st.success("✅ All functions are within comfort thresholds. No immediate action needed.")
            else:
                st.dataframe(attention_df, use_container_width=True, hide_index=True)

            # Exports
            st.markdown("#### 📤 Export")
            ex1, ex2, ex3 = st.columns(3)
            kpis_for_pptx = [
                ("Total activities", total),
                ("Approved", f"{approved} ({approved/max(total,1)*100:.0f}%)"),
                ("Pending", f"{pending} ({pending/max(total,1)*100:.0f}%)"),
                ("High weighting", high_w),
                ("Countries", countries),
            ]
            pptx_buf = build_pptx(narrative_plain, kpis_for_pptx, attention_df, top_func, top_country)
            ex1.download_button(
                "🎨 Download PowerPoint",
                data=pptx_buf,
                file_name=f"Activity_Calendar_ExecBrief_{date.today().isoformat()}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
            excel_buf = build_excel_export(view[[
                "StartDate", "EndDate", "Type", "Title", "Location", "Country",
                "Initiating Function", "Initiating Sub-Function",
                "Attendee Category", "Participating Function",
                "Internal/External", "Status", "Weighting", "Note"
            ]])
            ex2.download_button(
                "📊 Download Excel (current view)",
                data=excel_buf,
                file_name=f"Activity_Calendar_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            ex3.caption("📄 For PDF: press Ctrl+P → Save as PDF.")
    except Exception as e:
        st.error(f"Dashboard could not render: {e}")

# ============================================================
#  Tab: Calendar (List, Month, Week)
# ============================================================
with tab_calendar:
    chart_title_with_explainer("calendar", "###")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            cal_view = st.radio("Calendar view", ["List", "Month", "Week"], horizontal=True, key="cal_view")

            if cal_view == "List":
                st.caption(f"{len(view)} activities - {view['Country'].nunique()} countries")
                insight(f"Sortable list of {len(view)} activities. Hover any column header to sort. "
                        f"Pick an activity below for a detailed view.")
                display = view[["StartDate", "EndDate", "Type", "Title", "Location", "Country",
                                "Initiating Function", "Initiating Sub-Function",
                                "Attendee Category", "Participating Function", "Participating Sub-Function",
                                "Internal/External", "Status", "Weighting", "Note"]].reset_index(drop=True)
                st.dataframe(display, use_container_width=True, hide_index=True)

                excel_buf = build_excel_export(display)
                st.download_button("📊 Download Excel (filtered view)", data=excel_buf,
                                   file_name=f"Activity_Calendar_{date.today().isoformat()}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                    st.markdown(f"""
<div style='background:white; border-left:4px solid {FOUNDATION_ORANGE}; padding:16px 20px; border-radius:6px;'>
<h4 style='color:{FOUNDATION_ORANGE}; margin-top:0;'>{flag(row['Country'])} {row['Title']}</h4>
<p style='color:#666;'><strong>{row['Type']}</strong> - {row['Location']} - {row['StartDate'].strftime('%b %d')} to {row['EndDate'].strftime('%b %d, %Y')}</p>
<p><strong>Host:</strong> {row['Initiating Function']} ({row['Initiating Sub-Function']})<br>
<strong>Participating:</strong> {row['Participating Function'] if str(row['Participating Function']) != 'nan' else 'None'}<br>
<strong>Participating sub-functions:</strong> {row.get('Participating Sub-Function','') if str(row.get('Participating Sub-Function','')) != 'nan' else 'None'}<br>
<strong>Attendees:</strong> {row['Attendee Category']}<br>
<strong>Mode:</strong> {row['Internal/External']}<br>
<strong>Status:</strong> {row['Status']} - <strong>Weighting:</strong> {row['Weighting']}</p>
<p style='color:#444;'><em>{row['Note'] if str(row['Note']) != 'nan' else ''}</em></p>
</div>
""", unsafe_allow_html=True)

            elif cal_view == "Month":
                months = sorted(view["StartDate"].dt.to_period("M").unique())
                month_labels = {m: m.strftime("%B %Y") for m in months}
                pick_m = st.selectbox("Month", months, format_func=lambda m: month_labels[m], key="cal_month")
                year, month = pick_m.year, pick_m.month
                first_weekday, days_in_month = cal.monthrange(year, month)
                today = date.today()
                month_acts = view[(view["StartDate"].dt.year == year) & (view["StartDate"].dt.month == month)].copy()
                insight(f"{len(month_acts)} activities in {month_labels[pick_m]}. Hover any activity for full details.")

                dow_cols = st.columns(7)
                for i, dname in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
                    dow_cols[i].markdown(f"**{dname}**")
                pad = first_weekday; day = 1; done = False
                while not done:
                    week_cols = st.columns(7)
                    for i in range(7):
                        if pad > 0:
                            week_cols[i].markdown("<div class='cal-cell cal-outside'></div>", unsafe_allow_html=True)
                            pad -= 1
                        elif day <= days_in_month:
                            d = date(year, month, day)
                            day_acts = month_acts[month_acts["StartDate"].dt.day == day]
                            cls = "cal-cell"
                            if d == today: cls += " cal-today"
                            elif d.weekday() >= 5: cls += " cal-weekend"
                            inner = f"<span class='cal-day-num'>{day}</span>"
                            for _, a in day_acts.head(4).iterrows():
                                colr = TYPE_COLORS.get(a["Type"], "#999")
                                title_short = (a["Title"][:22] + "…") if len(str(a["Title"])) > 22 else a["Title"]
                                tooltip = activity_tooltip(a)
                                inner += f"<span class='cal-activity' style='background:{colr};' title='{tooltip}'>{title_short}</span>"
                            if len(day_acts) > 4:
                                inner += f"<span style='font-size:10px;color:#888;'>+{len(day_acts)-4} more</span>"
                            week_cols[i].markdown(f"<div class='{cls}'>{inner}</div>", unsafe_allow_html=True)
                            day += 1
                        else:
                            week_cols[i].markdown("<div class='cal-cell cal-outside'></div>", unsafe_allow_html=True)
                    if day > days_in_month: done = True

                st.markdown("##### Legend")
                legend_html = ""
                for t, c in TYPE_COLORS.items():
                    legend_html += f"<span class='cal-activity' style='background:{c}; padding:4px 10px;'>{t}</span> "
                st.markdown(legend_html, unsafe_allow_html=True)

            else:  # Week view
                weeks = sorted(view["StartDate"].dt.to_period("W").unique())
                week_labels = {w: f"Week of {w.start_time.strftime('%b %d, %Y')}" for w in weeks}
                pick_w = st.selectbox("Week", weeks, format_func=lambda w: week_labels[w], key="cal_week")
                start_of_week = pick_w.start_time.date()
                week_acts = view[
                    (view["StartDate"].dt.date >= start_of_week)
                    & (view["StartDate"].dt.date <= start_of_week + timedelta(days=6))
                ]
                insight(f"{len(week_acts)} activities in {week_labels[pick_w]}. Hover any activity for full details.")
                day_cols = st.columns(7)
                for i in range(7):
                    d = start_of_week + timedelta(days=i)
                    dname = d.strftime("%a %b %d")
                    day_acts = week_acts[week_acts["StartDate"].dt.date == d]
                    cls = "cal-cell"
                    if d == date.today(): cls += " cal-today"
                    elif d.weekday() >= 5: cls += " cal-weekend"
                    inner = f"<span class='cal-day-num'>{dname}</span>"
                    for _, a in day_acts.iterrows():
                        colr = TYPE_COLORS.get(a["Type"], "#999")
                        tooltip = activity_tooltip(a)
                        inner += f"<span class='cal-activity' style='background:{colr};' title='{tooltip}'>{a['Title']} {flag(a['Country'])}</span>"
                    day_cols[i].markdown(f"<div class='{cls}' style='min-height:240px;'>{inner}</div>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Calendar could not render: {e}")

# ============================================================
#  Tab: Heatmap (with cell borders)
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
            critical_hits = []; elevated_hits = []
            for func in pivot.columns:
                elev, crit = get_thresholds(func)
                peak = int(pivot[func].max())
                if peak >= crit: critical_hits.append((func, peak, crit))
                elif peak >= elev: elevated_hits.append((func, peak, elev))

            if critical_hits:
                insight(f"⚠️ Critical pressure: {len(critical_hits)} function(s) above their Critical thresholds. Action needed.")
            elif elevated_hits:
                insight(f"Elevated pressure: {len(elevated_hits)} function(s) above their Elevated thresholds. Monitor closely.")
            else:
                insight(f"Workload is balanced - all functions within their comfort thresholds.")

            # Heatmap with visible cell borders and gridlines
            fig = px.imshow(
                pivot.T,
                color_continuous_scale="Oranges",
                labels=dict(x="Week", y="Function", color="Activities"),
                aspect="auto",
                text_auto=True,
            )
            fig.update_traces(
                xgap=2, ygap=2,  # Cell separation
                textfont=dict(size=11, color="#2B2B2B"),
            )
            fig.update_layout(
                height=460, margin=dict(l=20, r=20, t=20, b=20),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(showgrid=True, gridcolor="#EEE", showline=True, linecolor="#D0D0D0"),
                yaxis=dict(showgrid=True, gridcolor="#EEE", showline=True, linecolor="#D0D0D0"),
            )
            st.plotly_chart(fig, use_container_width=True)
            st.caption(f"Enterprise default thresholds - Elevated >= {ENT_ELEV}, Critical >= {ENT_CRIT}.")

            st.markdown("##### Functions needing attention")
            rows = []
            for func in pivot.columns:
                elev, crit = get_thresholds(func)
                total = int(pivot[func].sum()); peak = int(pivot[func].max())
                if peak >= crit: status = "🔴 Critical"
                elif peak >= elev: status = "🟠 Elevated"
                else: status = "🟢 Balanced"
                action = ("Reschedule peak weeks" if peak >= crit
                          else ("Monitor peak weeks" if peak >= elev else "No action needed"))
                rows.append([func, total, peak, elev, crit, status, action])
            ft = pd.DataFrame(rows, columns=["Function", "Total", "Peak week", "Elevated", "Critical", "Status", "Action"])
            st.dataframe(ft.sort_values("Peak week", ascending=False), use_container_width=True, hide_index=True)

            st.markdown("##### Busiest weeks")
            week_totals = pivot.sum(axis=1).reset_index()
            week_totals.columns = ["Week", "Total"]
            week_totals = week_totals.sort_values("Total", ascending=False).head(10)
            week_totals["Week"] = pd.to_datetime(week_totals["Week"]).dt.strftime("Week of %b %d, %Y")
            st.dataframe(week_totals, use_container_width=True, hide_index=True)

            st.markdown("##### 🔍 Pressure point drill-down")
            d1, d2 = st.columns(2)
            with d1:
                pick_func = st.selectbox("Function", FUNCTIONS, key="hm_func")
            with d2:
                week_options = sorted(v["Week"].unique())
                week_labels = {w: pd.to_datetime(w).strftime("Week of %b %d, %Y") for w in week_options}
                pick_week = st.selectbox("Week", options=week_options, format_func=lambda w: week_labels[w], key="hm_week")
            cell_rows = v[(v["Initiating Function"] == pick_func) & (v["Week"] == pick_week)]
            if cell_rows.empty:
                st.caption(f"No {pick_func} activities in {week_labels[pick_week]}.")
            else:
                st.caption(f"{len(cell_rows)} {pick_func} activities in {week_labels[pick_week]}")
                st.dataframe(cell_rows[["Title", "Location", "Country", "StartDate", "EndDate", "Type",
                                         "Attendee Category", "Status", "Weighting"]],
                             use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Heatmap could not render: {e}")

# ============================================================
#  Tab: Gantt (with gridlines + alternating row bands)
# ============================================================
with tab_gantt:
    chart_title_with_explainer("gantt", "###")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            MAX_ROWS = 80
            g = view.sort_values("StartDate").head(MAX_ROWS).copy()
            g["Label"] = (g["Title"].astype(str) + " · " + g["Country"].astype(str)
                          + " · " + g["StartDate"].dt.strftime("%b %d"))
            if len(view) > MAX_ROWS:
                insight(f"Showing the first {MAX_ROWS} of {len(view)} activities (sorted by date). Use filters to narrow.")
            else:
                insight(f"Showing {len(g)} activities. Hover any bar for full details.")

            fig = px.timeline(
                g, x_start="StartDate", x_end="EndDate", y="Label",
                color="Initiating Function", color_discrete_map=FUNCTION_COLORS,
                hover_data={
                    "Location": True, "Type": True, "Initiating Sub-Function": True,
                    "Attendee Category": True, "Internal/External": True,
                    "Status": True, "Weighting": True, "Note": True,
                    "Label": False
                },
            )
            fig.update_yaxes(autorange="reversed")
            fig.update_traces(marker_line_color="white", marker_line_width=1)
            fig.update_layout(
                height=max(500, len(g) * 24),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis=dict(
                    showgrid=True, gridcolor="#E5E5E5", gridwidth=1,
                    showline=True, linecolor="#D0D0D0",
                    tickformat="%b %d", dtick="M1",
                ),
                yaxis=dict(
                    showgrid=True, gridcolor="#F0F0F0", gridwidth=1,
                    showline=True, linecolor="#D0D0D0",
                ),
                bargap=0.3,
            )
            # Alternating row shading for readability
            for i in range(0, len(g), 2):
                fig.add_hrect(
                    y0=i - 0.5, y1=i + 0.5,
                    fillcolor="#FAFAFA", layer="below", line_width=0,
                )
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
                insight(f"{flag(top_country)} {top_country} leads activity volume ({top_country_n} activities). "
                        f"{top_region} region carries {top_region_n/len(view)*100:.0f}% of total load.")

            st.markdown("##### Regional roll-up")
            chips_html = ""
            for _, r in by_region.iterrows():
                share = r["Activities"] / len(view) * 100
                chips_html += (f"<span class='region-chip'><strong>{r['Region']}</strong> · "
                               f"{r['Activities']} activities · {share:.0f}%</span>")
            st.markdown(chips_html, unsafe_allow_html=True)

            st.markdown("##### By country")
            by_country = (view.groupby("Country").size().reset_index(name="Activities")
                             .sort_values("Activities", ascending=True))
            by_country["Country with flag"] = by_country["Country"].apply(lambda c: f"{flag(c)} {c}")
            fig = px.bar(by_country, x="Activities", y="Country with flag", orientation="h",
                         color_discrete_sequence=[FOUNDATION_ORANGE], text="Activities")
            fig.update_traces(textposition="outside")
            fig.update_layout(
                height=max(400, len(by_country)*24), plot_bgcolor="white",
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis=dict(showgrid=True, gridcolor="#EEE"),
                yaxis=dict(showgrid=False),
            )
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Location could not render: {e}")

# ============================================================
#  Tab: Approvals
# ============================================================
with tab_approvals:
    st.subheader("✅ Approvals queue")
    try:
        all_acts = df.copy()
        pending_df = all_acts[all_acts["Status"] == "Pending"].copy()
        st.caption(f"{len(pending_df)} activities awaiting review (loaded from disk + Excel).")

        with st.expander("🔍 Status diagnostic"):
            status_counts = all_acts["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            st.dataframe(status_counts, use_container_width=True, hide_index=True)
            st.markdown("**Last 5 in dataset:**")
            st.dataframe(all_acts.tail(5)[["Title", "Status", "Initiating Function", "StartDate"]],
                         use_container_width=True, hide_index=True)

        if pending_df.empty:
            st.success("🎉 No activities pending. The queue is clear.")
        else:
            insight("Review each activity below. Approve to add to the live plan, or Decline to remove it.")
            for idx, row in pending_df.iterrows():
                with st.container(border=True):
                    c1, c2, c3 = st.columns([4, 2, 2])
                    with c1:
                        st.markdown(f"**{flag(row.get('Country',''))} {row['Title']}**")
                        st.caption(f"{row['Type']} · {row.get('Location','')} · "
                                   f"{pd.to_datetime(row['StartDate']).strftime('%b %d')} - "
                                   f"{pd.to_datetime(row['EndDate']).strftime('%b %d, %Y')}")
                        st.caption(f"Host: {row['Initiating Function']} ({row['Initiating Sub-Function']}) · "
                                   f"Weighting: {row['Weighting']}")
                        part = str(row.get("Participating Function", ""))
                        if part and part.lower() != "nan":
                            st.caption(f"Participating: {part}")
                        if str(row.get("Note","")) != "nan" and row.get("Note"):
                            st.caption(f"📝 {row['Note']}")
                    key = make_key(row.get("Title", ""), row.get("StartDate", ""))
                    with c2:
                        if st.button("✅ Approve", key=f"app_{idx}", type="primary"):
                            save_status_override(key, "Approved")
                            log_change("Approved", f"{row['Title']} ({row.get('Country','')})")
                            st.rerun()
                    with c3:
                        if st.button("❌ Decline", key=f"dec_{idx}"):
                            save_status_override(key, "Declined")
                            log_change("Declined", f"{row['Title']} ({row.get('Country','')})")
                            st.rerun()
    except Exception as e:
        st.error(f"Approvals could not render: {e}")

# ============================================================
#  Tab: Submit (with clearer Participating Sub-Function picker)
# ============================================================
with tab_submit:
    st.subheader("Submit a new activity")
    try:
        # Build sub-function list with parent function in brackets, grouped
        ALL_SUBS_LABELED = []
        for f in FUNCTIONS:
            for s in SUB_FUNCTIONS[f]:
                ALL_SUBS_LABELED.append(f"{s} ({f})")

        with st.form("submit", clear_on_submit=True):
            col1, col2 = st.columns(2)
            title    = col1.text_input("Activity name")
            a_type   = col2.selectbox("Type", TYPES)
            start    = col1.date_input("Start date")
            end      = col2.date_input("End date")
            location = col1.text_input("Location (City, Country)", placeholder="e.g., Kigali, Rwanda")
            func     = col2.selectbox("Initiating Function (Host)", FUNCTIONS)
            subfunc  = col1.selectbox("Initiating Sub-Function", SUB_FUNCTIONS[func])
            attendee = col2.selectbox("Attendee Category", ATTENDEE_CATEGORIES)

            st.markdown("##### Participating teams (optional)")
            part_funcs = st.multiselect("Participating Function(s)",
                                        [f for f in FUNCTIONS if f != func],
                                        help="Other functions involved as participants, not hosts.")
            part_subs_labeled = st.multiselect("Participating Sub-Function(s)",
                                                ALL_SUBS_LABELED,
                                                help="All 17 sub-functions across all functions are available here.")
            # Strip the "(Function)" suffix before saving
            part_subs = [s.rsplit(" (", 1)[0] for s in part_subs_labeled]

            col3, col4 = st.columns(2)
            delivery  = col3.selectbox("Internal/External", DELIVERY_MODE)
            weighting = col4.selectbox("Weighting", WEIGHTING, index=1)
            note      = st.text_area("Note", max_chars=250)

            submitted = st.form_submit_button("✅ Submit", type="primary")
            if submitted:
                if not title.strip():
                    st.error("Please enter an Activity name.")
                elif not location.strip():
                    st.error("Please enter a Location.")
                elif end < start:
                    st.error("End date cannot be before Start date.")
                else:
                    country = location.split(", ")[-1].strip() if ", " in location else location.strip()
                    region = COUNTRY_TO_REGION.get(country, "Other")
                    row_dict = {
                        "StartDate": pd.to_datetime(start).strftime("%Y-%m-%d"),
                        "EndDate": pd.to_datetime(end).strftime("%Y-%m-%d"),
                        "Type": a_type, "Title": title.strip(), "Location": location.strip(),
                        "Country": country, "Region": region,
                        "Initiating Function": func, "Initiating Sub-Function": subfunc,
                        "Attendee Category": attendee,
                        "Participating Function": ", ".join(part_funcs),
                        "Participating Sub-Function": ", ".join(part_subs),
                        "Internal/External": delivery,
                        "Month (Auto/Manual)": pd.to_datetime(start).strftime("%b"),
                        "Year (Auto/Manual)": pd.to_datetime(start).year,
                        "Note": note, "Status": "Pending", "Weighting": weighting
                    }
                    save_submission(row_dict)
                    log_change("Submitted", f"{title} ({country})")
                    st.success(f"✅ '{title}' saved to disk as Pending. Click Approvals tab to review.")
                    st.rerun()

        st.markdown("##### Persisted submissions (from disk)")
        subs = load_submissions()
        if subs.empty:
            st.caption("No persisted submissions yet.")
        else:
            st.dataframe(
                subs[["Title", "Type", "Location", "Initiating Function", "StartDate", "Status", "Weighting"]].tail(10),
                use_container_width=True, hide_index=True
            )
    except Exception as e:
        st.error(f"Submit form could not render: {e}")

# ============================================================
#  Tab: Mass Upload
# ============================================================
with tab_upload:
    st.subheader("Mass upload")
    try:
        c1, c2 = st.columns([3, 1])
        c1.write("Upload an Excel file matching the template.")
        template_buf = build_template_excel()
        c2.download_button("📥 Download template", data=template_buf,
                           file_name="Activity_Calendar_Template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        upload = st.file_uploader("Choose Excel file", type=["xlsx"])
        if upload:
            new_df = pd.read_excel(upload, sheet_name="Activities", header=1)
            st.success(f"Parsed {len(new_df)} rows. Use the Submit tab to add individual rows for now, or "
                       f"contact admin to bulk-load into persistence.")
    except Exception as e:
        st.error(f"Mass Upload could not render: {e}")

# ============================================================
#  Tab: Settings
# ============================================================
with tab_settings:
    st.subheader("⚙️ Workload comfort thresholds")
    try:
        st.markdown("##### Enterprise defaults")
        c1, c2, c3 = st.columns([1, 1, 1])
        new_elev = c1.number_input("Elevated", min_value=1, max_value=50,
                                   value=int(st.session_state.enterprise_elevated), key="ent_elev")
        new_crit = c2.number_input("Critical", min_value=1, max_value=50,
                                   value=int(st.session_state.enterprise_critical), key="ent_crit")
        if c3.button("Save", type="primary"):
            if new_crit <= new_elev:
                st.error("Critical must be > Elevated.")
            else:
                st.session_state.enterprise_elevated = new_elev
                st.session_state.enterprise_critical = new_crit
                log_change("Thresholds updated", f"Elevated={new_elev}, Critical={new_crit}")
                st.rerun()

        st.markdown("##### Audit log")
        if not st.session_state.audit_log:
            st.caption("No changes recorded yet.")
        else:
            st.dataframe(pd.DataFrame(st.session_state.audit_log), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Settings could not render: {e}")
