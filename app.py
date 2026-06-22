# ============================================================
#  Activity Calendar - Streamlit App - v0.5
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

# ============================================================
#  Page setup + Foundation theme
# ============================================================
st.set_page_config(page_title="Activity Calendar", page_icon="📅", layout="wide")
init_settings()

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
    st.caption(f"v0.5 - {len(df)} activities loaded - {df['Country'].nunique()} countries")

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
#  Tabs (Executive = exec-brief-first; Analyst = working tools first)
# ============================================================
if view_mode == "Executive":
    (tab_brief, tab_dashboard, tab_calendar, tab_heatmap, tab_gantt,
     tab_location, tab_approvals, tab_submit, tab_upload, tab_settings) = st.tabs(
        ["📋 Exec Brief", "📈 Dashboard", "📅 Calendar", "🔥 Heatmap", "📊 Gantt",
         "📍 Location", "✅ Approvals", "➕ Submit", "📤 Mass Upload", "⚙️ Settings"]
    )
else:
    (tab_calendar, tab_heatmap, tab_gantt, tab_location, tab_dashboard,
     tab_brief, tab_approvals, tab_submit, tab_upload, tab_settings) = st.tabs(
        ["📅 Calendar", "🔥 Heatmap", "📊 Gantt", "📍 Location", "📈 Dashboard",
         "📋 Exec Brief", "✅ Approvals", "➕ Submit", "📤 Mass Upload", "⚙️ Settings"]
    )

# ============================================================
#  Get current enterprise thresholds (may have been updated in Settings)
# ============================================================
ENT_ELEV, ENT_CRIT = get_thresholds()

# ============================================================
#  Tab: Exec Brief
# ============================================================
with tab_brief:
    try:
        total = len(view)
        approved = int((view["Status"] == "Approved").sum())
        pending  = int((view["Status"] == "Pending").sum())
        high_w   = int((view["Weighting"] == "High").sum())
        countries = view["Country"].nunique()

        if total > 0:
            top_func = view["Initiating Function"].value_counts().idxmax()
            top_func_n = int(view["Initiating Function"].value_counts().max())
            peak_month = view.assign(M=view["StartDate"].dt.to_period("M").astype(str))["M"].value_counts().idxmax()
            top_country = view["Country"].value_counts().idxmax()
            top_country_n = int(view["Country"].value_counts().max())
            high_pct = (view["Weighting"] == "High").mean() * 100
            narrative = (
                f"As of {date.today().strftime('%d %B %Y')}, the Foundation has <strong>{total} activities</strong> "
                f"planned across <strong>{countries} countries</strong>. {approved} are approved "
                f"({approved/total*100:.0f}%), {pending} pending review. <strong>{top_func}</strong> is "
                f"hosting the most work ({top_func_n} activities), and <strong>{flag(top_country)} {top_country}</strong> "
                f"leads activity volume ({top_country_n}). {high_pct:.0f}% of work is High weighting. "
                f"Peak month: <strong>{peak_month}</strong>."
            )
        else:
            narrative = "No activities match the current filters. Adjust filters to see a brief."

        st.markdown(f"""
<div class='exec-hero'>
<p style='color:#888; font-size:12px; margin:0 0 4px 0;'>EXECUTIVE BRIEF · {date.today().strftime('%d %B %Y')}</p>
<h2>What you need to know</h2>
<p>{narrative}</p>
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
            # Top chart - busiest months
            chart_title_with_explainer("busiest", "####")
            v = view.copy()
            v["Month"] = v["StartDate"].dt.to_period("M").astype(str)
            by_month = v.groupby(["Month", "Type"]).size().reset_index(name="Count")
            fig = px.bar(by_month, x="Month", y="Count", color="Type", barmode="stack",
                         color_discrete_map=TYPE_COLORS)
            fig.update_layout(plot_bgcolor="white", height=340,
                              margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)

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
                    attention.append((func, peak, "🔴 Critical", elev, crit))
                elif peak >= elev:
                    attention.append((func, peak, "🟠 Elevated", elev, crit))
            if attention:
                att_df = pd.DataFrame(attention, columns=["Function", "Peak week", "Status", "Elevated threshold", "Critical threshold"])
                st.dataframe(att_df, use_container_width=True, hide_index=True)
            else:
                st.success("✅ All functions are within comfort thresholds. No immediate action needed.")

            st.caption("Tip: Use browser print (Ctrl+P) for a clean A4 brief.")
    except Exception as e:
        st.error(f"Exec Brief could not render: {e}")

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
            fig.update_layout(plot_bgcolor="white", height=380, margin=dict(l=20, r=20, t=20, b=20))
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
#  Tab: Calendar (Month / Week / List)
# ============================================================
with tab_calendar:
    chart_title_with_explainer("calendar", "###")
    try:
        if view.empty:
            st.info("No activities match your filters.")
        else:
            cal_view = st.radio("Calendar view", ["Month", "Week", "List"], horizontal=True, key="cal_view")

            if cal_view == "Month":
                months = sorted(view["StartDate"].dt.to_period("M").unique())
                month_labels = {m: m.strftime("%B %Y") for m in months}
                pick_m = st.selectbox("Month", months, format_func=lambda m: month_labels[m], key="cal_month")

                year, month = pick_m.year, pick_m.month
                first_weekday, days_in_month = cal.monthrange(year, month)
                today = date.today()
                month_acts = view[
                    (view["StartDate"].dt.year == year) & (view["StartDate"].dt.month == month)
                ].copy()

                insight(f"{len(month_acts)} activities in {month_labels[pick_m]}. "
                        f"{month_acts['Country'].nunique()} countries.")

                dow_cols = st.columns(7)
                for i, dname in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
                    dow_cols[i].markdown(f"**{dname}**")

                pad = first_weekday
                day = 1
                done = False
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
                                inner += f"<span class='cal-activity' style='background:{colr};' title='{a['Title']} - {a['Country']}'>{title_short}</span>"
                            if len(day_acts) > 4:
                                inner += f"<span style='font-size:10px;color:#888;'>+{len(day_acts)-4} more</span>"
                            week_cols[i].markdown(f"<div class='{cls}'>{inner}</div>", unsafe_allow_html=True)
                            day += 1
                        else:
                            week_cols[i].markdown("<div class='cal-cell cal-outside'></div>", unsafe_allow_html=True)
                    if day > days_in_month:
                        done = True

                st.markdown("##### Legend")
                legend_html = ""
                for t, c in TYPE_COLORS.items():
                    legend_html += f"<span class='cal-activity' style='background:{c}; padding:4px 10px;'>{t}</span> "
                st.markdown(legend_html, unsafe_allow_html=True)

            elif cal_view == "Week":
                weeks = sorted(view["StartDate"].dt.to_period("W").unique())
                week_labels = {w: f"Week of {w.start_time.strftime('%b %d, %Y')}" for w in weeks}
                pick_w = st.selectbox("Week", weeks, format_func=lambda w: week_labels[w], key="cal_week")
                start_of_week = pick_w.start_time.date()
                week_acts = view[
                    (view["StartDate"].dt.date >= start_of_week)
                    & (view["StartDate"].dt.date <= start_of_week + timedelta(days=6))
                ]
                insight(f"{len(week_acts)} activities in {week_labels[pick_w]}.")

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
                        inner += f"<span class='cal-activity' style='background:{colr};'>{a['Title']} {flag(a['Country'])}</span>"
                    day_cols[i].markdown(f"<div class='{cls}' style='min-height:200px;'>{inner}</div>", unsafe_allow_html=True)

            else:
                st.caption(f"{len(view)} activities - {view['Country'].nunique()} countries")
                insight(f"Sortable list of {len(view)} activities. Pick one below for full details.")
                display = view[["StartDate", "EndDate", "Type", "Title", "Country",
                                "Initiating Function", "Initiating Sub-Function",
                                "Attendee Category", "Participating Function",
                                "Internal/External", "Status", "Weighting", "Note"]].reset_index(drop=True)
                st.dataframe(display, use_container_width=True, hide_index=True)

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
<h4 style='color:{FOUNDATION_ORANGE}; margin-top:0;'>{flag(row['Country'])} {row['Title']}</h4>
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
                        start = row["StartDate"]; end = row["EndDate"]
                        nearby = view[
                            (view["StartDate"] >= start - pd.Timedelta(days=7))
                            & (view["StartDate"] <= end + pd.Timedelta(days=7))
                            & (view["Title"] != row["Title"])
                        ]
                        st.markdown("##### Nearby activities (±7 days)")
                        if nearby.empty:
                            st.caption("No overlapping activities within ±7 days.")
                        else:
                            st.caption(f"{len(nearby)} other activities within ±7 days")
                            for _, n in nearby.head(8).iterrows():
                                st.markdown(f"- {flag(n['Country'])} **{n['Title']}** - {n['StartDate'].strftime('%b %d')} ({n['Initiating Function']})")
    except Exception as e:
        st.error(f"Calendar could not render: {e}")

# ============================================================
#  Tab: Heatmap (uses per-function thresholds from Settings)
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

            # Flag pressure using each function's own threshold
            critical_hits = []
            elevated_hits = []
            for func in pivot.columns:
                elev, crit = get_thresholds(func)
                peak = int(pivot[func].max())
                if peak >= crit:
                    critical_hits.append((func, peak, crit))
                elif peak >= elev:
                    elevated_hits.append((func, peak, elev))

            if critical_hits:
                insight(f"⚠️ Critical pressure: {len(critical_hits)} function(s) above their Critical thresholds. Action needed.")
            elif elevated_hits:
                insight(f"Elevated pressure: {len(elevated_hits)} function(s) above their Elevated thresholds. Monitor closely.")
            else:
                insight(f"Workload is balanced - all functions within their comfort thresholds.")

            fig = px.imshow(pivot.T, color_continuous_scale="Oranges",
                            labels=dict(x="Week", y="Function", color="Activities"),
                            aspect="auto")
            fig.update_layout(height=420, margin=dict(l=20, r=20, t=20, b=20), plot_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
            st.caption(f"Enterprise default thresholds - Elevated >= {ENT_ELEV}, Critical >= {ENT_CRIT}. Per-function overrides applied where set in Settings.")

            # Row totals
            st.markdown("##### Functions needing attention")
            rows = []
            for func in pivot.columns:
                elev, crit = get_thresholds(func)
                total = int(pivot[func].sum())
                peak = int(pivot[func].max())
                if peak >= crit: status = "🔴 Critical"
                elif peak >= elev: status = "🟠 Elevated"
                else: status = "🟢 Balanced"
                action = ("Reschedule or rebalance peak weeks" if peak >= crit
                          else ("Monitor peak weeks" if peak >= elev else "No action needed"))
                rows.append([func, total, peak, elev, crit, status, action])
            func_totals = pd.DataFrame(rows, columns=["Function", "Total", "Peak week", "Elevated", "Critical", "Status", "Recommended action"])
            func_totals = func_totals.sort_values("Peak week", ascending=False)
            st.dataframe(func_totals, use_container_width=True, hide_index=True)

            # Column totals
            st.markdown("##### Busiest weeks")
            week_totals = pivot.sum(axis=1).reset_index()
            week_totals.columns = ["Week", "Total"]
            week_totals = week_totals.sort_values("Total", ascending=False).head(10)
            week_totals["Week"] = pd.to_datetime(week_totals["Week"]).dt.strftime("Week of %b %d, %Y")
            st.dataframe(week_totals, use_container_width=True, hide_index=True)

            # Drill-down
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
                st.dataframe(cell_rows[["Title", "Country", "StartDate", "EndDate", "Type",
                                         "Attendee Category", "Status", "Weighting"]],
                             use_container_width=True, hide_index=True)
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
                insight(f"Showing the first {MAX_ROWS} of {len(view)} activities (sorted by date). Use filters to narrow further.")
            else:
                insight(f"Showing {len(g)} activities across {(g['EndDate'].max() - g['StartDate'].min()).days} days. Hover for details.")
            fig = px.timeline(g, x_start="StartDate", x_end="EndDate", y="Label",
                              color="Initiating Function", color_discrete_map=FUNCTION_COLORS,
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
                insight(f"{flag(top_country)} {top_country} leads activity volume ({top_country_n} activities). "
                        f"{top_region} region carries {top_region_n/len(view)*100:.0f}% of total load.")
            else:
                insight(f"{flag(top_country)} {top_country} leads activity volume ({top_country_n} activities).")

            st.markdown("##### Regional roll-up")
            chips_html = ""
            for _, r in by_region.iterrows():
                share = r["Activities"] / len(view) * 100
                chips_html += (f"<span class='region-chip'><strong>{r['Region']}</strong> · "
                               f"{r['Activities']} activities · {share:.0f}%</span>")
            st.markdown(chips_html, unsafe_allow_html=True)

            st.markdown("##### By country")
            comparison = st.checkbox("Show vs prior 90 days", value=False, key="loc_compare")
            by_country = (view.groupby("Country").size().reset_index(name="Current")
                             .sort_values("Current", ascending=True))
            by_country["Country with flag"] = by_country["Country"].apply(lambda c: f"{flag(c)} {c}")

            if comparison:
                if isinstance(f_dates, tuple) and len(f_dates) == 2:
                    cur_start, cur_end = f_dates
                else:
                    cur_start, cur_end = min_d, max_d
                prior_start = cur_start - timedelta(days=90)
                prior_end = cur_start - timedelta(days=1)
                prior = df[(df["StartDate"].dt.date >= prior_start) & (df["StartDate"].dt.date <= prior_end)]
                prior_counts = prior.groupby("Country").size().reset_index(name="Prior")
                merged = by_country.merge(prior_counts, on="Country", how="left").fillna(0)
                merged = merged.melt(id_vars=["Country", "Country with flag"],
                                     value_vars=["Current", "Prior"],
                                     var_name="Period", value_name="Activities")
                fig = px.bar(merged, x="Activities", y="Country with flag", color="Period",
                             orientation="h", barmode="group",
                             color_discrete_map={"Current": FOUNDATION_ORANGE, "Prior": "#CCCCCC"})
            else:
                fig = px.bar(by_country, x="Current", y="Country with flag", orientation="h",
                             color_discrete_sequence=[FOUNDATION_ORANGE], labels={"Current": "Activities"})
            fig.update_layout(height=max(400, len(by_country)*24), plot_bgcolor="white",
                              margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Location could not render: {e}")

# ============================================================
#  Tab: Approvals queue
# ============================================================
with tab_approvals:
    st.subheader("✅ Approvals queue")
    try:
        pending_df = st.session_state.activities[st.session_state.activities["Status"] == "Pending"].copy()
        st.caption(f"{len(pending_df)} activities awaiting review.")
        if pending_df.empty:
            st.success("🎉 No activities pending. The queue is clear.")
        else:
            insight("Review each activity below. Approve to add to the live plan, or Decline to remove it.")
            for idx, row in pending_df.iterrows():
                with st.container(border=True):
                    c1, c2, c3 = st.columns([4, 2, 2])
                    with c1:
                        st.markdown(f"**{flag(row.get('Country',''))} {row['Title']}**")
                        st.caption(f"{row['Type']} · {row['Country']} · "
                                   f"{pd.to_datetime(row['StartDate']).strftime('%b %d')} - "
                                   f"{pd.to_datetime(row['EndDate']).strftime('%b %d, %Y')}")
                        st.caption(f"Hosting: {row['Initiating Function']} ({row['Initiating Sub-Function']}) · "
                                   f"Weighting: {row['Weighting']}")
                        if str(row.get("Note","")) != "nan" and row.get("Note"):
                            st.caption(f"📝 {row['Note']}")
                    with c2:
                        if st.button("✅ Approve", key=f"app_{idx}", type="primary"):
                            st.session_state.activities.at[idx, "Status"] = "Approved"
                            log_change("Approved", f"{row['Title']} ({row['Country']})")
                            st.rerun()
                    with c3:
                        if st.button("❌ Decline", key=f"dec_{idx}"):
                            st.session_state.activities.at[idx, "Status"] = "Declined"
                            log_change("Declined", f"{row['Title']} ({row['Country']})")
                            st.rerun()
    except Exception as e:
        st.error(f"Approvals could not render: {e}")

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
                log_change("Submitted", f"{title} ({country})")
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
                    log_change("Mass upload", f"Loaded {len(new_df)} activities")
                    st.success(f"Loaded {len(new_df)} activities into the calendar.")
    except Exception as e:
        st.error(f"Mass Upload could not render: {e}")

# ============================================================
#  Tab: Settings (Thresholds + Audit log)
# ============================================================
with tab_settings:
    st.subheader("⚙️ Workload comfort thresholds")
    try:
        st.write("Set the activity volume per week at which a function is considered stretched (Elevated) or overloaded (Critical). "
                 "These drive the Heatmap RAG colours and the Exec Brief alerts.")

        with st.expander("How are defaults set?"):
            st.markdown(
                "- **Elevated** is an early-warning level - above this, teams typically feel stretched.\n"
                "- **Critical** is an overload level - above this, quality and decision-making suffer.\n"
                "- Defaults are based on average team capacity. Function leads can override per function below."
            )

        # Enterprise defaults
        st.markdown("##### Enterprise defaults")
        c1, c2, c3 = st.columns([1, 1, 1])
        new_elev = c1.number_input("Elevated (amber)", min_value=1, max_value=50,
                                   value=int(st.session_state.enterprise_elevated), key="ent_elev")
        new_crit = c2.number_input("Critical (red)", min_value=1, max_value=50,
                                   value=int(st.session_state.enterprise_critical), key="ent_crit")

        if c3.button("Save enterprise defaults", type="primary"):
            if new_crit <= new_elev:
                st.error("Critical must be greater than Elevated.")
            else:
                old = (st.session_state.enterprise_elevated, st.session_state.enterprise_critical)
                st.session_state.enterprise_elevated = new_elev
                st.session_state.enterprise_critical = new_crit
                log_change("Enterprise thresholds updated", f"Elevated {old[0]}→{new_elev}, Critical {old[1]}→{new_crit}")
                st.success("✅ Enterprise defaults updated.")
                st.rerun()

        # Per-function overrides
        st.markdown("##### Per-function overrides")
        st.caption("Leave a value blank to inherit the enterprise default.")

        for func in FUNCTIONS:
            cur = st.session_state.function_thresholds.get(func, {"elevated": None, "critical": None, "note": ""})
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns([2, 1, 1, 3])
                c1.markdown(f"**{func}**")
                ev = c2.text_input(f"Elevated", value="" if cur["elevated"] is None else str(cur["elevated"]),
                                   key=f"ev_{func}", label_visibility="collapsed", placeholder=f"default {st.session_state.enterprise_elevated}")
                cv = c3.text_input(f"Critical", value="" if cur["critical"] is None else str(cur["critical"]),
                                   key=f"cv_{func}", label_visibility="collapsed", placeholder=f"default {st.session_state.enterprise_critical}")
                note = c4.text_input(f"Note", value=cur.get("note", ""), key=f"nt_{func}",
                                     label_visibility="collapsed", placeholder="Optional note (e.g., lower during Q4 close)")

                def parse_int(x):
                    try: return int(x) if str(x).strip() else None
                    except: return None

                new_ev = parse_int(ev)
                new_cv = parse_int(cv)

                if (new_ev != cur["elevated"]) or (new_cv != cur["critical"]) or (note != cur.get("note", "")):
                    if c4.button(f"Save {func}", key=f"save_{func}"):
                        if new_ev is not None and new_cv is not None and new_cv <= new_ev:
                            st.error(f"{func}: Critical must be greater than Elevated.")
                        else:
                            st.session_state.function_thresholds[func] = {
                                "elevated": new_ev, "critical": new_cv, "note": note
                            }
                            log_change(f"Threshold updated", f"{func}: Elevated={new_ev}, Critical={new_cv}, Note='{note}'")
                            st.success(f"✅ {func} thresholds updated.")
                            st.rerun()

        # Audit log
        st.markdown("##### Audit log")
        if not st.session_state.audit_log:
            st.caption("No changes recorded yet.")
        else:
            audit_df = pd.DataFrame(st.session_state.audit_log)
            st.dataframe(audit_df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Settings could not render: {e}")
