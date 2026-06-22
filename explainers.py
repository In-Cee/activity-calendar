# ============================================================
#  Explainer text for charts and KPIs - v0.3
#  Plain-English content used by the ⓘ popovers across the app
# ============================================================

# ---- KPI tooltips (short, single sentences) ----
KPI_EXPLAINERS = {
    "Total activities": "All activities currently in the plan, regardless of status.",
    "Approved": "Activities confirmed and ready to proceed. Shown as count and share of total.",
    "Pending": "Activities awaiting review or approval. Shown as count and share of total.",
    "High weighting": "Activities tied to top strategic priorities. A healthy plan has roughly 30% or more at High.",
    "Countries": "Number of distinct countries with at least one activity in the current view.",
}

# ---- Chart explainers (longer, three sections) ----
CHART_EXPLAINERS = {
    "pressure": {
        "title": "Where is workload pressure building?",
        "what": "The number of activities each function is hosting per week.",
        "how": "Darker cells = busier weeks for that function. Watch for clusters of dark cells in the same column - those weeks affect the whole Foundation.",
        "why": "Helps us spot pressure points early and rebalance before quality or attendance suffers.",
        "thresholds": (
            "Comfort thresholds drive the colour intensity:\n\n"
            "- Amber (Elevated, >= 6 activities/week): Set based on average team capacity to deliver well alongside business-as-usual. Above this, teams typically feel stretched.\n\n"
            "- Red (Critical, >= 10 activities/week): Set based on observed concurrency limits across enabling functions. Above this, quality and decision-making start to suffer.\n\n"
            "These are enterprise-wide defaults and can be customised per function in a future Settings page."
        ),
    },
    "busiest": {
        "title": "When is the Foundation busiest?",
        "what": "Total activities planned each month, broken down by activity type.",
        "how": "Taller bars = busier months. Each colour block shows how much of the month is taken by a specific activity type.",
        "why": "Highlights heavy months for cross-functional load planning, retreat scheduling, and decision-making windows.",
    },
    "priority": {
        "title": "Are we focused on the right work?",
        "what": "The split of activities by strategic weighting (High, Medium, Low).",
        "how": "High = activities tied to top strategic priorities. Medium = important but not top tier. Low = supporting or business-as-usual.",
        "why": "Helps confirm time and effort are concentrated on the work that matters most. A plan dominated by Medium or Low is a prompt to re-prioritise.",
    },
    "location": {
        "title": "Where in our markets is activity concentrated?",
        "what": "Activities grouped by country and by region.",
        "how": "Longer bars = more activity in that country. The regional table shows volume rolled up by Foundation region.",
        "why": "Helps us anticipate travel load, country office capacity, and partner engagement intensity by market.",
    },
    "calendar": {
        "title": "Calendar (list view)",
        "what": "A sortable, filterable list of every activity in the current view.",
        "how": "Click column headers to sort. Use the sidebar filters to narrow the list.",
        "why": "Gives you a single source of truth for what's planned, when, where, and by whom.",
    },
    "gantt": {
        "title": "Activity timeline",
        "what": "Each activity drawn as a bar from start to end date, coloured by hosting function.",
        "how": "Hover any bar for full details. Bars are sorted by start date. Limited to the first 80 to stay readable - use filters to narrow if needed.",
        "why": "Helps spot overlapping activities and identify clusters that may need rescheduling.",
    },
}

# ---- Helper to render an explainer popover ----
def chart_explainer_markdown(key):
    """Return formatted markdown for a chart explainer popover."""
    e = CHART_EXPLAINERS.get(key, {})
    parts = []
    if "what" in e:
        parts.append(f"**What this shows**\n\n{e['what']}")
    if "how" in e:
        parts.append(f"**How to read it**\n\n{e['how']}")
    if "why" in e:
        parts.append(f"**Why it matters**\n\n{e['why']}")
    if "thresholds" in e:
        parts.append(f"**About the dotted lines / thresholds**\n\n{e['thresholds']}")
    return "\n\n---\n\n".join(parts)
