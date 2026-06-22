# ============================================================
#  Project lookups - v0.2
# ============================================================

TYPES = ["Meeting", "Training", "Process", "Retreat", "Event"]

FUNCTIONS = ["Programs", "Impact", "Comms", "IT", "Legal", "HR", "Finance"]

SUB_FUNCTIONS = {
    "Programs": ["Emergency Response", "Livelihoods", "Global Health", "Climate", "Education"],
    "Impact":   ["Research", "Monitoring & Evaluation", "Data Science"],
    "Comms":    ["Brand", "Content", "Public Affairs"],
    "IT":       ["Infrastructure", "Data Platform"],
    "Legal":    ["Contracts", "Risk"],
    "HR":       ["Learning & Development"],
    "Finance":  ["Accounting"],
}

ATTENDEE_CATEGORIES = ["All Foundation", "Extended Leadership", "Subset of function", "Leadership Team"]

DELIVERY_MODE = ["Internal", "External", "Hybrid"]

STATUS = ["Pending", "Approved", "Declined"]

WEIGHTING = ["High", "Medium", "Low"]

THRESHOLD_ELEVATED = 6
THRESHOLD_CRITICAL = 10

# ---- Foundation visual identity ----
FOUNDATION_ORANGE = "#F37021"
FOUNDATION_BG     = "#F7F5F2"
FOUNDATION_TEXT   = "#2B2B2B"
FOUNDATION_AMBER  = "#E8A33D"
FOUNDATION_RED    = "#C44536"
FOUNDATION_GREEN  = "#5A8A3D"

# Warm, muted palette for activity types (used in stacked charts)
TYPE_COLORS = {
    "Meeting":  "#2E7D8B",
    "Training": "#7A9E3F",
    "Process":  "#5B6E8C",
    "Retreat":  "#C66B3D",
    "Event":    "#D9A441",
}

# Warm palette for functions
FUNCTION_COLORS = {
    "Programs": "#F37021",
    "Impact":   "#2E7D8B",
    "Comms":    "#C66B3D",
    "IT":       "#5B6E8C",
    "Legal":    "#7A9E3F",
    "HR":       "#D9A441",
    "Finance":  "#A14D8A",
}

# Region grouping for the Location tab
COUNTRY_TO_REGION = {
    "Nigeria": "West Africa",
    "Ghana": "West Africa",
    "Senegal": "West Africa",
    "Cote d'Ivoire": "West Africa",
    "Côte d'Ivoire": "West Africa",
    "Mali": "West Africa",
    "Burkina Faso": "West Africa",
    "Kenya": "East & Southern Africa",
    "Tanzania": "East & Southern Africa",
    "Uganda": "East & Southern Africa",
    "Rwanda": "East & Southern Africa",
    "Ethiopia": "East & Southern Africa",
    "South Africa": "East & Southern Africa",
    "Malawi": "East & Southern Africa",
    "Zambia": "East & Southern Africa",
    "Mozambique": "East & Southern Africa",
    "DR Congo": "East & Southern Africa",
    "Morocco": "North Africa",
    "Egypt": "North Africa",
    "Tunisia": "North Africa",
    "Canada": "Global",
    "United States": "Global",
    "United Kingdom": "Global",
}
