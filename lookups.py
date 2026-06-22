# ============================================================
#  Project lookups — single source of truth
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

FOUNDATION_ORANGE = "#F37021"
FOUNDATION_BG     = "#F7F5F2"
FOUNDATION_TEXT   = "#2B2B2B"