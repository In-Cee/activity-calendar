# ============================================================
#  Settings store - v0.5
#  Manages workload thresholds and audit log
# ============================================================

import streamlit as st
from datetime import datetime
from lookups import FUNCTIONS, THRESHOLD_ELEVATED, THRESHOLD_CRITICAL


def init_settings():
    """Initialise threshold settings and audit log in session state."""
    if "enterprise_elevated" not in st.session_state:
        st.session_state.enterprise_elevated = THRESHOLD_ELEVATED
    if "enterprise_critical" not in st.session_state:
        st.session_state.enterprise_critical = THRESHOLD_CRITICAL
    if "function_thresholds" not in st.session_state:
        # Each function starts with default = enterprise-wide
        st.session_state.function_thresholds = {
            f: {"elevated": None, "critical": None, "note": ""} for f in FUNCTIONS
        }
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []


def get_thresholds(function_name=None):
    """Return (elevated, critical) for a function, or enterprise defaults if None or no override."""
    init_settings()
    if function_name is None:
        return st.session_state.enterprise_elevated, st.session_state.enterprise_critical

    f = st.session_state.function_thresholds.get(function_name, {})
    elev = f.get("elevated") if f.get("elevated") is not None else st.session_state.enterprise_elevated
    crit = f.get("critical") if f.get("critical") is not None else st.session_state.enterprise_critical
    return elev, crit


def log_change(action, detail, user="Radintshi Monyobo"):
    """Append a change to the audit log."""
    init_settings()
    st.session_state.audit_log.insert(0, {
        "When": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Who": user,
        "Action": action,
        "Detail": detail,
    })
