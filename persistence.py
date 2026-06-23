# ============================================================
#  Persistence layer - v0.6.3
#  Writes submissions and status changes to disk so they
#  survive session state resets.
# ============================================================

import os
import pandas as pd

SUBMISSIONS_FILE = "submissions.csv"
STATUS_OVERRIDES_FILE = "status_overrides.csv"


def load_submissions():
    """Load all previously persisted submissions."""
    if not os.path.exists(SUBMISSIONS_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_csv(SUBMISSIONS_FILE)
        if "StartDate" in df.columns:
            df["StartDate"] = pd.to_datetime(df["StartDate"], errors="coerce")
        if "EndDate" in df.columns:
            df["EndDate"] = pd.to_datetime(df["EndDate"], errors="coerce")
        return df
    except Exception:
        return pd.DataFrame()


def save_submission(row_dict):
    """Append a new submission row to the CSV file."""
    df_new = pd.DataFrame([row_dict])
    if os.path.exists(SUBMISSIONS_FILE):
        try:
            df_existing = pd.read_csv(SUBMISSIONS_FILE)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        except Exception:
            df_combined = df_new
    else:
        df_combined = df_new
    df_combined.to_csv(SUBMISSIONS_FILE, index=False)


def load_status_overrides():
    """Load any status changes applied to existing rows."""
    if not os.path.exists(STATUS_OVERRIDES_FILE):
        return {}
    try:
        df = pd.read_csv(STATUS_OVERRIDES_FILE)
        return dict(zip(df["Key"], df["Status"]))
    except Exception:
        return {}


def save_status_override(key, status):
    """Persist a status override by Title + StartDate key."""
    overrides = load_status_overrides()
    overrides[key] = status
    df = pd.DataFrame([{"Key": k, "Status": v} for k, v in overrides.items()])
    df.to_csv(STATUS_OVERRIDES_FILE, index=False)


def make_key(title, start_date):
    """Create a stable identifier for a row."""
    try:
        return f"{title}__{pd.to_datetime(start_date).strftime('%Y-%m-%d')}"
    except Exception:
        return f"{title}__{start_date}"
