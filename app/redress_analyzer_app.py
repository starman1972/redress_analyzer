import os
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st


# ------------------------------------
# Helpers
# ------------------------------------
def get_data_dir() -> Path:
    """Return the absolute path to the data directory.

    Assumes this file is located in: project_root/app/redress_app.py
    and data folder is: project_root/data
    """
    return Path(__file__).resolve().parent.parent / "data"


def list_excel_files(data_dir: Path) -> list[Path]:
    """List all .xlsx files in the data directory."""
    if not data_dir.exists():
        return []
    return sorted(data_dir.glob("*.xlsx"))


def read_official_mailing_date(file_path: Path) -> Optional[pd.Timestamp]:
    """Read OFFICIAL_MAILING_DATE from cell A1/B1 using pandas.

    Expects:
        A1 = "OFFICIAL_MAILING_DATE"
        B1 = date value (Excel date or string)
    """
    try:
        # Read first row (0) with no header
        meta_df = pd.read_excel(file_path, header=None, nrows=1)
    except Exception as e:
        st.error(f"Error reading metadata from file: {e}")
        return None

    if meta_df.shape[1] < 2:
        st.warning("Metadata row does not contain at least two columns (A and B).")
        return None

    label = str(meta_df.iat[0, 0]).strip()
    raw_date = meta_df.iat[0, 1]

    if label != "OFFICIAL_MAILING_DATE":
        st.warning(
            f'Cell A1 does not contain "OFFICIAL_MAILING_DATE" (found: "{label}").'
        )
        return None

    # Try to parse date
    if isinstance(raw_date, pd.Timestamp):
        return raw_date.normalize()
    try:
        parsed = pd.to_datetime(raw_date)
        if pd.isna(parsed):
            st.warning("Could not parse OFFICIAL_MAILING_DATE from cell B1.")
            return None
        return parsed.normalize()
    except Exception:
        st.warning("Could not parse OFFICIAL_MAILING_DATE from cell B1.")
        return None


def load_redress_data(file_path: Path) -> pd.DataFrame:
    """Load data starting from header row 4 (index=3)."""
    try:
        df = pd.read_excel(file_path, header=3)
    except Exception as e:
        st.error(f"Error reading data table from file: {e}")
        return pd.DataFrame()

    # Normalize column names (strip spaces, uppercase)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def compute_distribution(delta_series: pd.Series) -> pd.DataFrame:
    """Compute distribution and cumulative coverage for delta_days series."""
    dist = delta_series.value_counts().sort_index()
    cum = dist.cumsum()
    total = dist.sum()
    cum_pct = (cum / total * 100).round(2)

    dist_df = pd.DataFrame(
        {
            "Days until redress": dist.index,
            "Count": dist.values,
            "Cumulative count": cum.values,
            "Cumulative (%)": cum_pct.values,
        }
    )
    return dist_df


def days_for_target_coverage(dist_df: pd.DataFrame, target_pct: float) -> int:
    """Smallest days value where cumulative percentage >= target_pct."""
    if dist_df.empty:
        return 0
    mask = dist_df["Cumulative (%)"] >= target_pct
    if not mask.any():
        return int(dist_df["Days until redress"].max())
    # First index where condition is true
    idx = mask.idxmax()
    return int(dist_df.loc[idx, "Days until redress"])


def coverage_for_days(dist_df: pd.DataFrame, days: int) -> float:
    """Coverage percentage for a given waiting time in days."""
    if dist_df.empty:
        return 0.0
    mask = dist_df["Days until redress"] <= days
    if not mask.any():
        return 0.0
    covered = dist_df.loc[mask, "Count"].sum()
    total = dist_df["Count"].sum()
    return round(covered / total * 100, 2)


# ------------------------------------
# Streamlit App
# ------------------------------------
st.set_page_config(
    page_title="Redress Analyzer",
    layout="wide",
)

st.title("📬 Redress Analyzer")

st.markdown(
    """
This tool analyzes how many days pass between the **official mailing date**
and the **redress capture date** (when a mailing is marked as undeliverable).

All files are read from the local `data/` folder.  
Each file must contain:

- `OFFICIAL_MAILING_DATE` in cell **A1**, with the date in **B1**
- A table starting at row **4** with columns:
  - `GRUND_UNZUSTELLBARKEIT`
  - `REDRESSENERFASSUNG_DATUM`
"""
)

data_dir = get_data_dir()

if not data_dir.exists():
    st.error(f"Data folder not found: `{data_dir}`")
    st.stop()

files = list_excel_files(data_dir)

if not files:
    st.warning(f"No .xlsx files found in `{data_dir}`. Please add at least one file.")
    st.stop()

# ------------------------------------
# File selection
# ------------------------------------
st.sidebar.header("Settings")

file_labels = [f.name for f in files]
selected_label = st.sidebar.selectbox("Select a mailing file", options=file_labels)
selected_file = files[file_labels.index(selected_label)]

st.markdown(f"### 📄 Selected file: `{selected_label}`")

# ------------------------------------
# Read metadata (official mailing date)
# ------------------------------------
official_date = read_official_mailing_date(selected_file)

if official_date is None:
    st.warning(
        "OFFICIAL_MAILING_DATE could not be read from the file. "
        "Please select the official mailing date manually."
    )
    manual_date = st.date_input("Official mailing date (Day 0)")
    official_date = pd.to_datetime(manual_date)
else:
    st.success(f"Official mailing date (Day 0): **{official_date.date()}**")

# ------------------------------------
# Load main data
# ------------------------------------
df = load_redress_data(selected_file)

if df.empty:
    st.error("No data loaded from the selected file. Please check the file structure.")
    st.stop()

required_cols = {"GRUND_UNZUSTELLBARKEIT", "REDRESSENERFASSUNG_DATUM"}
missing = required_cols - set(df.columns)

if missing:
    st.error(
        "The following required columns are missing in the data table: "
        + ", ".join(sorted(missing))
    )
    st.write("Columns found:", list(df.columns))
    st.stop()

# Show a small preview
st.subheader("Data preview")
st.dataframe(df.head(), use_container_width=True)

# ------------------------------------
# Optional filtering by reason
# ------------------------------------
reason_col = "GRUND_UNZUSTELLBARKEIT"
redress_col = "REDRESSENERFASSUNG_DATUM"

unique_reasons = (
    df[reason_col].dropna().astype(str).sort_values().unique().tolist()
    if reason_col in df.columns
    else []
)

if unique_reasons:
    st.subheader("Filter (optional)")
    selected_reason = st.selectbox(
        "Filter by undeliverable reason",
        options=["(All reasons)"] + unique_reasons,
        index=0,
    )
    if selected_reason != "(All reasons)":
        df = df[df[reason_col].astype(str) == selected_reason]

# ------------------------------------
# Parse redress dates and compute delta_days
# ------------------------------------
st.subheader("Delta calculation")

df["redress_date"] = pd.to_datetime(df[redress_col], errors="coerce")

before_count = len(df)
df = df.dropna(subset=["redress_date"])
after_count = len(df)

if after_count == 0:
    st.error(
        "No valid redress dates found after parsing. "
        "Please check the `REDRESSENERFASSUNG_DATUM` column."
    )
    st.stop()

st.write(f"Records before date parsing: **{before_count}**")
st.write(f"Records with valid redress date: **{after_count}**")

# Compute delta in days
df["delta_days"] = (df["redress_date"] - official_date).dt.days

st.markdown(
    "By default, negative values (redress before official mailing date) "
    "are treated as data anomalies and can be excluded."
)
exclude_negative = st.checkbox("Exclude negative deltas", value=True)

if exclude_negative:
    df_valid = df[df["delta_days"] >= 0].copy()
else:
    df_valid = df.copy()

if df_valid.empty:
    st.error(
        "No valid records remain after applying filters (and possibly excluding negatives)."
    )
    st.stop()

total_valid = len(df_valid)
st.write(f"Records used for analysis: **{total_valid}**")

if total_valid < 30:
    st.warning(
        "Sample size is quite small (< 30). "
        "Statistics may not be very representative."
    )

# ------------------------------------
# Summary statistics
# ------------------------------------
st.subheader("Summary statistics (days until redress)")

delta = df_valid["delta_days"]

desc = delta.describe(percentiles=[0.5, 0.9, 0.95, 0.99])

stats_df = pd.DataFrame(
    {
        "Metric": [
            "Mean (average)",
            "Median (50%)",
            "90th percentile",
            "95th percentile",
            "99th percentile",
            "Minimum",
            "Maximum",
        ],
        "Value (days)": [
            round(float(desc["mean"]), 2),
            int(desc["50%"]),
            int(desc["90%"]),
            int(desc["95%"]),
            int(desc["99%"]),
            int(desc["min"]),
            int(desc["max"]),
        ],
    }
)

st.table(stats_df)

# ------------------------------------
# Distribution & cumulative coverage
# ------------------------------------
st.subheader("Distribution and cumulative coverage")

dist_df = compute_distribution(delta)

st.write("Distribution table:")
st.dataframe(dist_df, use_container_width=True)

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("**Frequency by days since official mailing date**")
    chart_data = dist_df.set_index("Days until redress")["Count"]
    st.bar_chart(chart_data)

with col_chart2:
    st.markdown("**Cumulative coverage (%)**")
    chart_data = dist_df.set_index("Days until redress")["Cumulative (%)"]
    st.line_chart(chart_data)

# ------------------------------------
# Business questions
# ------------------------------------
st.subheader("Business questions")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**1️⃣ How long to reach a certain coverage?**")
    target_pct = st.slider(
        "Target redress coverage (%)",
        min_value=80,
        max_value=100,
        value=95,
        step=1,
    )
    needed_days = days_for_target_coverage(dist_df, target_pct)
    st.success(
        f"To reach approximately **{target_pct}%** of all redresses, "
        f"you should wait at least **{needed_days} days** after the official mailing date."
    )

with col_b:
    st.markdown("**2️⃣ What coverage do we get with a given waiting time?**")
    max_days = int(dist_df["Days until redress"].max())
    wait_days = st.slider(
        "Planned waiting time (days after official mailing date)",
        min_value=0,
        max_value=max_days,
        value=min(10, max_days),
        step=1,
    )
    coverage = coverage_for_days(dist_df, wait_days)
    st.info(
        f"If you wait **{wait_days} days**, you will capture approximately **{coverage}%** "
        f"of all redresses."
    )

# ------------------------------------
# Textual summary
# ------------------------------------
st.subheader("Summary for internal communication")

median_days = int(desc["50%"])
p95_days = int(desc["95%"])
p99_days = int(desc["99%"])

reason_info = ""
if unique_reasons and selected_reason != "(All reasons)":
    reason_info = f" (reason filter: '{selected_reason}')"

summary_lines = [
    f"**File:** `{selected_label}`",
    f"**Official mailing date (Day 0):** {official_date.date()}",
    f"**Number of valid redresses used in this analysis{reason_info}:** {total_valid}",
    "",
    f"- 50% of all redresses are captured within **{median_days} days**",
    f"- 95% of all redresses are captured within **{p95_days} days**",
    f"- 99% of all redresses are captured within **{p99_days} days**",
    "",
    f"**Recommendation:**",
    f"For this mailing, you should wait at least **{p95_days} days** after the official mailing date",
    f"if you want to exclude approximately **95%** of all undeliverable addresses from the next mailing list.",
]

st.markdown("\n".join(summary_lines))

st.caption(
    "Note: Values are based on the current file, filters, and settings "
    "(e.g. exclusion of negative deltas)."
)
