from pathlib import Path
import re
from typing import Optional

import pandas as pd
import streamlit as st


# ------------------------------------
# Helpers
# ------------------------------------
def get_data_dir() -> Path:
    """Return the absolute path to the data directory."""
    return Path(__file__).resolve().parent.parent / "data"


def parse_date_from_filename(file_path: Path) -> Optional[pd.Timestamp]:
    """Extract YYYY-MM from filename suffix and return as Timestamp."""
    match = re.search(r"_(\d{4})[-_](\d{2})\.xlsx$", file_path.name)
    if not match:
        return None
    year, month = match.groups()
    try:
        return pd.Timestamp(year=int(year), month=int(month), day=1)
    except Exception:
        return None


def list_excel_files(data_dir: Path) -> list[Path]:
    """List .xlsx files ordered by date suffix (latest first)."""
    if not data_dir.exists():
        return []

    files = list(data_dir.glob("*.xlsx"))
    files_with_date: list[tuple[pd.Timestamp, Path]] = []
    files_without_date: list[Path] = []

    for file_path in files:
        parsed_date = parse_date_from_filename(file_path)
        if parsed_date is not None:
            files_with_date.append((parsed_date, file_path))
        else:
            files_without_date.append(file_path)

    ordered = [f for _, f in sorted(files_with_date, key=lambda x: x[0], reverse=True)]
    ordered.extend(sorted(files_without_date))
    return ordered


def read_official_mailing_date(file_path: Path) -> Optional[pd.Timestamp]:
    """Read OFFICIAL_MAILING_DATE from cell A1/B1 using pandas."""
    try:
        meta_df = pd.read_excel(file_path, header=None, nrows=1)
    except Exception as exc:
        st.error(f"Error reading metadata from file: {exc}")
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
    except Exception as exc:
        st.error(f"Error reading data table from file: {exc}")
        return pd.DataFrame()

    df.columns = [str(col).strip() for col in df.columns]
    return df


def compute_distribution(delta_series: pd.Series) -> pd.DataFrame:
    """Compute distribution and cumulative coverage for delta_days series."""
    dist = delta_series.value_counts().sort_index()
    cum = dist.cumsum()
    total = dist.sum()
    cum_pct = (cum / total * 100).round(2)

    return pd.DataFrame(
        {
            "Days until redress": dist.index,
            "Count": dist.values,
            "Cumulative count": cum.values,
            "Cumulative (%)": cum_pct.values,
        }
    )


def days_for_target_coverage(dist_df: pd.DataFrame, target_pct: float) -> int:
    """Smallest days value where cumulative percentage >= target_pct."""
    if dist_df.empty:
        return 0
    mask = dist_df["Cumulative (%)"] >= target_pct
    if not mask.any():
        return int(dist_df["Days until redress"].max())
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

# ------------------------------------
# Data / files
# ------------------------------------
data_dir = get_data_dir()

if not data_dir.exists():
    st.error(f"Data folder not found: `{data_dir}`")
    st.stop()

files = list_excel_files(data_dir)

if not files:
    st.warning(f"No .xlsx files found in `{data_dir}`. Please add at least one file.")
    st.stop()

st.sidebar.header("Mailings")
file_labels = [f.name for f in files]
selected_label = st.sidebar.radio(
    "Select a mailing file (latest first)", options=file_labels, index=0
)
selected_file = files[file_labels.index(selected_label)]

st.title("Redress Analyzer")
st.markdown(f"### Selected file: `{selected_label}`")

# ------------------------------------
# Read metadata (official mailing date)
# ------------------------------------
official_date = read_official_mailing_date(selected_file)

if official_date is None:
    with st.sidebar:
        st.warning(
            "OFFICIAL_MAILING_DATE could not be read from the file. "
            "Please select the official mailing date manually."
        )
        manual_date = st.date_input("Official mailing date (Day 0)")
    official_date = pd.to_datetime(manual_date)

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

selected_reason = "(All reasons)"
with st.sidebar:
    st.markdown("---")
    st.subheader("Filters")
    if unique_reasons:
        selected_reason = st.selectbox(
            "Filter by undeliverable reason",
            options=["(All reasons)"] + unique_reasons,
            index=0,
        )
    exclude_negative = st.checkbox(
        "Exclude negative deltas",
        value=True,
        help="Negative values mean redress before official mailing date.",
    )

if unique_reasons and selected_reason != "(All reasons)":
    df = df[df[reason_col].astype(str) == selected_reason]

# ------------------------------------
# Parse redress dates and compute delta_days
# ------------------------------------
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

st.write(
    f"Records parsed: **{before_count}** → valid redress dates: **{after_count}**"
)

df["delta_days"] = (df["redress_date"] - official_date).dt.days

df_valid = df[df["delta_days"] >= 0].copy() if exclude_negative else df.copy()

if df_valid.empty:
    st.error(
        "No valid records remain after applying filters (and possibly excluding negatives)."
    )
    st.stop()

total_valid = len(df_valid)
st.write(f"Records used for analysis: **{total_valid}**")

if total_valid < 30:
    st.warning(
        "Sample size is quite small (< 30). Statistics may not be very representative."
    )

# ------------------------------------
# Summary statistics
# ------------------------------------
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

# ------------------------------------
# Distribution & charts
# ------------------------------------
st.title("Redress Analyzer")

top_info = st.columns([2, 1])
with top_info[0]:
    st.markdown(
        f"**Official mailing date (Day 0):** `{official_date.date()}`"
    )
with top_info[1]:
    st.caption(f"File: `{selected_label}`")

st.subheader("Coverage at a glance")
dist_df = compute_distribution(delta)

col_metrics = st.columns(3)
col_metrics[0].metric("Median (days)", int(desc["50%"]))
col_metrics[1].metric("95th percentile (days)", int(desc["95%"]))
col_metrics[2].metric("99th percentile (days)", int(desc["99%"]))

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("**Frequency by days since official mailing date**")
    chart_data = dist_df.set_index("Days until redress")["Count"]
    st.bar_chart(chart_data, use_container_width=True)

with col_chart2:
    st.markdown("**Cumulative coverage (%)**")
    chart_data = dist_df.set_index("Days until redress")["Cumulative (%)"]
    st.line_chart(chart_data, use_container_width=True)

with st.expander("Show distribution table"):
    st.dataframe(dist_df, use_container_width=True)

with st.expander("Show summary statistics table"):
    st.table(stats_df)

# ------------------------------------
# Business questions
# ------------------------------------
st.subheader("Business questions")

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**How long to reach a certain coverage?**")
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
    st.markdown("**What coverage do we get with a given waiting time?**")
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
    "**Recommendation:**",
    f"For this mailing, you should wait at least **{p95_days} days** after the official mailing date",
    "if you want to exclude approximately **95%** of all undeliverable addresses from the next mailing list.",
]

st.markdown("\n".join(summary_lines))

st.caption(
    "Note: Values are based on the current file, filters, and settings "
    "(e.g. exclusion of negative deltas)."
)
