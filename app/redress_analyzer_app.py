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
        st.warning(f"[{file_path.name}] Could not read metadata: {exc}")
        return None

    if meta_df.shape[1] < 2:
        st.warning(f"[{file_path.name}] Metadata row does not contain at least two columns (A and B).")
        return None

    label = str(meta_df.iat[0, 0]).strip()
    raw_date = meta_df.iat[0, 1]

    if label != "OFFICIAL_MAILING_DATE":
        st.warning(f"[{file_path.name}] Cell A1 does not contain 'OFFICIAL_MAILING_DATE' (found: '{label}').")
        return None

    if isinstance(raw_date, pd.Timestamp):
        return raw_date.normalize()
    try:
        parsed = pd.to_datetime(raw_date)
        if pd.isna(parsed):
            st.warning(f"[{file_path.name}] Could not parse OFFICIAL_MAILING_DATE from cell B1.")
            return None
        return parsed.normalize()
    except Exception:
        st.warning(f"[{file_path.name}] Could not parse OFFICIAL_MAILING_DATE from cell B1.")
        return None


def load_redress_data(file_path: Path) -> pd.DataFrame:
    """Load data starting from header row 4 (index=3)."""
    try:
        df = pd.read_excel(file_path, header=3)
    except Exception as exc:
        st.warning(f"[{file_path.name}] Could not read table: {exc}")
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


def prepare_file_data(file_path: Path, prompt_manual: bool = False) -> Optional[dict]:
    """Load one mailing, parse dates, compute delta_days, and return enriched dataframe."""
    reason_col = "GRUND_UNZUSTELLBARKEIT"
    redress_col = "REDRESSENERFASSUNG_DATUM"

    official_date = read_official_mailing_date(file_path)
    if official_date is None and prompt_manual:
        manual_date = st.sidebar.date_input(f"Official mailing date for {file_path.stem}", key=f"manual_{file_path.stem}")
        official_date = pd.to_datetime(manual_date)
    if official_date is None:
        return None

    df = load_redress_data(file_path)
    required_cols = {reason_col, redress_col}
    if df.empty or required_cols - set(df.columns):
        st.warning(f"[{file_path.name}] Missing required columns: {required_cols - set(df.columns)}")
        return None

    df["redress_date"] = pd.to_datetime(df[redress_col], errors="coerce")
    before_count = len(df)
    df = df.dropna(subset=["redress_date"])
    after_count = len(df)

    df["delta_days"] = (df["redress_date"] - official_date).dt.days
    df["campaign"] = file_path.stem
    df["official_date"] = official_date
    df["reason"] = df[reason_col].astype(str)

    return {
        "campaign": file_path.stem,
        "official_date": official_date,
        "before_count": before_count,
        "after_count": after_count,
        "df": df,
    }


def summarize_by_campaign(df: pd.DataFrame) -> pd.DataFrame:
    """Return per-campaign summary stats for delta_days."""
    grouped = df.groupby("campaign")
    summary = grouped["delta_days"].agg(
        count="count",
        median="median",
        p90=lambda s: s.quantile(0.9),
        p95=lambda s: s.quantile(0.95),
        p99=lambda s: s.quantile(0.99),
        min="min",
        max="max",
    )
    first_dates = grouped["official_date"].first()
    summary = summary.join(first_dates.rename("official_date"))
    summary = summary.reset_index()
    return summary.sort_values("official_date", ascending=False)


# ------------------------------------
# Streamlit App
# ------------------------------------
st.set_page_config(
    page_title="Redress Analyzer",
    layout="wide",
)

data_dir = get_data_dir()
if not data_dir.exists():
    st.error(f"Data folder not found: `{data_dir}`")
    st.stop()

files = list_excel_files(data_dir)
if not files:
    st.warning(f"No .xlsx files found in `{data_dir}`. Please add at least one file.")
    st.stop()

# Sidebar controls
with st.sidebar:
    st.header("View")
    view_mode = st.radio(
        "Analysis mode",
        ["All mailings", "Single mailing"],
        index=0,
    )

    selected_label = None
    selected_file: Optional[Path] = None
    if view_mode == "Single mailing":
        st.markdown("---")
        st.subheader("Mailings")
        file_labels = [f.stem for f in files]
        selected_label = st.radio(
            "Select a mailing (latest first)",
            options=file_labels,
            index=0,
        )
        selected_file = files[file_labels.index(selected_label)]

# Main title
st.title("Redress Analyzer")

# ------------------------------------
# All-mailings view
# ------------------------------------
if view_mode == "All mailings":
    prepared_list = []
    for file_path in files:
        prepared = prepare_file_data(file_path, prompt_manual=False)
        if prepared is not None:
            prepared_list.append(prepared)

    if not prepared_list:
        st.error("No valid mailings could be loaded.")
        st.stop()

    df_all = pd.concat([p["df"] for p in prepared_list], ignore_index=True)

    all_reasons = sorted(
        df_all["reason"].dropna().unique().tolist()
    )

    with st.sidebar:
        st.markdown("---")
        st.subheader("Filters")
        selected_reason = st.selectbox(
            "Filter by undeliverable reason",
            options=["(All reasons)"] + all_reasons,
            index=0,
        )
        exclude_negative = st.checkbox(
            "Exclude negative deltas",
            value=True,
            help="Negative values mean redress before official mailing date.",
        )
        weighting_mode = st.radio(
            "Coverage weighting",
            options=["By records", "Average of campaigns"],
            index=0,
            help="Choose how to aggregate percentiles across mailings.",
        )

    df_filtered = df_all.copy()
    if selected_reason != "(All reasons)":
        df_filtered = df_filtered[df_filtered["reason"] == selected_reason]
    if exclude_negative:
        df_filtered = df_filtered[df_filtered["delta_days"] >= 0]

    if df_filtered.empty:
        st.warning("No data remains after applying filters.")
        st.stop()

    delta_series = df_filtered["delta_days"]
    dist_df = compute_distribution(delta_series)

    # Aggregate stats
    desc = delta_series.describe(percentiles=[0.5, 0.9, 0.95, 0.99])
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

    campaign_summary = summarize_by_campaign(df_filtered)
    avg_campaign_median = round(campaign_summary["median"].mean(), 2)
    avg_campaign_p95 = round(campaign_summary["p95"].mean(), 2)
    avg_campaign_p99 = round(campaign_summary["p99"].mean(), 2)

    info_cols = st.columns(4)
    with info_cols[0]:
        st.markdown(f"**Mailings loaded:** {len(prepared_list)}")
    with info_cols[1]:
        st.markdown(f"**Total records used:** {len(df_filtered)}")
    with info_cols[2]:
        st.markdown(f"**Median (all records):** {int(desc['50%'])} days")
    with info_cols[3]:
        st.markdown(f"**95th pct (all records):** {int(desc['95%'])} days")

    st.subheader("Coverage at a glance (all mailings)")
    kpi_cols = st.columns(3)
    kpi_cols[0].metric("Median (days)", int(desc["50%"]))
    kpi_cols[1].metric("95th percentile (days)", int(desc["95%"]))
    kpi_cols[2].metric("99th percentile (days)", int(desc["99%"]))

    if weighting_mode == "Average of campaigns":
        st.caption(
            f"Average of campaign medians/p95/p99 → median: {avg_campaign_median}, "
            f"p95: {avg_campaign_p95}, p99: {avg_campaign_p99}"
        )

    charts = st.columns(2)
    with charts[0]:
        st.markdown("**Frequency by days since official mailing date**")
        st.bar_chart(dist_df.set_index("Days until redress")["Count"], use_container_width=True)
    with charts[1]:
        st.markdown("**Cumulative coverage (%)**")
        st.line_chart(dist_df.set_index("Days until redress")["Cumulative (%)"], use_container_width=True)

    st.markdown("**Mailing comparison (median days)**")
    top_medians = campaign_summary.sort_values("median").set_index("campaign")["median"]
    st.bar_chart(top_medians, use_container_width=True)

    with st.expander("Show distribution table"):
        st.dataframe(dist_df, use_container_width=True)

    with st.expander("Show summary statistics table"):
        st.table(stats_df)

    with st.expander("Show per-mailing summary"):
        display_cols = ["campaign", "official_date", "count", "median", "p95", "p99", "min", "max"]
        st.dataframe(campaign_summary[display_cols], use_container_width=True)

    # Business questions for all mailings
    st.subheader("Business questions (all mailings)")
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("**How long to reach a certain coverage?**")
        target_pct = st.slider(
            "Target redress coverage (%)",
            min_value=50,
            max_value=100,
            value=90,
            step=1,
            key="all_target_pct",
        )
        needed_days = days_for_target_coverage(dist_df, target_pct)
        st.success(
            f"To reach approximately **{target_pct}%** of all redresses across mailings, "
            f"wait at least **{needed_days} days** after the official mailing date."
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
            key="all_wait_days",
        )
        coverage = coverage_for_days(dist_df, wait_days)
        st.info(
            f"If you wait **{wait_days} days**, you will capture approximately **{coverage}%** "
            f"of all redresses across mailings."
        )

# ------------------------------------
# Single-mailing view
# ------------------------------------
if view_mode == "Single mailing":
    if selected_file is None:
        st.stop()

    prepared = prepare_file_data(selected_file, prompt_manual=True)
    if prepared is None:
        st.error("The selected mailing could not be loaded. Please check the file.")
        st.stop()

    df = prepared["df"]
    after_count = prepared["after_count"]
    official_date = prepared["official_date"]

    unique_reasons = (
        df["reason"].dropna().astype(str).sort_values().unique().tolist()
        if "reason" in df.columns
        else []
    )

    with st.sidebar:
        st.markdown("---")
        st.subheader("Filters")
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

    df_filtered = df.copy()
    if selected_reason != "(All reasons)":
        df_filtered = df_filtered[df_filtered["reason"] == selected_reason]
    if exclude_negative:
        df_filtered = df_filtered[df_filtered["delta_days"] >= 0]

    if df_filtered.empty:
        st.warning("No data remains after applying filters.")
        st.stop()

    delta = df_filtered["delta_days"]
    desc = delta.describe(percentiles=[0.5, 0.9, 0.95, 0.99])
    total_valid = len(df_filtered)

    info_cols = st.columns(4)
    with info_cols[0]:
        st.markdown(f"**Official mailing date (Day 0):** `{official_date.date()}`")
    with info_cols[1]:
        st.markdown(f"**File:** `{selected_label}`")
    with info_cols[2]:
        st.markdown(f"**Valid redress dates:** {after_count}")
    with info_cols[3]:
        st.markdown(f"**Records used:** {total_valid}")

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
        st.table(
            pd.DataFrame(
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
        )

    # Business questions for single mailing
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

    # Textual summary
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
