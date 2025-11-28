# Redress Analyzer — Project Specification

## 1. Purpose

Internal Streamlit tool that analyzes how many days pass between the **official mailing date** of a campaign and the **redress capture date** (undeliverable event). It answers:

> How long do we wait after Day 0 until X% of all redresses have arrived, so we can clean the next mailing list?

The UI stays in English for wider team usage.

## 2. Minimal Data Model (Privacy Compliant)

Each Excel file contains only two columns starting at header row 4:

| GRUND_UNZUSTELLBARKEIT | REDRESSENERFASSUNG_DATUM |
|------------------------|--------------------------|
| Unknown at address     | 2025-03-18               |
| Moved / new address    | 2025-03-19               |

No personal data (addresses, IDs, ZIP) is present.

### Mandatory metadata

Row 1:

| A1                    | B1         |
|-----------------------|------------|
| OFFICIAL_MAILING_DATE | 2025-03-15 |

- `A1` must literally be `OFFICIAL_MAILING_DATE`
- `B1` contains the Day 0 date (Excel or ISO)
- Rows 2–3 can be blank

## 3. File Naming Convention

```
{CampaignName}_{YYYY-MM}.xlsx
```

Examples:

- `Magazin_043_2025-01.xlsx`
- `Magazin_044_2025-03.xlsx`
- `Must_Haves_2024-05.xlsx`

The suffix is used for chronological sorting (latest first).

## 4. Features

- **Single mailing view**
  - File selection (radio), ordered by date suffix
  - Reads Day 0 from A1/B1 (manual fallback if missing)
  - Optional filter by undeliverable reason
  - Optional exclusion of negative deltas
  - Distribution table (expander), summary stats (expander)
  - Charts: frequency and cumulative coverage
  - KPIs: median, 95th, 99th percentile
  - Business questions: days for X% coverage; coverage after N days
  - Textual summary with recommendation

- **All mailings view**
  - Loads all Excel files and aggregates them
  - Filters: reason, exclude negatives, weighting choice
    - Weighting: by records (default) or average of campaign percentiles (shown as caption)
  - KPIs: median/95th/99th across all filtered data
  - Charts: frequency, cumulative coverage, mailing comparison (median)
  - Tables: distribution, summary stats, per-mailing summary (median/p95/p99, counts, dates)
  - Business questions: days for X% coverage; coverage after N days across all mailings

## 5. Folder Structure

```
redress-analyzer/
+- app/
¦  +- redress_analyzer_app.py
+- data/
¦  +- Magazin_043_2025-01.xlsx
¦  +- Must_Haves_2024-05.xlsx
¦  +- ...
+- docs/
¦  +- PROJECT.md
+- requirements.txt
+- README.md
```

## 6. Dependencies

- streamlit
- pandas
- numpy
- openpyxl

## 7. Completion Criteria

- Can read all `.xlsx` in `/data`, parse Day 0, and required columns
- Handles single and all-mailings modes with filters
- Provides distributions, coverage charts, KPIs, and business-question sliders
- Produces recommendation-ready summary without exposing personal data
