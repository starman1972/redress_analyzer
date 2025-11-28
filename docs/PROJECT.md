# Redress Analyzer – Project Specification

## 1. Purpose

The Redress Analyzer is a Streamlit-based internal tool that analyzes how many days pass between the **official mailing date** of a magazine shipment and the **redress capture date** (the date when a mailing is marked as undeliverable by the provider).

Business question:

> **How long do we need to wait (after the official mailing date) until X% (e.g. 95%) of all redresses have arrived, so we can safely generate the next mailing list?**

This analysis prevents costs caused by sending the next magazine to addresses that are already known to be undeliverable.

The tool is fully in English because non-German-speaking colleagues will use it.


## 2. Minimal Data Model (Privacy Compliant)

Each Excel file must contain only **two columns** starting at **row 4**:

| GRUND_UNZUSTELLBARKEIT | REDRESSENERFASSUNG_DATUM |
|------------------------|--------------------------|
| Unknown at address     | 2025-03-18               |
| Moved / new address    | 2025-03-19               |
| ...                    | ...                      |

All personal or sensitive fields (addresses, ZIP, city, IDs, etc.) are removed.

### Mandatory metadata at the top:

Row 1 must contain:

| A1                      | B1          |
|-------------------------|-------------|
| OFFICIAL_MAILING_DATE   | 2025-03-15  |

- `A1 = OFFICIAL_MAILING_DATE` (literal string)
- `B1 = official mailing date` (Excel date or ISO format)

This defines **Day 0** for delta calculations.

Row 2–3 can be blank or unused.


## 3. File Naming Convention

Files in the `/data` folder must follow:

```

{MagazineName}_{YYY-MM}.xlsx

```

Examples:

```

Magazin_043_2025-01.xlsx
Magazin_044_2025-03.xlsx
Must_Haves_2024-05.xlsx

```

### Why this convention?

- Ensures correct chronological sorting
- Clear campaign identification
- Works automatically in the Streamlit dropdown


## 4. How the App Works (MVP)

### Step 1 — User selects a file
- Streamlit lists all `.xlsx` files in `data/`
- Example UI:  
  `Select a mailing file: [Magazin_043_2025-01.xlsx]`

### Step 2 — The app extracts metadata
- Reads `A1` → must equal `"OFFICIAL_MAILING_DATE"`
- Reads `B1` → date for baseline

### Step 3 — The app loads the table
- Reads the data starting at header row = 4 (`header=3` in pandas)
- Required columns:
  - `GRUND_UNZUSTELLBARKEIT`
  - `REDRESSENERFASSUNG_DATUM`

### Step 4 — Compute delta days
```

delta_days = REDRESSENERFASSUNG_DATUM - OFFICIAL_MAILING_DATE

```

### Step 5 — Output statistics
The app shows:

- Table: distribution of delta days
- Table: cumulative percentage
- Percentiles (50%, 90%, 95%, 99%)
- Bar chart (frequency)
- Line chart (cumulative)
- Two interactive controls:
  - “How long to reach X% coverage?”
  - “What coverage if we wait N days?”
- Summary block with recommendation text


## 5. Future Extensions (Phase 2)

- Filter by reason (Grund)
- Per-reason delta analysis
- Combine multiple files for a yearly view
- Export CSV or PDF summary
- Optional: Auto-detect column positions


## 6. Folder Structure

```

redress-analyzer/
│
├─ app/
│  └─ redress_app.py
│
├─ data/
│  ├─ Magazin_043_2025-01.xlsx
│  ├─ Magazin_044_2025-03.xlsx
│  ├─ Must_Haves_2024-05.xlsx
│  └─ ...
│
├─ docs/
│  └─ PROJECT.md
│
├─ requirements.txt
└─ README.md

```

## 7. Dependencies

```

streamlit
pandas
numpy
openpyxl

```

## 8. MVP Completion Criteria

The MVP is complete when the app can:

- Load a file from `/data`
- Read A1/B1 correctly
- Analyze redress deltas
- Produce:
  - Distribution
  - Cumulative curve
  - Percentiles
  - Recommendation text
- Without any manual input from the user
```