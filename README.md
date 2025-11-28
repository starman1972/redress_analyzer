# Redress Analyzer

Streamlit app to analyze undeliverable mailings ("redresses"). It calculates how many days pass between the official mailing date (Day 0) and each redress capture date, then answers how long to wait to reach a chosen coverage.

## Features
- Single mailing view: KPIs (median/95th/99th), frequency & cumulative charts, filters (reason, exclude negatives), business-question sliders.
- All mailings view: aggregates every file in `data/`, common filters, KPIs, comparison chart by mailing, and the same business-question sliders.
- Automatic parsing of Day 0 from `OFFICIAL_MAILING_DATE` in cell B1; manual fallback when missing.
- Required columns starting at header row 4: `GRUND_UNZUSTELLBARKEIT`, `REDRESSENERFASSUNG_DATUM`.

## Project structure
```
app/redress_analyzer_app.py
config/ (optional)
data/              # place .xlsx files here
docs/PROJECT.md
README.md
requirements.txt
```

## Getting started
1. Create/activate a virtual env (optional but recommended).
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Add your Excel files to `data/` using the naming convention `{CampaignName}_{YYYY-MM}.xlsx` and metadata `OFFICIAL_MAILING_DATE` in A1/B1.
4. Run the app:
   ```bash
   streamlit run app/redress_analyzer_app.py
   ```
5. Use the sidebar to switch between �All mailings� and �Single mailing�, apply filters, and read the charts/answers.

## Notes
- No personal data is stored or shown; files are reduced to reason + redress date plus the mailing date in the header.
- Negative deltas (redress before Day 0) can be excluded via checkbox.
- Weighting option in All-mailings view: by records (default) or average of campaign percentiles (shown as caption).
