# PMS Tracker

A self-updating website listing every PMS investment approach reported to APMI,
with returns, AUM and portfolio turnover.

- `apmi_scraper.py` — scrapes the APMI IA Performance and IA Turnover reports
- `.github/workflows/update-data.yml` — runs the scraper on the 10th, 20th and 28th of each month
- `docs/index.html` — the website (served by GitHub Pages)
- `docs/data.json` — the data the website reads (overwritten by each run)
- `docs/apmi_pms_data.xlsx` — the same data as an Excel file

Run locally: `pip install -r requirements.txt && python apmi_scraper.py --json docs/data.json`

Data source: Association of Portfolio Managers in India (apmiindia.org). For information only; not investment advice.
