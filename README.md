# RDZ Motors — Production Intelligence Dashboard

A modern, fully-local car factory production dashboard with an Excel automation
backend. Built for daily, monthly and yearly manufacturing reviews and live
presentations. Runs 100% offline in your browser — no server required.

![Theme](https://img.shields.io/badge/theme-black%20%2F%20white%20%2F%20deep%20purple-6d28d9)
![Python](https://img.shields.io/badge/python-3.9%2B-8b5cf6)
![License](https://img.shields.io/badge/license-MIT-a78bfa)

## ✨ Features

- **Excel automation** — one Python script builds a formatted `factory_data.xlsx`
  workbook (Daily / Monthly / Yearly sheets) with colored efficiency cells,
  SUM/AVERAGE totals and native Excel charts.
- **Modern dashboard** — dark black/deep-purple UI with the **Inter** font.
- **Info-page buttons** — switch between 📅 Daily, 🗓️ Monthly and 📊 Yearly views.
- **Real-time elements** — live clock, pulsing "LIVE FEED" badge and a live line
  output counter that keeps ticking.
- **KPI cards** — Units Produced, Target Attainment, Avg Efficiency, Defects /
  Quality, Downtime and Live Output.
- **Interactive charts** (Chart.js) — Production-vs-Target trend with a purple
  gradient fill, plus a Model Mix doughnut. Filter the trend by car model.
- **Load Excel** — edit the workbook, click the button, and the dashboard
  re-reads your file live via SheetJS.
- **Present mode** — printer-friendly output for meetings.

## 🚀 Quick start

```powershell
# 1. Install the one dependency
pip install openpyxl

# 2. Generate the Excel workbook + dashboard data
python generate_data.py

# 3. Open the dashboard
start dashboard.html
```

That's it. Double-clicking `dashboard.html` works too — it ships with generated
`data.js` so it renders instantly.

## 🔄 Updating the numbers

- **Re-run** `python generate_data.py` for fresh sample data, **or**
- **Edit** `factory_data.xlsx` with your real figures, open the dashboard and
  click **📁 Load Excel** to sync instantly.

## 📁 Project structure

```
generate_data.py     # Excel + data.js generator (the automation)
dashboard.html       # Self-contained dashboard (HTML/CSS/JS)
data.js              # Auto-generated dashboard data (do not edit by hand)
factory_data.xlsx    # Auto-generated Excel workbook
```

## 🛠 Requirements

- Python 3.9+
- [`openpyxl`](https://pypi.org/project/openpyxl/) (Excel export)
- A modern browser (Chart.js and SheetJS load from CDN on first use)

## 📄 License

MIT — see [LICENSE](LICENSE).
