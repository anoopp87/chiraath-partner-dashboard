# Chiraath Partner Dashboard (Excel → HTML)

This project reads **two tabs** from your Excel and generates a styled **partner dashboard** as a static web page.

## What it generates
- `dist/index.html` (dashboard)
- `dist/Chiraath-Business-Summary-Latest.xlsx` (download link on the dashboard)

## Setup (Anaconda / normal Python)
From this folder:

```bash
pip install -r requirements.txt
```

## Build the dashboard
```bash
python src/build_dashboard.py
```

Open `dist/index.html` in your browser.

## Weekly update flow
1. Replace `input/Chiraath - Business Summary - Latest.xlsx` with the latest file (same name)
2. Run:
   ```bash
   python src/build_dashboard.py
   ```
3. Publish `dist/` to GitHub Pages (or any static host)

## Notes
- The script expects sheets named `Summary` and `Dashboard`
- If you change the sheet names or move the KPI cells, edit the **Config** section in `src/build_dashboard.py`.
