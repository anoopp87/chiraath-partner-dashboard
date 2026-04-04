# Chiraath Partner Dashboard

Reads the Excel workbook and generates a static partner dashboard published via GitHub Pages.

---

## One-time setup

```bash
pip install -r requirements.txt
```

---

## Every time you update the data

**Step 1 — Replace the Excel file**

Drop the latest file into the `input/` folder, keeping the exact filename:
```
input/Chiraath - Business Summary.xlsx
```

**Step 2 — Rebuild the dashboard**

```bash
python src/build_dashboard.py
```

This writes the updated pages into `docs/`.

**Step 3 — Check it locally (optional)**

```bash
open docs/index.html
```

**Step 4 — Push to GitHub**

```bash
git add docs/ input/
git commit -m "Update dashboard - <month>"
git push
```

**Step 5 — Open the live site**

Wait ~30 seconds after pushing, then open:

👉 **https://anoopp87.github.io/chiraath-partner-dashboard/**

---

## Project structure

```
input/      ← drop updated Excel here (keep same filename)
src/        ← build script + HTML templates (edit these for layout changes)
docs/       ← generated output (committed and served by GitHub Pages)
```

## Notes

- Excel must have sheets named `Summary`, `Dashboard`, and `INVENTORY`
- If you rename sheets or move cells, update the **Config** section at the top of `src/build_dashboard.py`

---

## For developers

### Source files (`src/`)

**`build_dashboard.py`** — the only script that needs to be run. Does everything:
- Reads the Excel workbook using `openpyxl` and `pandas`
- Extracts KPIs, contribution/cash pool tables, monthly data, and inventory from specific cells/ranges
- Builds four Plotly charts (monthly bar, profit line, category qty stacked bar, pending value bar)
- Renders both HTML templates using Jinja2 and writes the output to `docs/`
- Cell/range addresses are all configurable at the top of the file under `# Config` — no need to dig into the logic for routine Excel layout changes

**`template.html`** — Jinja2 template for the main dashboard (`docs/index.html`):
- KPI cards (Total Purchases, Total Sales, Profit/Loss, Pending Stock Value)
- Contribution Summary table and Cash Pool Summary table with sticky headers, sticky first column, and green/red pay-vs-receive highlighting
- Four embedded Plotly charts
- All styling is plain CSS in a `<style>` block at the top; no external CSS framework

**`inventory_template.html`** — Jinja2 template for the inventory browser (`docs/inventory.html`):
- Search bar that filters inventory rows live in the browser (pure JS, no backend)
- Four inventory KPI cards (Total Purchased, In Stock, Low Stock, Units Sold)
- Rows highlighted amber for low stock (qty = 1), red for out of stock (qty = 0)

### Generated output (`docs/`)

| File | What it is |
|------|------------|
| `index.html` | Main partner dashboard — charts, KPIs, contribution and cash pool tables |
| `inventory.html` | Searchable inventory browser |
| `Chiraath-Business-Summary-Latest.xlsx` | Copy of the input Excel, served as the "Download" button on the dashboard |

`docs/` is committed to git and served directly by GitHub Pages — no build server needed.
