\
from __future__ import annotations

from pathlib import Path
from datetime import datetime, date
import math
import shutil

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from jinja2 import Template


# ---------------------------
# Config (edit if you rename tabs/cells)
# ---------------------------
SHEET_SUMMARY = "Summary"
SHEET_DASHBOARD = "Dashboard"

# Summary KPIs (cell addresses)
CELL_TOTAL_PURCHASES = "B3"
CELL_TOTAL_SALES_COMPLETED = "B4"
CELL_PROFIT_LOSS_COMPLETED = "B6"
CELL_PROFIT_STATUS = "B7"

# Dashboard cells
CELL_LAST_UPDATED = "B3"
CELL_PENDING_STOCK_VALUE = "E6"
CELL_QTY_SOLD = "G10"
CELL_QTY_PENDING = "G11"

# Dashboard tables (row/col ranges)
# Monthly table: header in row 9, data rows 10-21, cols A-D
MONTH_HEADER_ROW = 9
MONTH_DATA_ROWS = (10, 21)
MONTH_COLS = (1, 4)  # A-D

# Category qty: header row 46, data 47-49, cols F-H
CATQTY_HEADER_ROW = 46
CATQTY_DATA_ROWS = (47, 49)
CATQTY_COLS = (6, 8)  # F-H

# Pending value by category: header row 46, data 47-49, cols J-K
PENDVAL_HEADER_ROW = 46
PENDVAL_DATA_ROWS = (47, 49)
PENDVAL_COLS = (10, 11)  # J-K

# Contribution summary: header row 17, data 18-20, cols A-G
CONTRIB_HEADER_ROW = 17
CONTRIB_DATA_ROWS = (18, 20)
CONTRIB_COLS = (1, 7)  # A-G

# Cash pool summary: header row 25, data 26-28, cols A-F
CASH_HEADER_ROW = 25
CASH_DATA_ROWS = (26, 28)
CASH_COLS = (1, 7)  # A-F


# ---------------------------
# Helpers
# ---------------------------
def excel_serial_to_datetime(x) -> str:
    if isinstance(x, (datetime, date)):
        return x.strftime("%b %d, %Y")
    if isinstance(x, (int, float)) and x:
        d = datetime(1899, 12, 30) + pd.to_timedelta(float(x), unit="D")
        return d.strftime("%b %d, %Y")
    return "—"


def money0(x) -> str:
    try:
        if x is None:
            return "₹0"
        if isinstance(x, float) and math.isnan(x):
            return "₹0"
        return f"₹{float(x):,.0f}"
    except Exception:
        return str(x)


def money2(x) -> str:
    try:
        if x is None:
            return "₹0.00"
        if isinstance(x, float) and math.isnan(x):
            return "₹0.00"
        return f"₹{float(x):,.2f}"
    except Exception:
        return str(x)


def read_range(ws, min_row, max_row, min_col, max_col):
    return [
        [ws.cell(r, c).value for c in range(min_col, max_col + 1)]
        for r in range(min_row, max_row + 1)
    ]


def df_from_range(ws, header_row: int, data_rows: tuple[int, int], cols: tuple[int, int]) -> pd.DataFrame:
    headers = read_range(ws, header_row, header_row, cols[0], cols[1])[0]
    data = read_range(ws, data_rows[0], data_rows[1], cols[0], cols[1])
    return pd.DataFrame(data, columns=headers)


def format_df_currency(df: pd.DataFrame) -> tuple[list[dict], list[str]]:
    df2 = df.copy()
    for col in df2.columns:
        col_s = str(col)
        if "₹" in col_s or any(k in col_s.lower() for k in ["amount", "value", "paid", "balance", "target", "excess", "collected", "transferred", "share"]):
            df2[col] = df2[col].apply(money2)
    return df2.to_dict(orient="records"), list(df2.columns)


# ---------------------------
# Main builder
# ---------------------------
def build(input_xlsx: Path, template_path: Path, dist_dir: Path) -> None:
    dist_dir.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(input_xlsx, data_only=True)
    ws_sum = wb[SHEET_SUMMARY]
    ws_dash = wb[SHEET_DASHBOARD]

    # KPIs
    total_purchases = ws_sum[CELL_TOTAL_PURCHASES].value
    total_sales_completed = ws_sum[CELL_TOTAL_SALES_COMPLETED].value
    profit_loss = ws_sum[CELL_PROFIT_LOSS_COMPLETED].value
    profit_status = ws_sum[CELL_PROFIT_STATUS].value

    last_updated = excel_serial_to_datetime(ws_dash[CELL_LAST_UPDATED].value)
    pending_stock_value = ws_dash[CELL_PENDING_STOCK_VALUE].value
    qty_sold = ws_dash[CELL_QTY_SOLD].value
    qty_pending = ws_dash[CELL_QTY_PENDING].value

    # Tables
    contrib_df = df_from_range(ws_sum, CONTRIB_HEADER_ROW, CONTRIB_DATA_ROWS, CONTRIB_COLS)

    # Keep only Partner + last two columns
    contrib_df = contrib_df.iloc[:, [0, -2, -1]]


    cash_df = df_from_range(ws_sum, CASH_HEADER_ROW, CASH_DATA_ROWS, CASH_COLS)

    # Keep only Partner + last two columns
    cash_df = cash_df.iloc[:, [0, -2, -1]]

    month_df = df_from_range(ws_dash, MONTH_HEADER_ROW, MONTH_DATA_ROWS, MONTH_COLS)
    # Coerce numeric columns if present
    for c in ["Purchases (₹)", "Sales (₹)", "Profit (₹)"]:
        if c in month_df.columns:
            month_df[c] = pd.to_numeric(month_df[c], errors="coerce").fillna(0)

    cat_qty_df = df_from_range(ws_dash, CATQTY_HEADER_ROW, CATQTY_DATA_ROWS, CATQTY_COLS)
    for c in ["Qty Sold", "Qty Pending"]:
        if c in cat_qty_df.columns:
            cat_qty_df[c] = pd.to_numeric(cat_qty_df[c], errors="coerce").fillna(0)

    pending_val_df = df_from_range(ws_dash, PENDVAL_HEADER_ROW, PENDVAL_DATA_ROWS, PENDVAL_COLS)
    if "Pending Value (₹)" in pending_val_df.columns:
        pending_val_df["Pending Value (₹)"] = pd.to_numeric(pending_val_df["Pending Value (₹)"], errors="coerce").fillna(0)

    # Charts
    charts = {}
    PLOTLY_CONFIG = {"responsive": True, "displayModeBar": False}

    if {"Month", "Purchases (₹)", "Sales (₹)"}.issubset(set(month_df.columns)):
        fig_month = go.Figure()
        fig_month.add_trace(go.Bar(x=month_df["Month"], y=month_df["Purchases (₹)"], name="Purchases"))
        fig_month.add_trace(go.Bar(x=month_df["Month"], y=month_df["Sales (₹)"], name="Sales"))
        fig_month.update_layout(barmode="group", title="Monthly Purchases vs Sales")
        charts["month"] = pio.to_html(fig_month, include_plotlyjs="cdn", full_html=False, config=PLOTLY_CONFIG)
    else:
        charts["month"] = "<div style='color:#6b7280;font-size:13px'>Monthly chart unavailable (columns changed).</div>"

    if {"Month", "Profit (₹)"}.issubset(set(month_df.columns)):
        fig_profit = px.line(month_df, x="Month", y="Profit (₹)", markers=True, title="Monthly Profit (₹)")
        charts["profit"] = pio.to_html(fig_profit, include_plotlyjs=False, full_html=False, config=PLOTLY_CONFIG)
    else:
        charts["profit"] = "<div style='color:#6b7280;font-size:13px'>Profit chart unavailable (columns changed).</div>"

    if {"Category", "Qty Sold", "Qty Pending"}.issubset(set(cat_qty_df.columns)):
        fig_cat_qty = go.Figure()
        fig_cat_qty.add_trace(go.Bar(x=cat_qty_df["Category"], y=cat_qty_df["Qty Sold"], name="Qty Sold"))
        fig_cat_qty.add_trace(go.Bar(x=cat_qty_df["Category"], y=cat_qty_df["Qty Pending"], name="Qty Pending"))
        fig_cat_qty.update_layout(barmode="stack", title="Quantity by Category")
        charts["cat_qty"] = pio.to_html(fig_cat_qty, include_plotlyjs=False, full_html=False, config=PLOTLY_CONFIG)
    else:
        charts["cat_qty"] = "<div style='color:#6b7280;font-size:13px'>Category qty chart unavailable (columns changed).</div>"

    if {"Category", "Pending Value (₹)"}.issubset(set(pending_val_df.columns)):
        fig_pending_val = px.bar(pending_val_df, x="Category", y="Pending Value (₹)", title="Pending Stock Value by Category (₹)")
        charts["pending_val"] = pio.to_html(fig_pending_val, include_plotlyjs=False, full_html=False, config=PLOTLY_CONFIG)
    else:
        charts["pending_val"] = "<div style='color:#6b7280;font-size:13px'>Pending value chart unavailable (columns changed).</div>"

    contrib_records, contrib_cols = format_df_currency(contrib_df)
    cash_records, cash_cols = format_df_currency(cash_df)

    # Render HTML
    tpl = Template(template_path.read_text(encoding="utf-8"))
    excel_filename = "Chiraath-Business-Summary-Latest.xlsx"

    html = tpl.render(
        last_updated=last_updated,
        qty_sold=int(qty_sold) if qty_sold is not None else "—",
        qty_pending=int(qty_pending) if qty_pending is not None else "—",
        total_purchases=money0(total_purchases),
        total_sales_completed=money0(total_sales_completed),
        profit_loss=money0(profit_loss),
        profit_status=str(profit_status or ""),
        pending_stock_value=money2(pending_stock_value),
        chart_month=charts["month"],
        chart_profit=charts["profit"],
        chart_cat_qty=charts["cat_qty"],
        chart_pending_val=charts["pending_val"],
        contrib_cols=contrib_cols,
        contrib_records=contrib_records,
        cash_cols=cash_cols,
        cash_records=cash_records,
        excel_filename=excel_filename,
    )

    (dist_dir / "index.html").write_text(html, encoding="utf-8")

    # Copy excel for download button
    shutil.copyfile(input_xlsx, dist_dir / excel_filename)

    print(f"✅ Built dashboard: {dist_dir / 'index.html'}")
    print(f"⬇️  Excel download file: {dist_dir / excel_filename}")


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    input_xlsx = root / "input" / "Chiraath - Business Summary - Latest.xlsx"
    template_path = root / "src" / "template.html"
    dist_dir = root / "docs"

    if not input_xlsx.exists():
        raise FileNotFoundError(f"Missing input Excel at: {input_xlsx}")

    build(input_xlsx=input_xlsx, template_path=template_path, dist_dir=dist_dir)
