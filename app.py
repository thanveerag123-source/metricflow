import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# Excel export (charts + formatting)
import xlsxwriter


# =========================
# Column detection helpers
# =========================
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()


def _is_numeric_series(s: pd.Series) -> bool:
    x = pd.to_numeric(s, errors="coerce")
    return x.notna().mean() >= 0.6  # 60% numeric is enough


def auto_detect_columns(df: pd.DataFrame):
    cols = list(df.columns)
    norm = {_norm(c): c for c in cols}

    def find_any(keywords):
        for k in keywords:
            for n, orig in norm.items():
                if k in n:
                    return orig
        return None

    date_col = find_any(["inv date", "invoice date", "date", "order date", "payment date"])
    units_col = find_any(["qty", "quantity", "units", "unit sold", "units sold"])
    price_col = find_any(["unit price", "price", "price per unit", "rate"])
    region_col = find_any(["zone", "region", "area"])
    state_col = find_any(["state"])
    brand_col = find_any(["product", "brand", "item", "sku", "beverage"])

    # sales: avoid Sales Man / Salesperson etc.
    sales_candidates = []
    for c in cols:
        n = _norm(c)
        if any(k in n for k in ["sales", "revenue", "amount", "value", "total"]):
            if not any(bad in n for bad in ["sales man", "salesman", "sales person", "salesperson", "rep", "agent", "executive"]):
                sales_candidates.append(c)

    sales_col = None
    for c in sales_candidates:
        if _is_numeric_series(df[c]):
            sales_col = c
            break

    customer_col = find_any(["customer", "retailer", "client", "buyer", "store"])
    sales_person_col = find_any(["sales man", "salesman", "sales person", "salesperson", "executive", "rep", "agent"])

    return {
        "date": date_col,
        "sales": sales_col,
        "units": units_col,
        "price": price_col,
        "customer": customer_col,
        "sales_person": sales_person_col,
        "region": region_col,
        "state": state_col,
        "brand": brand_col,
    }


def to_num(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace(r"[\s,]", "", regex=True)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def load_best_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    best_name = None
    best_df = None
    best_score = -1

    for name in xls.sheet_names:
        df = xls.parse(name)
        if df is None or df.empty:
            continue

        m = auto_detect_columns(df)
        score = 0
        score += 4 if m["sales"] else 0
        score += 3 if m["date"] else 0
        score += 2 if m["customer"] else 0
        score += 2 if (m["region"] or m["state"]) else 0
        score += 2 if m["brand"] else 0
        score += 1 if m["units"] else 0
        score += 1 if m["price"] else 0
        score += min(5, int(len(df) / 200))

        if score > best_score:
            best_score = score
            best_name = name
            best_df = df.copy()

    return xls.sheet_names, best_name, best_df


# =========================
# Page setup
# =========================
st.set_page_config(
    page_title="MetricFlow",
    page_icon="📊",
    layout="wide",
)


# =========================
# Premium CSS (FIXES: top gap, accent, mobile, pointer)
# =========================
def inject_css(theme: str, accent_name: str):
    # Accent palettes (we also keep RGBA for smooth gradients)
    accents = {
        "Purple": ("#A855F7", "rgba(168,85,247,0.22)"),
        "Blue": ("#3B82F6", "rgba(59,130,246,0.22)"),
        "Emerald": ("#10B981", "rgba(16,185,129,0.22)"),
        "Rose": ("#F43F5E", "rgba(244,63,94,0.22)"),
    }
    accent_hex, accent_rgba = accents.get(accent_name, ("#A855F7", "rgba(168,85,247,0.22)"))

    if theme == "Dark":
        bg = "#070A12"
        panel = "rgba(255,255,255,0.06)"
        text = "#E5E7EB"
        muted = "#9CA3AF"
        border = "rgba(255,255,255,0.12)"
        shadow = "rgba(0,0,0,0.30)"
    else:
        bg = "#F7F7FB"
        panel = "rgba(15,23,42,0.06)"
        text = "#0F172A"
        muted = "#475569"
        border = "rgba(15,23,42,0.12)"
        shadow = "rgba(15,23,42,0.12)"

    st.markdown(
        f"""
        <style>
          :root {{
            --bg: {bg};
            --panel: {panel};
            --text: {text};
            --muted: {muted};
            --border: {border};
            --accent: {accent_hex};
            --accent_rgba: {accent_rgba};
            --radius: 18px;
          }}

          /* ===== FIX #1: Remove Streamlit top dark header gap ===== */
          header[data-testid="stHeader"] {{
            background: transparent !important;
          }}
          [data-testid="stToolbar"] {{
            visibility: hidden !important;
            height: 0px !important;
          }}
          [data-testid="stAppViewContainer"] {{
            background: transparent !important;
          }}
          .stApp {{
            background:
              radial-gradient(1200px 600px at 18% 10%, var(--accent_rgba), transparent 58%),
              radial-gradient(900px 520px at 80% 18%, rgba(59,130,246,0.14), transparent 56%),
              var(--bg);
            color: var(--text);
          }}

          /* Container spacing */
          .block-container {{
            padding-top: 1.25rem !important;
            padding-bottom: 2.25rem !important;
            max-width: 1200px;
          }}

          /* Pointer cursor (no I-beam on selects) */
          button, [role="button"], a {{
            cursor: pointer !important;
          }}
          [data-baseweb="select"] * {{
            cursor: pointer !important;
          }}
          [data-baseweb="select"] input {{
            caret-color: transparent !important;
          }}

          /* Hero */
          .hero {{
            padding: 22px 24px;
            border-radius: 22px;
            border: 1px solid var(--border);
            background:
              radial-gradient(900px 320px at 18% 0%, var(--accent_rgba), transparent 60%),
              radial-gradient(760px 320px at 82% 0%, rgba(59,130,246,0.16), transparent 60%),
              linear-gradient(135deg, rgba(255,255,255,0.14), rgba(255,255,255,0.06));
            box-shadow: 0 20px 60px {shadow};
            backdrop-filter: blur(14px);
          }}
          .muted {{
            color: var(--muted) !important;
          }}

          /* KPI Grid */
          .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 12px;
          }}
          .kpi {{
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 14px 14px;
            box-shadow: 0 10px 26px {shadow};
            backdrop-filter: blur(10px);
          }}
          .kpi-title {{
            font-size: 13px;
            color: var(--muted);
            margin-bottom: 6px;
          }}
          .kpi-value {{
            font-size: 22px;
            font-weight: 800;
            letter-spacing: 0.2px;
          }}

          /* Glass panels */
          .glass {{
            background: linear-gradient(135deg, rgba(255,255,255,0.10), rgba(255,255,255,0.05));
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px 16px;
            box-shadow: 0 12px 34px {shadow};
            backdrop-filter: blur(14px);
          }}

          /* Smooth buttons + accent glow */
          .stButton > button, .stDownloadButton > button {{
            border-radius: 14px !important;
            border: 1px solid var(--border) !important;
            background: linear-gradient(135deg, rgba(255,255,255,0.14), rgba(255,255,255,0.06)) !important;
            color: var(--text) !important;
            padding: 10px 14px !important;
            transition: transform .10s ease, box-shadow .18s ease, border-color .18s ease !important;
            box-shadow: 0 10px 26px {shadow} !important;
          }}
          .stButton > button:hover, .stDownloadButton > button:hover {{
            transform: translateY(-1px);
            border-color: var(--accent) !important;
            box-shadow: 0 16px 42px {shadow} !important;
          }}
          .stDownloadButton > button {{
            background: linear-gradient(135deg, var(--accent_rgba), rgba(59,130,246,0.18)) !important;
            border-color: rgba(255,255,255,0.18) !important;
          }}

          /* ===== FIX #3: Mobile responsiveness ===== */
          @media (max-width: 900px) {{
            .block-container {{ padding-left: 1rem !important; padding-right: 1rem !important; }}
            .kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
          }}
          @media (max-width: 520px) {{
            .kpi-grid {{ grid-template-columns: 1fr; }}
            .hero {{ padding: 18px 16px; }}
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    return accent_hex


# =========================
# Excel export
# =========================
def export_excel_report(df_export: pd.DataFrame, meta: dict, theme: str, accent_hex: str) -> bytes:
    output = io.BytesIO()
    m = meta["mapping"]

    # ensure numeric/date cleaned
    if m["sales"] and m["sales"] in df_export.columns:
        df_export[m["sales"]] = to_num(df_export[m["sales"]])
    if m["units"] and m["units"] in df_export.columns:
        df_export[m["units"]] = to_num(df_export[m["units"]])
    if m["price"] and m["price"] in df_export.columns:
        df_export[m["price"]] = to_num(df_export[m["price"]])
    if m["date"] and m["date"] in df_export.columns:
        df_export[m["date"]] = pd.to_datetime(df_export[m["date"]], errors="coerce")

    # Fallback sales
    sales_col = m["sales"]
    if sales_col and sales_col in df_export.columns:
        sales_sum = float(df_export[sales_col].sum(skipna=True))
        sales_nonnull = int(df_export[sales_col].notna().sum())
    else:
        sales_sum, sales_nonnull = 0.0, 0

    if (sales_sum == 0.0 or sales_nonnull == 0) and (m["units"] and m["price"]):
        df_export["_calc_sales"] = df_export[m["units"]] * df_export[m["price"]]
        sales_col = "_calc_sales"

    total_sales = float(df_export[sales_col].sum(skipna=True)) if sales_col else 0.0
    orders = int(len(df_export))
    units = float(df_export[m["units"]].sum(skipna=True)) if m["units"] else 0.0
    avg_price = float(df_export[m["price"]].mean(skipna=True)) if m["price"] else 0.0
    customers = int(df_export[m["customer"]].astype(str).nunique()) if m["customer"] else 0

    # Monthly trend
    monthly = None
    if m["date"] and sales_col:
        tmp = df_export.dropna(subset=[m["date"]]).copy()
        if not tmp.empty:
            tmp["Month"] = tmp[m["date"]].dt.to_period("M").dt.to_timestamp()
            monthly = tmp.groupby("Month", as_index=False)[sales_col].sum().sort_values("Month")

    reg_col = m["region"] or m["state"]
    by_region = None
    if reg_col and sales_col:
        by_region = df_export.groupby(reg_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(12)

    by_brand = None
    if m["brand"] and sales_col:
        by_brand = df_export.groupby(m["brand"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(12)

    top_customers = None
    if m["customer"] and sales_col:
        top_customers = df_export.groupby(m["customer"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(10)

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        # formats
        header_bg = "#111827"
        header_fg = "#FFFFFF"
        muted = "#64748B" if theme == "Light" else "#9CA3AF"
        card_bg = "#F1F5F9" if theme == "Light" else "#0F172A"
        dash_bg = "#FFFFFF" if theme == "Light" else "#0B1220"

        fmt_header = wb.add_format({"bold": True, "bg_color": header_bg, "font_color": header_fg})
        fmt_h1 = wb.add_format({"bold": True, "font_size": 18})
        fmt_muted = wb.add_format({"font_color": muted})
        fmt_money0 = wb.add_format({"num_format": "#,##0"})
        fmt_money2 = wb.add_format({"num_format": "#,##0.00"})
        fmt_card = wb.add_format({"bg_color": card_bg, "border": 1, "border_color": "#1F2937"})
        fmt_card_title = wb.add_format({"bold": True})
        fmt_card_value = wb.add_format({"bold": True, "font_size": 16})

        # Data
        df_export.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        ws_data.freeze_panes(1, 0)
        ws_data.write_row(0, 0, list(df_export.columns), fmt_header)

        # Summary
        summary = pd.DataFrame(
            {
                "Metric": ["Total Sales", "Units Sold", "Avg Price/Unit", "Orders (rows)", "Customers/Retailers"],
                "Value": [total_sales, units, avg_price, orders, customers],
            }
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)
        ws_sum = writer.sheets["Summary"]
        ws_sum.freeze_panes(1, 0)
        ws_sum.write_row(0, 0, ["Metric", "Value"], fmt_header)
        ws_sum.set_column(0, 0, 24)
        ws_sum.set_column(1, 1, 18, fmt_money2)

        # Optional sheets
        if monthly is not None:
            monthly.rename(columns={sales_col: "Sales"}).to_excel(writer, sheet_name="Monthly_Trend", index=False)
            ws = writer.sheets["Monthly_Trend"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Month", "Sales"], fmt_header)

        if by_region is not None:
            by_region.rename(columns={sales_col: "Sales"}).to_excel(writer, sheet_name="By_Region", index=False)
            ws = writer.sheets["By_Region"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, [reg_col, "Sales"], fmt_header)

        if by_brand is not None:
            by_brand.rename(columns={sales_col: "Sales"}).to_excel(writer, sheet_name="By_Brand", index=False)
            ws = writer.sheets["By_Brand"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, [m["brand"], "Sales"], fmt_header)

        if top_customers is not None:
            top_customers.rename(columns={sales_col: "Sales"}).to_excel(writer, sheet_name="Top_Customers", index=False)
            ws = writer.sheets["Top_Customers"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, [m["customer"], "Sales"], fmt_header)

        # Dashboard sheet
        ws_dash = wb.add_worksheet("Dashboard")
        ws_dash.hide_gridlines(2)  # gridlines OFF only here
        ws_dash.set_tab_color(accent_hex)
        ws_dash.set_column("A:A", 26)
        ws_dash.set_column("B:K", 16)
        ws_dash.set_default_row(20)

        ws_dash.write("A1", "Dashboard", fmt_h1)
        ws_dash.write("A2", f"Built from: {meta['source_name']} • Sheet: {meta['sheet']}", fmt_muted)

        def card(cell_range, title, value):
            ws_dash.merge_range(cell_range, "", fmt_card)
            tl = cell_range.split(":")[0]
            col = re.sub(r"[0-9]", "", tl)
            row = int(re.sub(r"[^0-9]", "", tl))
            ws_dash.write(f"{col}{row}", title, fmt_card_title)
            ws_dash.write(f"{col}{row+1}", value, fmt_card_value)

        card("A4:B6", "Total Sales", f"{total_sales:,.0f}")
        card("C4:D6", "Units Sold", f"{units:,.0f}")
        card("E4:F6", "Avg Price/Unit", f"{avg_price:,.2f}")
        card("G4:H6", "Orders", f"{orders:,}")
        card("I4:J6", "Customers", f"{customers:,}")

        ws_dash.write("A8", "Auto Insights:", fmt_muted)
        for i, line in enumerate(meta.get("insights", [])[:4]):
            ws_dash.write(8 + i, 0, f"• {line}", fmt_muted)

        # Charts area (3-4 charts)
        row_top = 13

        if monthly is not None and len(monthly) >= 2:
            chart = wb.add_chart({"type": "line"})
            n = len(monthly)
            chart.add_series({
                "categories": f"=Monthly_Trend!$A$2:$A${n+1}",
                "values": f"=Monthly_Trend!$B$2:$B${n+1}",
                "line": {"color": accent_hex, "width": 2.75},
            })
            chart.set_title({"name": "Monthly Sales Trend"})
            chart.set_legend({"none": True})
            chart.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"A{row_top}", chart)

        if by_region is not None and len(by_region) >= 2:
            chart2 = wb.add_chart({"type": "column"})
            n = len(by_region)
            chart2.add_series({
                "categories": f"=By_Region!$A$2:$A${n+1}",
                "values": f"=By_Region!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart2.set_title({"name": "Sales by Region/State"})
            chart2.set_legend({"none": True})
            chart2.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"F{row_top}", chart2)

        if by_brand is not None and len(by_brand) >= 2:
            chart3 = wb.add_chart({"type": "bar"})
            n = len(by_brand)
            chart3.add_series({
                "categories": f"=By_Brand!$A$2:$A${n+1}",
                "values": f"=By_Brand!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart3.set_title({"name": "Sales by Brand/Product"})
            chart3.set_legend({"none": True})
            chart3.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"A{row_top + 18}", chart3)

        if top_customers is not None and len(top_customers) >= 2:
            chart4 = wb.add_chart({"type": "column"})
            n = len(top_customers)
            chart4.add_series({
                "categories": f"=Top_Customers!$A$2:$A${n+1}",
                "values": f"=Top_Customers!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart4.set_title({"name": "Top Customers"})
            chart4.set_legend({"none": True})
            chart4.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"F{row_top + 18}", chart4)

    return output.getvalue()


# =========================
# App UI
# =========================
if "theme" not in st.session_state:
    st.session_state.theme = "Dark"
if "accent_name" not in st.session_state:
    st.session_state.accent_name = "Purple"

accent_hex = inject_css(st.session_state.theme, st.session_state.accent_name)

st.markdown(
    """
    <div class="hero">
      <div style="display:flex; align-items:center; justify-content:space-between; gap:16px; flex-wrap:wrap;">
        <div>
          <div style="font-size:34px; font-weight:900; letter-spacing:-0.4px;">📊 MetricFlow</div>
          <div class="muted" style="margin-top:6px;">
            Upload an Excel file → auto-detect columns → generate KPIs, charts, and a clean Excel dashboard (gridlines OFF on Dashboard sheet only).
          </div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.session_state.theme = st.selectbox("Theme", ["Dark", "Light"], index=0 if st.session_state.theme == "Dark" else 1)
with c2:
    st.session_state.accent_name = st.selectbox("Accent", ["Purple", "Blue", "Emerald", "Rose"],
                                                index=["Purple","Blue","Emerald","Rose"].index(st.session_state.accent_name))
with c3:
    st.markdown(
        '<div class="muted" style="padding-top:30px;">Tip: If Sales is missing, MetricFlow auto-calculates it using Quantity × Unit Price when possible.</div>',
        unsafe_allow_html=True,
    )

accent_hex = inject_css(st.session_state.theme, st.session_state.accent_name)

st.write("")
uploaded = st.file_uploader("Upload Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])

if not uploaded:
    st.info("Upload a file to generate a dashboard.")
    st.stop()

sheet_names, best_sheet, df = load_best_sheet(uploaded)
if df is None or df.empty:
    st.error("Couldn't find usable data in this file.")
    st.stop()

mapping = auto_detect_columns(df)
df2 = df.copy()

# Clean numeric/date
if mapping["sales"] and mapping["sales"] in df2.columns:
    df2[mapping["sales"]] = to_num(df2[mapping["sales"]])
if mapping["units"] and mapping["units"] in df2.columns:
    df2[mapping["units"]] = to_num(df2[mapping["units"]])
if mapping["price"] and mapping["price"] in df2.columns:
    df2[mapping["price"]] = to_num(df2[mapping["price"]])
if mapping["date"] and mapping["date"] in df2.columns:
    df2[mapping["date"]] = pd.to_datetime(df2[mapping["date"]], errors="coerce")

# Fallback sales
sales_col = mapping["sales"]
if sales_col and sales_col in df2.columns:
    sales_sum = float(df2[sales_col].sum(skipna=True))
    sales_nonnull = int(df2[sales_col].notna().sum())
else:
    sales_sum, sales_nonnull = 0.0, 0

if (sales_sum == 0.0 or sales_nonnull == 0) and mapping["units"] and mapping["price"]:
    df2["_calc_sales"] = df2[mapping["units"]] * df2[mapping["price"]]
    sales_col = "_calc_sales"

# KPIs
total_sales = float(df2[sales_col].sum(skipna=True)) if sales_col else 0.0
orders = int(len(df2))
units_sold = float(df2[mapping["units"]].sum(skipna=True)) if mapping["units"] else 0.0
avg_price = float(df2[mapping["price"]].mean(skipna=True)) if mapping["price"] else 0.0
customers = int(df2[mapping["customer"]].astype(str).nunique()) if mapping["customer"] else 0

# Insights
insights = []
if total_sales > 0:
    insights.append(f"Total sales recorded: {total_sales:,.0f} across {orders:,} rows.")

if mapping["brand"]:
    by_brand_tmp = df2.groupby(mapping["brand"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False)
    if not by_brand_tmp.empty:
        top = by_brand_tmp.iloc[0]
        insights.append(f"Top brand/product by sales: {top[mapping['brand']]} ({top[sales_col]:,.0f}).")

if mapping["customer"]:
    by_cust = df2.groupby(mapping["customer"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False)
    if len(by_cust) >= 3 and total_sales > 0:
        top3_share = by_cust.head(3)[sales_col].sum() / max(1.0, total_sales)
        insights.append(f"Top 3 customers contribute ~{top3_share:.0%} of total sales (concentration).")

reg_col = mapping["region"] or mapping["state"]
if reg_col:
    by_reg = df2.groupby(reg_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False)
    if not by_reg.empty:
        top = by_reg.iloc[0]
        insights.append(f"Strongest region/state: {top[reg_col]} ({top[sales_col]:,.0f}).")

# KPI cards
st.markdown(
    f"""
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-title">Total Sales</div><div class="kpi-value">{total_sales:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Units Sold</div><div class="kpi-value">{units_sold:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Avg Price / Unit</div><div class="kpi-value">{avg_price:,.2f}</div></div>
      <div class="kpi"><div class="kpi-title">Orders (rows)</div><div class="kpi-value">{orders:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Customers/Retailers</div><div class="kpi-value">{customers:,.0f}</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")
with st.expander("✨ View Insights", expanded=False):
    if insights:
        for line in insights:
            st.markdown(f"- {line}")
    else:
        st.info("Not enough fields to generate deeper insights. Add Date + Sales + Customer/Region/Brand for richer results.")

st.write("")
st.subheader("Dashboard")
st.caption("Charts only appear if required columns exist in your file.")

colA, colB = st.columns(2)
if mapping["date"]:
    tmp = df2.dropna(subset=[mapping["date"]]).copy()
    if not tmp.empty:
        tmp["Month"] = tmp[mapping["date"]].dt.to_period("M").dt.to_timestamp()
        monthly = tmp.groupby("Month", as_index=False)[sales_col].sum().sort_values("Month")
        with colA:
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.markdown("**Monthly Sales Trend**")
            st.line_chart(monthly.set_index("Month")[sales_col])
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        with colA:
            st.warning("Date column exists, but values couldn’t be parsed.")
else:
    with colA:
        st.warning("No usable date column detected for trend.")

reg_col = mapping["region"] or mapping["state"]
if reg_col:
    by_region = df2.groupby(reg_col, as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(12)
    with colB:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown(f"**Sales by {reg_col}**")
        st.bar_chart(by_region.set_index(reg_col)[sales_col])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colB:
        st.warning("No Region/State column detected.")

st.write("")
colC, colD = st.columns(2)

if mapping["customer"]:
    top_customers = df2.groupby(mapping["customer"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(10)
    with colC:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Top Customers/Retailers**")
        st.dataframe(top_customers, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colC:
        st.warning("No Customer/Retailer column detected.")

if mapping["brand"]:
    by_brand = df2.groupby(mapping["brand"], as_index=False)[sales_col].sum().sort_values(sales_col, ascending=False).head(12)
    with colD:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Brand/Product**")
        st.bar_chart(by_brand.set_index(mapping["brand"])[sales_col])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colD:
        st.warning("No Brand/Product column detected.")

st.write("")
st.subheader("Excel Export")
st.caption("Exports Data + Summary + Dashboard + optional tables. Dashboard gridlines OFF only.")

meta = {
    "source_name": uploaded.name,
    "sheet": best_sheet,
    "mapping": mapping,
    "insights": insights,
}

excel_bytes = export_excel_report(df2.copy(), meta, st.session_state.theme, accent_hex)

st.download_button(
    "⬇️ Download Excel Report",
    data=excel_bytes,
    file_name="MetricFlow_Dashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Advanced (optional): show auto-detected columns"):
    st.json(mapping)
    st.write(f"Auto-selected sheet: **{best_sheet}** (out of {sheet_names})")

st.write("")
st.subheader("Preview")
st.dataframe(df2.head(100), use_container_width=True)