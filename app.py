import io
import re
import numpy as np
import pandas as pd
import streamlit as st


# =========================
# Utils
# =========================
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()


def to_num(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace(r"[\s,]", "", regex=True)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def _is_numeric_series(s: pd.Series) -> bool:
    x = pd.to_numeric(s, errors="coerce")
    return x.notna().mean() >= 0.6


def hex_to_rgb(hex_color: str):
    h = hex_color.strip().lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def rgba(hex_color: str, a: float):
    r, g, b = hex_to_rgb(hex_color)
    return f"rgba({r},{g},{b},{a})"


def auto_detect_columns(df: pd.DataFrame):
    cols = list(df.columns)
    norm_map = {_norm(c): c for c in cols}

    def find_any(keywords):
        for k in keywords:
            for n, orig in norm_map.items():
                if k in n:
                    return orig
        return None

    date_col = find_any(["inv date", "invoice date", "date", "order date", "payment date", "created date"])
    units_col = find_any(["qty", "quantity", "units", "unit sold", "units sold", "no of units", "pieces"])
    price_col = find_any(["unit price", "price", "price per unit", "rate", "mrp", "selling price"])

    region_col = find_any(["zone", "region", "area", "territory", "sales zone"])
    state_col = find_any(["state", "province"])
    brand_col = find_any(["product", "brand", "item", "sku", "product name", "item name"])

    customer_col = find_any(["customer", "retailer", "client", "buyer", "store", "outlet", "shop"])
    sales_person_col = find_any(["sales man", "salesman", "sales person", "salesperson", "executive", "rep", "agent"])

    # Sales numeric column (avoid Sales Man)
    sales_candidates = []
    for c in cols:
        n = _norm(c)
        if any(k in n for k in ["sales", "revenue", "amount", "value", "total", "net", "gross"]):
            if not any(bad in n for bad in ["sales man", "salesman", "sales person", "salesperson", "rep", "agent", "executive"]):
                sales_candidates.append(c)

    sales_col = None
    for c in sales_candidates:
        if _is_numeric_series(df[c]):
            sales_col = c
            break

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


def prepare_df(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    out = df.copy()

    units = to_num(out[mapping["units"]]) if mapping.get("units") else pd.Series(np.nan, index=out.index)
    price = to_num(out[mapping["price"]]) if mapping.get("price") else pd.Series(np.nan, index=out.index)

    sales = None
    if mapping.get("sales"):
        s = to_num(out[mapping["sales"]])
        if s.notna().mean() >= 0.6:
            sales = s

    if sales is None and units.notna().mean() > 0 and price.notna().mean() > 0:
        sales = units * price

    out["__units__"] = units
    out["__price__"] = price
    out["__sales__"] = sales if sales is not None else pd.Series(np.nan, index=out.index)

    if mapping.get("customer"):
        out["__entity__"] = out[mapping["customer"]].astype(str)
    elif mapping.get("sales_person"):
        out["__entity__"] = out[mapping["sales_person"]].astype(str)
    else:
        out["__entity__"] = pd.Series(np.nan, index=out.index)

    if mapping.get("date"):
        out["__date__"] = pd.to_datetime(out[mapping["date"]], errors="coerce")
    else:
        out["__date__"] = pd.NaT

    return out


def score_sheet(df: pd.DataFrame) -> int:
    m = auto_detect_columns(df)
    score = 0
    score += 4 if m["sales"] else 0
    score += 3 if m["date"] else 0
    score += 2 if (m["customer"] or m["sales_person"]) else 0
    score += 2 if (m["region"] or m["state"]) else 0
    score += 2 if m["brand"] else 0
    score += 1 if m["units"] else 0
    score += 1 if m["price"] else 0
    score += min(5, int(len(df) / 200))
    return score


def load_best_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    best_name, best_df, best_score = None, None, -1
    for name in xls.sheet_names:
        df = xls.parse(name)
        if df is None or df.empty:
            continue
        sc = score_sheet(df)
        if sc > best_score:
            best_score, best_name, best_df = sc, name, df.copy()
    return xls.sheet_names, best_name, best_df


# =========================
# Premium CSS
# =========================
def inject_css(theme: str, accent_hex: str):
    if theme == "Dark":
        bg = "#070A12"
        text = "#E5E7EB"
        muted = "#9CA3AF"
        border = "rgba(255,255,255,0.10)"
        panel = "rgba(255,255,255,0.06)"
    else:
        bg = "#F6F7FB"
        text = "#0F172A"
        muted = "#475569"
        border = "rgba(0,0,0,0.10)"
        panel = "rgba(0,0,0,0.05)"

    # Accent-driven glow (THIS is what makes accent actually visible)
    glow1 = rgba(accent_hex, 0.22)
    glow2 = rgba(accent_hex, 0.14)

    st.markdown(
        f"""
        <style>
          :root {{
            --bg: {bg};
            --text: {text};
            --muted: {muted};
            --border: {border};
            --panel: {panel};
            --accent: {accent_hex};
            --radius: 18px;
          }}

          /* ✅ Fix the ugly dark top gap (Streamlit header) */
          header[data-testid="stHeader"] {{
            background: transparent !important;
          }}
          /* Sometimes the container has a background too */
          div[data-testid="stAppViewContainer"] {{
            background: transparent !important;
          }}
          /* Optional: hide Streamlit footer */
          footer {{ visibility: hidden; }}

          .stApp {{
            background:
              radial-gradient(1200px 600px at 18% 8%, {glow1}, transparent 60%),
              radial-gradient(900px 500px at 78% 18%, {glow2}, transparent 55%),
              var(--bg);
            color: var(--text);
          }}

          /* ✅ Pointer cursor (fixes “I” cursor) */
          button, [role="button"], a,
          label, input[type="checkbox"], input[type="radio"],
          [data-baseweb="select"] * , [data-baseweb="slider"] *,
          [data-testid="stSelectbox"] * {{
            cursor: pointer !important;
          }}

          /* Layout */
          .block-container {{
            padding-top: 1.2rem !important;
            padding-bottom: 2.8rem !important;
          }}

          .hero {{
            padding: 22px 26px;
            border-radius: 22px;
            border: 1px solid var(--border);
            background:
              radial-gradient(900px 340px at 18% 0%, {rgba(accent_hex,0.28)}, transparent 60%),
              linear-gradient(135deg, rgba(255,255,255,0.10), rgba(255,255,255,0.03));
            box-shadow: 0 20px 60px rgba(0,0,0,0.28);
          }}
          .muted {{ color: var(--muted) !important; }}

          .glass {{
            background: linear-gradient(135deg, rgba(255,255,255,0.08), rgba(255,255,255,0.03));
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.22);
            backdrop-filter: blur(14px);
          }}

          .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 12px;
          }}
          .kpi {{
            background: var(--panel);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 14px;
            box-shadow: 0 10px 26px rgba(0,0,0,0.18);
          }}
          .kpi-title {{
            font-size: 13px;
            color: var(--muted);
            margin-bottom: 6px;
          }}
          .kpi-value {{
            font-size: 22px;
            font-weight: 900;
          }}

          /* ✅ Mobile responsive */
          @media (max-width: 900px) {{
            .kpi-grid {{
              grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
          }}
          @media (max-width: 520px) {{
            .kpi-grid {{
              grid-template-columns: 1fr;
            }}
            .hero {{
              padding: 16px 16px;
            }}
          }}

          /* Smooth buttons */
          .stButton > button, .stDownloadButton > button {{
            border-radius: 14px !important;
            border: 1px solid var(--border) !important;
            background: linear-gradient(135deg, rgba(255,255,255,0.10), rgba(255,255,255,0.05)) !important;
            color: var(--text) !important;
            padding: 10px 14px !important;
            transition: transform .08s ease, box-shadow .15s ease !important;
            box-shadow: 0 10px 26px rgba(0,0,0,0.14) !important;
          }}
          .stButton > button:hover, .stDownloadButton > button:hover {{
            transform: translateY(-1px);
            box-shadow: 0 16px 34px rgba(0,0,0,0.20) !important;
          }}
          .stDownloadButton > button {{
            background: linear-gradient(135deg, {rgba(accent_hex,0.34)}, {rgba(accent_hex,0.20)}) !important;
            border-color: {rgba(accent_hex,0.35)} !important;
          }}

          /* Make selectboxes premium */
          [data-baseweb="select"] > div {{
            border-radius: 14px !important;
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================
# Excel Export
# =========================
def export_excel_report(df2: pd.DataFrame, meta: dict, theme: str, accent_hex: str) -> bytes:
    import xlsxwriter

    output = io.BytesIO()

    total_sales = float(np.nansum(df2["__sales__"].values))
    units_sold = float(np.nansum(df2["__units__"].values))
    avg_price = float(np.nanmean(df2["__price__"].values))
    orders = int(len(df2))
    entities = int(df2["__entity__"].nunique()) if df2["__entity__"].notna().any() else 0

    monthly = None
    if df2["__date__"].notna().any():
        tmp = df2.dropna(subset=["__date__"]).copy()
        tmp["Month"] = tmp["__date__"].dt.to_period("M").dt.to_timestamp()
        monthly = tmp.groupby("Month", as_index=False)["__sales__"].sum().sort_values("Month")

    reg_col = meta["mapping"].get("region") or meta["mapping"].get("state")
    by_region = None
    if reg_col and reg_col in df2.columns:
        by_region = df2.groupby(reg_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)

    brand_col = meta["mapping"].get("brand")
    by_brand = None
    if brand_col and brand_col in df2.columns:
        by_brand = df2.groupby(brand_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)

    top_entities = None
    if df2["__entity__"].notna().any():
        top_entities = df2.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(10)

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        if theme == "Dark":
            header_bg = "#111827"
            header_fg = "#FFFFFF"
            card_bg = "#0F172A"
            muted = "#9CA3AF"
            border = "#1F2937"
        else:
            header_bg = "#111827"
            header_fg = "#FFFFFF"
            card_bg = "#F1F5F9"
            muted = "#475569"
            border = "#CBD5E1"

        fmt_header = wb.add_format({"bold": True, "bg_color": header_bg, "font_color": header_fg})
        fmt_h1 = wb.add_format({"bold": True, "font_size": 18})
        fmt_muted = wb.add_format({"font_color": muted})
        fmt_money0 = wb.add_format({"num_format": "#,##0"})
        fmt_money2 = wb.add_format({"num_format": "#,##0.00"})
        fmt_card = wb.add_format({"bg_color": card_bg, "border": 1, "border_color": border})
        fmt_card_title = wb.add_format({"bold": True, "font_color": muted})
        fmt_card_value = wb.add_format({"bold": True, "font_size": 18})

        # Data
        df2.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        ws_data.freeze_panes(1, 0)
        ws_data.write_row(0, 0, list(df2.columns), fmt_header)

        # Summary
        summary = pd.DataFrame(
            {
                "Metric": ["Total Sales", "Units Sold", "Avg Price/Unit", "Orders (rows)", "Customers/Entities"],
                "Value": [total_sales, units_sold, avg_price, orders, entities],
            }
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)
        ws_sum = writer.sheets["Summary"]
        ws_sum.freeze_panes(1, 0)
        ws_sum.write_row(0, 0, ["Metric", "Value"], fmt_header)
        ws_sum.set_column(0, 0, 26)
        ws_sum.set_column(1, 1, 18, fmt_money2)

        if monthly is not None and len(monthly) > 0:
            monthly.rename(columns={"__sales__": "Sales"}).to_excel(writer, sheet_name="Monthly_Trend", index=False)
            ws = writer.sheets["Monthly_Trend"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Month", "Sales"], fmt_header)
            ws.set_column(0, 0, 16)
            ws.set_column(1, 1, 18, fmt_money0)

        if by_region is not None and len(by_region) > 0:
            by_region.rename(columns={"__sales__": "Sales"}).to_excel(writer, sheet_name="By_Region", index=False)
            ws = writer.sheets["By_Region"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, [reg_col, "Sales"], fmt_header)
            ws.set_column(0, 0, 22)
            ws.set_column(1, 1, 18, fmt_money0)

        if by_brand is not None and len(by_brand) > 0:
            by_brand.rename(columns={"__sales__": "Sales"}).to_excel(writer, sheet_name="By_Brand", index=False)
            ws = writer.sheets["By_Brand"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, [brand_col, "Sales"], fmt_header)
            ws.set_column(0, 0, 28)
            ws.set_column(1, 1, 18, fmt_money0)

        if top_entities is not None and len(top_entities) > 0:
            top_entities.rename(columns={"__entity__": "Entity", "__sales__": "Sales"}).to_excel(
                writer, sheet_name="Top_Entities", index=False
            )
            ws = writer.sheets["Top_Entities"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Entity", "Sales"], fmt_header)
            ws.set_column(0, 0, 30)
            ws.set_column(1, 1, 18, fmt_money0)

        # Dashboard
        ws_dash = wb.add_worksheet("Dashboard")
        ws_dash.hide_gridlines(2)
        ws_dash.set_tab_color(accent_hex)

        ws_dash.set_default_row(22)
        ws_dash.set_column("A:A", 24)
        ws_dash.set_column("B:B", 2)
        ws_dash.set_column("C:D", 18)
        ws_dash.set_column("E:E", 2)
        ws_dash.set_column("F:G", 18)
        ws_dash.set_column("H:H", 2)
        ws_dash.set_column("I:J", 18)

        ws_dash.write("A1", "Dashboard", fmt_h1)
        ws_dash.write("A2", f"Built from: {meta['source_name']} • Sheet: {meta['sheet']}", fmt_muted)

        def kpi_card(col_start, title, value, row_top=4):
            ws_dash.merge_range(row_top, col_start, row_top, col_start + 1, "", fmt_card)
            ws_dash.merge_range(row_top + 1, col_start, row_top + 2, col_start + 1, "", fmt_card)
            ws_dash.write(row_top, col_start, title, fmt_card_title)
            ws_dash.write(row_top + 1, col_start, value, fmt_card_value)

        kpi_card(0, "Total Sales", f"{total_sales:,.0f}")
        kpi_card(2, "Units Sold", f"{units_sold:,.0f}")
        kpi_card(5, "Avg Price/Unit", f"{avg_price:,.2f}")
        kpi_card(8, "Orders", f"{orders:,}")

        ws_dash.merge_range(7, 8, 7, 9, "", fmt_card)
        ws_dash.merge_range(8, 8, 9, 9, "", fmt_card)
        ws_dash.write(7, 8, "Customers/Entities", fmt_card_title)
        ws_dash.write(8, 8, f"{entities:,}", fmt_card_value)

        ws_dash.write("A11", "Auto Insights:", fmt_muted)
        insight_lines = meta.get("insights", [])[:4]
        for i, line in enumerate(insight_lines):
            ws_dash.write(11 + i, 0, f"• {line}", fmt_muted)

    return output.getvalue()


# =========================
# App
# =========================
st.set_page_config(page_title="MetricFlow", page_icon="📊", layout="wide")

ACCENTS = {
    "Purple": "#A855F7",
    "Blue": "#3B82F6",
    "Emerald": "#10B981",
    "Rose": "#F43F5E",
}

if "theme" not in st.session_state:
    st.session_state.theme = "Dark"
if "accent_name" not in st.session_state:
    st.session_state.accent_name = "Purple"

inject_css(st.session_state.theme, ACCENTS[st.session_state.accent_name])

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
    st.session_state.theme = st.selectbox(
        "Theme",
        ["Dark", "Light"],
        index=0 if st.session_state.theme == "Dark" else 1
    )
with c2:
    st.session_state.accent_name = st.selectbox(
        "Accent",
        list(ACCENTS.keys()),
        index=list(ACCENTS.keys()).index(st.session_state.accent_name)
    )
with c3:
    st.markdown(
        '<div class="muted" style="padding-top:30px;">Tip: If Sales is missing, MetricFlow auto-calculates it using Quantity × Unit Price when possible.</div>',
        unsafe_allow_html=True,
    )

# Re-inject after user changes theme/accent
inject_css(st.session_state.theme, ACCENTS[st.session_state.accent_name])

st.write("")

uploaded = st.file_uploader("Upload Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
if not uploaded:
    st.info("Upload a file to generate your dashboard.")
    st.stop()

sheet_names, best_sheet, df = load_best_sheet(uploaded)
if df is None or df.empty:
    st.error("Couldn't find usable data in this file.")
    st.stop()

mapping = auto_detect_columns(df)
df2 = prepare_df(df, mapping)

total_sales = float(np.nansum(df2["__sales__"].values))
units_sold = float(np.nansum(df2["__units__"].values))
avg_price = float(np.nanmean(df2["__price__"].values))
orders = int(len(df2))
entities = int(df2["__entity__"].nunique()) if df2["__entity__"].notna().any() else 0

# Insights
insights = []
if total_sales > 0:
    insights.append(f"Total sales recorded: {total_sales:,.0f} across {orders:,} rows.")

brand_col = mapping.get("brand")
if brand_col and brand_col in df2.columns:
    by_brand = df2.groupby(brand_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if not by_brand.empty:
        top = by_brand.iloc[0]
        insights.append(f"Top product/brand by sales: {top[brand_col]} ({top['__sales__']:,.0f}).")

reg_col = mapping.get("region") or mapping.get("state")
if reg_col and reg_col in df2.columns:
    by_reg = df2.groupby(reg_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if not by_reg.empty:
        top = by_reg.iloc[0]
        insights.append(f"Strongest region/state: {top[reg_col]} ({top['__sales__']:,.0f}).")

if df2["__entity__"].notna().any() and total_sales > 0:
    by_ent = df2.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if len(by_ent) >= 3:
        share = by_ent.head(3)["__sales__"].sum() / max(1.0, total_sales)
        insights.append(f"Top 3 customers/entities contribute ~{share:.0%} of total sales (concentration).")

st.markdown(
    f"""
    <div class="kpi-grid">
      <div class="kpi"><div class="kpi-title">Total Sales</div><div class="kpi-value">{total_sales:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Units Sold</div><div class="kpi-value">{units_sold:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Avg Price / Unit</div><div class="kpi-value">{avg_price:,.2f}</div></div>
      <div class="kpi"><div class="kpi-title">Orders (rows)</div><div class="kpi-value">{orders:,.0f}</div></div>
      <div class="kpi"><div class="kpi-title">Customers/Entities</div><div class="kpi-value">{entities:,.0f}</div></div>
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
        st.info("Upload richer data (Date + Sales/Qty/Price + Region/Brand/Customer) for deeper insights.")

st.write("")
st.subheader("Dashboard")
st.caption("Charts are generated only if the required columns exist in your file.")

colA, colB = st.columns(2)

if df2["__date__"].notna().any():
    tmp = df2.dropna(subset=["__date__"]).copy()
    tmp["Month"] = tmp["__date__"].dt.to_period("M").dt.to_timestamp()
    monthly = tmp.groupby("Month", as_index=False)["__sales__"].sum().sort_values("Month")
    with colA:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Monthly Sales Trend**")
        st.line_chart(monthly.set_index("Month")["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colA:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Monthly Sales Trend**")
        st.info("No usable date column detected.")
        st.markdown("</div>", unsafe_allow_html=True)

if reg_col and reg_col in df2.columns:
    by_region = df2.groupby(reg_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
    with colB:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown(f"**Sales by {reg_col}**")
        st.bar_chart(by_region.set_index(reg_col)["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colB:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Region/State**")
        st.info("No region/state column detected.")
        st.markdown("</div>", unsafe_allow_html=True)

st.write("")
colC, colD = st.columns(2)

if df2["__entity__"].notna().any():
    top_entities = df2.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(10)
    with colC:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Top Customers/Entities**")
        st.dataframe(top_entities.rename(columns={"__entity__": "Entity", "__sales__": "Sales"}), use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colC:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Top Customers/Entities**")
        st.info("No customer/entity column detected.")
        st.markdown("</div>", unsafe_allow_html=True)

if brand_col and brand_col in df2.columns:
    by_brand = df2.groupby(brand_col, as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
    with colD:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Brand/Product**")
        st.bar_chart(by_brand.set_index(brand_col)["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colD:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Brand/Product**")
        st.info("No brand/product column detected.")
        st.markdown("</div>", unsafe_allow_html=True)

st.write("")
st.subheader("Excel Export")
st.caption("Exports: Data + Summary + Dashboard. Gridlines OFF only on Dashboard sheet.")

meta = {"source_name": uploaded.name, "sheet": best_sheet, "mapping": mapping, "insights": insights}
excel_bytes = export_excel_report(df2, meta, st.session_state.theme, ACCENTS[st.session_state.accent_name])

st.download_button(
    "⬇️ Download MetricFlow Excel Report",
    data=excel_bytes,
    file_name="MetricFlow_Dashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Advanced (optional): show auto-detected columns", expanded=False):
    st.json(mapping)
    st.write(f"Auto-selected sheet: **{best_sheet}** (out of {sheet_names})")

st.write("")
st.subheader("Preview")
st.dataframe(df2.head(100), use_container_width=True)