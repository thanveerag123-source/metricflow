import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# Excel export (formatting + charts)
import xlsxwriter


# =========================
# Helpers: normalize + detect
# =========================
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()


def _norm_compact(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower()).strip()


def to_num(series: pd.Series) -> pd.Series:
    # Handles: 1,23,456 | ₹1,234.50 | 1 234 | etc.
    s = series.astype(str)
    s = s.str.replace(r"[\s,]", "", regex=True)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def try_parse_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def _is_numeric_series(s: pd.Series) -> bool:
    x = pd.to_numeric(s, errors="coerce")
    return x.notna().mean() >= 0.6


def detect_columns(df: pd.DataFrame):
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
    state_col = find_any(["state", "province"])
    brand_col = find_any(["product", "brand", "item", "sku", "beverage", "product name"])

    # sales detection should avoid "sales man / salesman"
    sales_candidates = []
    for c in cols:
        n = _norm(c)
        if any(k in n for k in ["sales", "revenue", "amount", "value", "total"]):
            if not any(bad in n for bad in ["sales man", "salesman", "sales person", "salesperson", "rep", "agent"]):
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


def score_sheet(df: pd.DataFrame) -> int:
    m = detect_columns(df)
    score = 0
    score += 4 if m["sales"] else 0
    score += 3 if m["date"] else 0
    score += 2 if m["customer"] else 0
    score += 2 if (m["region"] or m["state"]) else 0
    score += 2 if m["brand"] else 0
    score += 1 if m["units"] else 0
    score += 1 if m["price"] else 0
    score += min(5, int(len(df) / 200))
    return score


def load_best_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    best_name = None
    best_df = None
    best_score = -1

    for name in xls.sheet_names:
        df = xls.parse(name)
        if df is None or df.empty:
            continue
        sc = score_sheet(df)
        if sc > best_score:
            best_score = sc
            best_name = name
            best_df = df.copy()

    return xls.sheet_names, best_name, best_df


def prepare_df(df: pd.DataFrame, mapping: dict):
    out = df.copy()

    # numeric cleanup
    if mapping.get("units") and mapping["units"] in out.columns:
        out[mapping["units"]] = to_num(out[mapping["units"]])

    if mapping.get("price") and mapping["price"] in out.columns:
        out[mapping["price"]] = to_num(out[mapping["price"]])

    if mapping.get("sales") and mapping["sales"] in out.columns:
        out[mapping["sales"]] = to_num(out[mapping["sales"]])

    if mapping.get("date") and mapping["date"] in out.columns:
        out[mapping["date"]] = try_parse_date(out[mapping["date"]])

    # sales fallback
    sales_col = mapping.get("sales")
    sales_ok = False
    if sales_col and sales_col in out.columns:
        ssum = float(out[sales_col].sum(skipna=True))
        snn = int(out[sales_col].notna().sum())
        sales_ok = (ssum != 0.0 and snn > 0)

    if (not sales_ok) and mapping.get("units") and mapping.get("price"):
        if mapping["units"] in out.columns and mapping["price"] in out.columns:
            out["_calc_sales"] = out[mapping["units"]] * out[mapping["price"]]
            sales_col = "_calc_sales"

    out["__sales__"] = out[sales_col] if sales_col and sales_col in out.columns else np.nan
    out["__units__"] = out[mapping["units"]] if mapping.get("units") and mapping["units"] in out.columns else np.nan
    out["__price__"] = out[mapping["price"]] if mapping.get("price") and mapping["price"] in out.columns else np.nan

    # “entity” (customer/retailer OR salesperson)
    out["__entity__"] = None
    if mapping.get("customer") and mapping["customer"] in out.columns:
        out["__entity__"] = out[mapping["customer"]].astype(str)
    elif mapping.get("sales_person") and mapping["sales_person"] in out.columns:
        out["__entity__"] = out[mapping["sales_person"]].astype(str)

    # region-ish
    out["__geo__"] = None
    geo = mapping.get("region") or mapping.get("state")
    if geo and geo in out.columns:
        out["__geo__"] = out[geo].astype(str)

    # product-ish
    out["__product__"] = None
    if mapping.get("brand") and mapping["brand"] in out.columns:
        out["__product__"] = out[mapping["brand"]].astype(str)

    # date parts
    if mapping.get("date") and mapping["date"] in out.columns:
        out["__month__"] = pd.to_datetime(out[mapping["date"]], errors="coerce").dt.to_period("M").dt.to_timestamp()

    return out


# =========================
# Command parser (no API)
# =========================
def parse_command(cmd: str):
    """
    Very lightweight parser (no LLM).
    Supports:
      - "group by <col> monthly sum sales"
      - "group by <col1>, <col2> sum sales"
      - "top 10 <col> by sales"
      - "trend monthly sales"
    """
    if not cmd or not cmd.strip():
        return None

    c = cmd.strip().lower()

    out = {
        "mode": None,         # group / top / trend
        "group_cols": [],
        "agg": "sum",
        "metric": "__sales__",
        "top_n": 10,
        "time_grain": None,   # month
    }

    # trend
    if "trend" in c or "monthly" in c:
        out["mode"] = "trend"
        out["time_grain"] = "month"
        return out

    # top N
    m = re.search(r"top\s+(\d+)\s+(.+?)\s+by\s+(.+)", c)
    if m:
        out["mode"] = "top"
        out["top_n"] = int(m.group(1))
        out["group_cols"] = [m.group(2).strip()]
        return out

    # group by
    if "group by" in c:
        out["mode"] = "group"
        after = c.split("group by", 1)[1].strip()

        # monthly?
        if "monthly" in after or "month" in after:
            out["time_grain"] = "month"
            after = after.replace("monthly", "").replace("month", "").strip()

        # cols separated by comma
        cols_part = after
        # remove possible "sum sales" text
        cols_part = re.sub(r"\b(sum|mean|avg|average|count)\b.*$", "", cols_part).strip()
        group_cols = [x.strip() for x in cols_part.split(",") if x.strip()]
        out["group_cols"] = group_cols[:2]  # keep it simple: max 2 group cols
        return out

    return None


def resolve_columns_from_text(df: pd.DataFrame, text_cols: list):
    """
    Match user-typed names to actual columns.
    Uses compact normalization contains-match.
    """
    if not text_cols:
        return []

    cols = list(df.columns)
    cols_n = [_norm_compact(c) for c in cols]

    resolved = []
    for t in text_cols:
        tn = _norm_compact(t)
        best = None
        for i, cn in enumerate(cols_n):
            if tn and (tn in cn or cn in tn):
                best = cols[i]
                break
        if best is None:
            # fallback: try word-match
            t_words = set(_norm(t).split())
            for i, c in enumerate(cols):
                c_words = set(_norm(c).split())
                if len(t_words.intersection(c_words)) >= 1:
                    best = c
                    break
        if best:
            resolved.append(best)
    return resolved


# =========================
# Premium UI / CSS (mobile safe)
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
        panel = "rgba(0,0,0,0.04)"

    # IMPORTANT: remove the “top gap / dark bar”
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

          /* Kill Streamlit top header space */
          [data-testid="stHeader"] {{
            background: transparent !important;
            height: 0px !important;
          }}
          header {{ visibility: hidden; height: 0px; }}
          [data-testid="stToolbar"] {{ visibility: hidden; height: 0px; }}

          /* App background */
          .stApp {{
            background:
              radial-gradient(1100px 540px at 22% 10%, color-mix(in srgb, var(--accent) 22%, transparent), transparent 60%),
              radial-gradient(820px 460px at 72% 16%, rgba(59,130,246,0.14), transparent 60%),
              var(--bg);
            color: var(--text);
          }}

          /* Cursor pointer for interactive */
          button, [role="button"], a, select, input {{
            cursor: pointer !important;
          }}
          [data-baseweb="select"] * {{ cursor: pointer !important; }}
          [data-baseweb="slider"] * {{ cursor: pointer !important; }}

          .muted {{ color: var(--muted) !important; }}

          .hero {{
            padding: 22px 22px;
            border-radius: 22px;
            border: 1px solid var(--border);
            background:
              radial-gradient(900px 280px at 18% 0%, color-mix(in srgb, var(--accent) 30%, transparent), transparent 60%),
              radial-gradient(760px 300px at 80% 0%, rgba(59,130,246,0.20), transparent 60%),
              linear-gradient(135deg, rgba(255,255,255,0.10), rgba(255,255,255,0.03));
            box-shadow: 0 24px 70px rgba(0,0,0,0.26);
            backdrop-filter: blur(14px);
            overflow: hidden;
          }}

          /* Smooth buttons + accent */
          .stButton > button, .stDownloadButton > button {{
            border-radius: 14px !important;
            border: 1px solid var(--border) !important;
            background: linear-gradient(135deg, color-mix(in srgb, var(--accent) 22%, rgba(255,255,255,0.06)), rgba(255,255,255,0.05)) !important;
            color: var(--text) !important;
            padding: 10px 14px !important;
            transition: transform .08s ease, box-shadow .15s ease, border-color .15s ease !important;
            box-shadow: 0 12px 26px rgba(0,0,0,0.16) !important;
          }}
          .stButton > button:hover, .stDownloadButton > button:hover {{
            transform: translateY(-1px);
            box-shadow: 0 18px 34px rgba(0,0,0,0.22) !important;
            border-color: color-mix(in srgb, var(--accent) 40%, var(--border)) !important;
          }}

          /* KPI cards */
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
            box-shadow: 0 12px 26px rgba(0,0,0,0.16);
            backdrop-filter: blur(12px);
          }}
          .kpi-title {{
            font-size: 13px;
            color: var(--muted);
            margin-bottom: 6px;
          }}
          .kpi-value {{
            font-size: 22px;
            font-weight: 900;
            letter-spacing: 0.2px;
          }}

          /* Glass panels */
          .glass {{
            background: linear-gradient(135deg, rgba(255,255,255,0.08), rgba(255,255,255,0.03));
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px 16px;
            box-shadow: 0 16px 44px rgba(0,0,0,0.18);
            backdrop-filter: blur(14px);
          }}

          /* Mobile responsiveness */
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
              border-radius: 18px;
            }}
          }}

          /* Layout padding */
          .block-container {{
            padding-top: 1.0rem !important;
            padding-bottom: 2.0rem !important;
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================
# Excel export (Dashboard gridlines OFF only)
# =========================
def export_excel_report(df: pd.DataFrame, meta: dict, theme: str, accent_hex: str, cmd_plan: dict | None) -> bytes:
    output = io.BytesIO()

    # KPIs
    total_sales = float(pd.to_numeric(df["__sales__"], errors="coerce").sum(skipna=True))
    orders = int(len(df))
    units = float(pd.to_numeric(df["__units__"], errors="coerce").sum(skipna=True))
    avg_price = float(pd.to_numeric(df["__price__"], errors="coerce").mean(skipna=True))
    customers = int(df["__entity__"].astype(str).nunique()) if "__entity__" in df.columns else 0

    # Tables for export
    monthly = None
    if "__month__" in df.columns and df["__month__"].notna().any():
        tmp = df.dropna(subset=["__month__"]).copy()
        monthly = tmp.groupby("__month__", as_index=False)["__sales__"].sum().sort_values("__month__")
        monthly.rename(columns={"__month__": "Month", "__sales__": "Sales"}, inplace=True)

    by_geo = None
    if "__geo__" in df.columns and df["__geo__"].notna().any():
        by_geo = df.groupby("__geo__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
        by_geo.rename(columns={"__geo__": "Geo", "__sales__": "Sales"}, inplace=True)

    by_product = None
    if "__product__" in df.columns and df["__product__"].notna().any():
        by_product = df.groupby("__product__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
        by_product.rename(columns={"__product__": "Product", "__sales__": "Sales"}, inplace=True)

    top_entities = None
    if "__entity__" in df.columns and df["__entity__"].notna().any():
        top_entities = df.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(10)
        top_entities.rename(columns={"__entity__": "Entity", "__sales__": "Sales"}, inplace=True)

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        if theme == "Dark":
            header_bg = "#111827"
            header_fg = "#FFFFFF"
            muted = "#9CA3AF"
            card_bg = "#0F172A"
        else:
            header_bg = "#111827"
            header_fg = "#FFFFFF"
            muted = "#475569"
            card_bg = "#F1F5F9"

        fmt_header = wb.add_format({"bold": True, "bg_color": header_bg, "font_color": header_fg})
        fmt_h1 = wb.add_format({"bold": True, "font_size": 18})
        fmt_muted = wb.add_format({"font_color": muted})
        fmt_money0 = wb.add_format({"num_format": "#,##0"})
        fmt_money2 = wb.add_format({"num_format": "#,##0.00"})

        fmt_card = wb.add_format({"bg_color": card_bg, "border": 1, "border_color": "#1F2937"})
        fmt_card_title = wb.add_format({"bold": True})
        fmt_card_value = wb.add_format({"bold": True, "font_size": 16})

        # Data
        df.to_excel(writer, sheet_name="Data", index=False)
        ws_data = writer.sheets["Data"]
        ws_data.freeze_panes(1, 0)
        ws_data.write_row(0, 0, list(df.columns), fmt_header)

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
        if monthly is not None and len(monthly) >= 2:
            monthly.to_excel(writer, sheet_name="Monthly_Trend", index=False)
            ws = writer.sheets["Monthly_Trend"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Month", "Sales"], fmt_header)
            ws.set_column(0, 0, 16)
            ws.set_column(1, 1, 18, fmt_money0)

        if by_geo is not None and len(by_geo) >= 2:
            by_geo.to_excel(writer, sheet_name="By_Geo", index=False)
            ws = writer.sheets["By_Geo"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Geo", "Sales"], fmt_header)
            ws.set_column(0, 0, 22)
            ws.set_column(1, 1, 18, fmt_money0)

        if by_product is not None and len(by_product) >= 2:
            by_product.to_excel(writer, sheet_name="By_Product", index=False)
            ws = writer.sheets["By_Product"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Product", "Sales"], fmt_header)
            ws.set_column(0, 0, 28)
            ws.set_column(1, 1, 18, fmt_money0)

        if top_entities is not None and len(top_entities) >= 2:
            top_entities.to_excel(writer, sheet_name="Top_Entities", index=False)
            ws = writer.sheets["Top_Entities"]
            ws.freeze_panes(1, 0)
            ws.write_row(0, 0, ["Entity", "Sales"], fmt_header)
            ws.set_column(0, 0, 28)
            ws.set_column(1, 1, 18, fmt_money0)

        # Dashboard
        ws_dash = wb.add_worksheet("Dashboard")
        ws_dash.hide_gridlines(2)  # gridlines OFF only here
        ws_dash.set_tab_color(accent_hex)
        ws_dash.set_column("A:A", 26)
        ws_dash.set_column("B:K", 16)
        ws_dash.set_default_row(20)

        ws_dash.write("A1", "Dashboard", fmt_h1)
        ws_dash.write("A2", f"Built from: {meta['source_name']} • Sheet: {meta['sheet']}", fmt_muted)

        def money0(x): return f"{x:,.0f}"
        def money2(x): return f"{x:,.2f}"

        def card(rng, title, value):
            ws_dash.merge_range(rng, "", fmt_card)
            tl = rng.split(":")[0]
            col = re.sub(r"[0-9]", "", tl)
            row = int(re.sub(r"[^0-9]", "", tl))
            ws_dash.write(f"{col}{row}", title, fmt_card_title)
            ws_dash.write(f"{col}{row+1}", value, fmt_card_value)

        # KPI row
        card("A4:B6", "Total Sales", money0(total_sales))
        card("C4:D6", "Units Sold", money0(units))
        card("E4:F6", "Avg Price/Unit", money2(avg_price))
        card("G4:H6", "Orders", f"{orders:,}")
        card("I4:J6", "Customers", f"{customers:,}")

        # Insights area
        ws_dash.write("A8", "Auto Insights:", fmt_muted)
        for i, line in enumerate(meta.get("insights", [])[:5]):
            ws_dash.write(8 + i, 0, f"• {line}", fmt_muted)

        # Charts area (3–4 charts)
        row_chart_top = 13

        # 1) Monthly trend
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
            ws_dash.insert_chart(f"A{row_chart_top}", chart)

        # 2) Sales by Geo
        if by_geo is not None and len(by_geo) >= 2:
            chart2 = wb.add_chart({"type": "column"})
            n = len(by_geo)
            chart2.add_series({
                "categories": f"=By_Geo!$A$2:$A${n+1}",
                "values": f"=By_Geo!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart2.set_title({"name": "Sales by Region/State"})
            chart2.set_legend({"none": True})
            chart2.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"F{row_chart_top}", chart2)

        # 3) Sales by Product
        if by_product is not None and len(by_product) >= 2:
            chart3 = wb.add_chart({"type": "bar"})
            n = len(by_product)
            chart3.add_series({
                "categories": f"=By_Product!$A$2:$A${n+1}",
                "values": f"=By_Product!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart3.set_title({"name": "Sales by Product"})
            chart3.set_legend({"none": True})
            chart3.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"A{row_chart_top + 18}", chart3)

        # 4) Top entities
        if top_entities is not None and len(top_entities) >= 2:
            chart4 = wb.add_chart({"type": "column"})
            n = len(top_entities)
            chart4.add_series({
                "categories": f"=Top_Entities!$A$2:$A${n+1}",
                "values": f"=Top_Entities!$B$2:$B${n+1}",
                "fill": {"color": accent_hex},
                "border": {"none": True},
            })
            chart4.set_title({"name": "Top Customers/Retailers"})
            chart4.set_legend({"none": True})
            chart4.set_size({"width": 620, "height": 320})
            ws_dash.insert_chart(f"F{row_chart_top + 18}", chart4)

    return output.getvalue()


# =========================
# Streamlit setup
# =========================
st.set_page_config(
    page_title="MetricFlow",
    page_icon="📊",
    layout="wide",
)

if "theme" not in st.session_state:
    st.session_state.theme = "Dark"
if "accent_name" not in st.session_state:
    st.session_state.accent_name = "Purple"

ACCENTS = {
    "Purple": "#A855F7",
    "Blue": "#3B82F6",
    "Emerald": "#10B981",
    "Rose": "#F43F5E",
}

accent_hex = ACCENTS.get(st.session_state.accent_name, "#A855F7")
inject_css(st.session_state.theme, accent_hex)


# =========================
# HERO
# =========================
st.markdown(
    f"""
    <div class="hero">
      <div style="display:flex; align-items:flex-start; justify-content:space-between; gap:16px; flex-wrap:wrap;">
        <div>
          <div style="font-size:34px; font-weight:950; letter-spacing:-0.4px;">📊 MetricFlow</div>
          <div class="muted" style="margin-top:6px; max-width:860px;">
            Upload Excel → auto-detect columns → generate KPIs, charts, and a clean Excel dashboard (gridlines OFF on Dashboard sheet only).
          </div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")


# =========================
# Settings row (simple)
# =========================
c1, c2, c3 = st.columns([1, 1, 2])
with c1:
    st.session_state.theme = st.selectbox("Theme", ["Dark", "Light"], index=0 if st.session_state.theme == "Dark" else 1)
with c2:
    st.session_state.accent_name = st.selectbox("Accent", list(ACCENTS.keys()), index=list(ACCENTS.keys()).index(st.session_state.accent_name))
with c3:
    st.markdown(
        '<div class="muted" style="padding-top:30px;">Tip: If Sales is missing, MetricFlow auto-calculates it using Quantity × Unit Price (when possible).</div>',
        unsafe_allow_html=True,
    )

accent_hex = ACCENTS[st.session_state.accent_name]
inject_css(st.session_state.theme, accent_hex)

st.write("")


# =========================
# Multi-file uploader
# =========================
uploaded_files = st.file_uploader(
    "Upload Excel (.xlsx / .xlsm) — you can upload multiple files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload at least 1 Excel file to generate a dashboard.")
    st.stop()

mode = st.radio("How should MetricFlow use multiple files?", ["Use one file", "Combine all files (append rows)"], horizontal=True)

# pick a file (if single)
chosen_file = None
if mode == "Use one file":
    names = [f.name for f in uploaded_files]
    pick = st.selectbox("Select file", names, index=0)
    chosen_file = uploaded_files[names.index(pick)]
    active_files = [chosen_file]
else:
    active_files = uploaded_files


# =========================
# Load + combine
# =========================
all_frames = []
sheet_debug = []

for f in active_files:
    sheet_names, best_sheet, df = load_best_sheet(f)
    if df is None or df.empty:
        continue

    mapping = detect_columns(df)
    df2 = prepare_df(df, mapping)
    df2["__source_file__"] = f.name
    all_frames.append(df2)

    sheet_debug.append({
        "file": f.name,
        "best_sheet": best_sheet,
        "columns_detected": {k: v for k, v in mapping.items() if v is not None},
    })

if not all_frames:
    st.error("Couldn't find usable data in the uploaded file(s).")
    st.stop()

df_all = pd.concat(all_frames, ignore_index=True, sort=False)

# meta mapping from the first file (for display); export uses prepared internal columns anyway
meta_source = (chosen_file.name if chosen_file else f"{len(active_files)} files combined")
meta_sheet = "auto-best" if mode == "Combine all files (append rows)" else sheet_debug[0]["best_sheet"]


# =========================
# KPIs
# =========================
total_sales = float(pd.to_numeric(df_all["__sales__"], errors="coerce").sum(skipna=True))
orders = int(len(df_all))
units_sold = float(pd.to_numeric(df_all["__units__"], errors="coerce").sum(skipna=True))
avg_price = float(pd.to_numeric(df_all["__price__"], errors="coerce").mean(skipna=True))
customers = int(df_all["__entity__"].astype(str).nunique()) if df_all["__entity__"].notna().any() else 0

# Auto insights
insights = []
if total_sales > 0:
    insights.append(f"Total sales recorded: {total_sales:,.0f} across {orders:,} rows.")
if df_all["__product__"].notna().any():
    by_prod = df_all.groupby("__product__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if len(by_prod) > 0:
        insights.append(f"Top product by sales: {by_prod.iloc[0]['__product__']} ({by_prod.iloc[0]['__sales__']:,.0f}).")
if df_all["__entity__"].notna().any():
    by_ent = df_all.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if len(by_ent) >= 3 and total_sales > 0:
        top3_share = by_ent.head(3)["__sales__"].sum() / max(1.0, total_sales)
        insights.append(f"Top 3 customers contribute ~{top3_share:.0%} of total sales (high concentration).")
if df_all["__geo__"].notna().any():
    by_geo = df_all.groupby("__geo__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False)
    if len(by_geo) > 0:
        insights.append(f"Strongest region/state: {by_geo.iloc[0]['__geo__']} ({by_geo.iloc[0]['__sales__']:,.0f}).")
if mode == "Combine all files (append rows)":
    insights.append("Combined multiple files: use '__source_file__' to compare performance per file.")


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
        st.info("Upload a richer file (Date + Sales + Customer/Region/Product) to generate deeper insights.")

st.write("")


# =========================
# Command tab (no API)
# =========================
st.subheader("Command Mode (Optional)")
st.caption("Type a simple instruction to reshape the dashboard. Examples: "
           "`trend monthly sales` • `top 10 product name by sales` • `group by Sales Zone monthly` • `group by State, Product Name`")

cmd = st.text_input("Command", value="", placeholder="e.g., group by Sales Zone monthly")

cmd_plan = parse_command(cmd)
resolved_group_cols = []

if cmd_plan and cmd_plan.get("group_cols"):
    resolved_group_cols = resolve_columns_from_text(df_all, cmd_plan["group_cols"])

st.write("")


# =========================
# Dashboard
# =========================
st.subheader("Dashboard")
st.caption("Charts render automatically from your file. If your file lacks a column, that chart won’t show.")

colA, colB = st.columns(2)

# Trend
if "__month__" in df_all.columns and df_all["__month__"].notna().any():
    monthly = df_all.dropna(subset=["__month__"]).groupby("__month__", as_index=False)["__sales__"].sum().sort_values("__month__")
    with colA:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Monthly Sales Trend**")
        st.line_chart(monthly.set_index("__month__")["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colA:
        st.warning("No usable date column detected (needed for monthly trend).")

# Sales by Geo
if "__geo__" in df_all.columns and df_all["__geo__"].notna().any():
    geo = df_all.groupby("__geo__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
    with colB:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Region/State**")
        st.bar_chart(geo.set_index("__geo__")["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colB:
        st.warning("No Region/State column detected (needed for geo chart).")

st.write("")

colC, colD = st.columns(2)

# Top entities
if "__entity__" in df_all.columns and df_all["__entity__"].notna().any():
    top_entities = df_all.groupby("__entity__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(10)
    with colC:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Top Customers/Retailers**")
        st.dataframe(top_entities, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colC:
        st.warning("No Customer/Retailer (or Sales Person) column detected for top entities.")

# Sales by product
if "__product__" in df_all.columns and df_all["__product__"].notna().any():
    by_prod = df_all.groupby("__product__", as_index=False)["__sales__"].sum().sort_values("__sales__", ascending=False).head(12)
    with colD:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        st.markdown("**Sales by Product/Brand**")
        st.bar_chart(by_prod.set_index("__product__")["__sales__"])
        st.markdown("</div>", unsafe_allow_html=True)
else:
    with colD:
        st.warning("No Product/Brand column detected for product chart.")

st.write("")

# If command given, show an extra “Command Result” chart/table
if cmd_plan:
    st.subheader("Command Result")
    st.caption("This section is driven by your command. If the column name doesn’t match, adjust the wording.")

    # resolve metric (always sales for now)
    metric = "__sales__"

    if cmd_plan["mode"] == "trend":
        if "__month__" in df_all.columns and df_all["__month__"].notna().any():
            monthly = df_all.dropna(subset=["__month__"]).groupby("__month__", as_index=False)[metric].sum().sort_values("__month__")
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.markdown("**Command Trend: Monthly Sales**")
            st.line_chart(monthly.set_index("__month__")[metric])
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.warning("Your command needs a date column. Add Date / Invoice Date / Order Date in file.")
    elif cmd_plan["mode"] in ("group", "top"):
        gcols = resolved_group_cols

        # if user said monthly + has dates, auto add month
        if cmd_plan.get("time_grain") == "month" and "__month__" in df_all.columns and df_all["__month__"].notna().any():
            gcols = ["__month__"] + gcols

        if not gcols:
            st.warning("Command needs a valid column name. Example: `group by Sales Zone monthly`")
        else:
            out = df_all.dropna(subset=gcols).groupby(gcols, as_index=False)[metric].sum()

            if cmd_plan["mode"] == "top":
                # top by first group col
                out = out.sort_values(metric, ascending=False).head(cmd_plan["top_n"])

            st.markdown('<div class="glass">', unsafe_allow_html=True)
            st.markdown(f"**Grouped Result ({', '.join(gcols)})**")
            st.dataframe(out.sort_values(metric, ascending=False), use_container_width=True, hide_index=True)
            st.markdown("</div>", unsafe_allow_html=True)


# =========================
# Excel Export
# =========================
st.subheader("Excel Export")
st.caption("Exports a structured Excel file: Data + Summary + Dashboard + optional charts sheets. Dashboard gridlines OFF only.")

meta = {
    "source_name": meta_source,
    "sheet": meta_sheet,
    "insights": insights,
}

excel_bytes = export_excel_report(df_all, meta, st.session_state.theme, accent_hex, cmd_plan)

st.download_button(
    "⬇️ Download MetricFlow Excel Report",
    data=excel_bytes,
    file_name="MetricFlow_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Advanced (optional): debug auto-selection"):
    st.write("Selected / combined files:")
    st.json(sheet_debug)
    st.write("Tip: If a chart is missing, your file may not contain that column (Date / Region / Product / Customer).")

st.write("")
st.subheader("Preview (first 100 rows)")
st.dataframe(df_all.head(100), use_container_width=True)