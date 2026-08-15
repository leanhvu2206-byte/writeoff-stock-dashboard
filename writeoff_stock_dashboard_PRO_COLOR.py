# DASHBOARD VERSION: WRITE-OFF-STOCK-PRO-COLOR-V1
import os
import html as _html
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Write-Off Stock QC Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------
# STYLE
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {background: #F6F8FC;}
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FFFFFF 0%, #FAF8FF 100%);
        border-right: 1px solid #E6E2F5;
    }
    .block-container {padding-top: 1rem; padding-bottom: 3rem; max-width: 1900px;}

    .hero {
        background: linear-gradient(100deg, #3B0764 0%, #6D28D9 45%, #2563EB 100%);
        border-radius: 18px;
        padding: 22px 28px;
        margin-bottom: 18px;
        color: white;
        box-shadow: 0 10px 30px rgba(91, 61, 245, 0.22);
        border: 1px solid rgba(255,255,255,.18);
    }
    .hero h1 {font-size: 31px; margin: 0; font-weight: 900; letter-spacing: .25px;}
    .hero p {margin: 5px 0 0 0; opacity: .90; font-size: 15px; font-weight: 650;}

    .kpi-card {
        background: linear-gradient(145deg, #FFFFFF 0%, #FCFBFF 100%);
        border: 2px solid #DED6FF;
        border-left: 7px solid #6D28D9;
        border-radius: 16px;
        padding: 17px 18px;
        min-height: 125px;
        box-shadow: 0 7px 20px rgba(71, 47, 140, 0.09);
    }
    .kpi-label {font-size: 13px; color: #667085; font-weight: 850; text-transform: uppercase; letter-spacing: .2px;}
    .kpi-value {font-size: 29px; color: #25145F; font-weight: 950; margin-top: 8px; line-height: 1.05;}
    .kpi-note {font-size: 12px; color: #7A718F; margin-top: 8px; font-weight: 650;}

    div[data-testid="stPlotlyChart"] {
        background: linear-gradient(145deg, #FFFFFF 0%, #FCFBFF 100%);
        border: 2px solid #D8CCFF;
        border-radius: 18px;
        padding: 10px 12px 6px 12px;
        box-shadow: 0 8px 24px rgba(91, 61, 245, 0.11);
        margin-bottom: 16px;
    }
    div[data-testid="stPlotlyChart"]:hover {
        border-color: #7C3AED;
        box-shadow: 0 12px 30px rgba(91, 61, 245, 0.19);
        transition: all .20s ease-in-out;
    }

    .section-title {
        display: inline-block;
        background: linear-gradient(90deg, #5B21B6 0%, #7C3AED 50%, #2563EB 100%);
        color: #FFFFFF !important;
        border-radius: 12px;
        padding: 9px 15px;
        margin: 18px 0 11px 0;
        font-size: 17px !important;
        font-weight: 950 !important;
        letter-spacing: .45px;
        box-shadow: 0 6px 16px rgba(91, 61, 245, 0.20);
    }

    .big-table-wrap {
        background: #FFFFFF;
        border: 2px solid #DED6FF;
        border-radius: 15px;
        overflow-x: auto;
        box-shadow: 0 7px 20px rgba(71, 47, 140, 0.08);
        margin: 4px 0 16px 0;
    }
    .big-table {
        width: 100%;
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        background: #FFFFFF;
    }
    .big-table thead th {
        background: linear-gradient(180deg, #F5F2FF 0%, #EEE9FF 100%);
        color: #27165C;
        font-size: 17px;
        font-weight: 900;
        padding: 15px 13px;
        border-bottom: 1px solid #D9D1EF;
        border-right: 1px solid #E5E0F2;
        white-space: nowrap;
        text-align: left;
    }
    .big-table tbody td {
        color: #172B4D;
        font-size: 17px;
        font-weight: 650;
        padding: 15px 13px;
        border-bottom: 1px solid #E7ECF4;
        border-right: 1px solid #EEF1F5;
        white-space: nowrap;
    }
    .big-table tbody tr:nth-child(even) {background: #FBFAFF;}
    .big-table tbody tr:hover {background: #F3F0FF;}
    .big-table td.num {
        text-align: right;
        font-size: 24px !important;
        font-weight: 950 !important;
        color: #4C1D95 !important;
        font-variant-numeric: tabular-nums;
    }
    .big-table td.rank {
        text-align: center;
        font-size: 22px !important;
        font-weight: 950 !important;
        color: #6D28D9 !important;
    }
    .big-table td.decision {
        font-size: 17px !important;
        font-weight: 850 !important;
        color: #B42318 !important;
    }

    /* Sidebar widget typography */
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p {
        font-weight: 750 !important;
        color: #30235C !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# CONSTANTS / HELPERS
# ------------------------------------------------------------
DEFAULT_FILE = "Write off stock QC.xlsx"
AGING_ORDER = ["0-30", "31-60", "61-90", "91-180", "181-365", ">365"]
PALETTE = ["#6D28D9", "#2563EB", "#F97316", "#EC4899", "#10B981", "#EF4444", "#F59E0B", "#06B6D4"]
DECISION_COLORS = {
    "Write off": "#EF4444",
    "Adjustment": "#2563EB",
    "On hold": "#F59E0B",
    "Vendor returrn": "#10B981",
    "Vendor return": "#10B981",
}


def aging_bucket(days):
    if pd.isna(days):
        return "Unknown"
    d = int(days)
    if d <= 30:
        return "0-30"
    if d <= 60:
        return "31-60"
    if d <= 90:
        return "61-90"
    if d <= 180:
        return "91-180"
    if d <= 365:
        return "181-365"
    return ">365"


def fmt_num(v, decimals=0):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.{decimals}f}"


def fmt_vnd(v):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.0f}"


def kpi_card(label, value, note, accent="#6D28D9"):
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color:{accent}">
          <div class="kpi-label">{_html.escape(label)}</div>
          <div class="kpi-value">{_html.escape(value)}</div>
          <div class="kpi-note">{_html.escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_big_table(df, numeric_cols=None, rank_col=None, decision_col=None):
    numeric_cols = set(numeric_cols or [])
    work = df.copy()
    parts = ['<div class="big-table-wrap"><table class="big-table"><thead><tr>']
    for col in work.columns:
        parts.append(f'<th>{_html.escape(str(col))}</th>')
    parts.append('</tr></thead><tbody>')

    for _, row in work.iterrows():
        parts.append('<tr>')
        for col in work.columns:
            value = row[col]
            if col in numeric_cols:
                display = fmt_num(value)
                parts.append(f'<td class="num">{_html.escape(display)}</td>')
            elif rank_col and col == rank_col:
                display = "-" if pd.isna(value) else str(value)
                parts.append(f'<td class="rank">{_html.escape(display)}</td>')
            elif decision_col and col == decision_col:
                display = "-" if pd.isna(value) else str(value)
                parts.append(f'<td class="decision">{_html.escape(display)}</td>')
            else:
                display = "-" if pd.isna(value) else str(value)
                parts.append(f'<td>{_html.escape(display)}</td>')
        parts.append('</tr>')
    parts.append('</tbody></table></div>')
    st.markdown("".join(parts), unsafe_allow_html=True)


def style_figure(fig, title, height=430):
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", font=dict(size=20, color="#27165C"), x=0.02, xanchor="left"),
        height=height,
        margin=dict(l=20, r=20, t=70, b=35),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Arial", size=14, color="#33235E"),
        legend=dict(font=dict(size=14, color="#27165C"), title_font=dict(size=14, color="#27165C")),
        hoverlabel=dict(font_size=14),
    )
    fig.update_xaxes(
        tickfont=dict(size=13, color="#35245E", family="Arial Black"),
        title_font=dict(size=15, color="#27165C", family="Arial Black"),
        gridcolor="#ECEAF4",
        zeroline=False,
    )
    fig.update_yaxes(
        tickfont=dict(size=13, color="#35245E", family="Arial Black"),
        title_font=dict(size=15, color="#27165C", family="Arial Black"),
        gridcolor="#ECEAF4",
        zeroline=False,
    )
    return fig


def _parse_excel_date_series(series):
    """Parse normal Excel dates, text dates, and Excel serial numbers safely."""
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=False)

    # Handle Excel serial date values if any cells are stored as numbers.
    numeric = pd.to_numeric(series, errors="coerce")
    serial_mask = parsed.isna() & numeric.notna() & numeric.between(20000, 80000)
    if serial_mask.any():
        parsed.loc[serial_mask] = pd.to_datetime(
            numeric.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
        )
    return parsed


def load_data(source):
    # Intentionally not cached: when the Excel file in GitHub is replaced,
    # Streamlit always reads the latest workbook instead of a stale cached copy.
    df = pd.read_excel(source, sheet_name="Sheet1", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    required = [
        "Name", "Description", "Location On Hand", "Inventory Location",
        "Work Order Group", "Work Order Sub-Group", "Location Average Cost",
        "Location Total Value", "Decision", "Remark"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(missing))

    if "Date transfer" not in df.columns and "Column1" not in df.columns:
        raise ValueError("Missing date column: expected 'Date transfer' or 'Column1'.")

    # Normalize text
    for c in ["Name", "Description", "Inventory Location", "Work Order Group", "Work Order Sub-Group", "Decision", "Remark", "Memo"]:
        if c in df.columns:
            df[c] = df[c].fillna("Unknown").astype(str).str.strip()

    # Normalize numerics
    for c in ["Location On Hand", "Available", "Location Average Cost", "Location Total Value"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # IMPORTANT: use the real Date transfer column first.
    # Column1 is only a fallback for older workbook versions.
    transfer_date = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    if "Date transfer" in df.columns:
        transfer_date = _parse_excel_date_series(df["Date transfer"])
    if "Column1" in df.columns:
        fallback_date = _parse_excel_date_series(df["Column1"])
        transfer_date = transfer_date.fillna(fallback_date)

    df["Transfer Date"] = transfer_date
    today = pd.Timestamp.now().normalize()
    df["Aging Day"] = (today - df["Transfer Date"].dt.normalize()).dt.days.clip(lower=0)
    df["Aging Bucket"] = df["Aging Day"].apply(aging_bucket)
    df["Transfer Date Only"] = df["Transfer Date"].dt.date

    # Fix common spelling in source for display only
    df["Decision Display"] = df["Decision"].replace({"Vendor returrn": "Vendor return"})
    return df

# ------------------------------------------------------------
# DATA SOURCE
# ------------------------------------------------------------
st.sidebar.markdown("## 🎛️ FILTERS")
uploaded = st.sidebar.file_uploader("Upload Write off stock QC.xlsx", type=["xlsx"])

source = uploaded
if source is None:
    if os.path.exists(DEFAULT_FILE):
        source = DEFAULT_FILE
    else:
        st.info("Upload the Excel file from the sidebar, or place 'Write off stock QC.xlsx' beside this Python file.")
        st.stop()

try:
    df = load_data(source)
except Exception as e:
    st.error(f"Cannot read Excel file: {e}")
    st.stop()

raw_total_value = df["Location Total Value"].sum()
missing_transfer_dates = int(df["Transfer Date"].isna().sum())

# ------------------------------------------------------------
# FILTERS
# ------------------------------------------------------------
valid_dates = df["Transfer Date"].dropna()
if not valid_dates.empty:
    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    date_range = st.sidebar.date_input("Date Transfer", value=(min_date, max_date), min_value=min_date, max_value=max_date)
else:
    date_range = None

item_options = sorted(df["Name"].dropna().unique().tolist())
selected_items = st.sidebar.multiselect("Item", item_options, default=[])

decision_options = sorted(df["Decision Display"].dropna().unique().tolist())
selected_decisions = st.sidebar.multiselect("Decision", decision_options, default=decision_options)

group_options = sorted(df["Work Order Group"].dropna().unique().tolist())
selected_groups = st.sidebar.multiselect("Work Order Group", group_options, default=group_options)

subgroup_options = sorted(df["Work Order Sub-Group"].dropna().unique().tolist())
selected_subgroups = st.sidebar.multiselect("Work Order Sub-Group", subgroup_options, default=subgroup_options)

aging_options = [x for x in AGING_ORDER if x in set(df["Aging Bucket"])]
selected_aging = st.sidebar.multiselect("Aging Bucket", aging_options, default=aging_options)

st.sidebar.markdown("---")
top_n = st.sidebar.slider("Top N Items", min_value=1, max_value=20, value=10, step=1)
sort_metric = st.sidebar.selectbox("Top Item By", ["Location Total Value", "Location On Hand", "Aging Day"], index=0)

# Apply filters
f = df.copy()
if date_range and isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = date_range
    valid_date_mask = (
        (f["Transfer Date"].dt.date >= start_date)
        & (f["Transfer Date"].dt.date <= end_date)
    )
    # Keep rows whose transfer date cannot be parsed, so stock value is never
    # silently dropped from the dashboard total.
    f = f[valid_date_mask | f["Transfer Date"].isna()]
if selected_items:
    f = f[f["Name"].isin(selected_items)]
if selected_decisions:
    f = f[f["Decision Display"].isin(selected_decisions)]
if selected_groups:
    f = f[f["Work Order Group"].isin(selected_groups)]
if selected_subgroups:
    f = f[f["Work Order Sub-Group"].isin(selected_subgroups)]
if selected_aging:
    f = f[f["Aging Bucket"].isin(selected_aging)]

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------
st.markdown(
    f"""
    <div class="hero">
      <h1>📦 WRITE-OFF STOCK QC DASHBOARD</h1>
      <p>Stock value, quantity, decision status and aging overview • Refreshed {datetime.now().strftime('%d %b %Y %H:%M')}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if f.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

# ------------------------------------------------------------
# KPI
# ------------------------------------------------------------
total_value = f["Location Total Value"].sum()
total_qty = f["Location On Hand"].sum()
unique_items = f["Name"].nunique()
oldest_aging = f["Aging Day"].max()
writeoff_value = f.loc[f["Decision Display"].eq("Write off"), "Location Total Value"].sum()
onhold_value = f.loc[f["Decision Display"].eq("On hold"), "Location Total Value"].sum()

k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1: kpi_card("Total Stock Value", fmt_vnd(total_value), "VND • selected stock", "#6D28D9")
with k2: kpi_card("Qty On Hand", fmt_num(total_qty), "Total stock quantity", "#2563EB")
with k3: kpi_card("Unique Items", fmt_num(unique_items), "Item codes", "#EC4899")
with k4: kpi_card("Oldest Aging", fmt_num(oldest_aging), "Days", "#F97316")
with k5: kpi_card("Write-Off Value", fmt_vnd(writeoff_value), "VND • Decision = Write off", "#EF4444")
with k6: kpi_card("On-Hold Value", fmt_vnd(onhold_value), "VND • Decision = On hold", "#F59E0B")

# ------------------------------------------------------------
# DECISION + AGING
# ------------------------------------------------------------
st.markdown('<div class="section-title">📊 STOCK STATUS & AGING</div>', unsafe_allow_html=True)
left, right = st.columns([1, 1.35])

with left:
    dec = f.groupby("Decision Display", as_index=False).agg(
        Stock_Value=("Location Total Value", "sum"),
        Qty=("Location On Hand", "sum"),
        Items=("Name", "nunique"),
    ).sort_values("Stock_Value", ascending=False)
    fig = px.pie(
        dec,
        names="Decision Display",
        values="Stock_Value",
        hole=0.58,
        color="Decision Display",
        color_discrete_map=DECISION_COLORS,
    )
    fig.update_traces(
        textposition="inside",
        textinfo="percent+label",
        textfont=dict(size=15, family="Arial Black"),
        marker=dict(line=dict(color="white", width=2)),
        hovertemplate="<b>%{label}</b><br>Value: %{value:,.0f}<br>%{percent}<extra></extra>",
    )
    style_figure(fig, "STOCK VALUE BY DECISION", 450)
    st.plotly_chart(fig, use_container_width=True)

with right:
    aging = f.groupby("Aging Bucket", as_index=False).agg(
        Stock_Value=("Location Total Value", "sum"),
        Qty=("Location On Hand", "sum"),
    )
    aging["Aging Bucket"] = pd.Categorical(aging["Aging Bucket"], categories=AGING_ORDER, ordered=True)
    aging = aging.sort_values("Aging Bucket")
    fig = px.bar(
        aging,
        x="Aging Bucket",
        y="Stock_Value",
        color="Aging Bucket",
        color_discrete_sequence=PALETTE,
        text="Stock_Value",
    )
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", textfont=dict(size=14, family="Arial Black"), cliponaxis=False)
    fig.update_layout(showlegend=False)
    fig.update_yaxes(title="Stock Value (VND)")
    fig.update_xaxes(title="Aging Day")
    style_figure(fig, "STOCK VALUE BY AGING BUCKET", 450)
    st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# GROUP ANALYSIS + TREND
# ------------------------------------------------------------
st.markdown('<div class="section-title">🏭 WORK ORDER GROUP ANALYSIS</div>', unsafe_allow_html=True)
left, right = st.columns([1.1, 1])

with left:
    grp = f.groupby("Work Order Group", as_index=False).agg(
        Stock_Value=("Location Total Value", "sum"),
        Qty=("Location On Hand", "sum"),
        Items=("Name", "nunique"),
    ).sort_values("Stock_Value", ascending=False).head(12)
    fig = px.bar(
        grp.sort_values("Stock_Value"),
        x="Stock_Value",
        y="Work Order Group",
        orientation="h",
        color="Stock_Value",
        color_continuous_scale=["#60A5FA", "#7C3AED", "#EC4899"],
        text="Stock_Value",
    )
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="inside", textfont=dict(size=13, family="Arial Black", color="white"))
    fig.update_layout(coloraxis_showscale=False)
    fig.update_xaxes(title="Stock Value (VND)")
    style_figure(fig, "TOP WORK ORDER GROUPS BY STOCK VALUE", 500)
    st.plotly_chart(fig, use_container_width=True)

with right:
    trend = f.dropna(subset=["Transfer Date"]).groupby("Transfer Date", as_index=False).agg(
        Stock_Value=("Location Total Value", "sum"),
        Qty=("Location On Hand", "sum"),
    ).sort_values("Transfer Date")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=trend["Transfer Date"], y=trend["Stock_Value"], mode="lines+markers",
        name="Stock Value", line=dict(color="#7C3AED", width=4), marker=dict(size=7),
        fill="tozeroy", fillcolor="rgba(124,58,237,0.10)",
        hovertemplate="Date: %{x|%d-%b-%Y}<br>Value: %{y:,.0f}<extra></extra>",
    ))
    style_figure(fig, "STOCK VALUE BY TRANSFER DATE", 500)
    fig.update_yaxes(title="Stock Value (VND)")
    fig.update_xaxes(title="Transfer Date")
    st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# TOP ITEMS
# ------------------------------------------------------------
st.markdown(f'<div class="section-title">🏆 TOP {top_n} ITEM ANALYSIS</div>', unsafe_allow_html=True)

top = f.sort_values(sort_metric, ascending=False).head(top_n).copy()
bar = top.sort_values(sort_metric, ascending=True)
fig = px.bar(
    bar,
    x=sort_metric,
    y="Name",
    orientation="h",
    color=sort_metric,
    color_continuous_scale=["#38BDF8", "#6D28D9", "#F43F5E"],
    text=sort_metric,
    hover_data={"Description": True, "Decision Display": True, "Location Total Value": ":,.0f", "Location On Hand": ":,.0f", "Aging Day": True},
)
fig.update_traces(texttemplate="%{text:,.0f}", textposition="inside", textfont=dict(size=13, family="Arial Black", color="white"))
fig.update_layout(coloraxis_showscale=False)
style_figure(fig, f"TOP {top_n} ITEMS BY {sort_metric.upper()}", max(470, 46 * top_n + 120))
st.plotly_chart(fig, use_container_width=True)

# Big-number item table
top_table = top[["Name", "Description", "Decision Display", "Location On Hand", "Location Average Cost", "Location Total Value", "Aging Day"]].copy()
top_table.insert(0, "Rank", range(1, len(top_table) + 1))
top_table.columns = ["Rank", "Item", "Description", "Decision", "Qty On Hand", "Avg Cost", "Total Value", "Aging Day"]
render_big_table(
    top_table,
    numeric_cols=["Qty On Hand", "Avg Cost", "Total Value", "Aging Day"],
    rank_col="Rank",
    decision_col="Decision",
)

# ------------------------------------------------------------
# DECISION SUMMARY
# ------------------------------------------------------------
st.markdown('<div class="section-title">📋 DECISION SUMMARY</div>', unsafe_allow_html=True)
summary = f.groupby("Decision Display", as_index=False).agg(
    Items=("Name", "nunique"),
    Qty_On_Hand=("Location On Hand", "sum"),
    Stock_Value=("Location Total Value", "sum"),
    Oldest_Aging=("Aging Day", "max"),
)
summary["% of Value"] = summary["Stock_Value"] / summary["Stock_Value"].sum() * 100
summary = summary.sort_values("Stock_Value", ascending=False)
summary.columns = ["Decision", "Items", "Qty On Hand", "Stock Value", "Oldest Aging", "% of Value"]
# Format percent as text to keep table clean
summary["% of Value"] = summary["% of Value"].map(lambda x: f"{x:.1f}%")
render_big_table(summary, numeric_cols=["Items", "Qty On Hand", "Stock Value", "Oldest Aging"], decision_col="Decision")

# ------------------------------------------------------------
# DETAIL TABLE + EXPORT
# ------------------------------------------------------------
st.markdown('<div class="section-title">🔎 DETAIL DATA</div>', unsafe_allow_html=True)
show_cols = [
    "Name", "Description", "Location On Hand", "Inventory Location", "Work Order Group",
    "Work Order Sub-Group", "Location Average Cost", "Location Total Value", "Decision Display",
    "Remark", "Transfer Date", "Aging Day", "Aging Bucket"
]
detail = f[show_cols].copy()
detail.columns = [
    "Item", "Description", "Qty On Hand", "Inventory Location", "WO Group", "WO Sub-Group",
    "Avg Cost", "Total Value", "Decision", "Remark", "Transfer Date", "Aging Day", "Aging Bucket"
]

st.dataframe(
    detail,
    use_container_width=True,
    hide_index=True,
    height=480,
    column_config={
        "Qty On Hand": st.column_config.NumberColumn(format="%.0f"),
        "Avg Cost": st.column_config.NumberColumn(format="%.0f"),
        "Total Value": st.column_config.NumberColumn(format="%.0f"),
        "Aging Day": st.column_config.NumberColumn(format="%d"),
        "Transfer Date": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
    },
)

csv = detail.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Export filtered data to CSV",
    data=csv,
    file_name="writeoff_stock_filtered.csv",
    mime="text/csv",
    use_container_width=False,
)

st.caption(
    f"Source check: {len(df):,} rows loaded • Raw stock value: {raw_total_value:,.0f} VND • "
    f"Rows without parsed Transfer Date: {missing_transfer_dates:,}"
)
st.caption("All figures are calculated from the selected filters. Aging Day is recalculated dynamically from Transfer Date to today.")
