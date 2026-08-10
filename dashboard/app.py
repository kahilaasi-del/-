"""דשבורד BI לחשבשבת — Streamlit.

הרצה:
    streamlit run dashboard/app.py

הדשבורד קורא מ-data/extract/ (מה שנמשך ב-src.extract). אם אין נתונים,
תוצג הנחיה להריץ קודם את שלב המשיכה.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# מאפשר ייבוא מ-src כשמריצים דרך streamlit
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.config import EXTRACT_DIR  # noqa: E402

st.set_page_config(page_title="חשבשבת BI", page_icon="📊", layout="wide")

# פלטת צבעים עקבית ונגישה
PALETTE = ["#4C78A8", "#F58518", "#54A24B", "#E45756", "#72B7B2", "#B279A2"]


@st.cache_data
def load(name: str) -> pd.DataFrame | None:
    pq = EXTRACT_DIR / f"{name}.parquet"
    csv = EXTRACT_DIR / f"{name}.csv"
    if pq.exists():
        return pd.read_parquet(pq)
    if csv.exists():
        return pd.read_csv(csv)
    return None


def num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


st.title("📊 דשבורד חשבשבת BI")

sales = load("sales")
if sales is None or sales.empty:
    st.warning(
        "לא נמצאו נתוני מכירות. הרץ קודם:\n\n"
        "1. `python -m src.discover_schema`\n"
        "2. מלא את `config/mapping.yaml`\n"
        "3. `python -m src.extract`"
    )
    st.stop()

sales = sales.copy()
sales["line_total"] = num(sales.get("line_total", 0))
sales["quantity"] = num(sales.get("quantity", 0))
if "doc_date" in sales.columns:
    sales["doc_date"] = pd.to_datetime(sales["doc_date"], errors="coerce")

# ── סרגל צד: מסננים ──
st.sidebar.header("מסננים")
if "doc_date" in sales.columns and sales["doc_date"].notna().any():
    min_d, max_d = sales["doc_date"].min(), sales["doc_date"].max()
    date_range = st.sidebar.date_input("טווח תאריכים", (min_d, max_d))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        lo, hi = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        sales = sales[(sales["doc_date"] >= lo) & (sales["doc_date"] <= hi)]

if "agent" in sales.columns:
    agents = ["הכל"] + sorted(sales["agent"].dropna().astype(str).unique().tolist())
    pick = st.sidebar.selectbox("סוכן", agents)
    if pick != "הכל":
        sales = sales[sales["agent"].astype(str) == pick]

# ── כרטיסי KPI ──
c1, c2, c3, c4 = st.columns(4)
total = sales["line_total"].sum()
c1.metric("סה\"כ מכירות", f"{total:,.0f}")
c2.metric("מספר מסמכים", f"{sales['doc_id'].nunique():,}" if "doc_id" in sales.columns else "—")
c3.metric("כמות פריטים", f"{sales['quantity'].sum():,.0f}")
if "cost" in sales.columns:
    profit = total - (num(sales["cost"]) * sales["quantity"]).sum()
    c4.metric("רווח גולמי", f"{profit:,.0f}", f"{profit/total*100:.0f}%" if total else None)

st.divider()

# ── מגמת מכירות לאורך זמן ──
if "doc_date" in sales.columns and sales["doc_date"].notna().any():
    st.subheader("מגמת מכירות")
    trend = (sales.set_index("doc_date")["line_total"]
             .resample("M").sum().reset_index())
    fig = px.line(trend, x="doc_date", y="line_total", markers=True,
                  color_discrete_sequence=PALETTE)
    fig.update_layout(xaxis_title="", yaxis_title="מכירות", height=340)
    st.plotly_chart(fig, use_container_width=True)

col_a, col_b = st.columns(2)

# ── Top פריטים ──
item_col = "item_name" if "item_name" in sales.columns else ("item" if "item" in sales.columns else None)
if item_col:
    with col_a:
        st.subheader("פריטים מובילים")
        top = (sales.groupby(item_col)["line_total"].sum()
               .sort_values(ascending=False).head(10).reset_index())
        fig = px.bar(top, x="line_total", y=item_col, orientation="h",
                     color_discrete_sequence=PALETTE)
        fig.update_layout(xaxis_title="מכירות", yaxis_title="", height=380,
                          yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

# ── Top לקוחות ──
if "account" in sales.columns:
    with col_b:
        st.subheader("לקוחות מובילים")
        top = (sales.groupby("account")["line_total"].sum()
               .sort_values(ascending=False).head(10).reset_index())
        top["account"] = top["account"].astype(str)
        fig = px.bar(top, x="line_total", y="account", orientation="h",
                     color_discrete_sequence=[PALETTE[1]])
        fig.update_layout(xaxis_title="מכירות", yaxis_title="", height=380,
                          yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

# ── מכירות לפי סוכן ──
if "agent" in sales.columns:
    st.subheader("מכירות לפי סוכן")
    by_agent = (sales.groupby("agent")["line_total"].sum()
                .sort_values(ascending=False).reset_index())
    by_agent["agent"] = by_agent["agent"].astype(str)
    fig = px.bar(by_agent, x="agent", y="line_total", color_discrete_sequence=PALETTE)
    fig.update_layout(xaxis_title="", yaxis_title="מכירות", height=340)
    st.plotly_chart(fig, use_container_width=True)

st.caption("נתונים מתוך data/extract/ · לרענון: הרץ מחדש את src.extract ולחץ 'R'.")
