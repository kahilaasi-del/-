"""ניתוחים על הנתונים שנמשכו (data/extract/).

כולל:
  - profitability : רווחיות ומכירות לפי פריט/לקוח/סוכן
  - aging         : גיול חובות לקוחות
  - targets       : ביצוע מול יעדים
  - anomalies     : איתור רשומות חריגות/כפילויות

הרצה:
    python -m src.analysis --report all
    python -m src.analysis --report profitability
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
import yaml

from src.config import EXTRACT_DIR, ROOT, load_mapping


def _load(name: str) -> pd.DataFrame | None:
    """טוען דומיין שנמשך (Parquet אם קיים, אחרת CSV)."""
    pq = EXTRACT_DIR / f"{name}.parquet"
    csv = EXTRACT_DIR / f"{name}.csv"
    if pq.exists():
        return pd.read_parquet(pq)
    if csv.exists():
        return pd.read_csv(csv)
    return None


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


# ───────────────────────── רווחיות ─────────────────────────
def report_profitability() -> None:
    sales = _load("sales")
    if sales is None or sales.empty:
        print("אין נתוני מכירות. הרץ קודם python -m src.extract --domain sales")
        return

    df = sales.copy()
    df["line_total"] = _to_num(df.get("line_total", 0))
    df["quantity"] = _to_num(df.get("quantity", 0))
    has_cost = "cost" in df.columns
    if has_cost:
        df["cost_total"] = _to_num(df["cost"]) * df["quantity"]
        df["profit"] = df["line_total"] - df["cost_total"]

    def summarize(by: str, label: str) -> None:
        if by not in df.columns:
            return
        agg = {"line_total": "sum", "quantity": "sum"}
        if has_cost:
            agg["profit"] = "sum"
        g = df.groupby(by).agg(agg).sort_values("line_total", ascending=False)
        if has_cost:
            g["margin_%"] = (g["profit"] / g["line_total"].replace(0, pd.NA) * 100).round(1)
        print(f"\n── {label} (Top 15) ──")
        print(g.head(15).to_string())

    summarize("item_name" if "item_name" in df.columns else "item", "רווחיות לפי פריט")
    summarize("account", "מכירות לפי לקוח")
    summarize("agent", "מכירות לפי סוכן")


# ───────────────────────── גיול חובות ─────────────────────────
def report_aging(as_of: str | None = None) -> None:
    ledger = _load("ledger")
    if ledger is None or ledger.empty:
        print("אין נתוני כרטסת. הרץ קודם python -m src.extract --domain ledger")
        return

    df = ledger.copy()
    date_col = "value_date" if "value_date" in df.columns else "trans_date"
    if date_col not in df.columns or "account" not in df.columns:
        print("חסרות עמודות account/תאריך בכרטסת — בדוק את המיפוי.")
        return

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df["balance"] = _to_num(df.get("debit", 0)) - _to_num(df.get("credit", 0))
    ref = pd.to_datetime(as_of) if as_of else pd.Timestamp.today()
    df["age_days"] = (ref - df[date_col]).dt.days

    buckets = [(-10**9, 0, "טרם"), (0, 30, "0-30"), (30, 60, "31-60"),
               (60, 90, "61-90"), (90, 120, "91-120"), (120, 10**9, "120+")]

    def bucket(days: float) -> str:
        for lo, hi, name in buckets:
            if lo < days <= hi:
                return name
        return "120+"

    df["bucket"] = df["age_days"].apply(bucket)
    pivot = df.pivot_table(index="account", columns="bucket", values="balance",
                           aggfunc="sum", fill_value=0)
    total_col = 'סה"כ'
    pivot[total_col] = pivot.sum(axis=1)
    pivot = pivot[pivot[total_col].abs() > 0.5].sort_values(total_col, ascending=False)
    grand_total = pivot[total_col].sum()
    print(f"\n── גיול חובות ל-{ref.date()} (Top 20 יתרות) ──")
    print(pivot.head(20).to_string())
    print(f"\nסה\"כ יתרת חייבים: {grand_total:,.0f}")


# ───────────────────────── יעדים ─────────────────────────
def report_targets() -> None:
    sales = _load("sales")
    if sales is None or sales.empty:
        print("אין נתוני מכירות ליעדים.")
        return
    targets_path = ROOT / "config" / "targets.yaml"
    if not targets_path.exists():
        print("אין config/targets.yaml. העתק מ-targets.example.yaml והגדר יעדים.")
        return
    targets = yaml.safe_load(targets_path.read_text(encoding="utf-8")) or {}

    df = sales.copy()
    df["line_total"] = _to_num(df.get("line_total", 0))
    if "doc_date" in df.columns:
        df["doc_date"] = pd.to_datetime(df["doc_date"], errors="coerce")
        df["month"] = df["doc_date"].dt.month

    monthly = targets.get("monthly_total", {})
    if monthly and "month" in df.columns:
        actual = df.groupby("month")["line_total"].sum()
        print("\n── ביצוע מול יעד חודשי ──")
        print(f"{'חודש':>5} {'יעד':>12} {'בפועל':>12} {'%':>7}")
        for m in sorted(monthly):
            goal = monthly[m]
            act = actual.get(m, 0)
            pct = (act / goal * 100) if goal else 0
            print(f"{m:>5} {goal:>12,.0f} {act:>12,.0f} {pct:>6.0f}%")

    by_agent = targets.get("by_agent", {})
    if by_agent and "agent" in df.columns:
        actual = df.groupby("agent")["line_total"].sum()
        print("\n── ביצוע מול יעד לפי סוכן ──")
        for agent, goal in by_agent.items():
            act = actual.get(agent, 0)
            pct = (act / goal * 100) if goal else 0
            print(f"  {agent}: {act:,.0f} / {goal:,.0f}  ({pct:.0f}%)")


# ───────────────────────── אנומליות ─────────────────────────
def report_anomalies() -> None:
    sales = _load("sales")
    if sales is None or sales.empty:
        print("אין נתוני מכירות לבדיקת אנומליות.")
        return
    df = sales.copy()
    print("\n── בדיקות ואנומליות ──")

    # מחיר/כמות שליליים או אפס
    for col, label in [("quantity", "כמות"), ("unit_price", "מחיר יחידה"), ("line_total", "סכום שורה")]:
        if col in df.columns:
            vals = _to_num(df[col])
            neg = (vals < 0).sum()
            zero = (vals == 0).sum()
            print(f"  {label}: {neg:,} שליליים, {zero:,} אפסים")

    # שורות כפולות (אותו מסמך+פריט+כמות)
    key_cols = [c for c in ("doc_id", "item", "quantity") if c in df.columns]
    if len(key_cols) >= 2:
        dups = df.duplicated(subset=key_cols, keep=False).sum()
        print(f"  שורות חשודות ככפולות ({'+'.join(key_cols)}): {dups:,}")

    # חריגות מחיר לכל פריט (מעל 3 סטיות תקן)
    if {"item", "unit_price"}.issubset(df.columns):
        d = df.copy()
        d["unit_price"] = _to_num(d["unit_price"])
        stats = d.groupby("item")["unit_price"].agg(["mean", "std"])
        d = d.join(stats, on="item")
        d["z"] = (d["unit_price"] - d["mean"]) / d["std"].replace(0, pd.NA)
        outliers = d[d["z"].abs() > 3]
        print(f"  חריגות מחיר (|z|>3): {len(outliers):,}")
        if not outliers.empty:
            show = [c for c in ("doc_id", "item", "item_name", "unit_price", "mean") if c in outliers.columns]
            print(outliers[show].head(10).to_string(index=False))


REPORTS = {
    "profitability": report_profitability,
    "aging": report_aging,
    "targets": report_targets,
    "anomalies": report_anomalies,
}


def main() -> None:
    parser = argparse.ArgumentParser(description="ניתוחי BI על נתוני חשבשבת")
    parser.add_argument("--report", default="all", choices=[*REPORTS, "all"])
    args = parser.parse_args()

    if args.report == "all":
        for fn in REPORTS.values():
            fn()
    else:
        REPORTS[args.report]()


if __name__ == "__main__":
    main()
