"""משיכת נתונים מחשבשבת לפי config/mapping.yaml.

הקוד בונה שאילתות SELECT שממפות את שמות העמודות האמיתיים לשמות לוגיים
(alias), ושומר כל דומיין ל-Parquet + CSV תחת data/extract/.

הרצה:
    python -m src.extract                 # כל הדומיינים המוגדרים
    python -m src.extract --domain sales  # דומיין יחיד
"""
from __future__ import annotations

import argparse
import sys

import pandas as pd

from src.config import EXTRACT_DIR, ensure_dirs, load_mapping
from src.db import read_sql

# דומיינים שנמשכים כטבלה עם מיפוי עמודות
TABLE_DOMAINS = ("sales", "accounts", "ledger", "items")

PLACEHOLDER_MARK = "___"  # סימון בקובץ הדוגמה שטרם מולא


def _is_filled(value: str | None) -> bool:
    return bool(value) and PLACEHOLDER_MARK not in str(value)


def _build_query(table: str, columns: dict[str, str]) -> str:
    """בונה SELECT עם alias: real_col AS logical_name, רק לעמודות שמולאו."""
    selected = [
        f"[{real}] AS [{logical}]"
        for logical, real in columns.items()
        if _is_filled(real)
    ]
    if not selected:
        raise ValueError("לא הוגדרו עמודות (כולן עדיין placeholder).")
    return f"SELECT {', '.join(selected)} FROM {table}"


def extract_domain(name: str, spec: dict) -> pd.DataFrame | None:
    """מושך דומיין בודד לפי המיפוי שלו."""
    table = spec.get("table", "")
    columns = spec.get("columns", {}) or {}
    if not _is_filled(table):
        print(f"  ⚠ דומיין '{name}': טבלה לא מוגדרת ב-mapping.yaml — מדלג.")
        return None
    try:
        query = _build_query(table, columns)
    except ValueError as exc:
        print(f"  ⚠ דומיין '{name}': {exc} — מדלג.")
        return None

    print(f"  מושך '{name}' מ-{table}...")
    df = read_sql(query)
    print(f"    ✓ {len(df):,} שורות, {len(df.columns)} עמודות.")
    return df


def save(df: pd.DataFrame, name: str) -> None:
    """שומר Parquet (מהיר לניתוח) + CSV (לשליחה/פתיחה ידנית)."""
    parquet_path = EXTRACT_DIR / f"{name}.parquet"
    csv_path = EXTRACT_DIR / f"{name}.csv"
    try:
        df.to_parquet(parquet_path, index=False)
    except Exception as exc:  # noqa: BLE001  (pyarrow עלול לא להיות מותקן)
        print(f"    (Parquet נכשל: {exc} — שומר CSV בלבד)")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")  # utf-8-sig לפתיחה תקינה באקסל
    print(f"    נשמר: {csv_path.name}")


def main() -> None:
    parser = argparse.ArgumentParser(description="משיכת נתונים מחשבשבת")
    parser.add_argument("--domain", choices=TABLE_DOMAINS, help="דומיין יחיד (ברירת מחדל: הכל)")
    args = parser.parse_args()

    ensure_dirs()
    mapping = load_mapping()
    if not mapping:
        print("✗ אין config/mapping.yaml. העתק מ-mapping.example.yaml ומלא לפי דוח הסכימה.")
        sys.exit(1)

    domains = [args.domain] if args.domain else list(TABLE_DOMAINS)
    extracted = 0
    for name in domains:
        spec = mapping.get(name)
        if not spec:
            print(f"  ⚠ אין הגדרה לדומיין '{name}' ב-mapping.yaml — מדלג.")
            continue
        df = extract_domain(name, spec)
        if df is not None:
            save(df, name)
            extracted += 1

    print(f"\nהושלם. נמשכו {extracted} דומיינים ל-{EXTRACT_DIR}")


if __name__ == "__main__":
    main()
