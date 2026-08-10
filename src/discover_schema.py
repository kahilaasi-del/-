"""גילוי סכימה: סורק את מסד הנתונים של חשבשבת ומפיק קטלוג + ניחוש דומיינים.

מפני שמבנה חשבשבת SQL אינו סטנדרטי ומשתנה בין גרסאות, אנחנו לא מנחשים שמות
טבלאות בקוד. במקום זה סורקים את המסד האמיתי, ולפי מילות מפתח בשמות העמודות
מציעים אילו טבלאות מתאימות לכל דומיין. את התוצאה מאמתים ידנית ב-mapping.yaml.

הרצה:
    python -m src.discover_schema
"""
from __future__ import annotations

import json
from collections import defaultdict

import pandas as pd

from src.config import DATA_DIR, ensure_dirs
from src.db import read_sql

# מילות מפתח לזיהוי דומיינים — עברית ותבניות נפוצות בחשבשבת.
# ההתאמה היא רמז בלבד; יש לאמת ידנית.
DOMAIN_KEYWORDS: dict[str, list[str]] = {
    "documents_sales": [
        "מסמך", "חשבונית", "מכיר", "תעוד", "invoice", "doc", "sale",
        "מסמכים", "שורות מסמך", "פרטי מסמך",
    ],
    "journal_ledger": [
        "תנוע", "יומן", "כרטס", "אסמכת", "journal", "movement", "trans",
        "חובה", "זכות", "debit", "credit",
    ],
    "accounts": [
        "חשבון", "כרטיס", "לקוח", "ספק", "account", "customer", "supplier",
        "cardname", "כרטיסי חשבון",
    ],
    "inventory_items": [
        "פריט", "מלאי", "מק\"ט", "מקט", "item", "product", "stock", "sku",
        "מחירון", "כמות", "יחיד",
    ],
    "agents": [
        "סוכן", "agent", "salesman", "מפיץ", "נציג",
    ],
}


def list_tables() -> pd.DataFrame:
    """מחזיר טבלאות עם מספר שורות משוער (מ-sys.partitions, מהיר)."""
    query = """
    SELECT
        s.name  AS schema_name,
        t.name  AS table_name,
        SUM(CASE WHEN p.index_id IN (0,1) THEN p.rows ELSE 0 END) AS row_count
    FROM sys.tables t
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    JOIN sys.partitions p ON t.object_id = p.object_id
    GROUP BY s.name, t.name
    ORDER BY row_count DESC
    """
    return read_sql(query)


def list_columns() -> pd.DataFrame:
    """מחזיר את כל העמודות של כל הטבלאות."""
    query = """
    SELECT
        TABLE_SCHEMA  AS schema_name,
        TABLE_NAME    AS table_name,
        COLUMN_NAME   AS column_name,
        DATA_TYPE     AS data_type,
        ORDINAL_POSITION AS position
    FROM INFORMATION_SCHEMA.COLUMNS
    ORDER BY TABLE_NAME, ORDINAL_POSITION
    """
    return read_sql(query)


def score_domains(cols: pd.DataFrame) -> dict[str, list[tuple[str, int, list[str]]]]:
    """מנקד לכל טבלה עד כמה היא מתאימה לכל דומיין, לפי מילות מפתח בעמודות."""
    # מקבץ עמודות לפי טבלה
    by_table: dict[str, list[str]] = defaultdict(list)
    for _, r in cols.iterrows():
        by_table[f"{r['schema_name']}.{r['table_name']}"].append(str(r["column_name"]))

    results: dict[str, list[tuple[str, int, list[str]]]] = {d: [] for d in DOMAIN_KEYWORDS}
    for table, columns in by_table.items():
        joined = " ".join(columns).lower()
        table_lower = table.lower()
        for domain, keywords in DOMAIN_KEYWORDS.items():
            matched = [kw for kw in keywords if kw.lower() in joined or kw.lower() in table_lower]
            if matched:
                results[domain].append((table, len(matched), matched))
    # מיון לפי מספר התאמות
    for domain in results:
        results[domain].sort(key=lambda x: x[1], reverse=True)
    return results


def build_catalog(tables: pd.DataFrame, cols: pd.DataFrame) -> dict:
    """בונה קטלוג מלא: טבלה -> {row_count, columns[]}."""
    row_counts = {
        f"{r['schema_name']}.{r['table_name']}": int(r["row_count"])
        for _, r in tables.iterrows()
    }
    catalog: dict[str, dict] = {}
    for _, r in cols.iterrows():
        key = f"{r['schema_name']}.{r['table_name']}"
        entry = catalog.setdefault(key, {"row_count": row_counts.get(key, 0), "columns": []})
        entry["columns"].append({"name": r["column_name"], "type": r["data_type"]})
    return catalog


def write_report(tables: pd.DataFrame, domains: dict, catalog: dict) -> None:
    """כותב דוח קריא בעברית ל-schema_report.md."""
    lines: list[str] = []
    lines.append("# דוח סכימה — חשבשבת\n")
    lines.append(f"נמצאו **{len(catalog)}** טבלאות.\n")

    lines.append("## 20 הטבלאות הגדולות ביותר (לפי שורות)\n")
    lines.append("| טבלה | שורות | מס' עמודות |")
    lines.append("|------|-------|-----------|")
    for _, r in tables.head(20).iterrows():
        key = f"{r['schema_name']}.{r['table_name']}"
        ncols = len(catalog.get(key, {}).get("columns", []))
        lines.append(f"| {key} | {int(r['row_count']):,} | {ncols} |")
    lines.append("")

    lines.append("## מועמדים לפי דומיין (ניחוש — יש לאמת!)\n")
    domain_titles = {
        "documents_sales": "📄 מסמכים / מכירות",
        "journal_ledger": "📚 תנועות יומן / כרטסת",
        "accounts": "👥 חשבונות / לקוחות / ספקים",
        "inventory_items": "📦 פריטים / מלאי",
        "agents": "🧑‍💼 סוכנים",
    }
    for domain, title in domain_titles.items():
        lines.append(f"### {title}\n")
        candidates = domains.get(domain, [])[:5]
        if not candidates:
            lines.append("_לא נמצאו מועמדים אוטומטית — בדוק ידנית ב-catalog.json._\n")
            continue
        lines.append("| טבלה | התאמות | מילות מפתח שנמצאו |")
        lines.append("|------|--------|------------------|")
        for table, score, matched in candidates:
            rc = catalog.get(table, {}).get("row_count", 0)
            lines.append(f"| {table} ({rc:,} שורות) | {score} | {', '.join(matched)} |")
        lines.append("")

    lines.append("## השלב הבא\n")
    lines.append(
        "1. בדוק את המועמדים מול `catalog.json` (העמודות המלאות).\n"
        "2. מלא את `config/mapping.yaml` לפי הטבלאות והעמודות האמיתיות.\n"
        "3. אפשר לשלוח את הקובץ הזה ל-Claude כדי לבנות יחד את המיפוי.\n"
    )

    report_path = DATA_DIR / "schema_report.md"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ דוח נכתב: {report_path}")


def main() -> None:
    ensure_dirs()
    print("סורק טבלאות...")
    tables = list_tables()
    print(f"  נמצאו {len(tables)} טבלאות.")
    print("סורק עמודות...")
    cols = list_columns()
    print(f"  נמצאו {len(cols)} עמודות.")

    catalog = build_catalog(tables, cols)
    domains = score_domains(cols)

    catalog_path = DATA_DIR / "catalog.json"
    catalog_path.write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✓ קטלוג נכתב: {catalog_path}")

    write_report(tables, domains, catalog)
    print("\nהשלב הבא: מלא את config/mapping.yaml לפי data/schema_report.md")


if __name__ == "__main__":
    main()
