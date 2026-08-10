"""חיבור מאובטח (קריאה בלבד) ל-SQL Server של חשבשבת.

שימוש:
    from src.db import get_engine, read_sql
    df = read_sql("SELECT TOP 10 * FROM SomeTable")

בדיקת חיבור מהטרמינל:
    python -m src.db --test
"""
from __future__ import annotations

import sys
import urllib.parse

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

from src.config import DbSettings, load_db_settings

_engine: Engine | None = None


def _build_odbc_connection_string(cfg: DbSettings) -> str:
    """בונה מחרוזת חיבור ODBC לפי סוג האימות."""
    parts = [
        f"DRIVER={{{cfg.driver}}}",
        f"SERVER={cfg.server}",
        f"DATABASE={cfg.database}",
    ]
    if cfg.use_windows_auth:
        parts.append("Trusted_Connection=yes")
    else:
        parts.append(f"UID={cfg.user}")
        parts.append(f"PWD={cfg.password}")
    if cfg.trust_cert:
        parts.append("TrustServerCertificate=yes")
    # קריאה בלבד — רמז לשרת שאיננו כותבים
    parts.append("ApplicationIntent=ReadOnly")
    return ";".join(parts)


def get_engine(cfg: DbSettings | None = None) -> Engine:
    """מחזיר SQLAlchemy Engine (singleton)."""
    global _engine
    if _engine is not None:
        return _engine
    cfg = cfg or load_db_settings()
    odbc = _build_odbc_connection_string(cfg)
    url = "mssql+pyodbc:///?odbc_connect=" + urllib.parse.quote_plus(odbc)
    # pool_pre_ping מוודא שהחיבור חי לפני שימוש
    _engine = create_engine(url, pool_pre_ping=True, fast_executemany=False)
    return _engine


def read_sql(query: str, params: dict | None = None) -> pd.DataFrame:
    """מריץ שאילתת SELECT ומחזיר DataFrame. קריאה בלבד."""
    stripped = query.lstrip().lower()
    # שכבת הגנה: מונע כתיבה בטעות
    forbidden = ("insert ", "update ", "delete ", "drop ", "truncate ", "alter ", "exec ")
    if any(stripped.startswith(w) or f";{w}" in stripped for w in forbidden):
        raise ValueError("מותרות שאילתות קריאה בלבד (SELECT).")
    engine = get_engine()
    with engine.connect() as conn:
        return pd.read_sql(text(query), conn, params=params or {})


def test_connection() -> bool:
    """בודק חיבור בסיסי ומדפיס פרטים."""
    try:
        cfg = load_db_settings()
        print(f"מתחבר אל: {cfg.server} / {cfg.database}")
        print(f"אימות: {'Windows' if cfg.use_windows_auth else 'SQL (' + str(cfg.user) + ')'}")
        df = read_sql(
            "SELECT @@VERSION AS version, DB_NAME() AS db, "
            "(SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES) AS table_count"
        )
        row = df.iloc[0]
        print("\n✓ החיבור הצליח!")
        print(f"  מסד נתונים: {row['db']}")
        print(f"  מספר טבלאות: {row['table_count']}")
        print(f"  גרסה: {str(row['version']).splitlines()[0]}")
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"\n✗ החיבור נכשל: {exc}", file=sys.stderr)
        print("\nבדוק: כתובת שרת, שם DB, הרשאות, והתקנת ODBC Driver.", file=sys.stderr)
        return False


if __name__ == "__main__":
    if "--test" in sys.argv:
        sys.exit(0 if test_connection() else 1)
    print("שימוש: python -m src.db --test")
