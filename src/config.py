"""טעינת הגדרות מרכזית: משתני סביבה (.env) + מיפוי הסכימה (mapping.yaml)."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import yaml
from dotenv import load_dotenv

# נתיבי בסיס
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
EXTRACT_DIR = DATA_DIR / "extract"
CONFIG_DIR = ROOT / "config"
MAPPING_PATH = CONFIG_DIR / "mapping.yaml"

# טוען את .env אם קיים
load_dotenv(ROOT / ".env")


@dataclass
class DbSettings:
    """פרטי חיבור ל-SQL Server."""

    server: str
    database: str
    user: str | None
    password: str | None
    driver: str
    trust_cert: bool

    @property
    def use_windows_auth(self) -> bool:
        """אם לא הוגדרו משתמש/סיסמה — משתמשים ב-Windows Authentication."""
        return not (self.user and self.password)


def load_db_settings() -> DbSettings:
    """קורא את פרטי החיבור ממשתני הסביבה."""
    server = os.getenv("CHESHBSHEVET_SERVER", "").strip()
    database = os.getenv("CHESHBSHEVET_DATABASE", "").strip()
    if not server:
        raise RuntimeError(
            "חסר CHESHBSHEVET_SERVER ב-.env. העתק את .env.example ל-.env ומלא את הפרטים."
        )
    if not database:
        raise RuntimeError("חסר CHESHBSHEVET_DATABASE ב-.env (שם מסד הנתונים של החברה).")

    trust = os.getenv("CHESHBSHEVET_TRUST_CERT", "yes").strip().lower() in ("1", "yes", "true")
    return DbSettings(
        server=server,
        database=database,
        user=os.getenv("CHESHBSHEVET_USER", "").strip() or None,
        password=os.getenv("CHESHBSHEVET_PASSWORD", "").strip() or None,
        driver=os.getenv("CHESHBSHEVET_DRIVER", "ODBC Driver 17 for SQL Server").strip(),
        trust_cert=trust,
    )


def load_mapping() -> dict:
    """קורא את mapping.yaml. מחזיר dict ריק אם לא קיים עדיין."""
    if not MAPPING_PATH.exists():
        return {}
    with open(MAPPING_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def ensure_dirs() -> None:
    """מוודא שתיקיות הפלט קיימות."""
    DATA_DIR.mkdir(exist_ok=True)
    EXTRACT_DIR.mkdir(exist_ok=True)
