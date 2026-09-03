# -*- coding: utf-8 -*-
"""
סידור קטלוג מבצעים לפי קטגוריות ראשיות — עם SUN/STAR ראשונים.

לוקח גיליון קטלוג (כמו 'מבצעים ספטמבר 26'), מקבץ מחדש את הפריטים לפי
'קטגוריות ראשית', וכך שבתוך כל קטגוריה הפריטים המסומנים ב-priority = sun/star
מופיעים ראשונים. שאר הפריטים אחריהם. סדר הקטגוריות = סדר ההופעה המקורי
בגיליון (ו-#N/A / ריק בסוף). המאסטר המקורי לא נגע — נכתב קובץ חדש.

דוגמה:
    python -m promotions.organize_catalog \
        --input  "קובץ המבצעים.xlsx" \
        --sheet  "מבצעים ספטמבר 26" \
        --out    "data/מבצעים ספטמבר מסודר.xlsx" \
        --month  ספטמבר
"""
from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FILL_HEADER = PatternFill("solid", fgColor="305496")
FILL_CATEGORY = PatternFill("solid", fgColor="D9E1F2")   # פס קטגוריה בהיר
FILL_PRIORITY = PatternFill("solid", fgColor="FFF2CC")   # רקע עדין ל-sun/star

PRIORITY_FIRST = {"sun", "star"}
CAT_HEADER = "קטגוריות ראשית"
PRI_HEADER = "priority"
KEY_HEADER = "מפתח פריט"


def norm(v) -> str:
    return "" if v is None else str(v).replace("‏", "").strip()


def find_header_row(ws, key=KEY_HEADER, max_scan=15):
    """מאתר את שורת הכותרת — זו שמכילה 'מפתח פריט'."""
    for r in range(1, min(max_scan, ws.max_row) + 1):
        row_vals = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if key in row_vals:
            return r
    raise SystemExit(f"לא נמצאה שורת כותרת עם '{key}' ב-{max_scan} השורות הראשונות")


def keep_columns(ws, header_row):
    """מחזיר אינדקסי עמודות אמיתיות: כותרת לא-ריקה ושאינה כפילות '.1'."""
    keep = []
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(header_row, c).value)
        if h and not h.endswith(".1"):
            keep.append(c)
    return keep


def organize(input_path, sheet, month, add_visuals=True):
    wb = openpyxl.load_workbook(input_path, data_only=True)
    if sheet not in wb.sheetnames:
        raise SystemExit(f"גיליון '{sheet}' לא נמצא. גיליונות: {wb.sheetnames}")
    ws = wb[sheet]

    header_row = find_header_row(ws)
    cols = keep_columns(ws, header_row)
    headers = [norm(ws.cell(header_row, c).value) for c in cols]
    if CAT_HEADER not in headers:
        raise SystemExit(f"לא נמצאה עמודת '{CAT_HEADER}'. עמודות: {headers}")
    cat_pos = headers.index(CAT_HEADER)
    pri_pos = headers.index(PRI_HEADER) if PRI_HEADER in headers else None

    # שורות שמעל הכותרת (כותרות-על) — לשמר כפי שהן
    top_rows = []
    for r in range(1, header_row):
        top_rows.append([ws.cell(r, c).value for c in cols])

    # קריאת נתוני הפריטים
    data = []
    for r in range(header_row + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in cols]
        if all(v in (None, "") for v in vals):
            continue
        data.append(vals)

    # סדר קטגוריות = סדר הופעה מקורי; #N/A / ריק בסוף
    cat_order, seen = [], set()
    for vals in data:
        cat = norm(vals[cat_pos])
        if cat not in seen:
            seen.add(cat)
            cat_order.append(cat)

    def cat_rank(cat):
        c = norm(cat)
        if c in ("", "#N/A", "#n/a", "None"):
            return (2, 0)
        return (0, cat_order.index(c))

    def pri_rank(vals):
        p = norm(vals[pri_pos]).lower() if pri_pos is not None else ""
        return 0 if p in PRIORITY_FIRST else 1  # sun/star ראשונים

    # מיון יציב: קטגוריה -> (sun/star ראשון) -> סדר מקורי נשמר
    indexed = list(enumerate(data))
    indexed.sort(key=lambda t: (cat_rank(t[1][cat_pos]), pri_rank(t[1]), t[0]))
    ordered = [vals for _, vals in indexed]

    return headers, top_rows, ordered, cat_pos, pri_pos, add_visuals


def write_xlsx(out_path, sheet_title, headers, top_rows, rows, cat_pos, pri_pos, add_visuals):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.sheet_view.rightToLeft = True

    r = 1
    for tr in top_rows:  # כותרות-על
        for c, v in enumerate(tr, 1):
            ws.cell(r, c, v)
        ws.cell(r, 1).font = Font(bold=True)
        r += 1

    header_row_idx = r
    for c, h in enumerate(headers, 1):
        cell = ws.cell(r, c, h)
        cell.fill = FILL_HEADER
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = ws.cell(header_row_idx + 1, 1).coordinate
    r += 1

    prev_cat = object()
    for vals in rows:
        cat = norm(vals[cat_pos])
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v)
        if add_visuals:
            if cat != prev_cat:  # פס קטגוריה בתחילת קבוצה
                for c in range(1, len(headers) + 1):
                    ws.cell(r, c).fill = FILL_CATEGORY
                    ws.cell(r, c).font = Font(bold=True)
                prev_cat = cat
            elif pri_pos is not None and norm(vals[pri_pos]).lower() in PRIORITY_FIRST:
                ws.cell(r, pri_pos + 1).fill = FILL_PRIORITY
                ws.cell(r, pri_pos + 1).font = Font(bold=True)
        r += 1

    # רוחב עמודות
    for c, h in enumerate(headers, 1):
        width = max([len(str(h))] + [len(str(ws.cell(rr, c).value or "")) for rr in range(header_row_idx, r)])
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 10), 45)

    wb.save(out_path)
    return out_path


def main(argv=None):
    p = argparse.ArgumentParser(description="סידור קטלוג מבצעים לפי קטגוריות עם sun/star ראשונים")
    p.add_argument("--input", required=True, help="קובץ הקלט (XLSX)")
    p.add_argument("--sheet", required=True, help="שם הגיליון לסידור")
    p.add_argument("--out", required=True, help="נתיב פלט (XLSX חדש)")
    p.add_argument("--month", default="", help="שם החודש לכותרת הגיליון")
    p.add_argument("--no-visuals", action="store_true", help="בלי צביעה/פסי קטגוריה — רק סידור")
    args = p.parse_args(argv)

    headers, top_rows, rows, cat_pos, pri_pos, add_visuals = organize(
        args.input, args.sheet, args.month, add_visuals=not args.no_visuals
    )
    title = f"מבצעים {args.month} מסודר" if args.month else f"{args.sheet} מסודר"
    out = write_xlsx(args.out, title, headers, top_rows, rows, cat_pos, pri_pos, add_visuals)

    # סיכום
    from collections import Counter
    cats = Counter(norm(v[cat_pos]) for v in rows)
    sunstar = sum(1 for v in rows if pri_pos is not None and norm(v[pri_pos]).lower() in PRIORITY_FIRST)
    print(f"✅ נכתב: {out}")
    print(f"   {len(rows)} פריטים · {len(cats)} קטגוריות · {sunstar} מסומנים sun/star (ראשונים בכל קטגוריה)")


if __name__ == "__main__":
    main()
