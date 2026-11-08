# -*- coding: utf-8 -*-
"""
חילוץ פריטי המבצע מהמאסטר -> גיליון מוכן לסידור.

הזרימה החודשית המפושטת: ממלאים את עמודת "מבצעים" במאסטר (למי שבמבצע),
מייצאים את המאסטר (עם קטגוריות ראשית / priority / מלאי / מחיר / כשרות),
והסקריפט הזה מסנן רק את הפריטים שיש להם ערך בעמודת "מבצעים" ומוציא גיליון
נקי (שורת כותרת בשורה 1) — מוכן ל-organize_catalog.

דוגמה (שני שלבים):
    python -m promotions.extract_promos --master "המאסטר.xlsx" --out "data/נובמבר_מבצע.xlsx"
    python -m promotions.organize_catalog --input "data/נובמבר_מבצע.xlsx" \
        --sheet נובמבר --out "data/מבצעים נובמבר מסודר.xlsx" --month נובמבר
"""
from __future__ import annotations
import argparse
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

# העמודות שנשמרות לפלט (אם קיימות במאסטר)
COLS = ["קוד מיון פריט", "קטגוריות ראשית", "מפתח פריט", "שם פריט", "בר קוד", "priority",
        "מחיר קניה", "מחיר מכירה", "כמות בקרטון", "כמות במשטח", "פריט פקדון", "כשרות",
        "מבצעים", "תוקף 1 במוצר", "SNAPSHOT-55"]
PROMO_COL = "מבצעים"
KEY_COL = "מפתח פריט"


def find_header_row(ws, key=KEY_COL, max_scan=8):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        if key in [str(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]:
            return r
    raise SystemExit(f"לא נמצאה שורת כותרת עם '{key}'")


def main(argv=None):
    p = argparse.ArgumentParser(description="חילוץ פריטי מבצע מהמאסטר")
    p.add_argument("--master", required=True, help="קובץ המאסטר (XLSX)")
    p.add_argument("--sheet", help="שם גיליון המאסטר (ברירת מחדל: הראשון)")
    p.add_argument("--out", required=True, help="נתיב פלט (XLSX)")
    p.add_argument("--out-sheet", default="נובמבר", help="שם גיליון הפלט")
    args = p.parse_args(argv)

    wb = openpyxl.load_workbook(args.master, data_only=True)
    m = wb[args.sheet] if args.sheet else wb[wb.sheetnames[0]]
    hrow = find_header_row(m)
    hdr = [str(m.cell(hrow, c).value) for c in range(1, m.max_column + 1)]
    idx = {h: i + 1 for i, h in enumerate(hdr)}
    if PROMO_COL not in idx:
        raise SystemExit(f"לא נמצאה עמודת '{PROMO_COL}' במאסטר. עמודות: {hdr}")

    keep = [c for c in COLS if c in idx]
    out = openpyxl.Workbook(); ws = out.active
    ws.title = args.out_sheet[:31]; ws.sheet_view.rightToLeft = True
    ws.append(keep)

    n = 0
    for r in range(hrow + 1, m.max_row + 1):
        mb = m.cell(r, idx[PROMO_COL]).value
        if mb in (None, "") or str(mb).strip() in ("", "0", "0.0"):
            continue
        ws.append([m.cell(r, idx[h]).value for h in keep])
        n += 1

    for i, h in enumerate(keep, 1):
        w = max([len(str(h))] + [len(str(ws.cell(r, i).value or "")) for r in range(2, ws.max_row + 1)])
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 10), 40)

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    out.save(args.out)
    print(f"✅ נכתב: {args.out} | פריטי מבצע: {n}")
    print(f"   עכשיו: python -m promotions.organize_catalog --input \"{args.out}\" --sheet {args.out_sheet} --out \"data/מבצעים מסודר.xlsx\" --month <חודש>")


if __name__ == "__main__":
    main()
