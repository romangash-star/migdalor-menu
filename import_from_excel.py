#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
import_from_excel.py — מייבא את הפריטים לאתר מייצוא Excel של לוח Monday,
ומשייך לכל פריט את הקבצים שהורדת.

שימוש:
  python import_from_excel.py --excel "ייצוא.xlsx" --files-dir "C:/Users/roman/Downloads/monday"
  python import_from_excel.py --excel "ייצוא.xlsx" --files-dir "..." --dry-run

מה הוא עושה:
  1. קורא את הייצוא (קיבוץ לפי תחום, עמודות בעברית).
  2. לכל קישור קובץ ב-Monday מוציא את שם הקובץ, ומחפש אותו בתיקייה שהורדת.
  3. מעתיק את הקבצים שנמצאו לתיקיית files/ ומקשר אותם לפריט.
  4. שומר גיבוי של index.html ואז כותב את הנתונים החדשים.
"""

import argparse, json, re, shutil, sys, unicodedata, urllib.parse
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT       = Path(__file__).parent
HTML_FILE  = ROOT / "index.html"
FILES_DIR  = ROOT / "files"
BACKUP_DIR = ROOT / "backups"

# כותרת בייצוא  →  שם השדה באתר
COLUMNS = {
    "Name":                       "name",
    "תחום ראשי בתפריט":           "category",
    "שיוך מחלקתי":                "dept",
    "שם איש קשר להטמעה":          "contactName",
    "טלפון ליצירת קשר":           "phone",
    "אי מייל ליצירת קשר":         "email",
    "קהל יעד":                    "audience",
    "נושא תוכן מרכזי":            "topic",
    'האם הפריט קיים במאגר גפ"ן':  "gafan",
    "סוג פריט":                   "type",
}
FILE_COLUMN = "קובץ פירוט"
DESC_COLUMNS = ("תיאור", "פירוט", "תוכן")      # אופציונלי, אם קיים בייצוא

# בייצוא מופיעים לפעמים שמות תחום בוריאציות שונות. כאן הם מאוחדים לשם אחד,
# כדי שלא ייווצרו באתר תחומים כפולים בלי צבע.
CATEGORY_ALIASES = {
    "חזון ותפיסת יעוד":        "חזון ויעוד",
    "זהות וייעוד":             "חזון ויעוד",
    "כוח אדם":                 "כח אדם",
    "תכנית חינוכית וקהילתית":  "תכניות חינוכיות וקהילתיות",
    "תכניות חינוכיות":         "תכניות חינוכיות וקהילתיות",
    "תרבות ארגונים":           "תרבות ארגונית",
}

CATEGORY_ORDER = [
    "חזון ויעוד",
    "כח אדם",
    "תכניות חינוכיות וקהילתיות",
    "תרבות ארגונית",
    "שותפויות ומשאבים",
]

ITEM_FIELDS = ["name", "category", "dept", "contactName", "phone", "email",
               "audience", "topic", "gafan", "type", "description"]


def norm(text: str) -> str:
    """שם קובץ מנורמל להשוואה: בלי רווחים כפולים, אותיות קטנות, יוניקוד אחיד."""
    text = unicodedata.normalize("NFC", str(text)).strip().casefold()
    return re.sub(r"\s+", " ", text)


def index_downloads(folder: Path) -> dict:
    """ממפה שם קובץ מנורמל → הנתיב בפועל, כולל תיקיות משנה."""
    found = {}
    for path in folder.rglob("*"):
        if path.is_file():
            found.setdefault(norm(path.name), path)
            found.setdefault(norm(path.stem), path)   # גם בלי סיומת, ליתר ביטחון
    return found


def read_rows(excel: Path):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def parse_items(rows):
    """הייצוא בנוי משורת כותרות שחוזרת לכל קבוצה, ומעליה שם התחום."""
    header, items, current_group = None, [], ""
    for row in rows:
        cells = [("" if c is None else str(c).strip()) for c in row]
        if not any(cells):
            continue
        if "Name" in cells:
            header = cells
            continue
        if header is None:
            continue
        if cells[0] and not any(cells[1:]):
            current_group = cells[0]          # שורת כותרת של תחום
            continue

        def get(title):
            return cells[header.index(title)] if title in header and header.index(title) < len(cells) else ""

        name = get("Name")
        if not name:
            continue
        item = {field: "" for field in ITEM_FIELDS}
        for title, field in COLUMNS.items():
            item[field] = get(title)
        for title in DESC_COLUMNS:
            if title in header and get(title):
                item["description"] = get(title)
                break
        raw_category = (item["category"] or current_group).strip()
        item["category"] = CATEGORY_ALIASES.get(raw_category, raw_category)
        item["type"] = item["type"] or "מסמך המלצות"
        item["_links"] = [u.strip() for u in get(FILE_COLUMN).split(",") if u.strip().startswith("http")]
        items.append(item)
    return items


def attach_files(items, downloads: dict, copy: bool):
    """משייך לכל פריט את הקבצים שלו, ומעתיק לתיקיית האתר את מה שנמצא."""
    matched, missing = 0, []
    for item in items:
        files = []
        for url in item.pop("_links"):
            filename = urllib.parse.unquote(url.rsplit("/", 1)[-1]).strip()
            if not filename:
                continue
            local = downloads.get(norm(filename)) or downloads.get(norm(Path(filename).stem))
            entry = {"url": url, "name": filename, "localPath": None}
            if local:
                target = FILES_DIR / filename
                if copy:
                    FILES_DIR.mkdir(exist_ok=True)
                    if not target.exists() or target.stat().st_size != local.stat().st_size:
                        shutil.copy2(local, target)
                entry["localPath"] = f"files/{filename}"
                matched += 1
            else:
                missing.append(filename)
            files.append(entry)
        item["files"] = files
    return matched, missing


def write_html(data: dict):
    html = HTML_FILE.read_text(encoding="utf-8")
    payload = "const DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";"
    updated = re.sub(r"const DATA = \{.*?\};", lambda _: payload, html, count=1, flags=re.DOTALL)
    if updated == html:
        sys.exit("לא נמצא const DATA = {...} בתוך index.html")
    BACKUP_DIR.mkdir(exist_ok=True)
    backup = BACKUP_DIR / f"index-{datetime.now():%Y%m%d-%H%M%S}.html"
    backup.write_text(html, encoding="utf-8")
    HTML_FILE.write_text(updated, encoding="utf-8")
    return backup


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="קובץ הייצוא מ-Monday")
    ap.add_argument("--files-dir", help="התיקייה שאליה הורדת את הקבצים")
    ap.add_argument("--dry-run", action="store_true", help="רק להראות מה יקרה, בלי לשנות דבר")
    args = ap.parse_args()

    excel = Path(args.excel)
    if not excel.exists():
        sys.exit(f"לא נמצא הקובץ: {excel}")

    items = parse_items(read_rows(excel))
    if not items:
        sys.exit("לא נמצאו פריטים בייצוא. ודא שזה הקובץ הנכון.")

    downloads = {}
    if args.files_dir:
        folder = Path(args.files_dir)
        if not folder.exists():
            sys.exit(f"לא נמצאה תיקיית הקבצים: {folder}")
        downloads = index_downloads(folder)

    matched, missing = attach_files(items, downloads, copy=not args.dry_run)

    cats_in_data = {i["category"] for i in items if i["category"]}
    categories = [c for c in CATEGORY_ORDER if c in cats_in_data]
    categories += sorted(c for c in cats_in_data if c not in categories)
    data = {"categories": categories, "items": items}

    print(f"\nפריטים: {len(items)}   תחומים: {len(categories)}")
    for c in categories:
        print(f"   {c}: {sum(1 for i in items if i['category'] == c)}")
    unknown = [c for c in categories if c not in CATEGORY_ORDER]
    if unknown:
        print()
        print("תחומים שאינם מוכרים לאתר (יוצגו בלי צבע משלהם):")
        for c in unknown:
            print("   * " + c)
        print("   אפשר לתקן את השם בלוח, או להוסיף אותו ל-CATEGORY_ALIASES בסקריפט.")

    print(f"\nקבצים שקושרו: {matched}")
    if missing:
        print(f"קבצים שלא נמצאו בתיקייה: {len(missing)}")
        for name in missing[:15]:
            print("   ▸ " + name)
        if len(missing) > 15:
            print(f"   ... ועוד {len(missing) - 15}")
        print("   (הפריטים האלה עדיין יעבדו, אבל הקובץ לא יוצג באתר)")

    if args.dry_run:
        print("\n— בדיקה בלבד, לא שונה דבר. להרצה אמיתית: להסיר --dry-run —\n")
        return

    backup = write_html(data)
    print(f"\nהאתר עודכן. גיבוי המצב הקודם: {backup.name}")
    print("לפרסום:  git add -A  &&  git commit -m \"Import from Monday export\"  &&  git push\n")


if __name__ == "__main__":
    main()
