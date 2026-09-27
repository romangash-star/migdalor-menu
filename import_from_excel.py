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

import argparse, json, re, shutil, sys, unicodedata, urllib.error, urllib.parse, urllib.request
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
    "מחלקה רלבנטית":              "dept",
    "תיאור קצר":                  "description",
    "שם איש קשר להטמעה":          "contactName",
    "טלפון ליצירת קשר":           "phone",
    "אי מייל ליצירת קשר":         "email",
    "קהל יעד":                    "audience",
    "נושא תוכן מרכזי":            "topic",
    'האם הפריט קיים במאגר גפ"ן':  "gafan",
    "סוג פריט":                   "type",
}
FILE_COLUMN = "קובץ פירוט"

# עמודות שמכילות קישור לקובץ מעוצב (למשל "לינק לפידיאף מעוצב").
# הזיהוי גמיש, כדי שלא ייפול על שינוי קטן בשם העמודה.
def is_link_column(title: str) -> bool:
    t = title.strip().casefold()
    if title == FILE_COLUMN:
        return False
    return "מעוצב" in t or ("לינק" in t and "pdf" in t) or ("קישור" in t and "pdf" in t)
DESC_COLUMNS = ("תיאור", "פירוט", "תוכן")      # אופציונלי, אם קיים בייצוא

# בייצוא מופיעים לפעמים שמות תחום בוריאציות שונות. כאן הם מאוחדים לשם אחד,
# כדי שלא ייווצרו באתר תחומים כפולים בלי צבע.
CATEGORY_ALIASES = {
    "חזון ותפיסת יעוד":        "חזון ויעוד",
    "זהות ויעוד":              "חזון ויעוד",
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
    in_subitems = False
    for row in rows:
        cells = [("" if c is None else str(c).strip()) for c in row]
        if not any(cells):
            continue
        if "Name" in cells:
            # לתת-פריטים יש שורת כותרות משלהם, בלי עמודת התחום. מזהים לפי כך
            # שהיא חסרה, וכל מה שאחריה מדולג.
            if FILE_COLUMN in cells or "תחום ראשי בתפריט" in cells:
                header, in_subitems = cells, False
            else:
                in_subitems = True
            continue
        if header is None:
            continue
        if cells[0] == "Subitems":
            # Monday מייצא תת-פריטים עם עמודות משלהם. הם שייכים לפריט שמעליהם
            # ולא מיובאים כפריטים עצמאיים.
            in_subitems = True
            continue
        if cells[0] and not any(cells[1:]):
            current_group, in_subitems = cells[0], False   # שורת כותרת של תחום
            continue
        # שורת תת-פריט אין בה תחום; ברגע שחוזרת שורה עם תחום, זה שוב פריט רגיל.
        cat_idx = header.index("תחום ראשי בתפריט") if "תחום ראשי בתפריט" in header else -1
        has_category = 0 <= cat_idx < len(cells) and bool(cells[cat_idx])
        if in_subitems:
            if not has_category:
                continue
            in_subitems = False

        def get(title):
            return cells[header.index(title)] if title in header and header.index(title) < len(cells) else ""

        name = get("Name")
        if not name or name == "Subitems":      # שורת שירות של Monday, לא פריט
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
        # קישורי Canva הם עמודי עיצוב ולא קבצים, ולכן לא נלקחים לאתר.
        item["_pdf_links"] = [u.strip() for title in header if is_link_column(title)
                              for u in re.split(r"[,\s]+", get(title))
                              if u.strip().startswith("http") and "canva." not in u.lower()]
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


def direct_url(url: str) -> str:
    """קישור שיתוף של Google Drive מצביע על עמוד תצוגה. כאן הוא מומר לקישור
    שמחזיר את הקובץ עצמו, כדי שאפשר יהיה להוריד אותו."""
    m = re.search(r"drive\.google\.com/file/d/([\w-]+)", url) or \
        re.search(r"drive\.google\.com/open\?id=([\w-]+)", url) or \
        re.search(r"drive\.google\.com/uc\?[^\"]*id=([\w-]+)", url)
    if m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    m = re.search(r"docs\.google\.com/document/d/([\w-]+)", url)
    if m:
        return f"https://docs.google.com/document/d/{m.group(1)}/export?format=pdf"
    return url


def safe_name(text: str) -> str:
    return re.sub(r'[\\/?%*:|"<>]', "-", text).strip()[:120] or "file"


def download_links(items, save: bool, timeout: int = 60):
    """מוריד את הקבצים מעמודת הקישור המעוצב. קישורים שדורשים התחברות לא יירדו,
    והם יידווחו בסוף כדי שאפשר יהיה להוריד אותם ידנית."""
    done, blocked = 0, []
    for item in items:
        for url in item.pop("_pdf_links", []):
            if any(f.get("_source") == url for f in item.get("files", [])):
                continue
            target_name = safe_name(item["name"]) + " - מעוצב.pdf"
            target = FILES_DIR / target_name
            if not save:
                blocked.append((item["name"], url, "בדיקה בלבד"))
                continue
            if target.exists() and target.stat().st_size > 1000:
                item["files"].append({"url": url, "name": target_name,
                                      "localPath": f"files/{target_name}", "_source": url})
                done += 1
                continue
            try:
                req = urllib.request.Request(direct_url(url), headers={
                    "User-Agent": "Mozilla/5.0", "Accept": "application/pdf,*/*"})
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    ctype = (resp.headers.get("Content-Type") or "").lower()
                    disp = resp.headers.get("Content-Disposition") or ""
                    body = resp.read()
                if "pdf" not in ctype and not body[:5].startswith(b"%PDF"):
                    # כנראה דף התחברות או דף עיצוב, ולא הקובץ עצמו
                    blocked.append((item["name"], url, "הקישור לא מחזיר קובץ PDF"))
                    item["files"].append({"url": url, "name": target_name, "localPath": None, "_source": url})
                    continue
                m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)', disp)
                if m:
                    target_name = safe_name(urllib.parse.unquote(m.group(1)))
                    if not target_name.lower().endswith(".pdf"):
                        target_name += ".pdf"
                    target = FILES_DIR / target_name
                FILES_DIR.mkdir(exist_ok=True)
                target.write_bytes(body)
                item["files"].append({"url": url, "name": target_name,
                                      "localPath": f"files/{target_name}", "_source": url})
                done += 1
            except (urllib.error.URLError, urllib.error.HTTPError, OSError) as err:
                blocked.append((item["name"], url, str(err)[:60]))
                item["files"].append({"url": url, "name": target_name, "localPath": None, "_source": url})
    for item in items:
        for f in item.get("files", []):
            f.pop("_source", None)
    return done, blocked


def drop_unavailable(items):
    """מסיר קבצים שאין להם עותק מקומי והקישור אליהם מוגן ודורש התחברות ל-Monday,
    כדי שלא יופיעו באתר קבצים שאי אפשר לפתוח."""
    removed = []
    for item in items:
        kept = []
        for f in item.get("files", []):
            protected = "monday.com" in (f.get("url") or "").lower()
            if not f.get("localPath") and protected:
                removed.append((item["name"], f.get("name", "")))
            else:
                kept.append(f)
        item["files"] = kept
    return removed


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
    ap.add_argument("--skip-links", action="store_true", help="לדלג על הורדה מעמודת הקישור המעוצב")
    ap.add_argument("--keep-dead-links", action="store_true",
                    help="להשאיר באתר גם קבצים שאין להם עותק מקומי (קישור מוגן שלא ייפתח)")
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
    downloaded, blocked = download_links(items, save=not args.dry_run and not args.skip_links)
    dropped = [] if args.keep_dead_links else drop_unavailable(items)

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

    if downloaded or blocked:
        print()
        print(f"קבצים שהורדו מעמודת הקישור המעוצב: {downloaded}")
        if blocked:
            print(f"קישורים שלא ניתן היה להוריד אוטומטית: {len(blocked)}")
            for name, url, why in blocked[:10]:
                print(f"   * {name[:40]} — {why}")
                print(f"     {url}")
            if len(blocked) > 10:
                print(f"   ... ועוד {len(blocked) - 10}")
            print("   (אפשר להוריד אותם ידנית לתיקיית הקבצים ולהריץ שוב)")

    if dropped:
        print()
        print(f"קבצים שהושמטו מהאתר (אין עותק מקומי והקישור מוגן): {len(dropped)}")
        for name, fname in dropped[:10]:
            print(f"   * {fname[:60]}  ({name[:30]})")
        if len(dropped) > 10:
            print(f"   ... ועוד {len(dropped) - 10}")
        print("   (להוריד אותם לתיקיית הקבצים ולהריץ שוב, או --keep-dead-links כדי להשאיר בכל זאת)")

    if args.dry_run:
        print("\n— בדיקה בלבד, לא שונה דבר. להרצה אמיתית: להסיר --dry-run —\n")
        return

    backup = write_html(data)
    print(f"\nהאתר עודכן. גיבוי המצב הקודם: {backup.name}")
    print("לפרסום:  git add -A  &&  git commit -m \"Import from Monday export\"  &&  git push\n")


if __name__ == "__main__":
    main()
