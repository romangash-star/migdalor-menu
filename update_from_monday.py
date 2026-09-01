#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_from_monday.py  —  מעדכן את DATA ב-index.html מלוח Monday.com

שימוש:
  python update_from_monday.py --discover   # הצג את כל עמודות הלוח + ה-IDs שלהן
  python update_from_monday.py              # עדכן את index.html מ-Monday
"""

import json, re, sys, urllib.request, urllib.error
from pathlib import Path

# ════════════════════════════════════════════════════════════════════
#  הגדרות — ערוך כאן לפני הרצה ראשונה
# ════════════════════════════════════════════════════════════════════

API_TOKEN = "YOUR_API_TOKEN_HERE"
#   ↑ Monday → לחץ על תמונת הפרופיל → Developers → My Access Tokens → Copy

BOARD_ID  = "YOUR_BOARD_ID_HERE"
#   ↑ פתח את הלוח ב-Monday, תראה בURL: monday.com/boards/1234567890

# מיפוי: שם שדה ב-DATA  →  מזהה עמודה ב-Monday (column id)
# הרץ תחילה:  python update_from_monday.py --discover
# כדי לראות את כל ה-IDs. לאחר מכן עדכן את הרשימה כאן:
COLUMN_MAP = {
    "category":    "dropdown",   # קטגוריה  (dropdown / status)
    "type":        "status",     # סוג הפריט: ספק / מסמך המלצות  (status)
    "contactName": "text",       # שם איש/אשת קשר  (text)
    "phone":       "phone",      # טלפון  (phone)
    "email":       "email",      # אימייל  (email)
    "dept":        "text7",      # מחלקה / מגזר  (text)
    "audience":    "text4",      # קהל יעד  (text / dropdown)
    "topic":       "text5",      # נושא  (text / dropdown)
    "gafan":       "status4",    # גפ"ן  (status)
    "description": "long_text",  # תיאור / פירוט ארוך  (long_text) — הרץ --discover לאימות ה-ID
    "ready":       "status9",    # סטטוס מוכן להצגה  (status) — הרץ --discover לאימות ה-ID
}

# סדר הקטגוריות בתפריט
CATEGORY_ORDER = [
    "חזון ויעוד",
    "כח אדם",
    "תכניות חינוכיות וקהילתיות",
    "תרבות ארגונית",
    "שותפויות ומשאבים",
]

HTML_FILE = Path(__file__).parent / "index.html"

# ════════════════════════════════════════════════════════════════════
#  לוגיקה — אין צורך לשנות מכאן
# ════════════════════════════════════════════════════════════════════

ENDPOINT = "https://api.monday.com/v2"


def gql(query: str) -> dict:
    payload = json.dumps({"query": query}).encode("utf-8")
    req = urllib.request.Request(
        ENDPOINT,
        data=payload,
        headers={
            "Authorization": API_TOKEN,
            "Content-Type": "application/json",
            "API-Version": "2024-01",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read())
            if "errors" in result:
                for err in result["errors"]:
                    print(f"  שגיאת GraphQL: {err.get('message','')}")
                sys.exit(1)
            return result
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        print(f"שגיאת HTTP {e.code}: {body[:400]}")
        sys.exit(1)


def discover():
    print(f"\nשולח שאילתת גילוי ללוח {BOARD_ID}...")
    res = gql(f"{{ boards(ids:[{BOARD_ID}]) {{ name columns {{ id title type }} }} }}")
    boards = res.get("data", {}).get("boards", [])
    if not boards:
        print("לא נמצא לוח. בדוק את BOARD_ID ו-API_TOKEN.")
        sys.exit(1)
    board = boards[0]
    print(f"\nלוח: {board['name']}")
    print("─" * 55)
    print(f"{'ID':<25} {'סוג':<20} {'כותרת'}")
    print("─" * 55)
    for col in board["columns"]:
        print(f"  {col['id']:<23} {col['type']:<20} {col['title']}")
    print("─" * 55)
    print("\nהעתק את ה-IDs הרלוונטיים ל-COLUMN_MAP בתחילת הסקריפט, ואז הרץ שוב ללא --discover.\n")


def fetch_asset_urls(asset_ids: list) -> dict:
    """מחזיר {asset_id: public_url} לרשימת asset IDs"""
    if not asset_ids:
        return {}
    ids_str = ",".join(str(i) for i in asset_ids)
    res = gql(f"{{ assets(ids:[{ids_str}]) {{ id public_url name }} }}")
    out = {}
    for a in res.get("data", {}).get("assets", []):
        out[str(a["id"])] = {"url": a.get("public_url", ""), "name": a.get("name", "")}
    return out


def fetch_items() -> list:
    """משך את כל הפריטים מהלוח בדפים של 100"""
    items = []
    cursor = None
    page_num = 0

    while True:
        page_num += 1
        cursor_arg = f', cursor: "{cursor}"' if cursor else ""
        query = f"""
        {{
          boards(ids:[{BOARD_ID}]) {{
            items_page(limit:100{cursor_arg}) {{
              cursor
              items {{
                name
                column_values {{
                  id
                  text
                  value
                }}
              }}
            }}
          }}
        }}"""
        res   = gql(query)
        page  = res["data"]["boards"][0]["items_page"]
        batch = page.get("items", [])
        items += batch
        print(f"  עמוד {page_num}: {len(batch)} פריטים (סה\"כ {len(items)})")
        cursor = page.get("cursor")
        if not cursor or not batch:
            break

    return items


def collect_asset_ids(raw_items: list) -> list:
    """מאסף את כל ה-asset IDs מכל עמודות הקבצים"""
    ids = []
    for item in raw_items:
        for cv in item.get("column_values", []):
            val = cv.get("value")
            if not val or val == "null":
                continue
            try:
                parsed = json.loads(val)
            except (json.JSONDecodeError, TypeError):
                continue
            # עמודת file: {"files": [{"assetId": ..., "name": ...}, ...]}
            for f in parsed.get("files", []):
                aid = f.get("assetId") or f.get("asset_id")
                if aid:
                    ids.append(str(aid))
    return list(set(ids))


def parse_item(raw: dict, asset_map: dict) -> dict:
    """הופך פריט Monday גולמי לאובייקט DATA"""
    cols = {cv["id"]: cv for cv in raw.get("column_values", [])}

    def text(field_key: str) -> str:
        col_id = COLUMN_MAP.get(field_key, "")
        cv     = cols.get(col_id, {})
        return (cv.get("text") or "").strip()

    # קבצים — מכל עמודות file בפריט
    files = []
    for cv in raw.get("column_values", []):
        val = cv.get("value")
        if not val or val == "null":
            continue
        try:
            parsed = json.loads(val)
        except (json.JSONDecodeError, TypeError):
            continue
        for f in parsed.get("files", []):
            aid = str(f.get("assetId") or f.get("asset_id") or "")
            if aid and aid in asset_map:
                info = asset_map[aid]
                files.append({
                    "url":       info["url"],
                    "name":      info["name"] or f.get("name", ""),
                    "localPath": None,
                })

    return {
        "name":        raw["name"].strip(),
        "category":    text("category"),
        "dept":        text("dept"),
        "contactName": text("contactName"),
        "phone":       text("phone"),
        "email":       text("email"),
        "audience":    text("audience"),
        "topic":       text("topic"),
        "gafan":       text("gafan"),
        "type":        text("type") or "מסמך המלצות",
        "description": text("description"),
        "files":       files,
        "_ready":      text("ready"),
    }


def update_html(data: dict):
    if not HTML_FILE.exists():
        print(f"לא נמצא קובץ: {HTML_FILE}")
        sys.exit(1)
    html     = HTML_FILE.read_text(encoding="utf-8")
    new_data = "const DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";"
    updated  = re.sub(r"const DATA = \{.*?\};", new_data, html, count=1, flags=re.DOTALL)
    if updated == html:
        print("לא נמצא const DATA = {...} ב-index.html — בדוק שהקובץ תקין.")
        sys.exit(1)
    HTML_FILE.write_text(updated, encoding="utf-8")


def main():
    if "--discover" in sys.argv:
        discover()
        return

    if API_TOKEN == "YOUR_API_TOKEN_HERE" or BOARD_ID == "YOUR_BOARD_ID_HERE":
        print("שגיאה: הכנס API_TOKEN ו-BOARD_ID בתחילת הסקריפט.")
        sys.exit(1)

    print("═" * 50)
    print("  עדכון index.html מ-Monday.com")
    print("═" * 50)

    print("\n1. מושך פריטים מ-Monday...")
    raw_items = fetch_items()
    print(f"   סה\"כ: {len(raw_items)} פריטים")

    print("\n2. מושך URL-ים של קבצים...")
    asset_ids = collect_asset_ids(raw_items)
    print(f"   נמצאו {len(asset_ids)} קבצים")
    asset_map = fetch_asset_urls(asset_ids)
    print(f"   הומרו ל-URL: {len(asset_map)}")

    print("\n3. ממיר פריטים...")
    items = [parse_item(i, asset_map) for i in raw_items]
    items = [i for i in items if i["name"] and i["category"]]
    # ← הכנס כאן את הערך בעמודת הסטטוס שמסמן "מוכן להצגה", למשל: "מוכן"
    READY_STATUS = ""
    if READY_STATUS:
        before = len(items)
        items = [i for i in items if i.get("_ready","") == READY_STATUS]
        print(f"   סינון לפי סטטוס '{READY_STATUS}': {before} → {len(items)}")
    for i in items:
        i.pop("_ready", None)
    print(f"   {len(items)} פריטים תקינים")

    cats_in_data = {i["category"] for i in items}
    categories   = [c for c in CATEGORY_ORDER if c in cats_in_data]
    for c in sorted(cats_in_data):
        if c not in categories:
            categories.append(c)

    data = {"categories": categories, "items": items}

    print("\n4. מעדכן index.html...")
    update_html(data)
    print(f"   ✓ {len(items)} פריטים ב-{len(categories)} קטגוריות")

    print("\n═" * 50)
    print("הרץ כדי לפרסם לאתר:")
    print('  git add index.html')
    print('  git commit -m "Update data from Monday"')
    print('  git push origin main')
    print("═" * 50 + "\n")


if __name__ == "__main__":
    main()
