#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_download_helper.py — בונה עמוד עזר מקומי להורדה ידנית של הקבצים החסרים.

שימוש:
  python make_download_helper.py --excel "export_1790497669.xlsx"

נוצר קובץ download-helper.html. פותחים אותו בדפדפן, והוא מרכז את כל הקבצים
שחסרים לאתר: קישור לפתיחה, השם המדויק לשמירה, וסימון של מה שכבר ירד.
"""

import argparse, html, importlib.util, json
from pathlib import Path

ROOT = Path(__file__).parent
OUT  = ROOT / "download-helper.html"

imp_spec = importlib.util.spec_from_file_location("imp", ROOT / "import_from_excel.py")
imp = importlib.util.module_from_spec(imp_spec)
imp_spec.loader.exec_module(imp)


def collect(excel: Path):
    items = imp.parse_items(imp.read_rows(excel))
    have = imp.index_downloads(imp.FILES_DIR) if imp.FILES_DIR.exists() else {}
    rows = []
    for item in items:
        needs = []
        for url in item["_links"]:                       # מצורפים מ-Monday
            import urllib.parse
            name = urllib.parse.unquote(url.rsplit("/", 1)[-1]).strip()
            if not name:
                continue
            done = imp.norm(name) in have or imp.norm(Path(name).stem) in have
            needs.append({"url": url, "save_as": name, "kind": "מצורף מ-Monday", "done": done})
        for url in item["_pdf_links"]:                   # PDF מעוצב מ-Drive
            save_as = imp.safe_name(item["name"]) + ".pdf"
            done = imp.find_designed_pdf(item["name"], have) is not None
            needs.append({"url": url, "save_as": save_as, "kind": "PDF מעוצב", "done": done})
        if needs:
            rows.append({"item": item["name"], "category": item["category"], "needs": needs})
    return rows


PAGE = """<!DOCTYPE html>
<html lang="he" dir="rtl"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>הורדת הקבצים לאתר</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Heebo',Arial,sans-serif;background:#f0f2de;color:#1a3340;line-height:1.6;font-size:16px}
header{background:linear-gradient(135deg,#00586f,#1c7d8f);color:#fff;padding:22px 24px}
h1{font-size:1.5rem;font-weight:700;letter-spacing:.04em}
.wrap{max-width:1000px;margin:0 auto;padding:22px 24px 60px}
.box{background:#fff;border-radius:12px;padding:18px 20px;margin-bottom:16px;box-shadow:0 4px 18px rgba(0,88,111,.1)}
.box b{color:#00586f}
code{background:#eef2f2;padding:2px 7px;border-radius:5px;font-size:.88rem;direction:ltr;display:inline-block}
.bar{height:12px;background:#dde5e0;border-radius:20px;overflow:hidden;margin:10px 0 6px}
.bar span{display:block;height:100%;background:#38d1d6;width:0;transition:width .3s}
.item{background:#fff;border-radius:12px;margin-bottom:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,88,111,.08)}
.item h2{font-size:1.02rem;padding:12px 18px;background:#00586f;color:#fff;font-weight:700}
.file{display:flex;align-items:center;gap:12px;padding:12px 18px;border-bottom:1px solid #eef2f2;flex-wrap:wrap}
.file:last-child{border-bottom:none}
.file.done{background:#f3faf4}
.file input[type=checkbox]{width:20px;height:20px;flex-shrink:0;cursor:pointer}
.info{flex:1;min-width:240px}
.kind{font-size:.72rem;font-weight:700;color:#4A556C}
.name{font-weight:700;word-break:break-all}
.actions{display:flex;gap:8px}
a.btn,button.btn{font:inherit;font-size:.85rem;font-weight:700;border:none;border-radius:8px;padding:8px 14px;
 cursor:pointer;text-decoration:none;display:inline-block}
a.btn{background:#00586f;color:#fff}
button.btn{background:#eef2f2;color:#00586f}
.hide-done .file.done{display:none}
.toolbar{display:flex;gap:14px;align-items:center;margin-bottom:14px;flex-wrap:wrap}
.count{font-weight:700}
</style></head><body>
<header><h1>הורדת הקבצים לאתר</h1></header>
<div class="wrap">
  <div class="box">
    <b>איך עובדים:</b> פותחים כל קובץ בלחיצה על "פתיחה", מורידים אותו,
    ושומרים בדיוק בשם שמופיע כאן, לתוך התיקייה:
    <div style="margin:8px 0"><code>__FILES_DIR__</code></div>
    אפשר להעתיק את השם המדויק בלחיצה על "העתקת השם". מה שכבר נמצא בתיקייה מסומן אוטומטית.
    בסיום מריצים:
    <div style="margin:8px 0"><code>python import_from_excel.py --excel "__EXCEL__" --files-dir files</code></div>
  </div>
  <div class="box">
    <div class="count" id="count"></div>
    <div class="bar"><span id="bar"></span></div>
    <div class="toolbar">
      <label><input type="checkbox" id="hideDone"> להסתיר את מה שכבר ירד</label>
      <button class="btn" id="resetBtn">איפוס הסימונים</button>
    </div>
  </div>
  <div id="list"></div>
</div>
<script>
const ROWS = __ROWS__;
const KEY = 'migdalor_downloads';
const marked = new Set(JSON.parse(localStorage.getItem(KEY) || '[]'));
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

function render(){
  document.getElementById('list').innerHTML = ROWS.map(row => `
    <div class="item">
      <h2>${esc(row.item)}</h2>
      ${row.needs.map(f => {
        const done = f.done || marked.has(f.url);
        return `<div class="file${done ? ' done' : ''}">
          <input type="checkbox" data-url="${esc(f.url)}" ${done ? 'checked' : ''} ${f.done ? 'disabled' : ''}>
          <div class="info">
            <div class="kind">${esc(f.kind)}${f.done ? ' — כבר קיים בתיקייה' : ''}</div>
            <div class="name">${esc(f.save_as)}</div>
          </div>
          <div class="actions">
            <a class="btn" href="${esc(f.url)}" target="_blank" rel="noopener">פתיחה</a>
            <button class="btn copy" data-name="${esc(f.save_as)}">העתקת השם</button>
          </div>
        </div>`;
      }).join('')}
    </div>`).join('');

  document.querySelectorAll('.file input[type=checkbox]').forEach(cb => cb.onchange = () => {
    cb.checked ? marked.add(cb.dataset.url) : marked.delete(cb.dataset.url);
    localStorage.setItem(KEY, JSON.stringify([...marked]));
    render();
  });
  document.querySelectorAll('button.copy').forEach(b => b.onclick = async () => {
    await navigator.clipboard.writeText(b.dataset.name);
    b.textContent = 'הועתק';
    setTimeout(() => b.textContent = 'העתקת השם', 1200);
  });

  const all = ROWS.flatMap(r => r.needs);
  const done = all.filter(f => f.done || marked.has(f.url)).length;
  document.getElementById('count').textContent = `${done} מתוך ${all.length} קבצים ירדו`;
  document.getElementById('bar').style.width = (all.length ? done / all.length * 100 : 0) + '%';
}
document.getElementById('hideDone').onchange = e =>
  document.body.classList.toggle('hide-done', e.target.checked);
document.getElementById('resetBtn').onclick = () => {
  marked.clear(); localStorage.removeItem(KEY); render();
};
render();
</script></body></html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    args = ap.parse_args()
    excel = Path(args.excel)
    if not excel.exists():
        raise SystemExit(f"לא נמצא הקובץ: {excel}")

    rows = collect(excel)
    page = (PAGE.replace("__ROWS__", json.dumps(rows, ensure_ascii=False))
                .replace("__FILES_DIR__", html.escape(str(imp.FILES_DIR)))
                .replace("__EXCEL__", html.escape(excel.name)))
    OUT.write_text(page, encoding="utf-8")

    total = sum(len(r["needs"]) for r in rows)
    have = sum(1 for r in rows for f in r["needs"] if f["done"])
    print(f"נוצר {OUT.name}: {total} קבצים, מתוכם {have} כבר קיימים בתיקייה.")
    print(f"לפתיחה: {OUT}")


if __name__ == "__main__":
    main()
