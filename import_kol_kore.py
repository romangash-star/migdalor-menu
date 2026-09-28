# -*- coding: utf-8 -*-
"""ממיר את ייצוא לוח "קול קורא לתפריט רשויות" לפריטים במבנה של האתר."""
import json, io, re, sys

CATEGORY = {
    'תכניות חינוכיות וקהילתיות': 'תכניות חינוכיות וקהילתיות',
    'חזון ותפיסת יעוד':          'חזון ויעוד',
    'כח אדם':                    'כח אדם',
    'תרבות ארגונים':             'תרבות ארגונית',
    'שותפויות ומשאבים':          'שותפויות ומשאבים',
}
TOPIC = {'יחסים בין בקבוצות': 'יחסים בין קבוצות'}

def phone(v):
    """972501234567 -> 050-123-4567, כדי שגם חיוג וגם וואטסאפ יעבדו."""
    d = re.sub(r'\D', '', v or '')
    if d.startswith('972'):
        d = '0' + d[3:]
    if len(d) == 10 and d.startswith('0'):
        return f'{d[:3]}-{d[3:6]}-{d[6:]}'
    if len(d) == 9 and d.startswith('0'):
        return f'{d[:2]}-{d[2:5]}-{d[5:]}'
    return v.strip()

def doc_urls(v):
    return [u.strip().rstrip(',;') for u in re.split(r'(?=https?://)', v or '') if u.strip().startswith('http')]

def convert(rows):
    items, skipped = [], []
    for r in rows:
        name = (r['Name'] or '').strip()
        cat  = CATEGORY.get((r['תחום בתפריט'] or '').strip())
        if not name or not cat:
            skipped.append({'name': name, 'why': 'תחום לא מזוהה: ' + (r['תחום בתפריט'] or '(ריק)')})
            continue
        items.append({
            'name': name,
            'category': cat,
            'dept': (r['מחלקה רלבנטית'] or '').strip(),
            'org': (r['שם הארגון/הספק המבצע'] or '').strip(),
            'contactName': (r['שם איש הקשר'] or '').strip(),
            'phone': phone(r['טלפון ליצירת קשר']),
            'email': (r['מייל ליצירת קשר'] or '').strip(),
            'audience': ', '.join(a.strip() for a in re.split(r'\s*,\s*', r['קהל יעד'] or '') if a.strip()),
            'topic': TOPIC.get((r['נושא מרכזי של הפריט המוצע'] or '').strip(), (r['נושא מרכזי של הפריט המוצע'] or '').strip()),
            'gafan': (r['האם הפריט קיים במאגר גפ"ן'] or '').strip(),
            'type': (r['סוג פריט'] or 'ספק').strip(),
            'description': re.sub(r'\s+\n', '\n', (r['תיאור ההצעה'] or '').strip()),
            'notes': (r['דגשים/הערות'] or '').strip(),
            'website': (r['לינק לאתר הספק/הפעילות'] or '').strip(),
            'files': [],
            '_docs': doc_urls(r['מסמכים רלוונטיים']),
        })
    return items, skipped

if __name__ == '__main__':
    d = json.load(io.open('kol_kore.json', encoding='utf-8'))
    items, skipped = convert(d['rows'][:70])
    io.open('new_items.json', 'w', encoding='utf-8').write(json.dumps(items, ensure_ascii=False, indent=1))
    sys.stdout.reconfigure(encoding='utf-8')
    print('converted:', len(items), '| skipped:', len(skipped))
    for s in skipped:
        print('   -', s['name'][:40], '|', s['why'])
