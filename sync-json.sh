#!/usr/bin/env bash
# ═══════════════════════════════════════════════════════════════════
#  استبيان يعمر — سحب الاستجابات وحفظها ملفات JSON بجانب index.html
#
#  الاستخدام:
#     ./sync-json.sh                  سحب من الخادم مباشرة
#     ./sync-json.sh --from x.json    من ملف تصدير نزّلته يدوياً
#
#  عدّل القيمتين بالأسفل مرة واحدة. الملفات تُكتب في مجلد yamur-backup
#  داخل نفس مجلد index.html.
# ═══════════════════════════════════════════════════════════════════
set -euo pipefail

SCRIPT_URL='https://script.google.com/macros/s/AKfycbzHsr2GIvhTw6AiFsfWR-9Q1Y-nzs5QHzZRGsDRmWbBxbKf_GAsPo3grpBONGg2g1dA9w/exec'
ADMIN_KEY='CHANGE-yaamur-2030'      # نفس ADMIN_KEY في apps-script.gs

HERE="$(cd "$(dirname "$0")" && pwd)"
OUT_DIR="$HERE/yamur-backup"
SRC=''

while [ $# -gt 0 ]; do
  case "$1" in
    --from) SRC="${2:-}"; shift 2 ;;
    --out)  OUT_DIR="${2:-}"; shift 2 ;;
    -h|--help) sed -n '2,12p' "$0"; exit 0 ;;
    *) echo "وسيط غير معروف: $1" >&2; exit 2 ;;
  esac
done

TMP="$(mktemp -t yamur-export)"
trap 'rm -f "$TMP"' EXIT

if [ -n "$SRC" ]; then
  [ -f "$SRC" ] || { echo "✗ لا يوجد ملف: $SRC" >&2; exit 1; }
  cp "$SRC" "$TMP"
  echo "المصدر: $SRC"
else
  case "$ADMIN_KEY" in
    ''|CHANGE-ME-*) echo "✗ عدّل ADMIN_KEY في أعلى هذا الملف أولاً — ويجب أن يطابق apps-script.gs." >&2; exit 1 ;;
  esac
  case "$SCRIPT_URL" in
    *"/exec") : ;;
    *) echo "✗ SCRIPT_URL يجب أن ينتهي بـ /exec" >&2; exit 1 ;;
  esac
  echo "جارٍ السحب من الخادم…"
  curl -fsSL --max-time 90 --get "$SCRIPT_URL" \
    --data-urlencode 'action=export' \
    --data-urlencode "key=$ADMIN_KEY" -o "$TMP" \
    || { echo "✗ تعذّر الوصول إلى الخادم. تحقّق من الاتصال ومن رابط النشر." >&2; exit 1; }
fi

OUT_DIR="$OUT_DIR" python3 - "$TMP" <<'PY'
# -*- coding: utf-8 -*-
import json, os, sys, io

out = os.environ['OUT_DIR']
try:
    data = json.load(io.open(sys.argv[1], encoding='utf-8'))
except Exception as e:
    sys.exit('✗ رد غير مفهوم من الخادم (ليس JSON): %s' % e)

if not isinstance(data, dict) or data.get('ok') is not True:
    sys.exit('✗ رفض الخادم الطلب: %s' % (data.get('error', 'سبب غير معروف') if isinstance(data, dict) else data))

rows = [r for r in (data.get('responses') or []) if isinstance(r, dict)]
os.makedirs(out, exist_ok=True)

def write(path, text):
    """يكتب فقط عند الاختلاف — يعيد new / updated / same"""
    if os.path.exists(path):
        if io.open(path, encoding='utf-8').read() == text: return 'same'
        state = 'updated'
    else:
        state = 'new'
    io.open(path, 'w', encoding='utf-8').write(text)
    return state

tally = {'new': 0, 'updated': 0, 'same': 0}
skipped = []
for r in rows:
    ref = str(r.get('ref') or '').strip()
    if not ref or '/' in ref or ref.startswith('.'):
        skipped.append(ref or '(بلا رقم مرجعي)')
        continue
    tally[write(os.path.join(out, ref + '.json'),
                json.dumps(r, ensure_ascii=False, indent=2) + '\n')] += 1

master = write(os.path.join(out, '_all-responses.json'),
               json.dumps(rows, ensure_ascii=False, indent=2) + '\n')

print('')
print('  المجلد        : %s' % out)
print('  الاستجابات    : %d' % len(rows))
print('  ملفات جديدة   : %d' % tally['new'])
print('  ملفات محدَّثة  : %d' % tally['updated'])
print('  بلا تغيير     : %d' % tally['same'])
print('  الملف الجامع  : _all-responses.json (%s)'
      % {'new': 'أُنشئ', 'updated': 'حُدِّث', 'same': 'بلا تغيير'}[master])
if skipped:
    print('  تُخطّيت       : %d (رقم مرجعي غير صالح: %s)' % (len(skipped), '، '.join(skipped[:5])))
if not rows:
    print('\n  لا توجد استجابات بعد — أو أن النسخ الاحتياطية في Drive لم تُهيَّأ.')
print('')
PY
