/**
 * ═══════════════════════════════════════════════════════════════════════
 *  استبيان يعمر — حفظ الاستجابات + نسخ احتياطية مستقلة
 *  اتحاد جمعيات العناية بالمساجد
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  كل استجابة تُحفظ في 4 أماكن مستقلة:
 *    1) صف في Google Sheet                        (للعرض والتحليل)
 *    2) ملف JSON مستقل لكل استجابة في Drive        (YM-XXXX.json)
 *    3) ملف JSON جامع يضم كل الاستجابات            (_all-responses.json)
 *    4) نسخة بالبريد الإلكتروني مع مرفق JSON       (اختياري — خارج Drive تماماً)
 *
 *  فقدان أي واحد منها لا يفقد البيانات.
 *
 *  ─── خطوات التركيب (مرة واحدة) ───
 *  1) افتح Google Sheet جديد  →  Extensions  →  Apps Script
 *  2) امسح المحتوى الافتراضي والصق هذا الملف كاملاً
 *  3) عدّل الإعدادات بالأسفل: SECRET و ADMIN_KEY و BACKUP_EMAIL
 *  4) Deploy → New deployment → Type: Web app
 *       Execute as: Me   |   Who has access: Anyone     ← مهم جداً
 *  5) انسخ رابط النشر المنتهي بـ /exec والصقه في CONFIG.SCRIPT_URL داخل index.html
 *  6) للتأكد: افتح الرابط في المتصفح — يجب أن يظهر {"ok":true,...}
 *
 *  أول استجابة سيطلب Google صلاحيات (Drive + Gmail) — اقبلها مرة واحدة.
 *  أو شغّل الدالة testBackup مرة من المحرر لإعطاء الصلاحيات مقدماً.
 *
 *  ملاحظة: عند أي تعديل لاحق على هذا السكربت لا بد من
 *  Deploy → Manage deployments → Edit → Version: New version → Deploy
 */

/* ══════════════════ الإعدادات ══════════════════ */

/** يجب أن تطابق CONFIG.SECRET في index.html.
 *  تحذير: هذه ليست كلمة سر حقيقية — أي شخص يفتح كود الصفحة يراها.
 *  وظيفتها منع الإرسال العابر فقط، لا حماية البيانات. */
var SECRET = 'YAAMUR-2026-data8';

/** مفتاح الإدارة لتصدير كل البيانات — لا يوضع في index.html أبداً.
 *  اجعله طويلاً وعشوائياً. */
var ADMIN_KEY = 'CHANGE-ME-LONG-RANDOM-ADMIN-ONLY';

/** بريد لإرسال نسخة من كل استجابة مع مرفق JSON.
 *  اتركه فارغاً '' لتعطيل الإرسال. */
var BACKUP_EMAIL = '';

var SHEET_NAME     = 'الاستجابات';
var BACKUP_FOLDER  = 'yamur-backup';
var MASTER_FILE    = '_all-responses.json';
var TIMEZONE       = 'Asia/Riyadh';
var REF_COL        = 2;      /* عمود الرقم المرجعي — يمنع التكرار */
var MAX_CELL       = 1000;   /* أقصى طول لخلية نصية */
var MAX_COLS       = 200;    /* حد أعلى لعدد الأعمدة */

/* ══════════════════ قائمة داخل الجدول ══════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('يعمر')
    .addItem('تصدير نسخة JSON الآن', 'backupNow')
    .addItem('إعادة بناء النسخة من الجدول', 'rebuildBackupFromSheet')
    .addItem('فتح مجلد النسخ الاحتياطية', 'showBackupFolderLink')
    .addSeparator()
    .addItem('عرض قائمة الاستجابات', 'listResponses')
    .addItem('حذف استجابة برقمها المرجعي', 'promptDeleteResponse')
    .addToUi();
}

/* ══════════════════ فحص الجاهزية والتصدير ══════════════════ */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  /* تصدير كل الاستجابات — يحتاج مفتاح الإدارة، وليس SECRET */
  if (p.action === 'export') {
    if (!ADMIN_KEY || String(p.key || '') !== ADMIN_KEY) {
      return reply({ ok: false, error: 'unauthorized' });
    }
    var all = readMaster();
    return reply({ ok: true, count: all.length, responses: all });
  }

  return reply({
    ok: true,
    service: 'yamur-survey',
    sheet: SHEET_NAME,
    responses: countRows(),
    time: stamp()
  });
}

/* ══════════════════ استلام الاستجابة وحفظها ══════════════════ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return reply({ ok: false, error: 'الخادم مشغول حالياً. أعد المحاولة بعد لحظات.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply({ ok: false, fatal: true, error: 'طلب فارغ.' });
    }

    var body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return reply({ ok: false, fatal: true, error: 'صيغة البيانات غير صحيحة.' });
    }

    if (String(body.secret || '') !== SECRET) {
      return reply({ ok: false, fatal: true, error: 'غير مصرَّح. تأكّد من تطابق الكلمة السرية.' });
    }

    var headers = body.headers;
    var values  = body.values;

    if (!Array.isArray(headers) || !Array.isArray(values)) {
      return reply({ ok: false, fatal: true, error: 'بنية البيانات غير صحيحة.' });
    }
    if (headers.length !== values.length) {
      return reply({ ok: false, fatal: true,
        error: 'عدد الأعمدة (' + headers.length + ') لا يطابق عدد القيم (' + values.length + ').' });
    }
    if (headers.length === 0 || headers.length > MAX_COLS) {
      return reply({ ok: false, fatal: true, error: 'عدد الأعمدة غير مقبول.' });
    }

    var ref = clean(String(body.ref || ''), 40);
    if (!ref) {
      return reply({ ok: false, fatal: true, error: 'الرقم المرجعي مفقود.' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) sh = ss.insertSheet(SHEET_NAME);

    var head = headers.map(function (h) { return clean(String(h), 200); });
    ensureHeader(sh, head);

    /* منع التكرار: نفس الرقم المرجعي لا يُكتب مرتين */
    var last = sh.getLastRow();
    if (last > 1) {
      var refs = sh.getRange(2, REF_COL, last - 1, 1).getValues();
      for (var i = 0; i < refs.length; i++) {
        if (String(refs[i][0]).trim() === ref) {
          return reply({ ok: true, row: i + 2, duplicate: true });
        }
      }
    }

    /* تنقية القيم */
    var row = values.map(function (v, i) {
      if (i === 0) return stamp();                     /* وقت الإرسال من الخادم */
      if (v === null || v === undefined) return '';
      if (typeof v === 'number') return isFinite(v) ? v : '';
      if (typeof v === 'boolean') return v;
      return guard(clean(String(v), MAX_CELL));
    });

    /* ① الكتابة في الجدول */
    var target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    /* التحقق من نجاح الكتابة فعلياً قبل الرد بالنجاح */
    var written = String(sh.getRange(target, REF_COL).getValue()).trim();
    if (written !== ref) {
      return reply({ ok: false, error: 'تعذّر تأكيد حفظ الصف. أعد المحاولة.' });
    }

    /* ②③④ النسخ الاحتياطية — فشلها لا يُفقد البيانات لأن الجدول حُفظ */
    var bk = saveBackup(ref, head, row, target);

    var res = { ok: true, row: target, backup: bk.ok };
    if (!bk.ok) res.backupWarning = bk.error;
    return reply(res);

  } catch (err) {
    return reply({ ok: false, error: 'خطأ في الخادم: ' + (err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ══════════════════ النسخ الاحتياطية ══════════════════ */

/** يحوّل العناوين والقيم إلى كائن مقروء */
function toRecord(ref, head, row, sheetRow) {
  var data = {};
  for (var i = 0; i < head.length; i++) {
    var v = row[i];
    /* إزالة الفاصلة العليا التي أُضيفت لأجل الجدول فقط */
    if (typeof v === 'string' && v.charAt(0) === "'") v = v.substring(1);
    data[head[i]] = v;
  }
  return { ref: ref, savedAt: stamp(), sheetRow: sheetRow, data: data };
}

/** يحفظ الاستجابة في ملف JSON مستقل + الملف الجامع + بريد اختياري.
 *  كل طبقة مستقلة تماماً: فشل إحداها لا يمنع الأخريات من العمل. */
function saveBackup(ref, head, row, sheetRow) {
  var rec, content, name, errs = [];

  try {
    rec     = toRecord(ref, head, row, sheetRow);
    content = JSON.stringify(rec, null, 2);
    name    = ref + '.json';
  } catch (e0) {
    return { ok: false, error: 'record: ' + e0 };
  }

  var folder = null;
  try { folder = getFolder(); }
  catch (e1) { errs.push('folder: ' + e1); }

  if (folder) {
    /* ② ملف مستقل لكل استجابة */
    try {
      var hit = folder.getFilesByName(name);
      if (hit.hasNext()) hit.next().setContent(content);
      else folder.createFile(name, content, 'application/json');
    } catch (e2) { errs.push('file: ' + e2); }

    /* ③ الملف الجامع — يُحاول حتى لو فشل الملف المستقل */
    try { upsertMaster(folder, rec); }
    catch (e3) { errs.push('master: ' + e3); }
  }

  /* ④ نسخة بالبريد — خارج Drive، تُحاول حتى لو فشل Drive كله */
  if (BACKUP_EMAIL) {
    try {
      MailApp.sendEmail({
        to: BACKUP_EMAIL,
        subject: 'استبيان يعمر — استجابة جديدة ' + ref,
        body: 'وصلت استجابة جديدة.\n\n' +
              'الرقم المرجعي: ' + ref + '\n' +
              'الجمعية: ' + (rec.data['اسم الجمعية'] || '') + '\n' +
              'الوقت: ' + rec.savedAt + '\n' +
              'صف الجدول: ' + sheetRow + '\n\n' +
              'نسخة JSON كاملة في المرفق.',
        attachments: [{ fileName: name, content: content, mimeType: 'application/json' }]
      });
    } catch (e4) { errs.push('email: ' + e4); }
  }

  return errs.length ? { ok: false, error: errs.join(' | ') } : { ok: true };
}

/** يضيف الاستجابة إلى الملف الجامع أو يحدّثها إن وُجدت */
function upsertMaster(folder, rec) {
  var all = [];
  var hit = folder.getFilesByName(MASTER_FILE);
  var file = hit.hasNext() ? hit.next() : null;

  if (file) {
    try {
      var parsed = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
      if (Array.isArray(parsed)) all = parsed;
    } catch (e) {
      /* ملف تالف: نحفظه جانباً بدل أن نطمس عليه */
      file.setName(MASTER_FILE.replace('.json', '') + '-corrupt-' + fileStamp() + '.json');
      file = null;
    }
  }

  var found = false;
  for (var i = 0; i < all.length; i++) {
    if (all[i] && all[i].ref === rec.ref) { all[i] = rec; found = true; break; }
  }
  if (!found) all.push(rec);

  var out = JSON.stringify(all, null, 2);
  if (file) file.setContent(out);
  else folder.createFile(MASTER_FILE, out, 'application/json');
}

/** يقرأ الملف الجامع */
function readMaster() {
  try {
    var hit = getFolder().getFilesByName(MASTER_FILE);
    if (!hit.hasNext()) return [];
    var parsed = JSON.parse(hit.next().getBlob().getDataAsString('UTF-8'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

/** يجد مجلد النسخ أو ينشئه */
function getFolder() {
  var hit = DriveApp.getFoldersByName(BACKUP_FOLDER);
  return hit.hasNext() ? hit.next() : DriveApp.createFolder(BACKUP_FOLDER);
}

/* ══════════════════ دوال تُشغَّل من المحرر أو القائمة ══════════════════ */

/** يعيد بناء كل النسخ الاحتياطية من الجدول — استخدمها إن تعطّلت النسخ فترة */
function rebuildBackupFromSheet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return log_('لا توجد استجابات في الجدول.');

  var vals   = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var head   = vals[0];
  var folder = getFolder();
  var all    = [];

  for (var r = 1; r < vals.length; r++) {
    var ref = String(vals[r][REF_COL - 1]).trim();
    if (!ref) continue;
    var rec = toRecord(ref, head, vals[r], r + 1);
    all.push(rec);
    var name = ref + '.json';
    var content = JSON.stringify(rec, null, 2);
    var hit = folder.getFilesByName(name);
    if (hit.hasNext()) hit.next().setContent(content);
    else folder.createFile(name, content, 'application/json');
  }

  var out = JSON.stringify(all, null, 2);
  var mh  = folder.getFilesByName(MASTER_FILE);
  if (mh.hasNext()) mh.next().setContent(out);
  else folder.createFile(MASTER_FILE, out, 'application/json');

  return log_('أُعيد بناء ' + all.length + ' استجابة في مجلد ' + BACKUP_FOLDER);
}

/** يحدّث الملف الجامع ويرسله بالبريد إن كان مُهيَّأً */
function backupNow() {
  rebuildBackupFromSheet();
  if (BACKUP_EMAIL) {
    var all = readMaster();
    MailApp.sendEmail({
      to: BACKUP_EMAIL,
      subject: 'استبيان يعمر — نسخة كاملة (' + all.length + ' استجابة)',
      body: 'نسخة احتياطية كاملة بتاريخ ' + stamp() + '.',
      attachments: [{
        fileName: 'yamur-all-' + fileStamp() + '.json',
        content: JSON.stringify(all, null, 2),
        mimeType: 'application/json'
      }]
    });
    return log_('أُرسلت نسخة كاملة (' + all.length + ' استجابة) إلى ' + BACKUP_EMAIL);
  }
  return log_('حُدِّثت النسخة في Drive. لتفعيل الإرسال بالبريد اضبط BACKUP_EMAIL.');
}

/** يطبع رابط مجلد النسخ */
function showBackupFolderLink() {
  return log_('مجلد النسخ الاحتياطية: ' + getFolder().getUrl());
}

/** يعرض كل الاستجابات: الرقم المرجعي والجمعية والوقت — لمعرفة ما يجب حذفه */
function listResponses() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return log_('لا توجد استجابات.');

  var vals = sh.getRange(1, 1, sh.getLastRow(), 4).getValues();
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    out.push((r + 1) + '  |  ' + vals[r][1] + '  |  ' + vals[r][2] + '  |  ' + vals[r][0]);
  }
  return log_('الصف | الرقم المرجعي | الجمعية | الوقت\n\n' + out.join('\n'));
}

/** يسأل عن الرقم المرجعي ثم يحذف الاستجابة من الجدول ومن النسخ */
function promptDeleteResponse() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('حذف استجابة',
    'اكتب الرقم المرجعي المطلوب حذفه (مثال: YM-1A2B3C4D).\nسيُحذف من الجدول ومن النسخ الاحتياطية.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  deleteResponse(res.getResponseText().trim());
}

/** يحذف استجابة واحدة برقمها المرجعي من الجدول ومن ملفات النسخ */
function deleteResponse(ref) {
  ref = String(ref || '').trim();
  if (!ref) return log_('لم يُحدَّد رقم مرجعي.');

  var removed = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) {
    var refs = sh.getRange(2, REF_COL, sh.getLastRow() - 1, 1).getValues();
    for (var i = refs.length - 1; i >= 0; i--) {      /* من الأسفل لأعلى حتى لا تتغيّر الأرقام */
      if (String(refs[i][0]).trim() === ref) {
        sh.deleteRow(i + 2);
        removed.push('صف الجدول ' + (i + 2));
      }
    }
    SpreadsheetApp.flush();
  }

  var folder = null;
  try { folder = getFolder(); }
  catch (e1) { removed.push('تحذير: تعذّر فتح مجلد النسخ (' + e1 + ')'); }

  if (folder) {
    /* حذف الملف المستقل */
    try {
      var hit = folder.getFilesByName(ref + '.json');
      while (hit.hasNext()) { hit.next().setTrashed(true); removed.push('ملف ' + ref + '.json'); }
    } catch (e2) { removed.push('تحذير: تعذّر حذف الملف المستقل (' + e2 + ')'); }

    /* تنظيف الملف الجامع — مستقل عن الخطوة السابقة */
    try {
      var all = readMaster().filter(function (r) { return !r || r.ref !== ref; });
      var out = JSON.stringify(all, null, 2);
      var mh  = folder.getFilesByName(MASTER_FILE);
      if (mh.hasNext()) mh.next().setContent(out);
      else folder.createFile(MASTER_FILE, out, 'application/json');
      removed.push('الملف الجامع');
    } catch (e3) { removed.push('تحذير: تعذّر تحديث الملف الجامع (' + e3 + ')'); }
  }

  return log_(removed.length
    ? 'حُذف ' + ref + ' من: ' + removed.join('، ')
    : 'لم يُعثر على ' + ref + '.');
}

/** شغّلها مرة من المحرر لإعطاء الصلاحيات مقدماً واختبار النسخ */
function testBackup() {
  var folder = getFolder();
  var name = '_test.json';
  var hit = folder.getFilesByName(name);
  var content = JSON.stringify({ test: true, time: stamp() }, null, 2);
  if (hit.hasNext()) hit.next().setContent(content);
  else folder.createFile(name, content, 'application/json');
  return log_('النسخ الاحتياطية تعمل. المجلد: ' + folder.getUrl());
}

/* ══════════════════ أدوات مساعدة ══════════════════ */

function ensureHeader(sh, head) {
  var lastCol = sh.getLastColumn();
  var existing = (sh.getLastRow() > 0 && lastCol > 0)
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  var empty = existing.length === 0 || existing.join('').replace(/\s/g, '') === '';

  if (empty) {
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return;
  }
  if (head.length > existing.length) {
    var extra = head.slice(existing.length);
    sh.getRange(1, existing.length + 1, 1, extra.length).setValues([extra]);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
}

function countRows() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    return (sh && sh.getLastRow() > 1) ? sh.getLastRow() - 1 : 0;
  } catch (e) { return 0; }
}

/** حذف محارف التحكّم ومحارف اتجاه النص وتحديد الطول */
function clean(s, max) {
  var ctrl = new RegExp('[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\u007F]', 'g');
  var bidi = new RegExp('[\\u200B-\\u200F\\u202A-\\u202E\\u2066-\\u2069\\uFEFF]', 'g');
  return String(s).replace(ctrl, '').replace(bidi, '').slice(0, max || MAX_CELL);
}

/** منع تنفيذ الصيغ داخل الجدول */
function guard(s) {
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function stamp() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
function fileStamp() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmmss');
}

function log_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
