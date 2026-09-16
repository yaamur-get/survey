/**
 * ═══════════════════════════════════════════════════════════════════════
 *  جمعية يعمر للعناية بالمساجد — نماذج المعرض التطوعي
 *  جامعة الملك فهد للبترول والمعادن
 *
 *  خادم واحد يخدم النموذجين ويكتب كلاً منهما في ورقة مستقلة:
 *      volunteer  →  ورقة «الفرص التطوعية»
 *      capstone   →  ورقة «مشاريع التخرج»
 *
 *  كل استجابة تُحفظ في ثلاثة أماكن مستقلة:
 *    ① صف في ورقة النموذج                              (للعرض والتحليل)
 *    ② ملف JSON مستقل لكل استجابة في Drive             (YV-XXXX.json)
 *    ③ ملف JSON جامع لكل نموذج                         (_volunteer.json)
 *  فقدان أي واحد منها لا يفقد البيانات.
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  ─── التركيب (مرة واحدة) ───
 *  ١) افتح Google Sheet جديداً  →  Extensions  →  Apps Script
 *  ٢) امسح المحتوى الافتراضي والصق هذا الملف كاملاً
 *  ٣) عدّل SECRET و ADMIN_KEY بالأسفل
 *  ٤) Deploy → New deployment → Type: Web app
 *        Execute as: Me   |   Who has access: Anyone      ← مهم جداً
 *  ٥) انسخ رابط النشر المنتهي بـ /exec والصقه في الحقل scriptUrl
 *     داخل الملفين volunteer.html و capstone.html (الرابط نفسه للاثنين)
 *  ٦) للتأكّد: افتح الرابط في المتصفح — يجب أن يظهر {"ok":true,...}
 *
 *  أول استجابة سيطلب Google صلاحية Drive — اقبلها مرة واحدة،
 *  أو شغّل الدالة testBackup من المحرر لإعطائها مقدماً.
 *
 *  تنبيه: عند أي تعديل لاحق على هذا السكربت لا بدّ من
 *  Deploy → Manage deployments → Edit → Version: New version → Deploy
 */

/* ══════════════════ الإعدادات ══════════════════ */

/** يجب أن تطابق الحقل secret في volunteer.html و capstone.html.
 *  تحذير: هذه ليست كلمة سر حقيقية — من يفتح كود الصفحة يراها.
 *  وظيفتها منع الإرسال العابر فقط، لا حماية البيانات. */
var SECRET = 'YAAMUR-KFUPM-2026';

/** مفتاح الإدارة لتصدير كل البيانات — لا يوضع في صفحات HTML أبداً. */
var ADMIN_KEY = 'CHANGE-ME-LONG-RANDOM-ADMIN-ONLY';

/** النماذج المسموح بها وأوراقها — لا يُكتب في ورقة خارج هذه القائمة. */
var FORMS = {
  volunteer: { sheet: 'الفرص التطوعية', label: 'الفرص التطوعية', master: '_volunteer.json' },
  capstone:  { sheet: 'مشاريع التخرج',  label: 'مشاريع التخرج',  master: '_capstone.json'  }
};

var BACKUP_FOLDER = 'yaamur-kfupm-backup';
var TIMEZONE      = 'Asia/Riyadh';
var REF_COL       = 2;      /* عمود الرقم المرجعي — يمنع التكرار */
var MAX_CELL      = 1000;   /* أقصى طول لخلية نصية */
var MAX_COLS      = 100;    /* حد أعلى لعدد الأعمدة */

/* ══════════════════ قائمة داخل الجدول ══════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('يعمر — نماذج المعرض')
    .addItem('عرض عدد الاستجابات', 'showCounts')
    .addItem('إعادة بناء النسخ من الجداول', 'rebuildBackupFromSheets')
    .addItem('فتح مجلد النسخ الاحتياطية', 'showBackupFolderLink')
    .addSeparator()
    .addItem('حذف استجابة برقمها المرجعي', 'promptDeleteResponse')
    .addToUi();
}

/* ══════════════════ فحص الجاهزية والتصدير ══════════════════ */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  /* تصدير الاستجابات — يحتاج مفتاح الإدارة، وليس SECRET */
  if (p.action === 'export') {
    if (!ADMIN_KEY || String(p.key || '') !== ADMIN_KEY) {
      return reply({ ok: false, error: 'unauthorized' });
    }
    var id = String(p.form || '');
    if (!FORMS[id]) return reply({ ok: false, error: 'form غير معروف. استخدم volunteer أو capstone.' });
    var all = readMaster(FORMS[id].master);
    return reply({ ok: true, form: id, count: all.length, responses: all });
  }

  var counts = {};
  for (var k in FORMS) counts[k] = countRows(FORMS[k].sheet);

  return reply({
    ok: true,
    service: 'yaamur-kfupm-forms',
    forms: counts,
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

    /* توجيه النموذج إلى ورقته — من القائمة البيضاء فقط */
    var formId = String(body.form || '');
    var form   = FORMS[formId];
    if (!form) {
      return reply({ ok: false, fatal: true, error: 'نموذج غير معروف: ' + formId });
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
    var sh = ss.getSheetByName(form.sheet);
    if (!sh) sh = ss.insertSheet(form.sheet);

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

    /* ① الكتابة في الورقة */
    var target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    /* لا نردّ بالنجاح قبل تأكيد الكتابة فعلياً */
    var written = String(sh.getRange(target, REF_COL).getValue()).trim();
    if (written !== ref) {
      return reply({ ok: false, error: 'تعذّر تأكيد حفظ الصف. أعد المحاولة.' });
    }

    /* ②③ النسخ الاحتياطية — فشلها لا يُفقد البيانات لأن الورقة حُفظت */
    var bk = saveBackup(formId, ref, head, row, target);

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
function toRecord(formId, ref, head, row, sheetRow) {
  var data = {};
  for (var i = 0; i < head.length; i++) {
    var v = row[i];
    /* إزالة الفاصلة العليا التي أُضيفت لأجل الجدول فقط */
    if (typeof v === 'string' && v.charAt(0) === "'") v = v.substring(1);
    data[head[i]] = v;
  }
  return { form: formId, ref: ref, savedAt: stamp(), sheetRow: sheetRow, data: data };
}

/** كل طبقة مستقلة تماماً: فشل إحداها لا يمنع الأخرى من العمل. */
function saveBackup(formId, ref, head, row, sheetRow) {
  var rec, content, name, errs = [];

  try {
    rec     = toRecord(formId, ref, head, row, sheetRow);
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

    /* ③ الملف الجامع للنموذج — يُحاول حتى لو فشل الملف المستقل */
    try { upsertMaster(folder, FORMS[formId].master, rec); }
    catch (e3) { errs.push('master: ' + e3); }
  }

  return errs.length ? { ok: false, error: errs.join(' | ') } : { ok: true };
}

/** يضيف الاستجابة إلى الملف الجامع أو يحدّثها إن وُجدت */
function upsertMaster(folder, masterName, rec) {
  var all = [];
  var hit = folder.getFilesByName(masterName);
  var file = hit.hasNext() ? hit.next() : null;

  if (file) {
    try {
      var parsed = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
      if (Array.isArray(parsed)) all = parsed;
    } catch (e) {
      /* ملف تالف: نحفظه جانباً بدل أن نطمس عليه */
      file.setName(masterName.replace('.json', '') + '-corrupt-' + fileStamp() + '.json');
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
  else folder.createFile(masterName, out, 'application/json');
}

function readMaster(masterName) {
  try {
    var hit = getFolder().getFilesByName(masterName);
    if (!hit.hasNext()) return [];
    var parsed = JSON.parse(hit.next().getBlob().getDataAsString('UTF-8'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function getFolder() {
  var hit = DriveApp.getFoldersByName(BACKUP_FOLDER);
  return hit.hasNext() ? hit.next() : DriveApp.createFolder(BACKUP_FOLDER);
}

/* ══════════════════ دوال تُشغَّل من المحرر أو القائمة ══════════════════ */

function showCounts() {
  var lines = [];
  for (var k in FORMS) lines.push(FORMS[k].label + ': ' + countRows(FORMS[k].sheet) + ' استجابة');
  return log_(lines.join('\n'));
}

/** يعيد بناء كل النسخ الاحتياطية من الجداول — استخدمها إن تعطّلت النسخ فترة */
function rebuildBackupFromSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = getFolder();
  var done = [];

  for (var k in FORMS) {
    var sh = ss.getSheetByName(FORMS[k].sheet);
    if (!sh || sh.getLastRow() < 2) { done.push(FORMS[k].label + ': لا توجد استجابات'); continue; }

    var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var head = vals[0];
    var all  = [];

    for (var r = 1; r < vals.length; r++) {
      var ref = String(vals[r][REF_COL - 1] || '').trim();
      if (!ref) continue;
      all.push(toRecord(k, ref, head, vals[r], r + 1));
    }

    var out = JSON.stringify(all, null, 2);
    var hit = folder.getFilesByName(FORMS[k].master);
    if (hit.hasNext()) hit.next().setContent(out);
    else folder.createFile(FORMS[k].master, out, 'application/json');

    done.push(FORMS[k].label + ': ' + all.length + ' استجابة');
  }

  return log_('أُعيد بناء النسخ.\n' + done.join('\n') + '\n\nالمجلد: ' + folder.getUrl());
}

function showBackupFolderLink() {
  return log_('مجلد النسخ الاحتياطية:\n' + getFolder().getUrl());
}

function promptDeleteResponse() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('حذف استجابة', 'اكتب الرقم المرجعي (مثال: YV-1A2B3C4D):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var ref = String(res.getResponseText() || '').trim();
  if (!ref) return log_('لم تُدخل رقماً مرجعياً.');
  return deleteResponse(ref);
}

/** يحذف الاستجابة من ورقتها ومن النسخ الاحتياطية */
function deleteResponse(ref) {
  ref = String(ref).trim();
  if (!ref) return log_('الرقم المرجعي مطلوب.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hitForm = null, hitRow = 0;

  for (var k in FORMS) {
    var sh = ss.getSheetByName(FORMS[k].sheet);
    if (!sh || sh.getLastRow() < 2) continue;
    var refs = sh.getRange(2, REF_COL, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < refs.length; i++) {
      if (String(refs[i][0]).trim() === ref) { hitForm = k; hitRow = i + 2; break; }
    }
    if (hitForm) { sh.deleteRow(hitRow); break; }
  }

  if (!hitForm) return log_('لم يُعثر على الرقم المرجعي ' + ref + ' في أي ورقة.');

  /* إزالة النسخ الاحتياطية أيضاً */
  try {
    var folder = getFolder();
    var f = folder.getFilesByName(ref + '.json');
    while (f.hasNext()) f.next().setTrashed(true);

    var all = readMaster(FORMS[hitForm].master).filter(function (r) { return r && r.ref !== ref; });
    var m = folder.getFilesByName(FORMS[hitForm].master);
    if (m.hasNext()) m.next().setContent(JSON.stringify(all, null, 2));
  } catch (e) { /* الصف حُذف من الورقة على كل حال */ }

  return log_('حُذفت الاستجابة ' + ref + ' من ورقة «' + FORMS[hitForm].label + '» (الصف ' + hitRow + ').');
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

function countRows(sheetName) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
