/* ═══════════════════════════════════════════════════════════════
   محرّك النماذج — جمعية يعمر للعناية بالمساجد
   يقرأ المخطَّط من window.YFORM ويبني نموذجاً متعدّد الشاشات:
   حفظ تلقائي على الجهاز · تحقّق من المدخلات · مراجعة · إرسال
   إلى Google Sheets عبر Apps Script مع إعادة محاولة.
   ملف واحد يخدم كل النماذج — لا تُكرَّر هذه الشِّفرة في الصفحات.
   ═══════════════════════════════════════════════════════════════ */
(function () {
'use strict';

var F = window.YFORM;
if (!F) { return; }

var MAX_TEXT = 600;

/* ═══════════ أدوات نصّية ═══════════ */
function esc(s){
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }

/* الأرقام العربية والفارسية إلى لاتينية */
var RE_ARDIG = new RegExp('[\\u0660-\\u0669\\u06F0-\\u06F9]', 'g');
function normDigits(s){
  return String(s == null ? '' : s).replace(RE_ARDIG, function(d){
    var c = d.charCodeAt(0);
    return String(c >= 0x06F0 ? c - 0x06F0 : c - 0x0660);
  });
}
/* إزالة محارف التحكّم والاتجاه — تفسد الخلايا وتخفي محتوى */
var RE_CTRL = new RegExp('[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\u007F]', 'g');
var RE_BIDI = new RegExp('[\\u200B-\\u200F\\u202A-\\u202E\\u2060-\\u2064\\u2066-\\u2069\\uFEFF]', 'g');
function cleanText(raw, max){
  var s = String(raw == null ? '' : raw).replace(RE_CTRL,'').replace(RE_BIDI,'');
  return s.slice(0, max || MAX_TEXT);
}
/* منع تنفيذ الصيغ داخل Google Sheets أو Excel */
function safeCell(v){
  if (typeof v !== 'string') return v;
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}

/* ═══════════ التحقّق ═══════════ */
var RE_LETTER = new RegExp('[\\u0621-\\u064AA-Za-z]');

function vText(v, f){
  var s = cleanText(v, f.max || 200).trim().replace(/\s{2,}/g,' ');
  if (!s) return { ok:false, msg:'هذا الحقل مطلوب.' };
  if (s.length < 2) return { ok:false, msg:'الإجابة قصيرة جداً.' };
  return { ok:true, val:s };
}
function vName(v, f){
  var s = cleanText(v, f.max || 120).trim().replace(/\s{2,}/g,' ');
  if (s.length < 5 || s.split(' ').length < 2)
    return { ok:false, msg:'اكتب الاسم كاملاً — الاسم الأول واسم الأب واسم العائلة.' };
  if (!RE_LETTER.test(s)) return { ok:false, msg:'الاسم غير صحيح.' };
  return { ok:true, val:s };
}
function vEmail(v){
  var s = cleanText(v, 100).trim();
  if (!/^[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9][A-Za-z0-9.-]{0,62}\.[A-Za-z]{2,12}$/.test(s))
    return { ok:false, msg:'صيغة البريد غير صحيحة. مثال: s202012345@kfupm.edu.sa' };
  return { ok:true, val:s.toLowerCase() };
}
function vTel(v){
  var s = normDigits(v).replace(/\D/g,'');
  if (/^00966\d{9}$/.test(s))     s = '0' + s.slice(5);
  else if (/^966\d{9}$/.test(s))  s = '0' + s.slice(3);
  if (!/^05\d{8}$/.test(s))
    return { ok:false, msg:'رقم الجوال يجب أن يكون ١٠ أرقام ويبدأ بـ 05. مثال: 0551234567' };
  return { ok:true, val:s };
}
function vNum(v, f){
  var s = normDigits(v).replace(/[^\d]/g,'');
  if (!s) return { ok:false, msg:'اكتب رقماً.' };
  var n = parseInt(s, 10);
  var lo = (f.min == null ? 0 : f.min), hi = (f.max == null ? 999 : f.max);
  if (!isFinite(n) || n < lo || n > hi)
    return { ok:false, msg:'اكتب رقماً بين ' + lo + ' و' + hi + '.' };
  return { ok:true, val:n };
}

var VALD = { name:vName, email:vEmail, tel:vTel };

/* ═══════════ التخزين ═══════════ */
var STORE_KEY = F.storeKey;
var SENT_KEY  = F.storeKey + '_sent';
var D = {};
try { D = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch(e){ D = {}; }
if (typeof D !== 'object' || D === null) D = {};

var saveTimer = null, savedTimer = null, wiped = false;

function persist(quiet){
  if (wiped) return;
  try{
    D._idx = idx;
    D._at  = new Date().toISOString();
    localStorage.setItem(STORE_KEY, JSON.stringify(D));
    if (!quiet) flashSaved('✓ محفوظ على جهازك');
  }catch(e){ /* التخزين ممتلئ أو محظور — نتابع دون توقّف */ }
}
function autosave(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(function(){ persist(); }, 400);
}
function flashSaved(txt){
  var el = document.getElementById('savedlbl'); if (!el) return;
  el.textContent = txt;
  el.classList.add('on');
  clearTimeout(savedTimer);
  savedTimer = setTimeout(function(){ el.classList.remove('on'); }, 1800);
}
function wipeAndReload(){
  wiped = true;
  clearTimeout(saveTimer);
  try{ localStorage.removeItem(STORE_KEY); }catch(e){}
  location.reload();
}
function resetAll(){
  if (!confirm('سيتم مسح جميع إجاباتك المحفوظة على هذا الجهاز والبدء من جديد. هل تريد المتابعة؟')) return;
  wipeAndReload();
}

/* رقم مرجعي فريد للاستجابة */
function makeRef(){
  try{
    var a = new Uint8Array(4);
    window.crypto.getRandomValues(a);
    return F.refPrefix + '-' + Array.prototype.map.call(a, function(b){
      return ('0' + b.toString(16)).slice(-2);
    }).join('').toUpperCase();
  }catch(e){
    return F.refPrefix + '-' + String(Date.now()).slice(-8);
  }
}
if (!D._ref) D._ref = makeRef();

/* سجلّ ما أُرسل من هذا الجهاز — الأجهزة في المعرض مشتركة بين عدّة طلاب */
function readSent(){
  try{
    var a = JSON.parse(localStorage.getItem(SENT_KEY) || '[]');
    return Array.isArray(a) ? a : [];
  }catch(e){ return []; }
}
function recordSent(){
  try{
    var a = readSent().filter(function(r){ return r && r.ref !== D._ref; });
    a.push({ ref:D._ref, name:cleanText(D[F.nameKey] || '', 120).trim(), at:new Date().toISOString() });
    localStorage.setItem(SENT_KEY, JSON.stringify(a.slice(-50)));
  }catch(e){ /* لا يمنع الإرسال */ }
}

/* ═══════════ بناء الشاشات ═══════════ */
var SCREENS = [{ type:'intro' }];
F.steps.forEach(function(st, i){ SCREENS.push({ type:'step', si:i }); });
SCREENS.push({ type:'review' });
SCREENS.push({ type:'done' });

var DONE_IDX = SCREENS.length - 1;
var idx = 0, sending = false;

/* كل الحقول بالترتيب — بما فيها حقول «أخرى» المولَّدة */
function allFields(){
  var out = [];
  F.steps.forEach(function(st){
    st.fields.forEach(function(f){
      out.push(f);
      if (f.other) out.push({ k:f.k + '_other', t:'تحديد «أخرى» — ' + f.t, col:f.colOther || (f.col || f.t) + ' (أخرى)', type:'text', _gen:true, parent:f });
    });
  });
  return out;
}
function fieldByKey(k){
  var all = allFields();
  for (var i = 0; i < all.length; i++) if (all[i].k === k) return all[i];
  return null;
}

/* هل الحقل ظاهر الآن؟ */
function visible(f){
  if (f._gen) return visible(f.parent) && hasOther(f.parent);
  return !f.showIf || !!f.showIf(D);
}
function hasOther(f){
  var v = D[f.k];
  if (Array.isArray(v)) return v.indexOf('أخرى') > -1;
  return v === 'أخرى';
}

/* ═══════════ رسم الحقول ═══════════ */
function lblFor(f){
  var opt = (f.req === false) ? '<span class="opt-lbl">(اختياري)</span>' : '';
  return '<span class="field-lbl" id="l-' + esc(f.k) + '">' + esc(f.t) + opt + '</span>';
}
function hintFor(f){
  return f.hint ? '<span class="field-hint">' + esc(f.hint) + '</span>' : '';
}
var TICK = '<svg width="10" height="10" viewBox="0 0 12 12" fill="none" aria-hidden="true">'
         + '<path d="M2 6.2L4.8 9L10 3" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

function drawField(f){
  var v = D[f.k];
  var body = '';

  if (f.type === 'textarea'){
    var tv = cleanText(v || '', f.max || MAX_TEXT);
    body = '<div class="ta-wrap" id="w-' + esc(f.k) + '">'
         + '<textarea class="ta" id="in-' + esc(f.k) + '" data-k="' + esc(f.k) + '" rows="3"'
         + ' maxlength="' + (f.max || MAX_TEXT) + '" aria-labelledby="l-' + esc(f.k) + '"'
         + ' placeholder="' + esc(f.ph || '') + '">' + esc(tv) + '</textarea>'
         + '<div class="ta-foot" id="c-' + esc(f.k) + '">' + tv.length + ' / ' + (f.max || MAX_TEXT) + '</div>'
         + '</div>';

  } else if (f.type === 'select'){
    var o = '<option value="">— اختر —</option>';
    f.opts.forEach(function(x){
      o += '<option value="' + esc(x) + '"' + (v === x ? ' selected' : '') + '>' + esc(x) + '</option>';
    });
    body = '<div class="sel-wrap"><select class="field-sel" id="in-' + esc(f.k) + '" data-k="' + esc(f.k) + '"'
         + ' aria-labelledby="l-' + esc(f.k) + '">' + o + '</select></div>';

  } else if (f.type === 'radio' || f.type === 'checks'){
    var multi = (f.type === 'checks');
    var sel   = multi ? (Array.isArray(v) ? v : []) : [v];
    var mark  = multi ? 'box' : 'dot';
    var o2 = f.opts.map(function(x){
      var on = sel.indexOf(x) > -1;
      return '<span class="opt' + (on ? ' on' : '') + '" role="' + (multi ? 'checkbox' : 'radio') + '"'
           + ' tabindex="0" aria-checked="' + (on ? 'true' : 'false') + '"'
           + ' data-opt="' + esc(f.k) + '" data-val="' + esc(x) + '">'
           + '<span class="opt-mark ' + mark + '">' + (multi ? TICK : '') + '</span>'
           + esc(x) + '</span>';
    }).join('');
    body = '<div class="opts' + (f.cols ? ' col' : '') + '" id="w-' + esc(f.k) + '" role="'
         + (multi ? 'group' : 'radiogroup') + '" aria-labelledby="l-' + esc(f.k) + '">' + o2 + '</div>';

  } else if (f.type === 'number'){
    body = '<div class="num-wrap"><input class="field-inp" id="in-' + esc(f.k) + '" data-k="' + esc(f.k) + '"'
         + ' type="text" inputmode="numeric" autocomplete="off" aria-labelledby="l-' + esc(f.k) + '"'
         + ' maxlength="3" placeholder="' + esc(f.ph || '0') + '" value="' + esc(v == null ? '' : v) + '">'
         + (f.unit ? '<span class="num-unit">' + esc(f.unit) + '</span>' : '') + '</div>';

  } else if (f.type === 'consent'){
    return '<div class="field' + (visible(f) ? '' : ' hidden') + '" id="f-' + esc(f.k) + '">'
         + '<label class="agree' + (D[f.k] ? ' on' : '') + '" id="w-' + esc(f.k) + '" data-consent="' + esc(f.k) + '">'
         + '<span class="chk' + (D[f.k] ? ' on' : '') + '" role="checkbox" tabindex="0"'
         + ' aria-checked="' + (D[f.k] ? 'true' : 'false') + '">'
         + '<svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">'
         + '<path d="M2 6.2L4.8 9L10 3" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"/></svg>'
         + '</span><span class="agree-lbl">' + esc(f.t) + '</span></label>'
         + '<div class="field-note" id="n-' + esc(f.k) + '"></div></div>';

  } else {
    var t = (f.type === 'email') ? 'email' : (f.type === 'tel' ? 'tel' : 'text');
    var ltr = (f.type === 'email' || f.type === 'tel') ? ' dir="ltr" style="text-align:right"' : '';
    var dl  = f.list ? ' list="dl-' + esc(f.k) + '"' : '';
    body = '<input class="field-inp" id="in-' + esc(f.k) + '" data-k="' + esc(f.k) + '" type="' + t + '"'
         + ' maxlength="' + (f.max || 200) + '" autocomplete="' + esc(f.ac || 'off') + '"'
         + ' aria-labelledby="l-' + esc(f.k) + '"' + ltr + dl
         + (f.type === 'tel' ? ' inputmode="numeric"' : '')
         + ' placeholder="' + esc(f.ph || '') + '" value="' + esc(v == null ? '' : v) + '">';
    if (f.list){
      body += '<datalist id="dl-' + esc(f.k) + '">'
            + f.list.map(function(x){ return '<option value="' + esc(x) + '"></option>'; }).join('')
            + '</datalist>';
    }
  }

  return '<div class="field' + (visible(f) ? '' : ' hidden') + '" id="f-' + esc(f.k) + '">'
       + lblFor(f) + hintFor(f) + body
       + '<div class="field-note" id="n-' + esc(f.k) + '"></div></div>';
}

/* ═══════════ الشاشات ═══════════ */
function renderIntro(){
  var pts = (F.intro.points || []).map(function(p){ return '<li>' + p + '</li>'; }).join('');
  var resume = '';
  if (D._idx > 0 && D._idx < DONE_IDX && !D.submitted){
    resume = '<div class="resume-box">لديك إجابات محفوظة على هذا الجهاز — <b>ستُستأنف من حيث توقّفت</b>.</div>';
  }
  return '<div class="eyebrow">' + esc(F.eyebrow) + '</div>'
       + '<h1 class="scr-title">' + esc(F.intro.title || 'قبل البدء') + '</h1>'
       + resume
       + '<p class="lede">' + F.intro.lede + '</p>'
       + (pts ? '<div class="plain-h">' + esc(F.intro.pointsTitle || 'قبل أن تبدأ') + '</div><ul class="plain-list">' + pts + '</ul>' : '')
       + '<button class="btn-primary" data-act="next">' + esc(F.intro.cta || 'ابدأ') + '</button>';
}

function renderStep(q){
  var st = F.steps[q.si];
  return '<div class="eyebrow">' + esc(st.eyebrow || F.eyebrow) + '</div>'
       + '<h1 class="scr-title">' + esc(st.title) + '</h1>'
       + (st.sub ? '<p class="scr-sub tight">' + esc(st.sub) + '</p>' : '')
       + '<div class="fields">' + st.fields.map(function(f){
            var h = drawField(f);
            if (f.other){
              h += drawField({ k:f.k + '_other', t:f.otherLabel || 'حدّد «أخرى»', type:'text',
                               ph:f.otherPh || '', max:120, _gen:true, parent:f });
            }
            return h;
         }).join('') + '</div>'
       + '<div class="err-msg" id="err-step"></div>'
       + '<button class="btn-primary" data-act="next">التالي</button>'
       + '<button class="btn-back" data-act="prev">رجوع</button>';
}

function showVal(f){
  var v = D[f.k];
  if (f.type === 'consent') return v ? 'نعم — موافق' : '';
  if (Array.isArray(v)) return v.join('، ');
  return (v == null ? '' : String(v)).trim();
}

function renderReview(){
  var html = '';
  F.steps.forEach(function(st, si){
    var rows = '';
    st.fields.forEach(function(f){
      if (!visible(f)) return;
      var val = showVal(f);
      var isTxt = (f.type === 'textarea' || f.type === 'checks' || (val && val.length > 34));
      /* نصّ الموافقة فقرة كاملة — نعرض تسمية قصيرة بدلاً منها في المراجعة */
      var lbl = f.revT || (f.type === 'consent' ? (f.col || 'الموافقة') : f.t);
      rows += '<div class="rev-row"><span class="rev-k">' + esc(lbl) + '</span>'
            + '<span class="rev-v' + (isTxt ? ' txt' : '') + (val ? '' : ' empty') + '">'
            + (val ? esc(val) : '—') + '</span></div>';
      if (f.other && hasOther(f)){
        var ov = (D[f.k + '_other'] || '').trim();
        rows += '<div class="rev-row"><span class="rev-k">' + esc(f.otherLabel || 'تحديد «أخرى»') + '</span>'
              + '<span class="rev-v txt' + (ov ? '' : ' empty') + '">' + (ov ? esc(ov) : '—') + '</span></div>';
      }
    });
    if (!rows) return;
    html += '<div class="rev-group"><div class="rev-h"><span>' + esc(st.title) + '</span>'
          + '<button class="rev-edit" data-act="jump" data-to="' + (si + 1) + '">تعديل</button></div>'
          + rows + '</div>';
  });

  return '<div class="eyebrow">الخطوة الأخيرة</div>'
       + '<h1 class="scr-title">مراجعة قبل الإرسال</h1>'
       + '<p class="scr-sub tight">راجع بياناتك ثم أرسل. لا تصل البيانات إلى الجمعية إلا بعد تأكيد الإرسال.</p>'
       + html
       + '<div class="err-msg" id="err-send"></div>'
       + '<button class="btn-primary" id="btn-send" data-act="send">إرسال النموذج</button>'
       + '<button class="btn-back" id="btn-back-rev" data-act="prev">رجوع</button>'
       + '<button class="btn-back" style="font-size:.72rem" data-act="reset">مسح الإجابات والبدء من جديد</button>';
}

function renderDone(){
  var prev = readSent().filter(function(r){ return r && r.ref && r.ref !== D._ref; });
  var log = '';
  if (prev.length){
    log = '<div class="sent-log"><span class="sent-log-h">نماذج سابقة أُرسلت من هذا الجهاز (' + prev.length + ')</span>'
        + prev.slice().reverse().map(function(r){
            return '<div class="sent-log-row"><span>' + esc(r.name || '—') + '</span><b>' + esc(r.ref) + '</b></div>';
          }).join('')
        + '</div>';
  }
  return '<div class="done-wrap">'
       + '<div class="done-mark"><svg width="26" height="26" viewBox="0 0 32 32" fill="none" aria-hidden="true">'
       + '<path d="M5 16L12.5 23.5L27 8" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg></div>'
       + '<h1 class="done-title">' + esc(F.done.title) + '</h1>'
       + '<p class="done-sub">' + F.done.sub + '</p>'
       + '<div class="ref-box">الرقم المرجعي لطلبك<b>' + esc(D._ref) + '</b></div>'
       + log
       + '<button class="btn-primary" style="margin-top:1.4rem" data-act="new">تعبئة نموذج جديد</button>'
       + '</div>';
}

function render(q){
  var h = '';
  if (q.type === 'intro')  h = renderIntro();
  if (q.type === 'step')   h = renderStep(q);
  if (q.type === 'review') h = renderReview();
  if (q.type === 'done')   h = renderDone();
  return '<div class="' + (q.type === 'review' ? 'wide' : '') + '">' + h + '</div>';
}

/* ═══════════ التحقّق ═══════════ */
function checkField(f){
  var v = D[f.k];

  if (f.type === 'consent')
    return v ? { ok:true, val:true } : { ok:false, msg:'يلزم الموافقة للمتابعة.' };

  if (f.type === 'checks'){
    var a = Array.isArray(v) ? v : [];
    if (f.req === false) return { ok:true, val:a };
    if (!a.length) return { ok:false, msg:'اختر خياراً واحداً على الأقل.' };
    if (f.maxPick && a.length > f.maxPick)
      return { ok:false, msg:'اختر ' + f.maxPick + ' خيارات كحدٍّ أقصى.' };
    return { ok:true, val:a };
  }

  if (f.type === 'radio' || f.type === 'select'){
    if (f.req === false && !v) return { ok:true, val:'' };
    if (!v) return { ok:false, msg:'اختر أحد الخيارات.' };
    return { ok:true, val:v };
  }

  if (f.type === 'number'){
    if (f.req === false && !String(v || '').trim()) return { ok:true, val:'' };
    return vNum(v, f);
  }

  /* نصّي */
  if (f.req === false){
    var s = cleanText(v || '', f.max || MAX_TEXT).trim();
    if (!s) return { ok:true, val:'' };
    if (f.vald && VALD[f.vald]) return VALD[f.vald](v, f);
    return { ok:true, val:s.replace(/\s{2,}/g,' ') };
  }
  if (f.vald && VALD[f.vald]) return VALD[f.vald](v, f);
  return vText(v, f);
}

function markField(k, msg){
  var n = document.getElementById('n-' + k);
  var w = document.getElementById('w-' + k) || document.getElementById('in-' + k);
  if (n){ n.textContent = msg || ''; n.classList.toggle('on', !!msg); }
  if (w) w.classList.toggle('bad', !!msg);
}

function validateStep(si){
  var st = F.steps[si], first = null, count = 0;
  st.fields.forEach(function(f){
    if (!visible(f)){ markField(f.k, ''); return; }
    var r = checkField(f);
    markField(f.k, r.ok ? '' : r.msg);
    if (!r.ok){ count++; if (!first) first = f.k; }
    else if (r.val !== undefined && f.type !== 'checks' && f.type !== 'consent') D[f.k] = r.val;

    if (f.other){
      var ok2 = f.k + '_other';
      if (visible({ _gen:true, parent:f })){
        var ov = cleanText(D[ok2] || '', 120).trim();
        if (!ov){ markField(ok2, 'اكتب ما تقصده بـ «أخرى».'); count++; if (!first) first = ok2; }
        else { markField(ok2, ''); D[ok2] = ov; }
      } else { markField(ok2, ''); }
    }
  });
  return first ? { first:first, count:count } : null;
}

function showStepErr(res){
  var e = document.getElementById('err-step');
  if (e){
    e.textContent = res.count > 1
      ? 'يُرجى تصحيح ' + res.count + ' حقول قبل المتابعة.'
      : 'يُرجى تصحيح الحقل المُعلَّم قبل المتابعة.';
    e.classList.add('on');
  }
  var el = document.getElementById('in-' + res.first) || document.getElementById('w-' + res.first);
  if (el && el.scrollIntoView) el.scrollIntoView({ behavior:'smooth', block:'center' });
  if (el && el.focus) { try{ el.focus({ preventScroll:true }); }catch(e2){} }
}
function hideStepErr(){
  var e = document.getElementById('err-step');
  if (e) e.classList.remove('on');
}

/* ═══════════ التنقّل ═══════════ */
function next(){
  if (sending) return;
  var q = SCREENS[idx];
  if (q.type === 'review'){ submitAll(); return; }
  if (q.type === 'done') return;
  if (q.type === 'step'){
    var bad = validateStep(q.si);
    if (bad){ showStepErr(bad); return; }
  }
  persist(true);
  navigate(1);
}
function prev(){ if (!sending) navigate(-1); }
function jumpTo(t){
  if (sending || t < 0 || t >= SCREENS.length) return;
  navigate(t - idx);
}

function navigate(dir){
  var ni = idx + dir;
  if (ni < 0 || ni >= SCREENS.length) return;

  var stage = document.getElementById('stage');
  var olds  = Array.prototype.slice.call(stage.children);
  olds.forEach(function(o){ o.classList.remove('active','entering'); o.classList.add('exiting'); });
  var old = olds[olds.length - 1] || null;

  idx = ni;
  var tmp = document.createElement('div');
  tmp.innerHTML = render(SCREENS[idx]);
  var el = tmp.firstElementChild;
  el.classList.add('screen','entering');
  stage.appendChild(el);

  var wrap = stage.parentElement;
  if (old) wrap.style.minHeight = old.offsetHeight + 'px';

  requestAnimationFrame(function(){ requestAnimationFrame(function(){
    el.classList.remove('entering');
    el.classList.add('active');
    setTimeout(function(){
      olds.forEach(function(o){ if (o.parentNode === stage) stage.removeChild(o); });
      wrap.style.minHeight = '';
      focusFirst();
    }, 300);
  }); });

  updateChrome();
  persist(true);
  try{ window.scrollTo({ top:0, behavior:'smooth' }); }catch(e){ window.scrollTo(0,0); }
}

/* أنواع تُكتب بلوحة المفاتيح — وحدها تستحق التركيز التلقائي */
var TYPED = { text:1, tel:1, email:1, number:1, textarea:1 };

function focusFirst(){
  var q = SCREENS[idx];
  if (q.type !== 'step') return;
  if (!window.matchMedia('(min-width:760px)').matches) return;
  var st = F.steps[q.si];
  for (var i = 0; i < st.fields.length; i++){
    var f = st.fields[i];
    if (!visible(f) || !TYPED[f.type]) continue;
    var el = document.getElementById('in-' + f.k);
    if (el){ try{ el.focus({ preventScroll:true }); }catch(e){} return; }
  }
}

/* حقول الخطوة النشطة بالترتيب — لتنقّل Enter بينها */
function stepInputs(){
  var st = document.getElementById('stage');
  var sc = st ? st.querySelector('.screen.active') : null;
  if (!sc) return [];
  return Array.prototype.filter.call(
    sc.querySelectorAll('.field:not(.hidden) input[data-k], .field:not(.hidden) select[data-k]'),
    function(el){ return el.offsetParent !== null; }
  );
}

/* أرقام هندية عربية — لتتّسق لوحة الشاشة مع عناوين الخطوات */
var AR_DIG = ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'];
function ard(n){ return String(n).replace(/\d/g, function(d){ return AR_DIG[+d]; }); }

function updateChrome(){
  document.getElementById('pfill').style.width = (idx / (SCREENS.length - 1) * 100) + '%';
  var q = SCREENS[idx], lbl = '';
  if (q.type === 'step')        lbl = 'خطوة ' + ard(q.si + 1) + ' من ' + ard(F.steps.length);
  else if (q.type === 'review') lbl = 'مراجعة';
  else if (q.type === 'done')   lbl = 'تم';
  document.getElementById('steplbl').textContent = lbl;
  document.body.classList.toggle('wide', q.type === 'review');
}

/* ═══════════ تحديث ظهور الحقول المشروطة ═══════════ */
function refreshVisibility(){
  var q = SCREENS[idx];
  if (q.type !== 'step') return;
  F.steps[q.si].fields.forEach(function(f){
    var el = document.getElementById('f-' + f.k);
    if (el) el.classList.toggle('hidden', !visible(f));
    if (f.other){
      var eo = document.getElementById('f-' + f.k + '_other');
      if (eo) eo.classList.toggle('hidden', !(visible(f) && hasOther(f)));
    }
  });
}

/* ═══════════ معالجات الإدخال ═══════════ */
function onInput(el){
  var k = el.getAttribute('data-k');
  var f = fieldByKey(k);
  if (!f) return;

  if (f.type === 'number'){
    var c = normDigits(el.value).replace(/[^\d]/g,'');
    if (el.value !== c){
      var atEnd = el.selectionStart === el.value.length;
      el.value = c;
      if (atEnd){ try{ el.setSelectionRange(c.length, c.length); }catch(e){} }
    }
    D[k] = c;
  } else if (f.type === 'textarea'){
    D[k] = cleanText(el.value, f.max || MAX_TEXT);
    var c2 = document.getElementById('c-' + k);
    if (c2){
      var mx = f.max || MAX_TEXT;
      c2.textContent = D[k].length + ' / ' + mx;
      c2.classList.toggle('over', D[k].length >= mx);
    }
  } else {
    D[k] = el.value;
  }
  markField(k, '');
  hideStepErr();
  autosave();
}

function onSelect(el){
  var k = el.getAttribute('data-k');
  D[k] = el.value;
  markField(k, '');
  hideStepErr();
  refreshVisibility();
  persist(true);
}

function onOpt(el){
  var k = el.getAttribute('data-opt');
  var val = el.getAttribute('data-val');
  var f = fieldByKey(k);
  if (!f) return;

  if (f.type === 'checks'){
    var a = Array.isArray(D[k]) ? D[k].slice() : [];
    var i = a.indexOf(val);
    if (i > -1) a.splice(i, 1); else a.push(val);
    D[k] = a;
  } else {
    D[k] = (D[k] === val && f.req === false) ? '' : val;
  }

  var box = document.getElementById('w-' + k);
  if (box){
    var sel = Array.isArray(D[k]) ? D[k] : [D[k]];
    Array.prototype.forEach.call(box.querySelectorAll('.opt'), function(o){
      var on = sel.indexOf(o.getAttribute('data-val')) > -1;
      o.classList.toggle('on', on);
      o.setAttribute('aria-checked', on ? 'true' : 'false');
    });
  }
  markField(k, '');
  hideStepErr();
  refreshVisibility();
  persist(true);
}

function onConsent(el){
  var k = el.getAttribute('data-consent');
  D[k] = !D[k];
  el.classList.toggle('on', !!D[k]);
  var c = el.querySelector('.chk');
  if (c){ c.classList.toggle('on', !!D[k]); c.setAttribute('aria-checked', D[k] ? 'true' : 'false'); }
  el.classList.remove('bad');
  markField(k, '');
  hideStepErr();
  persist(true);
}

/* ═══════════ بناء الأعمدة والقيم ═══════════ */
function buildHeaders(){
  var h = ['وقت الإرسال','الرقم المرجعي'];
  allFields().forEach(function(f){ h.push(cleanText(f.col || f.t, 200)); });
  return h;
}
function buildValues(){
  var v = ['', D._ref];
  allFields().forEach(function(f){
    if (!visible(f)){ v.push(''); return; }
    if (f.type === 'consent'){ v.push(D[f.k] ? 'نعم' : 'لا'); return; }
    if (f.type === 'number'){
      var n = parseInt(normDigits(D[f.k] || ''), 10);
      v.push(isFinite(n) ? n : '');
      return;
    }
    var raw = Array.isArray(D[f.k]) ? D[f.k].join('، ') : (D[f.k] == null ? '' : String(D[f.k]));
    v.push(safeCell(cleanText(raw, f.max || MAX_TEXT).trim()));
  });
  return v;
}

/* فحص نهائي شامل — لا نرسل بيانات ناقصة */
function finalCheck(){
  var bad = [];
  F.steps.forEach(function(st, si){
    var hit = false;
    st.fields.forEach(function(f){
      if (!visible(f)) return;
      if (!checkField(f).ok) hit = true;
      if (f.other && hasOther(f) && !cleanText(D[f.k + '_other'] || '', 120).trim()) hit = true;
    });
    if (hit) bad.push({ title:st.title, i:si + 1 });
  });
  return bad;
}

/* ═══════════ الإرسال ═══════════ */
function postJSON(url, obj, ms){
  var ctl = new AbortController();
  var timer = setTimeout(function(){ ctl.abort(); }, ms);
  /* text/plain يجعل الطلب "بسيطاً" فلا يحتاج preflight من CORS */
  return fetch(url, {
    method:'POST',
    headers:{ 'Content-Type':'text/plain;charset=utf-8' },
    body:JSON.stringify(obj),
    redirect:'follow',
    signal:ctl.signal
  }).then(function(res){
    clearTimeout(timer);
    return res.text().then(function(txt){
      if (!res.ok) throw new Error('رفض الخادم الطلب (HTTP ' + res.status + ').');
      var data;
      try{ data = JSON.parse(txt); }
      catch(e){ throw new Error('ردّ غير مفهوم من الخادم — تأكّد من رابط النشر وأن صلاحية الوصول «لأي شخص».'); }
      return data;
    });
  }, function(err){
    clearTimeout(timer);
    if (err && err.name === 'AbortError') throw new Error('انتهت مدة الانتظار. تحقّق من الاتصال بالإنترنت.');
    throw new Error('تعذّر الوصول إلى الخادم. تحقّق من الاتصال بالإنترنت ثم أعد المحاولة.');
  });
}

function setSending(on, txt){
  sending = on;
  var b = document.getElementById('btn-send'), bk = document.getElementById('btn-back-rev');
  if (b){
    b.disabled = on;
    b.innerHTML = on ? '<span class="spin"></span> ' + esc(txt || 'جارٍ الإرسال…') : 'إرسال النموذج';
  }
  if (bk) bk.disabled = on;
  var eds = document.querySelectorAll('.rev-edit');
  for (var i = 0; i < eds.length; i++) eds[i].disabled = on;
}
function sendErr(msg){
  var e = document.getElementById('err-send');
  if (e){ e.innerHTML = msg; e.classList.add('on'); }
}

function submitAll(){
  if (sending) return;
  var e = document.getElementById('err-send');
  if (e) e.classList.remove('on');

  var bad = finalCheck();
  if (bad.length){
    sendErr('لا يمكن الإرسال — توجد حقول ناقصة أو غير صالحة في: '
          + esc(bad.map(function(b){ return b.title; }).join('، '))
          + '. اضغط «تعديل» بجانب القسم وأكمله.');
    return;
  }
  if (!F.scriptUrl || F.scriptUrl.indexOf('/exec') === -1){
    sendErr('الربط بجدول البيانات غير مُهيَّأ بعد. <b>لم تُرسل البيانات</b> — إجاباتك محفوظة على جهازك ولن تُفقد.');
    return;
  }

  var payload = {
    secret:  F.secret,
    form:    F.id,
    sheet:   F.sheet,
    ref:     D._ref,
    headers: buildHeaders(),
    values:  buildValues()
  };

  setSending(true);
  var lastErr = 'تعذّر الإرسال.', attempt = 0;

  function attemptSend(){
    attempt++;
    if (attempt > 1) setSending(true, 'إعادة المحاولة ' + attempt + ' من 3…');
    return postJSON(F.scriptUrl, payload, 30000).then(function(res){
      if (res && res.ok === true){
        recordSent();
        D.submitted = true;
        D.row = res.row || '';
        persist(true);
        setSending(false);
        navigate(DONE_IDX - idx);
        return true;
      }
      lastErr = (res && res.error) ? String(res.error) : 'ردّ غير متوقّع من الخادم.';
      if (res && res.fatal === true) return false;
      return retry();
    }, function(err){
      lastErr = (err && err.message) ? err.message : String(err);
      return retry();
    });
  }
  function retry(){
    if (attempt >= 3) return false;
    return sleep(1200 * attempt).then(attemptSend);
  }

  attemptSend().then(function(done){
    if (done) return;
    setSending(false);
    sendErr(esc(lastErr) + '<br>إجاباتك محفوظة على جهازك ولم تُفقد — اضغط «إرسال النموذج» للمحاولة مرّة أخرى.');
  });
}

/* ═══════════ استرجاع طارئ — بلا واجهة ═══════════
   إن تعذّر وصول استجابة إلى الخادم: افتح Console واكتب downloadBackup() */
window.downloadBackup = function(){
  try{
    var h = buildHeaders(), v = buildValues(), ans = {};
    h.forEach(function(k, i){
      var x = v[i];
      if (i === 0) x = D._at || '';
      if (typeof x === 'string' && x.charAt(0) === "'") x = x.substring(1);
      ans[k] = x;
    });
    var obj = { form:F.id, ref:D._ref, generatedAt:new Date().toISOString(),
                submittedToServer:!!D.submitted, sheetRow:D.row || null, answers:ans };
    var blob = new Blob([JSON.stringify(obj, null, 2)], { type:'application/json;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = F.id + '-' + (D._ref || 'response') + '.json';
    document.body.appendChild(a); a.click();
    setTimeout(function(){ if (a.parentNode) a.parentNode.removeChild(a); URL.revokeObjectURL(url); }, 1500);
    flashSaved('تم تنزيل النسخة');
  }catch(e){ alert('تعذّر تنزيل النسخة على هذا المتصفح.'); }
};

/* ═══════════ المظهر ═══════════ */
function toggleTheme(){
  var cur = document.documentElement.getAttribute('data-theme');
  var nxt = (cur === 'dark') ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', nxt);
  try{ localStorage.setItem('yaamur_theme', nxt); }catch(e){}
}

/* ═══════════ الأحداث ═══════════ */
var ACTS = {
  next: next, prev: prev, send: submitAll, reset: resetAll,
  theme: toggleTheme, 'new': wipeAndReload,
  jump: function(el){ jumpTo(parseInt(el.getAttribute('data-to'), 10)); }
};

document.addEventListener('click', function(ev){
  var t = ev.target;
  var opt = t.closest ? t.closest('[data-opt]') : null;
  if (opt){ onOpt(opt); return; }
  var cns = t.closest ? t.closest('[data-consent]') : null;
  if (cns){ ev.preventDefault(); onConsent(cns); return; }
  var act = t.closest ? t.closest('[data-act]') : null;
  if (act){
    var fn = ACTS[act.getAttribute('data-act')];
    if (fn) { ev.preventDefault(); fn(act); }
  }
});

document.addEventListener('input', function(ev){
  var t = ev.target;
  if (!t || !t.getAttribute) return;
  if (t.tagName === 'SELECT') return;
  if (t.getAttribute('data-k')) onInput(t);
});
document.addEventListener('change', function(ev){
  var t = ev.target;
  if (t && t.tagName === 'SELECT' && t.getAttribute('data-k')) onSelect(t);
});
document.addEventListener('blur', function(ev){
  var t = ev.target;
  if (!t || !t.getAttribute || !t.getAttribute('data-k')) return;
  var f = fieldByKey(t.getAttribute('data-k'));
  if (!f || !f.vald || !String(D[f.k] || '').trim()) return;
  var r = checkField(f);
  markField(f.k, r.ok ? '' : r.msg);
  if (r.ok && r.val !== undefined){ D[f.k] = r.val; t.value = r.val; persist(true); }
}, true);

/* المسافة والإدخال يفعّلان الخيارات عند التنقّل بلوحة المفاتيح */
document.addEventListener('keydown', function(ev){
  var t = ev.target;
  if (ev.key === ' ' || ev.key === 'Spacebar'){
    if (t && t.closest){
      var o = t.closest('[data-opt]'); if (o){ ev.preventDefault(); onOpt(o); return; }
      var c = t.closest('[data-consent]'); if (c){ ev.preventDefault(); onConsent(c); return; }
    }
  }
  if (ev.key !== 'Enter' || ev.shiftKey) return;
  if (t && (t.tagName === 'TEXTAREA' || t.tagName === 'BUTTON')) return;
  if (t && t.closest){
    var o2 = t.closest('[data-opt]'); if (o2){ ev.preventDefault(); onOpt(o2); return; }
    var c2 = t.closest('[data-consent]'); if (c2){ ev.preventDefault(); onConsent(c2); return; }
  }
  ev.preventDefault();
  if (SCREENS[idx] && SCREENS[idx].type === 'step' && t &&
      (t.tagName === 'INPUT' || t.tagName === 'SELECT')){
    var list = stepInputs();
    var i = list.indexOf(t);
    if (i > -1 && i < list.length - 1){ list[i + 1].focus(); return; }
  }
  next();
});

window.addEventListener('beforeunload', function(){ clearTimeout(saveTimer); persist(true); });
document.addEventListener('visibilitychange', function(){
  if (document.hidden){ clearTimeout(saveTimer); persist(true); }
});

/* ═══════════ الإقلاع ═══════════ */
(function boot(){
  try{
    var th = localStorage.getItem('yaamur_theme');
    if (th) document.documentElement.setAttribute('data-theme', th);
  }catch(e){}

  /* الاستئناف من آخر موضع — لا نستأنف على شاشة «تم» بعد إرسال مكتمل */
  var start = 0;
  if (typeof D._idx === 'number' && D._idx > 0 && D._idx < DONE_IDX && !D.submitted) start = D._idx;
  idx = 0;

  document.title = F.pageTitle || F.title;
  var stage = document.getElementById('stage');
  var tmp = document.createElement('div');
  tmp.innerHTML = render(SCREENS[0]);
  var el = tmp.firstElementChild;
  el.classList.add('screen','active');
  stage.appendChild(el);
  updateChrome();

  if (start > 0) setTimeout(function(){ jumpTo(start); }, 60);
})();

})();
