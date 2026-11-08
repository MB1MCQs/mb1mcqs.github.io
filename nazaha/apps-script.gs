/**
 * مسابقة نزاهة — حفظ النتائج في جدول Google
 * الصقي هذا الكود في: الإضافات ← Apps Script، ثم انشريه كتطبيق ويب.
 */
const SHEET_NAME = 'المشاركات';
const HEADERS = ['الوقت', 'الصفة', 'الاسم', 'اسم الابنة', 'الصف', 'الشعبة', 'الدرجة', 'من', 'المعرّف'];
const GRADES = ['أول ثانوي', 'ثاني ثانوي', 'ثالث ثانوي'];
const TOTAL = 15;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// يمنع تحويل النص إلى معادلة داخل الجدول
function clean_(v, max) {
  let s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doGet() {
  const values = sheet_().getDataRange().getValues().slice(1);
  const entries = values.filter(r => r[2]).map(r => ({
    at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    role: r[1] === 'ولية أمر' ? 'parent' : 'student',
    name: String(r[2]).replace(/^'/, ''),
    child: String(r[3] || '').replace(/^'/, ''),
    grade: String(r[4]),
    section: String(r[5]),
    score: Number(r[6]),
    total: Number(r[7]) || TOTAL
  }));
  return json_({ entries: entries });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const role = d.role === 'parent' ? 'parent' : 'student';
  const name = clean_(d.name, 60);
  const child = role === 'parent' ? clean_(d.child, 60) : '';
  const grade = String(d.grade || '');
  const section = parseInt(d.section, 10);
  const score = parseInt(d.score, 10);
  const id = clean_(d.id, 40);

  if (name.length < 3 || !id) return json_({ ok: false, error: 'invalid' });
  if (GRADES.indexOf(grade) < 0) return json_({ ok: false, error: 'invalid' });
  if (!(section >= 1 && section <= 8)) return json_({ ok: false, error: 'invalid' });
  if (!(score >= 0 && score <= TOTAL)) return json_({ ok: false, error: 'invalid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 9, last - 1, 1).getValues().map(r => String(r[0]));
      if (ids.indexOf(id) >= 0) return json_({ ok: false, error: 'duplicate' });
    }
    sh.appendRow([new Date(), role === 'parent' ? 'ولية أمر' : 'طالبة', name, child, grade, section, score, TOTAL, id]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
