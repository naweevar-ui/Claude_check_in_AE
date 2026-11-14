/**
 * ระบบเช็คชื่อเข้าอบรม — ฝั่ง Google Sheet (Google Apps Script)
 * เวอร์ชัน 1.7.1 (14/11/2026)
 *
 * วิธีใช้โดยย่อ (รายละเอียดอยู่ใน README.md)
 * 1. เปิด Google Sheet > ส่วนขยาย > Apps Script แล้ววางโค้ดนี้แทนของเดิม
 * 2. เลือกฟังก์ชัน setup แล้วกด Run หนึ่งครั้ง (อนุญาตสิทธิ์ตามที่ Google ถาม)
 * 3. Deploy > New deployment > Web app
 *    Execute as: Me / Who has access: Anyone
 * 4. คัดลอก URL ที่ลงท้ายด้วย /exec ไปใส่ใน config.js
 */

const TZ = 'Asia/Bangkok';
const SH_PEOPLE = 'รายชื่อ';
const SH_LOG = 'บันทึกเช็คชื่อ';
const SH_SET = 'ตั้งค่า';
const VERSION = '1.7.1';
const SH_DATA = 'Data';   // แท็บสรุปแบบง่าย: timestamp | name | note (มา / สาย)
const PEOPLE_HEAD = ['รหัส', 'ชื่อ-นามสกุล', 'บริษัท', 'โทร', 'อีเมล', 'LINE', 'ติดต่อผ่าน'];
const LOG_HEAD = ['วันที่', 'รหัส', 'ชื่อ-นามสกุล', 'เวลาเข้า', 'สถานะ', 'สาย (นาที)',
  'วิธียืนยัน', 'ลิงก์ลายเซ็น/รูป', 'ตัวอย่างภาพ', 'บันทึกเมื่อ', 'timestamp'];
const DEFAULT_START = '12:00';
const FOLDER_NAME = 'ลายเซ็นและรูปเช็คชื่อ';

/* ---------- ตั้งค่าครั้งแรก ---------- */
function setup() {
  const ppl = sheet_(SH_PEOPLE, PEOPLE_HEAD);
  ppl.getRange(1, 1, 1, PEOPLE_HEAD.length).setValues([PEOPLE_HEAD]).setFontWeight('bold');
  ppl.getRange('D:D').setNumberFormat('@');   // เก็บเบอร์โทรเป็นข้อความ ไม่ให้เลข 0 หาย
  const log = sheet_(SH_LOG, LOG_HEAD);
  log.hideColumns(LOG_HEAD.length);
  const set = sheet_(SH_SET, null);
  if (!set.getRange('A1').getValue()) {
    set.getRange('A1:B1').setValues([['เวลาเริ่มเรียน', DEFAULT_START]]);
    set.getRange('B1').setNumberFormat('@').setValue(DEFAULT_START);
    set.getRange('A1').setFontWeight('bold');
  }
  folder_();
  SpreadsheetApp.getActive().toast('ตั้งค่าเรียบร้อย ใส่รายชื่อในชีต "รายชื่อ" แล้ว Deploy เป็น Web app ได้เลย');
}

/* ---------- รับคำขอจากหน้าเว็บ ---------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'state') {
      const date = checkDate_(p.date);
      return json_({ ok: true, version: VERSION, start: getStart_(), people: readPeople_(), checks: readChecks_(date) });
    }
    return json_({ ok: true, version: VERSION, message: 'ระบบเช็คชื่อพร้อมใช้งาน' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      case 'checkin': return json_(checkin_(body));
      case 'uncheck': return json_(uncheck_(body));
      case 'addPeople': return json_(addPeople_(body));
      case 'removePerson': return json_(removePerson_(body));
      case 'setStart': return json_(setStart_(body));
      default: return json_({ ok: false, error: 'ไม่รู้จักคำสั่งนี้' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ---------- เช็คชื่อ ---------- */
function checkin_(b) {
  const date = checkDate_(b.date);
  const person = readPeople_().find(function (p) { return p.id === String(b.pid); });
  if (!person) throw new Error('ไม่พบรายชื่อนี้');
  const kind = b.kind === 'photo' ? 'photo' : 'sig';
  const now = new Date();                       // ใช้เวลาของเซิร์ฟเวอร์ Google เป็นเวลาเข้าเรียน
  const at = now.getTime();
  const late = lateMinutes_(now);

  let viewUrl = '', thumb = '';
  if (typeof b.img === 'string' && /^data:image\/(png|jpeg|webp);base64,/.test(b.img)) {
    if (b.img.length > 3000000) throw new Error('ไฟล์ภาพใหญ่เกินไป');
    const mime = b.img.slice(5, b.img.indexOf(';'));
    const bytes = Utilities.base64Decode(b.img.split(',')[1]);
    const ext = mime === 'image/png' ? 'png' : 'jpg';
    const name = date + '_' + person.name + '_' + (kind === 'sig' ? 'ลายเซ็น' : 'รูป') + '.' + ext;
    const file = folder_().createFile(Utilities.newBlob(bytes, mime, name));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {
      // บางองค์กรปิดการแชร์ด้วยลิงก์ ภาพยังถูกเก็บไว้ใน Drive แต่คนอื่นอาจมองไม่เห็นภาพตัวอย่าง
    }
    viewUrl = file.getUrl();
    thumb = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w600';
  }

  removeLogRows_(date, person.id);
  const log = sheet_(SH_LOG, LOG_HEAD);
  log.appendRow([
    "'" + date, person.id, person.name,
    "'" + Utilities.formatDate(now, TZ, 'HH:mm:ss'),
    late > 0 ? 'มาสาย' : 'ตรงเวลา', late,
    kind === 'sig' ? 'เซ็นชื่อ' : 'ถ่ายรูป',
    viewUrl, thumb ? '=IMAGE("' + thumb + '")' : '',
    "'" + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), at
  ]);
  if (thumb) log.setRowHeight(log.getLastRow(), 60);
  // บันทึกสรุปลงแท็บ Data ด้วย (ถ้ามีแท็บนี้ในชีต)
  const data = SpreadsheetApp.getActive().getSheetByName(SH_DATA);
  if (data) {
    data.appendRow(["'" + Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm'), person.name, late > 0 ? 'สาย' : 'มา']);
  }
  return { ok: true, check: { pid: person.id, at: at, kind: kind, img: thumb } };
}

function uncheck_(b) {
  const date = checkDate_(b.date);
  removeLogRows_(date, String(b.pid));
  return { ok: true };
}

/* ---------- รายชื่อ ---------- */
function addPeople_(b) {
  const clean = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); };
  const list = (b.people || (b.names || []).map(function (n) { return { name: n }; }))
    .map(function (v) { return typeof v === 'string' ? { name: v } : v; })
    .map(function (v) {
      return { name: clean(v.name), company: clean(v.company), phone: clean(v.phone),
        email: clean(v.email), line: clean(v.line), via: clean(v.via) };
    })
    .filter(function (v) { return v.name && v.name.length <= 200; });
  const sh = sheet_(SH_PEOPLE, PEOPLE_HEAD);
  readPeople_();                                   // เติมรหัสให้แถวที่ยังไม่มี
  const vals = sh.getDataRange().getValues();
  const rowOf = {};
  for (let r = 1; r < vals.length; r++) {
    const n = clean(vals[r][1]);
    if (n) rowOf[n] = r + 1;
  }
  const keys = ['company', 'phone', 'email', 'line', 'via'];
  const ids = []; let added = 0, updated = 0;
  list.forEach(function (v) {
    const row = rowOf[v.name];
    if (row) {
      const cur = sh.getRange(row, 1, 1, PEOPLE_HEAD.length).getValues()[0];
      let changed = false;
      keys.forEach(function (k, i) { if (v[k] && v[k] !== clean(cur[i + 2])) { cur[i + 2] = v[k]; changed = true; } });
      if (changed) { sh.getRange(row, 1, 1, PEOPLE_HEAD.length).setValues([cur]); updated++; }
      ids.push(String(cur[0]));
      return;
    }
    const id = newId_();
    sh.appendRow([id, v.name, v.company, v.phone, v.email, v.line, v.via]);
    rowOf[v.name] = sh.getLastRow();
    ids.push(id); added++;
  });
  return { ok: true, added: added, updated: updated, ids: ids, people: readPeople_() };
}

function removePerson_(b) {
  const id = String(b.id || '');
  const sh = sheet_(SH_PEOPLE, PEOPLE_HEAD);
  const vals = sh.getDataRange().getValues();
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][0]) === id) sh.deleteRow(r + 1);
  }
  return { ok: true };
}

function readPeople_() {
  const sh = sheet_(SH_PEOPLE, PEOPLE_HEAD);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rng = sh.getRange(2, 1, last - 1, PEOPLE_HEAD.length);
  const vals = rng.getValues();
  let changed = false;
  const out = [];
  vals.forEach(function (row) {
    const name = String(row[1] || '').replace(/\s+/g, ' ').trim();
    if (!name) return;
    if (!row[0]) { row[0] = newId_(); changed = true; }   // เติมรหัสให้ชื่อที่พิมพ์ลงชีตเอง
    const t = function (v) { return String(v == null ? '' : v).trim(); };
    out.push({ id: String(row[0]), name: name, company: t(row[2]), phone: t(row[3]),
      email: t(row[4]), line: t(row[5]), via: t(row[6]) });
  });
  if (changed) rng.setValues(vals);
  return out;
}

/* ---------- เวลาเริ่มเรียน ---------- */
function setStart_(b) {
  const v = String(b.start || '');
  if (!/^\d{2}:\d{2}$/.test(v)) throw new Error('รูปแบบเวลาไม่ถูกต้อง');
  const set = sheet_(SH_SET, null);
  set.getRange('A1').setValue('เวลาเริ่มเรียน');
  set.getRange('B1').setNumberFormat('@').setValue(v);
  return { ok: true, start: v };
}

function getStart_() {
  const v = sheet_(SH_SET, null).getRange('B1').getValue();
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})[:.](\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : DEFAULT_START;
}

function lateMinutes_(d) {
  const parts = getStart_().split(':').map(Number);
  const now = Utilities.formatDate(d, TZ, 'HH:mm').split(':').map(Number);
  return Math.max(0, (now[0] * 60 + now[1]) - (parts[0] * 60 + parts[1]));
}

/* ---------- อ่านบันทึก ---------- */
function readChecks_(date) {
  const sh = sheet_(SH_LOG, LOG_HEAD);
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  const vals = sh.getRange(2, 1, last - 1, LOG_HEAD.length).getValues();
  const formulas = sh.getRange(2, 9, last - 1, 1).getFormulas();
  vals.forEach(function (row, i) {
    if (cellDate_(row[0]) !== date) return;
    const m = String(formulas[i][0] || '').match(/IMAGE\("([^"]+)"\)/);
    out[String(row[1])] = {
      pid: String(row[1]),
      at: Number(row[10]) || 0,
      kind: row[6] === 'ถ่ายรูป' ? 'photo' : 'sig',
      img: m ? m[1] : ''
    };
  });
  return out;
}

function removeLogRows_(date, pid) {
  const sh = sheet_(SH_LOG, LOG_HEAD);
  const vals = sh.getDataRange().getValues();
  for (let r = vals.length - 1; r >= 1; r--) {
    if (cellDate_(vals[r][0]) === date && String(vals[r][1]) === pid) sh.deleteRow(r + 1);
  }
}

/* ---------- ตัวช่วย ---------- */
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (_) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function cellDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

function checkDate_(d) {
  const s = String(d || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('วันที่ไม่ถูกต้อง');
  return s;
}

function newId_() {
  return 'p' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
