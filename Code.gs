/**
 * ระบบลงทะเบียนเลือกวิชาเรียนเสรี - Google Apps Script (Web App)
 * ทำงานคู่กับ index.html
 *
 * GET  -> { status: "success", students: [...] }
 * POST -> body เป็น JSON (ส่งแบบ text/plain) มี action = register | update | delete
 */

// ถ้าสร้างสคริปต์จากเมนู Extensions > Apps Script ใน Google Sheet ให้ปล่อยว่างไว้
// ถ้าสร้างสคริปต์แยก ให้ใส่ ID ของ Google Sheet (ข้อความยาว ๆ ใน URL ระหว่าง /d/ และ /edit)
const SPREADSHEET_ID = '';

const SHEET_NAME = 'Registrations';
const MAX_QUOTA = 35; // โควตาต่อวิชา ต้องตรงกับ maxQuota ใน index.html

// ลำดับคอลัมน์ในชีต: id, name, class, email, courseCode, timestamp
const HEADERS = ['รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ห้องเรียน', 'อีเมล', 'รหัสวิชา', 'เวลาลงทะเบียน'];

// รหัสวิชาที่อนุญาต ต้องตรงกับ INITIAL_COURSES ใน index.html
const VALID_COURSES = [
  'ค30294', 'ค30295', 'ท30208', 'ท30210', 'อ30209', 'ญ30209', 'จ30207',
  'ว30201', 'ว30296', 'ว30202', 'ส30235', 'ส30257', 'ง30208', 'ง30227',
  'ง30272', 'ง30283', 'ง30278', 'ง30281', 'ศ30239', 'ศ30240', 'พ30202'
];

/* ---------- Entry points ---------- */

function doGet(e) {
  try {
    // ใช้ข้อมูลจากแคชถ้ามี (ตอบเร็วขึ้นมากเมื่อมีคนเปิดเว็บพร้อมกันหลายคน)
    const cached = cacheGet_();
    if (cached) {
      return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
    }
    const out = JSON.stringify({ status: 'success', students: readAll_() });
    cachePut_(out);
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // กันหลายคนกดลงทะเบียนพร้อมกันแล้วโควตาเกิน
  } catch (err) {
    // code: 'BUSY' ทำให้หน้าเว็บลองส่งซ้ำให้อัตโนมัติ
    return json_({ status: 'error', code: 'BUSY', message: 'ระบบกำลังมีผู้ใช้งานจำนวนมาก กรุณาลองใหม่อีกครั้ง' });
  }

  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let result;
    switch (body.action) {
      case 'register': result = register_(body); break;
      case 'update':   result = update_(body);   break;
      case 'delete':   result = delete_(body);   break;
      default:         result = { status: 'error', message: 'ไม่รู้จักคำสั่ง: ' + body.action };
    }
    if (result.status === 'success') cacheClear_(); // ข้อมูลเปลี่ยนแล้ว ล้างแคชเก่า
    return json_(result);
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Actions ---------- */

function register_(d) {
  const s = clean_(d);
  const err = validate_(s);
  if (err) return { status: 'error', message: err };

  const rows = readAll_();
  if (rows.some(r => r.id === s.id)) {
    return { status: 'error', message: 'รหัสนักเรียน ' + s.id + ' นี้เคยลงทะเบียนเรียบร้อยแล้ว' };
  }
  if (countInCourse_(rows, s.courseCode) >= MAX_QUOTA) {
    return { status: 'error', message: 'ขออภัย รายวิชานี้ที่นั่งเต็มโควตา (' + MAX_QUOTA + ' คน) แล้ว' };
  }

  const sheet = getSheet_();
  const row = sheet.getLastRow() + 1;
  writeRow_(sheet, row, s);
  return { status: 'success' };
}

function update_(d) {
  const s = clean_(d);
  const origId = String(d.origId || '').trim();
  const err = validate_(s);
  if (err) return { status: 'error', message: err };

  const sheet = getSheet_();
  const rows = readAll_();
  const idx = rows.findIndex(r => r.id === origId);
  if (idx === -1) return { status: 'error', message: 'ไม่พบข้อมูลรหัส ' + origId };

  if (s.id !== origId && rows.some(r => r.id === s.id)) {
    return { status: 'error', message: 'รหัสนักเรียน ' + s.id + ' ซ้ำในระบบ' };
  }
  if (rows[idx].courseCode !== s.courseCode && countInCourse_(rows, s.courseCode) >= MAX_QUOTA) {
    return { status: 'error', message: 'รายวิชานี้เต็มโควตา (' + MAX_QUOTA + ' คน) แล้ว' };
  }

  s.timestamp = rows[idx].timestamp; // คงเวลาลงทะเบียนเดิมไว้
  writeRow_(sheet, idx + 2, s);      // +2 เพราะแถวที่ 1 เป็นหัวตาราง
  return { status: 'success' };
}

function delete_(d) {
  const id = String(d.id || '').trim();
  const rows = readAll_();
  const idx = rows.findIndex(r => r.id === id);
  if (idx === -1) return { status: 'error', message: 'ไม่พบข้อมูลรหัส ' + id };

  getSheet_().deleteRow(idx + 2);
  return { status: 'success' };
}

/* ---------- Helpers ---------- */

let _sheet = null; // เก็บไว้ใช้ซ้ำภายในการรันครั้งเดียว

function getSheet_() {
  if (_sheet) return _sheet;
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // ตั้งทุกคอลัมน์เป็นข้อความ กันรหัสที่ขึ้นต้นด้วย 0 หาย (ทำครั้งเดียวตอนสร้างชีต ไม่ทำทุกคำขอเพราะช้า)
    sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setNumberFormat('@');
  }
  _sheet = sheet;
  return sheet;
}

function readAll_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      id: String(r[0]).trim(),
      name: String(r[1]),
      class: String(r[2]),
      email: String(r[3]),
      courseCode: String(r[4]),
      timestamp: String(r[5])
    }));
}

function writeRow_(sheet, row, s) {
  sheet.getRange(row, 1, 1, HEADERS.length)
    .setNumberFormat('@')
    .setValues([[s.id, s.name, s.class, s.email, s.courseCode, s.timestamp]]);
}

function clean_(d) {
  const id = String(d.id || '').trim();
  return {
    id: id,
    name: String(d.name || '').trim(),
    class: String(d.class || '').trim(),
    email: 'br' + id + '@br.ac.th', // สร้างอีเมลฝั่งเซิร์ฟเวอร์เอง ไม่เชื่อค่าจากหน้าเว็บ
    courseCode: String(d.courseCode || '').trim(),
    timestamp: String(d.timestamp || new Date().toLocaleString('th-TH'))
  };
}

function validate_(s) {
  if (!/^\d{5}$/.test(s.id)) return 'รหัสนักเรียนต้องเป็นตัวเลข 5 หลัก';
  if (!s.name) return 'กรุณากรอกชื่อ-นามสกุล';
  if (!/^ม\.[456]\/\d+$/.test(s.class)) return 'ห้องเรียนไม่ถูกต้อง';
  if (VALID_COURSES.indexOf(s.courseCode) === -1) return 'รหัสวิชาไม่ถูกต้อง';
  return '';
}

function countInCourse_(rows, code) {
  return rows.filter(r => r.courseCode === code).length;
}

/* ---------- Cache (เก็บคำตอบของ doGet ชั่วคราว) ---------- */

const CACHE_KEY = 'students_json_v1';
const CACHE_TTL = 30;     // วินาที
const CACHE_CHUNK = 25000; // ตัวอักษรต่อชิ้น (ขีดจำกัดของ CacheService คือ 100KB ต่อค่า และภาษาไทยใช้ 3 ไบต์ต่อตัว)

function cacheGet_() {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(CACHE_KEY + ':n') || 0);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(CACHE_KEY + ':' + i);
  const parts = cache.getAll(keys);
  let out = '';
  for (let i = 0; i < n; i++) {
    const part = parts[CACHE_KEY + ':' + i];
    if (part == null) return null;
    out += part;
  }
  return out;
}

function cachePut_(str) {
  const cache = CacheService.getScriptCache();
  const n = Math.ceil(str.length / CACHE_CHUNK);
  const obj = {};
  obj[CACHE_KEY + ':n'] = String(n);
  for (let i = 0; i < n; i++) obj[CACHE_KEY + ':' + i] = str.substr(i * CACHE_CHUNK, CACHE_CHUNK);
  cache.putAll(obj, CACHE_TTL);
}

function cacheClear_() {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(CACHE_KEY + ':n') || 0);
  const keys = [CACHE_KEY + ':n'];
  for (let i = 0; i < n; i++) keys.push(CACHE_KEY + ':' + i);
  cache.removeAll(keys);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งเดียวจากเมนู Run เพื่อสร้างชีตและขออนุญาตเข้าถึง Spreadsheet */
function setup() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setNumberFormat('@');
}
