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
    return json_({ status: 'success', students: readAll_() });
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
    switch (body.action) {
      case 'register': return json_(register_(body));
      case 'update':   return json_(update_(body));
      case 'delete':   return json_(delete_(body));
      default:         return json_({ status: 'error', message: 'ไม่รู้จักคำสั่ง: ' + body.action });
    }
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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งเดียวจากเมนู Run เพื่อสร้างชีตและขออนุญาตเข้าถึง Spreadsheet */
function setup() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setNumberFormat('@');
}
