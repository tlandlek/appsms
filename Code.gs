/**
 * ============================================================
 *  ระบบส่ง SMS แจ้งการมาเรียนผ่าน Thaibulksms
 *  Google Apps Script (Backend)
 * ============================================================
 *
 *  วิธีติดตั้ง:
 *  1) สร้าง Google Sheets ตั้งชื่อ sheet ว่า "students"
 *     คอลัมน์: A=studentId  B=name  C=parentPhone  D=parentName
 *
 *  2) Extensions → Apps Script → วางโค้ดนี้ทั้งหมด → Save
 *
 *  3) Project Settings → Script Properties เพิ่ม 3 ค่า:
 *        THAIBULKSMS_KEY     = api key
 *        THAIBULKSMS_SECRET  = api secret
 *        THAIBULKSMS_SENDER  = sender name (เช่น "School")
 *
 *  4) Deploy → New deployment → Web app
 *        Execute as: Me  |  Who has access: Anyone
 *     Copy URL ไปวางในหน้าเว็บ ⚙️ ตั้งค่า
 *
 *  ตัวแปรในข้อความกำหนดเอง:
 *     {name} {id} {date} {time} {school} {parentName}
 * ============================================================
 */

const SHEET_STUDENTS = 'students';
const SHEET_LOG      = 'log';
const SHEET_ATTEND   = 'attendance';

// ─── Router ───────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'ping')           return jsonOut({ ok: true, studentCount: countStudents() });
    if (action === 'send')           return jsonOut(handleSend(body));
    if (action === 'getStudents')    return jsonOut(handleGetStudents());
    if (action === 'saveStudents')   return jsonOut(handleSaveStudents(body));
    if (action === 'addStudent')     return jsonOut(handleAddStudent(body));
    if (action === 'updateStudent')  return jsonOut(handleUpdateStudent(body));
    if (action === 'deleteStudent')  return jsonOut(handleDeleteStudent(body));
    if (action === 'getLogs')        return jsonOut(handleGetLogs(body));
    if (action === 'saveAttendance') return jsonOut(handleSaveAttendance(body));
    if (action === 'getAttendance')  return jsonOut(handleGetAttendance(body));

    return jsonOut({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err.message || err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Sheet helpers ────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name || SHEET_STUDENTS);
  if (!sh) throw new Error(`ไม่พบ sheet "${name || SHEET_STUDENTS}"`);
  return sh;
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function countStudents() {
  const sh = getSheet(SHEET_STUDENTS);
  return Math.max(0, sh.getLastRow() - 1);
}

// ─── Find student by ID ───────────────────────────────────────
function findStudent(studentId) {
  const sh   = getSheet(SHEET_STUDENTS);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('ยังไม่มีข้อมูลนักเรียนในชีต');
  const rows   = sh.getRange(2, 1, last - 1, 4).getValues();
  const target = String(studentId).trim();
  for (const r of rows) {
    if (String(r[0]).trim() === target) {
      return {
        studentId:   String(r[0]).trim(),
        name:        String(r[1] || '').trim(),
        parentPhone: normalizePhone(r[2]),
        parentName:  String(r[3] || '').trim(),
      };
    }
  }
  throw new Error(`ไม่พบรหัสนักเรียน ${target}`);
}

// Returns 1-based row number or -1
function findStudentRow(studentId) {
  const sh   = getSheet(SHEET_STUDENTS);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids    = sh.getRange(2, 1, last - 1, 1).getValues();
  const target = String(studentId).trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 2;
  }
  return -1;
}

function normalizePhone(raw) {
  let p = String(raw || '').replace(/\D/g, '');
  if (!p) throw new Error('เบอร์ผู้ปกครองว่าง');
  if (p.startsWith('66')) p = '0' + p.slice(2);
  if (p.length !== 10 || !p.startsWith('0')) {
    throw new Error(`รูปแบบเบอร์ไม่ถูกต้อง: ${raw}`);
  }
  return p;
}

// ─── Message builder ──────────────────────────────────────────
function buildMessage(student, statusType, schoolName, customTemplate) {
  const today  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');
  const time   = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  const school = schoolName || 'โรงเรียน';
  const name   = student.name || `รหัส ${student.studentId}`;

  if (statusType === 'custom' && customTemplate) {
    return customTemplate
      .replace(/\{name\}/g,       name)
      .replace(/\{id\}/g,         student.studentId)
      .replace(/\{date\}/g,       today)
      .replace(/\{time\}/g,       time)
      .replace(/\{school\}/g,     school)
      .replace(/\{parentName\}/g, student.parentName || 'ผู้ปกครอง');
  }

  switch (statusType) {
    case 'arrive': return `[${school}] แจ้งผู้ปกครอง: ${name} เดินทางถึงโรงเรียนแล้ว เวลา ${time} น. วันที่ ${today}`;
    case 'leave':  return `[${school}] แจ้งผู้ปกครอง: ${name} ออกจากโรงเรียนแล้ว เวลา ${time} น. วันที่ ${today}`;
    case 'absent': return `[${school}] แจ้งผู้ปกครอง: ${name} ไม่มาเรียนวันที่ ${today} กรุณาติดต่อโรงเรียน`;
    case 'late':   return `[${school}] แจ้งผู้ปกครอง: ${name} มาสาย เวลา ${time} น. วันที่ ${today}`;
    case 'sick':   return `[${school}] แจ้งผู้ปกครอง: ${name} ลาป่วย วันที่ ${today}`;
    case 'leave2': return `[${school}] แจ้งผู้ปกครอง: ${name} ลากิจ วันที่ ${today}`;
    default:       return `[${school}] แจ้งผู้ปกครอง: ${name} วันที่ ${today} เวลา ${time} น.`;
  }
}

// ─── Send ─────────────────────────────────────────────────────
function handleSend(body) {
  const student = findStudent(body.studentId);
  const message = buildMessage(student, body.statusType, body.schoolName, body.customTemplate);
  const api     = sendThaibulksms(student.parentPhone, message);
  logSend(student, body.statusType, message, api);
  return { ok: true, studentName: student.name, parentPhone: student.parentPhone, message, api };
}

// ─── Thaibulksms API ──────────────────────────────────────────
function sendThaibulksms(msisdn, message) {
  const props  = PropertiesService.getScriptProperties();
  const key    = props.getProperty('THAIBULKSMS_KEY');
  const secret = props.getProperty('THAIBULKSMS_SECRET');
  const sender = props.getProperty('THAIBULKSMS_SENDER') || 'School';
  if (!key || !secret) throw new Error('ยังไม่ได้ตั้งค่า THAIBULKSMS_KEY/SECRET ใน Script Properties');

  const res = UrlFetchApp.fetch('https://api-v2.thaibulksms.com/sms', {
    method:  'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(`${key}:${secret}`), Accept: 'application/json' },
    payload: { msisdn, message, sender, force: 'standard' },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  let parsed = null;
  try { parsed = JSON.parse(text); } catch (_) {}
  if (code < 200 || code >= 300) throw new Error(`Thaibulksms error ${code}: ${text}`);
  return parsed || { raw: text };
}

// ─── Log ──────────────────────────────────────────────────────
function logSend(student, statusType, message, api) {
  const sh = getOrCreateSheet(SHEET_LOG,
    ['timestamp', 'studentId', 'name', 'parentPhone', 'statusType', 'message', 'apiResponse']);
  sh.appendRow([new Date(), student.studentId, student.name, student.parentPhone, statusType, message, JSON.stringify(api)]);
}

// ─── Get all students ─────────────────────────────────────────
function handleGetStudents() {
  const sh   = getSheet(SHEET_STUDENTS);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, students: [] };
  const rows     = sh.getRange(2, 1, last - 1, 4).getValues();
  const students = rows
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      studentId:   String(r[0]).trim(),
      name:        String(r[1] || '').trim(),
      parentPhone: String(r[2] || '').trim(),
      parentName:  String(r[3] || '').trim(),
    }));
  return { ok: true, students };
}

// ─── Bulk save (Excel import) ─────────────────────────────────
function handleSaveStudents(body) {
  const sh      = getSheet(SHEET_STUDENTS);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow - 1);

  const students = (body.students || []).filter(s => String(s.studentId || '').trim());
  if (students.length > 0) {
    const rows = students.map(s => [
      String(s.studentId || '').trim(),
      String(s.name || '').trim(),
      String(s.parentPhone || '').trim(),
      String(s.parentName || '').trim(),
    ]);
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  return { ok: true, count: students.length };
}

// ─── Add student ──────────────────────────────────────────────
function handleAddStudent(body) {
  if (findStudentRow(body.studentId) !== -1) {
    throw new Error(`รหัส ${body.studentId} มีอยู่แล้ว`);
  }
  getSheet(SHEET_STUDENTS).appendRow([
    String(body.studentId || '').trim(),
    String(body.name      || '').trim(),
    String(body.parentPhone || '').trim(),
    String(body.parentName  || '').trim(),
  ]);
  return { ok: true };
}

// ─── Update student ───────────────────────────────────────────
function handleUpdateStudent(body) {
  const row = findStudentRow(body.studentId);
  if (row === -1) throw new Error(`ไม่พบรหัส ${body.studentId}`);
  getSheet(SHEET_STUDENTS).getRange(row, 1, 1, 4).setValues([[
    String(body.studentId   || '').trim(),
    String(body.name        || '').trim(),
    String(body.parentPhone || '').trim(),
    String(body.parentName  || '').trim(),
  ]]);
  return { ok: true };
}

// ─── Delete student ───────────────────────────────────────────
function handleDeleteStudent(body) {
  const row = findStudentRow(body.studentId);
  if (row === -1) throw new Error(`ไม่พบรหัส ${body.studentId}`);
  getSheet(SHEET_STUDENTS).deleteRow(row);
  return { ok: true };
}

// ─── Get logs ─────────────────────────────────────────────────
function handleGetLogs(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return { ok: true, logs: [] };

  const limit    = Math.min(body.limit || 50, 200);
  const lastRow  = sh.getLastRow();
  const startRow = Math.max(2, lastRow - limit + 1);
  const count    = lastRow - startRow + 1;
  const rows     = sh.getRange(startRow, 1, count, 6).getValues();

  const logs = rows.reverse().map(r => ({
    timestamp:   r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : '',
    studentId:   String(r[1] || ''),
    name:        String(r[2] || ''),
    parentPhone: String(r[3] || ''),
    statusType:  String(r[4] || ''),
    message:     String(r[5] || ''),
  }));
  return { ok: true, logs };
}

// ─── Save attendance ──────────────────────────────────────────
function handleSaveAttendance(body) {
  const date    = String(body.date || '');
  const records = body.records || [];
  const sh      = getOrCreateSheet(SHEET_ATTEND,
    ['date', 'studentId', 'name', 'parentPhone', 'status']);

  // Remove existing rows for this date (iterate bottom-up)
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) {
      if (String(dates[i][0]) === date) sh.deleteRow(i + 2);
    }
  }

  if (records.length > 0) {
    const newRows = records.map(r => [
      date,
      String(r.studentId || ''),
      String(r.name || ''),
      String(r.parentPhone || ''),
      String(r.status || 'มาเรียน'),
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  }
  return { ok: true, saved: records.length };
}

// ─── Get attendance ───────────────────────────────────────────
function handleGetAttendance(body) {
  const date = String(body.date || '');
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_ATTEND);
  if (!sh || sh.getLastRow() < 2) return { ok: true, records: [] };

  const lastRow = sh.getLastRow();
  const rows    = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const records = rows
    .filter(r => String(r[0]) === date)
    .map(r => ({
      studentId:   String(r[1] || ''),
      name:        String(r[2] || ''),
      parentPhone: String(r[3] || ''),
      status:      String(r[4] || 'มาเรียน'),
    }));
  return { ok: true, records };
}
