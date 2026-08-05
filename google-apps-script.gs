/**
 * Apps Script — ตัวกลางระหว่างเว็บแอพ Attendance Control กับ Google Sheet
 * ชีตปลายทาง: ใส่ URL ชีตของคุณเอง (สคริปต์นี้ทำงานกับชีตที่มันถูกติดตั้งอยู่)
 *
 * ทำ 2 หน้าที่
 *   1) เขียนตารางผลลัพธ์ (Matrix / Summary / Daily rate / Warning list) ลงชีตให้คนทั่วไปเปิดดู
 *   2) เก็บ "ข้อมูลดิบที่ประมวลผลแล้ว" ไว้ให้หน้างานดึงไปแสดงในเว็บ — แอดมินอัปครั้งเดียว ทุกเครื่องเห็นเหมือนกัน
 *
 * ── ติดตั้ง (ทำครั้งเดียว) ─────────────────────────────────────────────
 * 1. เปิดชีตปลายทาง → Extensions → Apps Script
 * 2. วางไฟล์นี้ทับ Code.gs → เปลี่ยนค่า TOKEN ด้านล่างเป็นรหัสของคุณเอง → Save
 * 3. Deploy → New deployment → type: Web app
 *      Execute as: Me          Who has access: Anyone
 * 4. คัดลอก Web app URL (ลงท้าย /exec) ไปวางในหน้า "Export & Sync" ของ Attendance Admin
 *    และวาง URL เดียวกันในหน้างาน (Attendance Viewer จะถามครั้งแรกครั้งเดียว)
 *
 * หมายเหตุ: ทุกครั้งที่แก้ไฟล์นี้ ต้อง Deploy → Manage deployments → Edit → Version: New
 */

var TOKEN = 'CHANGE-ME-1234';        // ตั้งรหัสเอง แล้วใส่ให้ตรงกันในเว็บแอพ
var DATA_SHEET = '_data';            // แท็บซ่อนสำหรับเก็บข้อมูลดิบ (ห้ามลบ)
var CHUNK = 40000;                   // ขนาดต่อเซลล์ (ลิมิตชีตคือ 50,000 ตัวอักษร)

/* ── ผู้ใช้และสิทธิ์การเข้าถึง ──────────────────────────────────────────
   รายชื่อผู้ใช้อยู่ในชีตชื่อ DATA เรียงคอลัมน์ตามนี้ (แถวแรกเป็นหัวตาราง)

     A: ชื่อ            B: ยูสเซอร์         C: รหัส          D: กะที่ดูแล

   ช่อง "กะที่ดูแล" ใส่ได้หลายกะคั่นด้วยจุลภาค เช่น  กะเช้า, กะบ่าย
   ถ้าใส่คำว่า  ทั้งหมด  หรือ  ALL  คนนั้นจะเป็นแอดมิน เห็นทุกกะและจัดการผู้ใช้ได้

   ถ้าชีต DATA อยู่คนละไฟล์กับชีตนี้ ให้ใส่ ID ของไฟล์นั้นที่ USER_SHEET_ID
   (เว้นว่างไว้ = ใช้ชีตเดียวกับที่ติดตั้งสคริปต์นี้)

   ความปลอดภัย: การตรวจรหัสทำที่นี่ทั้งหมด รหัสผ่านไม่เคยถูกส่งไปที่เบราว์เซอร์
   และผู้ใช้ทั่วไปจะได้รับเฉพาะข้อมูลของกะที่ตัวเองดูแลเท่านั้น              */
var USER_SHEET = 'DATA';
var USER_SHEET_ID = '';              // เช่น '1mA69kP1t3AQysSgdBpx57MPdFIlH9xttAvM25Fw1T0w'

function userBook_() {
  return USER_SHEET_ID ? SpreadsheetApp.openById(USER_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function userSheet_() {
  var ss = userBook_();
  var sh = ss.getSheetByName(USER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USER_SHEET);
    sh.appendRow(['ชื่อ', 'ยูสเซอร์', 'รหัส', 'กะที่ดูแล']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#201e1d').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function norm_(v) { return String(v == null ? '' : v).trim(); }
function isAdminShifts_(s) { return /ทั้งหมด|^all$/i.test(norm_(s)); }
function splitShifts_(s) {
  return norm_(s).split(/[,;/|]/).map(function (x) { return x.trim(); }).filter(String);
}

function readUsers_() {
  var sh = userSheet_();
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var u = norm_(v[i][1]);
    if (!u) continue;
    out.push({ row: i + 1, name: norm_(v[i][0]), user: u, pass: norm_(v[i][2]), shiftText: norm_(v[i][3]) });
  }
  return out;
}

// ตรวจรหัส — คืน null ถ้าไม่ผ่าน ไม่บอกว่าผิดที่ยูสเซอร์หรือรหัส
function auth_(user, pass) {
  var u = norm_(user), p = norm_(pass);
  if (!u || !p) return null;
  var list = readUsers_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].user.toLowerCase() === u.toLowerCase() && list[i].pass === p) return list[i];
  }
  return null;
}
function publicUser_(u) {
  return { name: u.name, user: u.user, shiftText: u.shiftText,
    shifts: splitShifts_(u.shiftText), isAdmin: isAdminShifts_(u.shiftText) };
}

// เหลือเฉพาะพนักงานในกะที่คนนี้ดูแล — ตัดข้อมูลกะอื่นทิ้งตั้งแต่ที่เซิร์ฟเวอร์
function filterDataset_(ds, who) {
  if (!ds || who.isAdmin) return ds;
  var allow = who.shifts.map(function (s) { return s.toLowerCase(); });
  if (!allow.length) return { v: ds.v, month: ds.month, dates: ds.dates, shifts: ds.shifts, employees: [], savedAt: ds.savedAt };
  var shiftNames = (ds.shifts || []).map(function (s) { return String(s || '').toLowerCase(); });
  var okIdx = {};
  for (var i = 0; i < shiftNames.length; i++) {
    for (var j = 0; j < allow.length; j++) {
      if (allow[j] && shiftNames[i].indexOf(allow[j]) >= 0) { okIdx[i] = true; break; }
    }
  }
  var emps = (ds.employees || []).filter(function (e) {
    var cells = e.c || [];
    for (var k = 0; k < cells.length; k++) if (okIdx[cells[k][3]]) return true;
    return false;
  });
  var out = {};
  for (var key in ds) out[key] = ds[key];
  out.employees = emps;
  return out;
}

var COLOR_HEX = {
  red: '#ec3013', yellow: '#e8a400', gray: '#d7d3d3',
  'gray-half': '#eae7e7', ok: '#ffffff', out: '#f3f2f2'
};

/* ── รับข้อมูลจากแอดมิน ────────────────────────────────────────────────── */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    // คำสั่งจัดการผู้ใช้ใช้รหัสแอดมินยืนยัน ไม่ต้องใช้ TOKEN
    if (body.action === 'addUser' || body.action === 'delUser') return manageUsers_(body);
    if (body.token !== TOKEN) return json({ ok: false, error: 'token ไม่ถูกต้อง' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = [];

    // 1) เขียนตารางที่คนอ่านได้
    Object.keys(body.tabs || {}).forEach(function (name) {
      var rows = body.tabs[name] || [];
      var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
      sheet.clear();
      if (!rows.length) return;
      var width = Math.max.apply(null, rows.map(function (r) { return r.length; }));
      var norm = rows.map(function (r) {
        var out = r.slice();
        while (out.length < width) out.push('');
        return out;
      });
      sheet.getRange(1, 1, norm.length, width).setValues(norm);
      sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#201e1d').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      var colors = (body.colors || {})[name];
      if (colors) {
        for (var r = 0; r < colors.length; r++) {
          var line = colors[r] || [];
          for (var c = 0; c < line.length; c++) {
            var hex = COLOR_HEX[line[c]];
            if (hex) sheet.getRange(r + 2, c + 1).setBackground(hex);
          }
        }
      }
      written.push(name + ' (' + norm.length + ' แถว)');
    });

    // 2) เก็บข้อมูลดิบไว้ให้หน้างานดึง
    if (body.dataset && body.month) {
      saveDataset_(ss, body.month, body.dataset);
      written.push('dataset ' + body.month);
    }
    if (body.meta) saveMeta_(ss, body.meta);

    var log = ss.getSheetByName('Sync log') || ss.insertSheet('Sync log');
    if (log.getLastRow() === 0) log.appendRow(['เวลา', 'รอบข้อมูล', 'สิ่งที่เขียน', 'หมายเหตุ']);
    log.appendRow([new Date(), body.month || body.period || '', written.join(', '), body.note || '']);

    return json({ ok: true, written: written, months: listMonths_(ss) });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ── หน้างานดึงข้อมูล ──────────────────────────────────────────────────
   ?action=months            → รายชื่อเดือนที่มี
   ?action=dataset&month=... → ข้อมูลของเดือนนั้น (ไม่ใส่ month = เดือนล่าสุด)
   ?action=meta              → เกณฑ์เตือน / กะ / ประวัติหนังสือ            */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var p = (e && e.parameter) || {};
    var action = p.action || 'months';

    // ── ล็อกอิน ─────────────────────────────────────────────────────────
    if (action === 'login') {
      var who = auth_(p.user, p.pass);
      if (!who) return json({ ok: false, error: 'ยูสเซอร์หรือรหัสไม่ถูกต้อง' });
      return json({ ok: true, me: publicUser_(who) });
    }

    // ── รายชื่อผู้ใช้ (เฉพาะแอดมิน) ────────────────────────────────────
    if (action === 'users') {
      var admin = auth_(p.user, p.pass);
      if (!admin || !isAdminShifts_(admin.shiftText)) return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
      return json({ ok: true, users: readUsers_().map(function (u) {
        return { name: u.name, user: u.user, shiftText: u.shiftText, isAdmin: isAdminShifts_(u.shiftText) };
      }) });
    }

    // ── ทุก action ที่เหลือต้องล็อกอินก่อน ─────────────────────────────
    var me = auth_(p.user, p.pass);
    if (!me) return json({ ok: false, error: 'ต้องล็อกอินก่อน', needLogin: true });
    var pub = publicUser_(me);

    if (action === 'months') return json({ ok: true, months: listMonths_(ss), me: pub });
    if (action === 'meta') return json({ ok: true, meta: readMeta_(ss), me: pub });
    if (action === 'dataset') {
      var months = listMonths_(ss);
      var month = p.month || (months.length ? months[months.length - 1].month : '');
      if (!month) return json({ ok: false, error: 'ยังไม่มีข้อมูลในชีต' });
      // กรองที่นี่ ไม่ส่งข้อมูลกะอื่นออกไปเลย
      return json({ ok: true, month: month, me: pub, dataset: filterDataset_(readDataset_(ss, month), pub) });
    }
    return json({ ok: true, message: 'Attendance sync endpoint พร้อมใช้งาน', me: pub });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ── แอดมินเพิ่ม/ลบ/แก้ผู้ใช้จากหน้าเว็บ ────────────────────────────────
   เรียกด้วย POST { adminUser, adminPass, action, ... }                    */
function manageUsers_(body) {
  var admin = auth_(body.adminUser, body.adminPass);
  if (!admin || !isAdminShifts_(admin.shiftText)) return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
  var sh = userSheet_();
  var act = body.action;

  if (act === 'addUser') {
    var u = norm_(body.user);
    if (!u || !norm_(body.pass)) return json({ ok: false, error: 'ต้องมียูสเซอร์และรหัส' });
    var exists = readUsers_().filter(function (x) { return x.user.toLowerCase() === u.toLowerCase(); });
    if (exists.length) {
      sh.getRange(exists[0].row, 1, 1, 4).setValues([[norm_(body.name), u, norm_(body.pass), norm_(body.shiftText)]]);
      return json({ ok: true, updated: true, user: u });
    }
    sh.appendRow([norm_(body.name), u, norm_(body.pass), norm_(body.shiftText)]);
    return json({ ok: true, added: true, user: u });
  }

  if (act === 'delUser') {
    var t = norm_(body.user).toLowerCase();
    if (t === norm_(admin.user).toLowerCase()) return json({ ok: false, error: 'ลบตัวเองไม่ได้' });
    var found = readUsers_().filter(function (x) { return x.user.toLowerCase() === t; });
    if (!found.length) return json({ ok: false, error: 'ไม่พบยูสเซอร์นี้' });
    sh.deleteRow(found[0].row);
    return json({ ok: true, deleted: true, user: found[0].user });
  }

  return json({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + act });
}

/* ── เก็บ/อ่านข้อมูลดิบในแท็บซ่อน ────────────────────────────────────── */
function dataSheet_(ss) {
  var sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.appendRow(['key', 'part', 'payload']);
    sh.hideSheet();
  }
  return sh;
}
function deleteKey_(sh, key) {
  var values = sh.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) if (values[i][0] === key) sh.deleteRow(i + 1);
}
function readKey_(sh, key) {
  var values = sh.getDataRange().getValues();
  var parts = [];
  for (var i = 1; i < values.length; i++) if (values[i][0] === key) parts.push([values[i][1], values[i][2]]);
  if (!parts.length) return null;
  parts.sort(function (a, b) { return a[0] - b[0]; });
  return parts.map(function (x) { return x[1]; }).join('');
}
function writeKey_(sh, key, text) {
  deleteKey_(sh, key);
  var rows = [];
  for (var i = 0; i * CHUNK < text.length; i++) rows.push([key, i, text.substr(i * CHUNK, CHUNK)]);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}
function saveDataset_(ss, month, dataset) {
  var sh = dataSheet_(ss);
  writeKey_(sh, 'ds:' + month, JSON.stringify(dataset));
  var idx = JSON.parse(readKey_(sh, 'index') || '[]').filter(function (x) { return x.month !== month; });
  idx.push({ month: month, employees: (dataset.employees || []).length, days: (dataset.dates || []).length, savedAt: new Date().toISOString() });
  idx.sort(function (a, b) { return a.month < b.month ? -1 : 1; });
  while (idx.length > 24) { var drop = idx.shift(); deleteKey_(sh, 'ds:' + drop.month); }
  writeKey_(sh, 'index', JSON.stringify(idx));
}
function readDataset_(ss, month) {
  var raw = readKey_(dataSheet_(ss), 'ds:' + month);
  return raw ? JSON.parse(raw) : null;
}
function listMonths_(ss) { return JSON.parse(readKey_(dataSheet_(ss), 'index') || '[]'); }
function saveMeta_(ss, meta) { writeKey_(dataSheet_(ss), 'meta', JSON.stringify(meta)); }
function readMeta_(ss) { return JSON.parse(readKey_(dataSheet_(ss), 'meta') || 'null'); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
