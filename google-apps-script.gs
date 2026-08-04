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

var COLOR_HEX = {
  red: '#ec3013', yellow: '#e8a400', gray: '#d7d3d3',
  'gray-half': '#eae7e7', ok: '#ffffff', out: '#f3f2f2'
};

/* ── รับข้อมูลจากแอดมิน ────────────────────────────────────────────────── */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
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
    if (action === 'months') return json({ ok: true, months: listMonths_(ss) });
    if (action === 'meta') return json({ ok: true, meta: readMeta_(ss) });
    if (action === 'dataset') {
      var months = listMonths_(ss);
      var month = p.month || (months.length ? months[months.length - 1].month : '');
      if (!month) return json({ ok: false, error: 'ยังไม่มีข้อมูลในชีต' });
      return json({ ok: true, month: month, dataset: readDataset_(ss, month) });
    }
    return json({ ok: true, message: 'Attendance sync endpoint พร้อมใช้งาน' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
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
