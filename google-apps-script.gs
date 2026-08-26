/**
 * Apps Script — ตัวกลางระหว่างเว็บแอพ Attendance Control กับ Google Sheet
 * ชีตปลายทาง: ใส่ URL ชีตของคุณเอง (สคริปต์นี้ทำงานกับชีตที่มันถูกติดตั้งอยู่)
 *
 * ทำ 2 หน้าที่
 *   1) เขียนตารางผลลัพธ์ (Matrix / Summary / Daily rate / Warning list) ลงชีตให้คนทั่วไปเปิดดู
 *   2) เก็บ "ข้อมูลดิบที่ประมวลผลแล้ว" ไว้ให้หน้างานดึงไปแสดงในเว็บ — แอดมินอัปครั้งเดียว ทุกเครื่องเห็นเหมือนกัน
 *
 * ── ใช้ชีตเดียวจบ (แนะนำ) ──────────────────────────────────────────────
 * ติดตั้งสคริปต์นี้ในชีตที่มีแท็บ DATA อยู่ แล้วปล่อย USER_SHEET_ID ว่างไว้
 * สคริปต์จะอ่านรายชื่อผู้ใช้จากแท็บ DATA และสร้างแท็บเก็บข้อมูลการมาทำงาน
 * (_data, Matrix, Summary, Sync log) ไว้ในไฟล์เดียวกันนั้นเอง
 *
 * ถ้าจำเป็นต้องแยกไฟล์จริง ๆ ค่อยใส่ USER_SHEET_ID ชี้ไปไฟล์ที่มีแท็บ DATA
 *
 * ── ติดตั้ง (ทำครั้งเดียว) ─────────────────────────────────────────────
 * 1. เปิดชีตที่มีแท็บ DATA → Extensions → Apps Script
 * 2. วางไฟล์นี้ทับ Code.gs → Save (ไม่ต้องแก้อะไรเลย)
 * 3. Deploy → New deployment → type: Web app
 *      Execute as: Me          Who has access: Anyone
 * 4. คัดลอก Web app URL (ลงท้าย /exec) ไปวางในหน้า "Export & Sync" ของ Attendance Admin
 *    และวาง URL เดียวกันในหน้างาน (Attendance Viewer จะถามครั้งแรกครั้งเดียว)
 *
 * หมายเหตุ: ทุกครั้งที่แก้ไฟล์นี้ ต้อง Deploy → Manage deployments → Edit → Version: New
 */

// ไม่ต้องแก้บรรทัดนี้แล้ว — ระบบใช้การล็อกอินแอดมินยืนยันตัวแทน TOKEN
// เก็บไว้เผื่อเครื่องมือเก่าที่ยังส่ง token แบบเดิมมาเท่านั้น
var TOKEN = 'CHANGE-ME-1234';
var DATA_SHEET = '_data';            // แท็บซ่อนสำหรับเก็บข้อมูลดิบ (ห้ามลบ)
var CHUNK = 40000;                   // ขนาดต่อเซลล์ (ลิมิตชีตคือ 50,000 ตัวอักษร)

/* ── ผู้ใช้และสิทธิ์การเข้าถึง ──────────────────────────────────────────
   รายชื่อผู้ใช้อยู่ในชีตชื่อ DATA เรียงคอลัมน์ตามนี้ (แถวแรกเป็นหัวตาราง)

     A: ชื่อ    B: ยูสเซอร์    C: รหัส    D: กะที่ดูแล    E: เมนูที่เข้าได้

   D "กะที่ดูแล" คุมว่าเห็นข้อมูลของใคร — ใส่ชื่อกะให้ตรงกับที่ตั้งไว้ในเว็บ
     ใส่ได้หลายกะคั่นจุลภาค เช่น  03:00-12:00, 18:00-03:00
     ใส่คำว่า  ทั้งหมด  = เห็นข้อมูลทุกกะ

   E "เมนูที่เข้าได้" คุมว่าเข้าหน้าไหนได้ — เว้นว่างไว้ก็ได้
     เว้นว่าง + กะ "ทั้งหมด"  = เข้าได้ทุกเมนู (เป็นแอดมินเต็ม)
     เว้นว่าง + ดูแลบางกะ     = เข้าได้เฉพาะเมนูดูข้อมูล อัปโหลดและตั้งค่าไม่ได้
     ระบุเอง                  = เช่น  แดชบอร์ด, ตารางสี, หนังสือเตือน

   เว้น USER_SHEET_ID ว่างไว้ = อ่านแท็บ DATA จากไฟล์เดียวกับที่ติดตั้งสคริปต์
   ซึ่งเป็นวิธีที่แนะนำ ไม่ต้องตั้งค่าอะไรเพิ่ม
   ใส่ ID เฉพาะกรณีที่แท็บ DATA อยู่คนละไฟล์กับที่ติดตั้งสคริปต์จริง ๆ

   ความปลอดภัย: การตรวจรหัสทำที่นี่ทั้งหมด รหัสผ่านไม่เคยถูกส่งไปที่เบราว์เซอร์
   และผู้ใช้ทั่วไปจะได้รับเฉพาะข้อมูลของกะที่ตัวเองดูแลเท่านั้น              */
var USER_SHEET = 'DATA';
var USER_SHEET_ID = '';              // วาง ID ของไฟล์ที่มีชีต DATA (ตัวอักษรยาว ๆ กลาง URL ของชีต)

function userBook_() {
  return USER_SHEET_ID ? SpreadsheetApp.openById(USER_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function userSheet_() {
  var ss = userBook_();
  var sh = ss.getSheetByName(USER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USER_SHEET);
    sh.appendRow(['ชื่อ', 'ยูสเซอร์', 'รหัส', 'กะที่ดูแล', 'เมนูที่เข้าได้']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#201e1d').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function norm_(v) { return String(v == null ? '' : v).trim(); }
function isAdminShifts_(s) { return /ทั้งหมด|^all$/i.test(norm_(s)); }
function splitShifts_(s) {
  return norm_(s).split(/[,;/|]/).map(function (x) { return x.trim(); }).filter(String);
}

/* อ่านรายชื่อผู้ใช้ทีเดียวแล้วจำไว้ 6 ชั่วโมง
   ทุกคำขอที่ต้องยืนยันตัวตนจะเรียกฟังก์ชันนี้ ถ้าอ่านชีตใหม่ทุกครั้งจะเสียเวลา
   หลายวินาทีต่อคำขอ — ล้างแคชทันทีเมื่อมีการเพิ่มหรือลบผู้ใช้              */
var USER_CACHE_KEY = 'users_v1';
function dropUserCache_() { try { CacheService.getScriptCache().remove(USER_CACHE_KEY); } catch (e) {} }

function readUsers_() {
  try {
    var hit = CacheService.getScriptCache().get(USER_CACHE_KEY);
    if (hit) return JSON.parse(hit);
  } catch (e) {}
  var fresh = readUsersFromSheet_();
  try { CacheService.getScriptCache().put(USER_CACHE_KEY, JSON.stringify(fresh), 21600); } catch (e) {}
  return fresh;
}

function readUsersFromSheet_() {
  var sh = userSheet_();
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var u = norm_(v[i][1]);
    if (!u) continue;
    out.push({ row: i + 1, name: norm_(v[i][0]), user: u, pass: norm_(v[i][2]),
      shiftText: norm_(v[i][3]), menuText: norm_(v[i][4]) });
  }
  return out;
}

/* ── เมนูที่แต่ละคนเข้าได้ (คอลัมน์ E) ──────────────────────────────────
   เว้นว่าง   = ใช้ค่าเริ่มต้น (ดูกะทั้งหมด → ทุกเมนู, ดูบางกะ → เมนูดูอย่างเดียว)
   "ทั้งหมด"  = เข้าได้ทุกเมนู
   นอกนั้น    = ใส่ชื่อเมนูคั่นจุลภาค เช่น  แดชบอร์ด, ตารางสี, รายพนักงาน     */
var ALL_MENUS = ['upload', 'dash', 'grid', 'emp', 'hub', 'warn', 'archive', 'export', 'rules', 'users'];
var MENU_LABEL = {
  upload: 'อัปโหลดไฟล์', dash: 'แดชบอร์ด', grid: 'ตารางสี', emp: 'รายพนักงาน',
  hub: 'รายตำแหน่ง', warn: 'หนังสือเตือน', archive: 'คลังสถิติย้อนหลัง',
  export: 'ส่งออกไฟล์', rules: 'ตั้งค่ากฎสี', users: 'ผู้ใช้และสิทธิ์'
};
/* ค่าเริ่มต้นของคนที่ดูแลเฉพาะบางกะ — ดูได้แต่แก้ระบบไม่ได้
   ไม่มี "รายตำแหน่ง" เพราะเป็นการสรุปข้ามทั้งฮับ คนที่ดูแลกะเดียวไม่ควรเห็นภาพรวมทั้งหมด
   ถ้าอยากให้ใครเห็น ใส่คำว่า hub ลงคอลัมน์ "เมนูที่เข้าได้" ของคนนั้นได้เป็นราย ๆ ไป */
var VIEWER_MENUS = ['dash', 'grid', 'emp', 'warn', 'archive', 'export'];

function menusOf_(u) {
  var t = norm_(u.menuText);
  if (!t) return isAdminShifts_(u.shiftText) ? ALL_MENUS.slice() : VIEWER_MENUS.slice();
  if (/ทั้งหมด|^all$/i.test(t)) return ALL_MENUS.slice();
  var want = t.split(/[,;/|]/).map(function (x) { return x.trim().toLowerCase(); }).filter(String);
  var out = [];
  ALL_MENUS.forEach(function (k) {
    for (var i = 0; i < want.length; i++) {
      if (want[i] === k || want[i] === String(MENU_LABEL[k]).toLowerCase()) { out.push(k); return; }
    }
  });
  return out;
}

// ตรวจรหัส — คืน null ถ้าไม่ผ่าน ไม่บอกว่าผิดที่ยูสเซอร์หรือรหัส
/* หารายชื่อจากที่จำไว้ก่อน ถ้าไม่เจอค่อยไปอ่านชีตใหม่แล้วลองอีกที
   เพราะรายชื่อถูกจำไว้ 6 ชั่วโมง ถ้าแอดมินไปพิมพ์เพิ่มคนในชีตเอง
   คนใหม่จะล็อกอินไม่ได้จนกว่าจะครบเวลา ทั้งที่แถวอยู่ในชีตเรียบร้อยแล้ว
   ยอมอ่านชีตเพิ่มหนึ่งครั้งเฉพาะตอนล็อกอินไม่ผ่าน แลกกับไม่ต้องรอ 6 ชั่วโมง */
function auth_(user, pass) {
  var u = norm_(user), p = norm_(pass);
  if (!u || !p) return null;
  var hit = matchUser_(readUsers_(), u, p);
  if (hit) return hit;
  var fresh = readUsersFromSheet_();
  try { CacheService.getScriptCache().put(USER_CACHE_KEY, JSON.stringify(fresh), 21600); } catch (e) {}
  return matchUser_(fresh, u, p);
}
function matchUser_(list, u, p) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].user.toLowerCase() === u.toLowerCase() && list[i].pass === p) return list[i];
  }
  return null;
}
function publicUser_(u) {
  return { name: u.name, user: u.user, shiftText: u.shiftText,
    shifts: splitShifts_(u.shiftText), isAdmin: isAdminShifts_(u.shiftText),
    menus: menusOf_(u), menuText: norm_(u.menuText) };
}

/* ── บัตรผ่านชั่วคราว (token) ───────────────────────────────────────────
   ล็อกอินสำเร็จแล้วได้บัตรผ่านสุ่มไป 1 ใบ อายุ 12 ชั่วโมง คำขอหลังจากนั้น
   ใช้บัตรใบนี้แทนรหัสผ่าน — รหัสผ่านจริงจึงถูกส่งแค่ครั้งเดียวตอนล็อกอิน
   และไม่ไปโผล่ใน URL หรือประวัติเบราว์เซอร์ ถ้าบัตรหลุดก็ลบทิ้งได้
   (ล็อกอินได้หลายเครื่องพร้อมกัน เพราะแต่ละเครื่องถือบัตรคนละใบ)          */
var TOKEN_TTL_MS = 12 * 60 * 60 * 1000;              // ปกติ 12 ชั่วโมง
var TOKEN_TTL_REMEMBER_MS = 30 * 24 * 60 * 60 * 1000; // ติ๊ก "จำฉันไว้" 30 วัน

function tokenStore_() { return PropertiesService.getScriptProperties(); }
function newToken_(user, remember) {
  var t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  var ttl = remember ? TOKEN_TTL_REMEMBER_MS : TOKEN_TTL_MS;
  tokenStore_().setProperty('tk:' + t, JSON.stringify({ user: user, exp: Date.now() + ttl }));
  return t;
}
function userByToken_(token) {
  var t = norm_(token);
  if (!t) return null;
  var raw = tokenStore_().getProperty('tk:' + t);
  if (!raw) return null;
  var rec;
  try { rec = JSON.parse(raw); } catch (e) { return null; }
  if (!rec || !rec.exp || Date.now() > rec.exp) { tokenStore_().deleteProperty('tk:' + t); return null; }
  var list = readUsers_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].user.toLowerCase() === String(rec.user).toLowerCase()) return list[i];
  }
  return null;   // ผู้ใช้ถูกลบออกจากชีตแล้ว บัตรใช้ไม่ได้ทันที
}
function dropToken_(token) {
  var t = norm_(token);
  if (t) tokenStore_().deleteProperty('tk:' + t);
}
// รับได้ทั้งบัตรผ่าน และรหัสผ่านตรง ๆ (เผื่อเรียกจากที่อื่น)
function whoIs_(p) {
  return p.token ? userByToken_(p.token) : auth_(p.user, p.pass);
}

// เหลือเฉพาะพนักงานในกะที่คนนี้ดูแล — ตัดข้อมูลกะอื่นทิ้งตั้งแต่ที่เซิร์ฟเวอร์
function filterDataset_(ds, who) {
  if (!ds || who.isAdmin) return ds;
  var allow = who.shifts.map(function (s) { return s.toLowerCase(); });
  if (!allow.length) return { v: ds.v, month: ds.month, dates: ds.dates, shifts: ds.shifts, employees: [], savedAt: ds.savedAt };
  /* เทียบชื่อกะแบบตรงตัว ไม่ใช่แค่มีคำนั้นอยู่ข้างใน
     เพราะชื่อกะเป็นช่วงเวลา ถ้าเทียบแบบมีคำอยู่ข้างใน คนที่ดูแล "03:00"
     จะไปตรงกับ "18:00-03:00" ด้วย ทำให้เห็นข้อมูลเกินสิทธิ์
     ยังเผื่อกรณีพิมพ์ชื่อเดิมจากไฟล์มา จึงยอมให้ตรงแบบมีคำอยู่ข้างในได้
     ต่อเมื่อคำที่พิมพ์ยาวพอจะไม่กำกวม (ไม่ใช่แค่เวลาเดียว)               */
  var shiftNames = (ds.shifts || []).map(function (s) { return String(s || '').toLowerCase().trim(); });
  var okIdx = {};
  for (var i = 0; i < shiftNames.length; i++) {
    for (var j = 0; j < allow.length; j++) {
      var want = String(allow[j] || '').toLowerCase().trim();
      if (!want) continue;
      var hit = shiftNames[i] === want;
      if (!hit && want.length >= 8 && shiftNames[i].indexOf(want) >= 0) hit = true;
      if (hit) { okIdx[i] = true; break; }
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

    // ── ล็อกอิน ─────────────────────────────────────────────────────────
    // ใช้ POST เพื่อไม่ให้รหัสผ่านไปโผล่ใน URL, log ของเซิร์ฟเวอร์ หรือประวัติเบราว์เซอร์
    if (body.action === 'login') {
      var who = auth_(body.user, body.pass);
      if (!who) return json({ ok: false, error: 'ยูสเซอร์หรือรหัสไม่ถูกต้อง' });
      return json({ ok: true, me: publicUser_(who), token: newToken_(who.user, !!body.remember),
        remember: !!body.remember });
    }

    // คำสั่งจัดการผู้ใช้ ยืนยันด้วยบัตรผ่านหรือรหัสของแอดมิน
    if (body.action === 'addUser' || body.action === 'delUser') return manageUsers_(body);

    /* ── ใครส่งข้อมูลเข้ามาเขียนลงชีตได้บ้าง ────────────────────────────
       ใช้บัตรผ่านของแอดมินที่ล็อกอินอยู่เป็นตัวยืนยัน ไม่ต้องตั้งรหัสแยกอีก
       บัตรผ่านเกิดตอนล็อกอิน หมดอายุเองได้ และไม่เคยอยู่ในโค้ดสาธารณะ
       จึงปลอดภัยกว่าการฝังรหัสตายตัวไว้ในไฟล์                            */
    var sender = body.token ? userByToken_(body.token) : null;
    if (!sender && TOKEN !== 'CHANGE-ME-1234' && body.token === TOKEN) sender = { shiftText: 'ทั้งหมด' };
    if (!sender || !isAdminShifts_(sender.shiftText)) {
      return json({ ok: false, error: 'ต้องล็อกอินเป็นแอดมินก่อนจึงจะส่งข้อมูลขึ้นชีตได้', needLogin: true });
    }

    /* ── ล้างข้อมูลฝั่งเซิร์ฟเวอร์ทั้งหมด ────────────────────────────────
       ข้อมูลกระจายอยู่หลายที่ (แท็บ _data, ตารางที่คนอ่านได้, ชีตสำรองรายปี)
       ถ้าไม่มีคำสั่งเดียวจบ แอดมินต้องไล่ลบเองทีละที่แล้วลืมบางที่เสมอ      */
    if (body.action === 'resetAll') {
      // หยิบชีตปัจจุบันตรงนี้เอง เพราะตัวแปร ss ถูกประกาศไว้ท้ายฟังก์ชัน
      // ถ้าอ้างก่อนจะได้ค่าว่างแล้วพังทั้งคำสั่ง
      var rss = SpreadsheetApp.getActiveSpreadsheet();
      var wiped = [];
      // 1) ข้อมูลดิบที่หน้างานดึงไปแสดง
      var dsh = dataSheet_(rss);
      listMonths_(rss).forEach(function (x) { deleteKey_(dsh, 'ds:' + x.month); });
      deleteKey_(dsh, 'index');
      dropSnapshot_(rss);            // ของสำรองไว้กดย้อนก็ต้องหายไปด้วย
      wiped.push('ข้อมูลที่หน้างานดึงไปแสดง');
      // 2) ตารางที่คนเปิดอ่านได้ในชีตนี้
      ['Matrix', 'Summary by position', 'Daily rate', 'Warning list', 'Sync log'].forEach(function (n) {
        var t = rss.getSheetByName(n);
        if (t) { t.clear(); wiped.push(n); }
      });
      // 3) ชีตสำรองรายปีทั้งสองไฟล์ — ลบทุกแท็บที่ชื่อเป็นตัวเลขปี
      ['sch', 'att'].forEach(function (kind) {
        var id = sheetIdOf_(cfgGet_(CFG_KEYS[kind]));
        if (!id) return;
        try {
          var bk = SpreadsheetApp.openById(id);
          // เก็บชื่อแท็บที่จะลบไว้ก่อน แล้วค่อยลบทีหลัง
          // ถ้าลบระหว่างวนรายการ ตัวถัดไปจะกลายเป็นแท็บที่ถูกลบไปแล้ว
          var years = [];
          bk.getSheets().forEach(function (t) { if (/^\d{4}$/.test(t.getName())) years.push(t.getName()); });
          years.forEach(function (nm) {
            var t = bk.getSheetByName(nm);
            if (!t) return;
            if (bk.getSheets().length > 1) bk.deleteSheet(t); else t.clear();
            wiped.push(bk.getName() + ' แท็บ ' + nm);
          });
          if (!years.length) wiped.push(bk.getName() + ' ไม่มีแท็บรายปีให้ลบ');
        } catch (e2) { wiped.push('ชีตสำรอง ' + kind + ' ลบไม่สำเร็จ: ' + String(e2)); }
      });
      return json({ ok: true, wiped: wiped });
    }

    /* ── ย้อนการอัปรอบล่าสุด ────────────────────────────────────────────
       ถอนวันที่ของรอบที่เพิ่งอัปออกจากชีตสำรองก่อน แล้วค่อยเอาของเดิมคืน
       ถ้าเดือนนั้นเดิมยังไม่มีอะไรเลย ก็ลบทิ้งให้ว่างเหมือนก่อนอัป
       ตารางที่คนอ่านได้ (Matrix ฯลฯ) หน้าเว็บจะเขียนทับให้เองหลังได้ของคืน */
    if (body.action === 'undoSync') {
      var uss = SpreadsheetApp.getActiveSpreadsheet();
      var snap = readSnapshot_(uss);
      if (!snap || !snap.month) return json({ ok: false, error: 'ไม่มีรอบอัปที่ย้อนกลับได้' });
      var usheet = dataSheet_(uss);
      var undone = [];
      var uyear = String(snap.month).slice(0, 4);
      ['att', 'sch'].forEach(function (kind) {
        try {
          var r = dropDatesFromBackup_(kind, uyear, snap.dates || []);
          if (r) undone.push('ถอน ' + r.ถอนวันที่ + ' วันออกจาก ' + r.ไฟล์);
        } catch (e3) { undone.push('ชีตสำรอง ' + kind + ' ถอนไม่สำเร็จ: ' + String(e3)); }
      });
      var back = null;
      if (snap.had) {
        var prevRaw = readKey_(usheet, 'snapds');
        if (!prevRaw) return json({ ok: false, error: 'ของเดิมหายไปจากชีต ย้อนกลับไม่ได้' });
        writeKey_(usheet, 'ds:' + snap.month, prevRaw);
        back = JSON.parse(prevRaw);
        undone.push('คืนข้อมูลเดือน ' + snap.month + ' เป็นชุดก่อนหน้า');
      } else {
        var res0 = dropMonth_(uss, snap.month);
        undone = undone.concat(res0.done);
      }
      dropSnapshot_(uss);
      return json({ ok: true, month: snap.month, undone: undone,
        dataset: back ? filterDataset_(back, publicUser_(sender)) : null,
        months: listMonths_(uss) });
    }

    /* ── ลบข้อมูลทั้งเดือนออกจากชีต ─────────────────────────────────────
       แก้ของเก่าด้วยมือในแท็บ Matrix ไม่มีผล เพราะหน้าเว็บอ่านจากแท็บ _data
       จึงต้องมีคำสั่งลบให้ครบทุกที่ในครั้งเดียว                            */
    if (body.action === 'deleteMonth') {
      if (!body.month) return json({ ok: false, error: 'ไม่ได้บอกว่าจะลบเดือนไหน' });
      var mss = SpreadsheetApp.getActiveSpreadsheet();
      var res1 = dropMonth_(mss, body.month);
      var msnap = readSnapshot_(mss);
      if (msnap && msnap.month === body.month) dropSnapshot_(mss);   // ย้อนไปหาของที่ลบแล้วไม่ได้
      return json({ ok: true, month: body.month, done: res1.done, months: res1.months });
    }

    /* ── จำตารางกะไว้บนชีต ──────────────────────────────────────────────
       เดิมจำไว้ในเครื่องที่อัป พออัปจากอีกเครื่องก็ต้องหาไฟล์ตารางกะมาอัปใหม่
       เก็บไว้ตรงกลางแล้วเครื่องไหนอัปไฟล์เวลาทำงานก็หยิบไปใช้ได้เหมือนกัน   */
    if (body.action === 'setSchedule') {
      var sss = SpreadsheetApp.getActiveSpreadsheet();
      if (body.table && body.table.length) {
        writeKey_(dataSheet_(sss), 'schtab', JSON.stringify({
          name: norm_(body.name), rows: body.table, at: new Date().toISOString() }));
        return json({ ok: true, saved: body.table.length });
      }
      deleteKey_(dataSheet_(sss), 'schtab');       // ส่งมาว่าง = สั่งลืมตารางกะ
      return json({ ok: true, saved: 0 });
    }

    // ตั้งค่าชีตสำรอง — แอดมินวาง URL ในหน้าเว็บ เก็บไว้ฝั่ง Google ไม่อยู่ในโค้ด
    if (body.action === 'setBackup') {
      if (body.sch !== undefined) cfgSet_(CFG_KEYS.sch, sheetIdOf_(body.sch));
      if (body.att !== undefined) cfgSet_(CFG_KEYS.att, sheetIdOf_(body.att));
      return json({ ok: true, sch: cfgGet_(CFG_KEYS.sch), att: cfgGet_(CFG_KEYS.att) });
    }
    if (body.action === 'getBackup') {
      return json({ ok: true, sch: cfgGet_(CFG_KEYS.sch), att: cfgGet_(CFG_KEYS.att) });
    }
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

    // 2) เก็บข้อมูลดิบไว้ให้หน้างานดึง — เก็บของเดิมไว้ก่อนหนึ่งชุดเผื่อกดย้อน
    if (body.dataset && body.month) {
      keepSnapshot_(ss, body.month, body.dataset);
      saveDataset_(ss, body.month, body.dataset);
      written.push('dataset ' + body.month);
    }

    // 3) สำรองแบบย่อลงชีตแยก แบ่งแท็บตามปี (ขึ้นปีใหม่สร้างแท็บใหม่เอง)
    var year = String(body.month || '').slice(0, 4) || String(new Date().getFullYear());
    ['att', 'sch'].forEach(function (kind) {
      var rows = body.backup && body.backup[kind];
      if (!rows || !rows.length) return;
      try {
        var r = saveBackup_(kind, year, rows);
        if (r) written.push('สำรอง ' + (kind === 'att' ? 'attendance' : 'ตารางกะ')
          + ' → ' + r.ไฟล์ + ' แท็บ ' + r.แท็บ + ' (' + r.ผล.แถว + ' แถว)');
      } catch (err2) {
        written.push('สำรอง ' + kind + ' ไม่สำเร็จ: ' + String(err2));
      }
    });
    if (body.meta) saveMeta_(ss, body.meta);

    var log = ss.getSheetByName('Sync log') || ss.insertSheet('Sync log');
    if (log.getLastRow() === 0) log.appendRow(['เวลา', 'รอบข้อมูล', 'สิ่งที่เขียน', 'หมายเหตุ']);
    log.appendRow([new Date(), body.month || body.period || '', written.join(', '), body.note || '']);

    /* ส่งข้อมูลที่เพิ่งเก็บลงชีตกลับไปด้วยเลย หน้าเว็บจะได้แสดงของบนชีตจริง
       ไม่ใช่ของที่ตัวเองเพิ่งคำนวณ — ถ้าเขียนไม่ครบจะเห็นทันทีตั้งแต่ตอนอัป  */
    var justSaved = body.month ? filterDataset_(readDataset_(ss, body.month), publicUser_(sender)) : null;
    var snapw = readSnapshot_(ss);
    return json({ ok: true, written: written, months: listMonths_(ss),
      month: body.month || '', dataset: justSaved,
      // บอกกลับไปเลยว่ารอบนี้ย้อนกลับเป็นอะไร ปุ่มย้อนจะได้ขึ้นข้อความถูกตั้งแต่แรก
      undo: snapw ? { month: snapw.month, at: snapw.at, had: !!snapw.had } : null });
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

    // เช็คว่าเว็บแอปเปิดให้เข้าถึงได้จริงไหม ใช้ตอนตั้งค่าครั้งแรก
    if (action === 'ping') return json({ ok: true, message: 'เชื่อมต่อได้', needsLogin: true });

    /* ── เครื่องมือวินิจฉัยตอนตั้งค่า ────────────────────────────────────
       บอกว่าสคริปต์กำลังอ่านไฟล์ไหน แท็บไหน เจอกี่แถว และยูสเซอร์ที่ถามหา
       มีอยู่ไหม โดยไม่เปิดเผยรหัสผ่าน ต้องใส่ TOKEN ถึงจะเรียกได้
       เรียกแบบ: ...exec?action=diag&t=TOKEN ของคุณ&user=ยูสเซอร์ที่จะเช็ค    */
    if (action === 'diag') {
      // ใช้บัตรผ่านของแอดมิน หรือยูสเซอร์+รหัสแอดมิน ไม่ผูกกับ TOKEN อีกแล้ว
      var dadmin = whoIs_(p);
      if (!dadmin || !isAdminShifts_(dadmin.shiftText)) {
        return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
      }
      var book = userBook_();
      var tab = book.getSheetByName(USER_SHEET);
      var head = tab ? tab.getRange(1, 1, 1, 4).getValues()[0].map(norm_) : [];
      var list = readUsers_();
      var out = {
        ok: true,
        ไฟล์ที่อ่าน: book.getName(),
        idไฟล์: book.getId(),
        อ่านจากไฟล์เดียวกับสคริปต์: !USER_SHEET_ID,
        มีแท็บDATA: !!tab,
        หัวตารางแถว1: head,
        จำนวนแถวที่อ่านได้: list.length,
        ยูสเซอร์ทั้งหมด: list.map(function (u) { return u.user; })
      };
      var want = norm_(p.user);
      if (want) {
        var hit = null;
        for (var i = 0; i < list.length; i++) {
          if (list[i].user.toLowerCase() === want.toLowerCase()) { hit = list[i]; break; }
        }
        out.เช็คยูสเซอร์ = want;
        out.เจอยูสเซอร์นี้ = !!hit;
        if (hit) {
          var rawPass = String(tab.getRange(hit.row, 3).getValue());
          out.ช่องรหัสมีค่า = !!hit.pass;
          out.รหัสมีช่องว่างหัวท้าย = rawPass !== rawPass.trim();
          out.ช่องชื่อ = hit.name;
          out.ช่องกะที่ดูแล = hit.shiftText;
          out.เป็นแอดมิน = isAdminShifts_(hit.shiftText);
          out.อยู่แถวที่ = hit.row;
        }
      }
      return json(out);
    }

    // ออกจากระบบ — ทิ้งบัตรผ่านใบนั้น
    if (action === 'logout') { dropToken_(p.token); return json({ ok: true }); }

    // ── รายชื่อผู้ใช้ (เฉพาะแอดมิน) ────────────────────────────────────
    if (action === 'users') {
      var admin = whoIs_(p);
      if (!admin || !isAdminShifts_(admin.shiftText)) return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
      return json({ ok: true, users: readUsers_().map(function (u) {
        return { name: u.name, user: u.user, shiftText: u.shiftText, menuText: u.menuText,
          menus: menusOf_(u), isAdmin: isAdminShifts_(u.shiftText) };
      }), allMenus: ALL_MENUS.map(function (k) { return { key: k, label: MENU_LABEL[k] }; }) });
    }

    // ── ทุก action ที่เหลือต้องล็อกอินก่อน ─────────────────────────────
    var me = whoIs_(p);
    if (!me) return json({ ok: false, error: 'ต้องล็อกอินก่อน', needLogin: true });
    var pub = publicUser_(me);

    /* รวมสามคำขอที่หน้าเว็บต้องใช้ตอนเปิด (months + dataset + meta) ไว้ในรอบเดียว
       เดิมยิงทีละรอบต่อกัน แต่ละรอบใช้เวลาสองถึงสามวินาที รวมแล้วรอนานเกินจำเป็น */
    if (action === 'bootstrap') {
      var months = listMonths_(ss);
      var pick = p.month || (months.length ? months[months.length - 1].month : '');
      var snapb = pub.isAdmin ? readSnapshot_(ss) : null;
      return json({ ok: true, me: pub, months: months, month: pick,
        meta: readMeta_(ss),
        // บอกว่ามีรอบอัปให้ย้อนกลับไหม แอดมินจะได้เห็นปุ่มย้อนจากเครื่องไหนก็ได้
        undo: snapb ? { month: snapb.month, at: snapb.at, had: !!snapb.had } : null,
        dataset: pick ? filterDataset_(readDataset_(ss, pick), pub) : null });
    }

    /* ตารางกะที่จำไว้ — ดึงตอนจะอัปไฟล์เท่านั้น ไม่พ่วงไปกับ bootstrap
       เพราะเป็นตารางใหญ่ ถ้าส่งทุกครั้งที่เปิดหน้าจะทำให้เข้าเว็บช้าโดยเปล่าประโยชน์ */
    if (action === 'schedule') {
      if (!pub.isAdmin) return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
      var srow = readKey_(dataSheet_(ss), 'schtab');
      return json({ ok: true, schedule: srow ? JSON.parse(srow) : null });
    }

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
  var admin = body.token ? userByToken_(body.token) : auth_(body.adminUser, body.adminPass);
  if (!admin || !isAdminShifts_(admin.shiftText)) return json({ ok: false, error: 'ต้องเป็นแอดมินเท่านั้น' });
  var sh = userSheet_();
  var act = body.action;

  if (act === 'addUser') {
    var u = norm_(body.user);
    if (!u || !norm_(body.pass)) return json({ ok: false, error: 'ต้องมียูสเซอร์และรหัส' });
    // อ่านชีตสด ๆ ไม่ใช้ของที่จำไว้ ไม่งั้นคนที่พิมพ์เพิ่มในชีตเองจะถูกเพิ่มซ้ำอีกแถว
    var exists = readUsersFromSheet_().filter(function (x) { return x.user.toLowerCase() === u.toLowerCase(); });
    if (exists.length) {
      sh.getRange(exists[0].row, 1, 1, 5).setValues([[norm_(body.name), u, norm_(body.pass), norm_(body.shiftText), norm_(body.menuText)]]);
      dropUserCache_();
      return json({ ok: true, updated: true, user: u });
    }
    sh.appendRow([norm_(body.name), u, norm_(body.pass), norm_(body.shiftText), norm_(body.menuText)]);
    dropUserCache_();
    return json({ ok: true, added: true, user: u });
  }

  if (act === 'delUser') {
    var t = norm_(body.user).toLowerCase();
    if (t === norm_(admin.user).toLowerCase()) return json({ ok: false, error: 'ลบตัวเองไม่ได้' });
    // อ่านสด ๆ เหมือนกัน เลขแถวที่จำไว้อาจไม่ตรงแล้วถ้ามีคนแก้ชีตด้วยมือ
    var found = readUsersFromSheet_().filter(function (x) { return x.user.toLowerCase() === t; });
    if (!found.length) return json({ ok: false, error: 'ไม่พบยูสเซอร์นี้' });
    sh.deleteRow(found[0].row);
    dropUserCache_();
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
/* ── เก็บของเดิมไว้หนึ่งชุดก่อนเขียนทับ ────────────────────────────────
   เผื่ออัปผิดไฟล์หรือผิดเดือน จะได้กดย้อนกลับได้จากเครื่องไหนก็ได้
   เก็บแค่รอบล่าสุดรอบเดียว รอบเก่ากว่านั้นให้ลบข้อมูลทั้งเดือนทิ้งแทน     */
function keepSnapshot_(ss, month, newDataset) {
  var sh = dataSheet_(ss);
  var prev = readKey_(sh, 'ds:' + month);
  writeKey_(sh, 'snap', JSON.stringify({
    month: month,
    had: !!prev,
    // วันที่ของรอบที่กำลังจะเขียน ใช้ตอนย้อนกลับเพื่อรู้ว่าต้องถอนวันไหนออก
    dates: (newDataset && newDataset.dates) || [],
    at: new Date().toISOString(),
  }));
  if (prev) writeKey_(sh, 'snapds', prev); else deleteKey_(sh, 'snapds');
}
function readSnapshot_(ss) {
  var raw = readKey_(dataSheet_(ss), 'snap');
  return raw ? JSON.parse(raw) : null;
}
function dropSnapshot_(ss) {
  var sh = dataSheet_(ss);
  deleteKey_(sh, 'snap');
  deleteKey_(sh, 'snapds');
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

/* ── ชีตสำรองแยกไฟล์ แบ่งแท็บตามปี ──────────────────────────────────────
   เก็บ ID ของชีตสำรองไว้ใน Script Properties ฝั่ง Google ไม่ได้อยู่ในโค้ด
   แอดมินกรอก URL ชีตในหน้า Export & Sync ครั้งเดียว ระบบจำให้เอง

   แต่ละไฟล์จะมีแท็บชื่อเป็นปี เช่น 2026, 2027 ขึ้นปีใหม่สร้างแท็บใหม่เอง
   ข้อมูลที่เขียนลงเป็นแบบย่อ เก็บเฉพาะที่จำเป็น ไม่เอาข้อความยาว ๆ
   จากไฟล์ต้นฉบับ เช่น "in:08:45 out:20:32 ลาพักร้อนครึ่งวันเช้า"
   จะเหลือแค่รหัสสั้น ๆ อ่านง่ายและไฟล์ไม่บวม                              */
var CFG_KEYS = { sch: 'backup_sheet_sch', att: 'backup_sheet_att' };

function cfgGet_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function cfgSet_(k, v) { PropertiesService.getScriptProperties().setProperty(k, String(v || '')); }

// รับได้ทั้ง URL เต็มและ ID เปล่า ๆ
function sheetIdOf_(s) {
  var t = norm_(s);
  var m = t.match(/\/d\/([a-zA-Z0-9-_]{20,})/);
  return m ? m[1] : t;
}

function yearTab_(ss, year) {
  var name = String(year);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(1);          // ปีล่าสุดอยู่ซ้ายสุด หาง่าย
  }
  return sh;
}

/* รวมข้อมูลใหม่เข้ากับของเดิมในแท็บปีนั้น
   rows = [{ id, name, position, days: { '2026-08-01': 'รหัส', ... } }]
   เขียนทับเฉพาะวันที่ส่งมา วันอื่นในแท็บยังอยู่ครบ                        */
/* Google Sheets แปลงข้อความอย่าง 2026-08-01 ให้กลายเป็นเซลล์ชนิดวันที่เอง
   พออ่านกลับมาจึงได้ Date ไม่ใช่ข้อความเดิม ถ้าไม่แปลงกลับ ระบบจะนึกว่าเป็น
   วันใหม่แล้วสร้างคอลัมน์ซ้ำเพิ่มทุกครั้งที่อัป จนคอลัมน์บานปลาย            */
function dateKey_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = norm_(v);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  var d = new Date(s);                       // เผื่อเคยถูกบันทึกเป็นข้อความวันที่แบบยาว
  if (s && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function mergeYear_(sh, rows) {
  var old = sh.getDataRange().getValues();
  var byId = {}, dates = {};
  if (old.length > 1) {
    var head = old[0];
    for (var c = 3; c < head.length; c++) { var d = dateKey_(head[c]); if (d) dates[d] = true; }
    for (var r = 1; r < old.length; r++) {
      var id = norm_(old[r][0]);
      if (!id) continue;
      var rec = { id: id, name: norm_(old[r][1]), position: norm_(old[r][2]), days: {} };
      for (var c2 = 3; c2 < head.length; c2++) {
        var dd = dateKey_(head[c2]), vv = norm_(old[r][c2]);
        if (dd && vv) rec.days[dd] = vv;     // วันเดียวกันเขียนทับกัน ไม่แตกคอลัมน์
      }
      byId[id] = rec;
    }
  }
  /* พนักงานที่หลุดออกจากไฟล์แล้ว (ลาออก หรือเคยเผลออัปข้อมูลตัวอย่างขึ้นมา)
     จะค้างอยู่ตลอดไปถ้าเก็บทุกแถวเดิมไว้หมด — จึงตัดคนที่ไม่มีในไฟล์รอบนี้
     ออกจากช่วงวันที่กำลังอัป ส่วนวันอื่นที่เขาเคยมีข้อมูลยังอยู่ครบ         */
  var incoming = {};
  (rows || []).forEach(function (n) { if (norm_(n.id)) incoming[norm_(n.id)] = true; });
  var touched = {};
  (rows || []).forEach(function (n) {
    Object.keys(n.days || {}).forEach(function (d) { touched[d] = true; });
  });
  Object.keys(byId).forEach(function (id) {
    if (incoming[id]) return;
    Object.keys(touched).forEach(function (d) { delete byId[id].days[d]; });
    // ไม่เหลือข้อมูลวันไหนเลย = ไม่ใช่พนักงานของชุดนี้ ตัดทิ้ง
    if (!Object.keys(byId[id].days).length) delete byId[id];
  });

  (rows || []).forEach(function (n) {
    var id = norm_(n.id);
    if (!id) return;
    if (!byId[id]) byId[id] = { id: id, name: norm_(n.name), position: norm_(n.position), days: {} };
    if (n.name) byId[id].name = norm_(n.name);
    if (n.position) byId[id].position = norm_(n.position);
    Object.keys(n.days || {}).forEach(function (d) {
      dates[d] = true;
      byId[id].days[d] = String(n.days[d] == null ? '' : n.days[d]);
    });
  });

  var allDates = Object.keys(dates).sort();
  var header = ['รหัสพนักงาน', 'ชื่อ', 'ตำแหน่ง'].concat(allDates);
  var out = [header];
  Object.keys(byId).sort().forEach(function (id) {
    var rec = byId[id];
    var line = [rec.id, rec.name, rec.position];
    allDates.forEach(function (d) { line.push(rec.days[d] || ''); });
    out.push(line);
  });

  sh.clear();
  // บังคับให้แถวหัวเป็นข้อความล้วน ไม่งั้น Sheets จะแปลง 2026-08-01 เป็นชนิดวันที่
  // แล้วรอบหน้าอ่านกลับมาไม่ตรง จนเกิดคอลัมน์ซ้ำ
  sh.getRange(1, 1, 1, header.length).setNumberFormat('@');
  sh.getRange(1, 1, out.length, header.length).setValues(out);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold')
    .setBackground('#201e1d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  return { แถว: out.length - 1, คอลัมน์วันที่: allDates.length };
}

function saveBackup_(kind, year, rows) {
  var id = sheetIdOf_(cfgGet_(CFG_KEYS[kind]));
  if (!id) return null;                       // ยังไม่ได้ตั้งชีตสำรองไว้ ข้ามไป
  var ss = SpreadsheetApp.openById(id);
  var res = mergeYear_(yearTab_(ss, year), rows);
  return { ไฟล์: ss.getName(), แท็บ: String(year), ผล: res };
}

/* ถอนวันที่ทั้งชุดออกจากแท็บปีในชีตสำรอง — ใช้ตอนย้อนการอัปหรือลบทั้งเดือน
   ลบทั้งคอลัมน์ของวันนั้น คนที่ไม่เหลือข้อมูลวันไหนเลยก็ตัดแถวออกด้วย     */
function dropDatesFromBackup_(kind, year, dates) {
  var id = sheetIdOf_(cfgGet_(CFG_KEYS[kind]));
  if (!id || !dates || !dates.length) return null;
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName(String(year));
  if (!sh) return null;
  var old = sh.getDataRange().getValues();
  if (old.length < 2) return null;

  var kill = {};
  dates.forEach(function (d) { var k = dateKey_(d); if (k) kill[k] = true; });
  var head = old[0];
  var keep = [0, 1, 2];                       // รหัส / ชื่อ / ตำแหน่ง เก็บไว้เสมอ
  for (var c = 3; c < head.length; c++) {
    if (!kill[dateKey_(head[c])]) keep.push(c);
  }
  var out = [];
  for (var r = 0; r < old.length; r++) {
    var line = keep.map(function (c) { return old[r][c]; });
    // แถวหัวตารางเก็บไว้เสมอ ส่วนแถวคนต้องเหลือข้อมูลอย่างน้อยหนึ่งวัน
    if (r === 0 || line.slice(3).some(function (v) { return norm_(v); })) out.push(line);
  }
  sh.clear();
  sh.getRange(1, 1, 1, keep.length).setNumberFormat('@');
  sh.getRange(1, 1, out.length, keep.length).setValues(out);
  sh.getRange(1, 1, 1, keep.length).setFontWeight('bold')
    .setBackground('#201e1d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  return { ไฟล์: ss.getName(), แท็บ: String(year),
    ถอนวันที่: head.length - keep.length, เหลือแถว: out.length - 1 };
}

/* ลบข้อมูลทั้งเดือนออกจากทุกที่ที่เก็บไว้ */
function dropMonth_(ss, month) {
  var sh = dataSheet_(ss);
  var ds = readDataset_(ss, month);
  var dates = (ds && ds.dates) || [];
  deleteKey_(sh, 'ds:' + month);
  var idx = JSON.parse(readKey_(sh, 'index') || '[]').filter(function (x) { return x.month !== month; });
  writeKey_(sh, 'index', JSON.stringify(idx));
  var done = ['ข้อมูลเดือน ' + month];
  var year = String(month).slice(0, 4);
  ['att', 'sch'].forEach(function (kind) {
    try {
      var r = dropDatesFromBackup_(kind, year, dates);
      if (r) done.push('ชีตสำรอง ' + r.ไฟล์ + ' แท็บ ' + r.แท็บ + ' (ถอน ' + r.ถอนวันที่ + ' วัน)');
    } catch (e) { done.push('ชีตสำรอง ' + kind + ' ถอนไม่สำเร็จ: ' + String(e)); }
  });
  // ไม่เหลือเดือนไหนเลย = ล้างตารางที่คนอ่านได้ด้วย ไม่งั้นจะค้างของเดือนที่ลบไปแล้ว
  if (!idx.length) {
    ['Matrix', 'Summary by position', 'Daily rate', 'Warning list'].forEach(function (n) {
      var t = ss.getSheetByName(n);
      if (t) { t.clear(); done.push('ล้าง ' + n); }
    });
  }
  return { done: done, months: idx };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
