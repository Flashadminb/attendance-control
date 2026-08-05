// HCM attendance engine — xlsx reading, cell classification, roll-ups, demo data.

/* ── xlsx (zip + sheetXML) reader, no dependencies ───────────────────────── */
async function unzip(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('ไม่ใช่ไฟล์ .xlsx');
  const n = dv.getUint16(eocd + 10, true);
  let off = dv.getUint32(eocd + 16, true);
  const dir = {};
  for (let i = 0; i < n; i++) {
    const nl = dv.getUint16(off + 28, true), el = dv.getUint16(off + 30, true), cl = dv.getUint16(off + 32, true);
    const name = new TextDecoder().decode(buf.slice(off + 46, off + 46 + nl));
    dir[name] = { lho: dv.getUint32(off + 42, true), method: dv.getUint16(off + 10, true), csize: dv.getUint32(off + 20, true) };
    off += 46 + nl + el + cl;
  }
  const out = {};
  for (const [name, f] of Object.entries(dir)) {
    const nl = dv.getUint16(f.lho + 26, true), el = dv.getUint16(f.lho + 28, true);
    const s = f.lho + 30 + nl + el;
    const raw = buf.slice(s, s + f.csize);
    out[name] = f.method === 0
      ? new TextDecoder().decode(raw)
      : await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
  }
  return out;
}

const colIndex = (ref) => {
  let n = 0;
  for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

export async function readSheet(file) {
  const buf = new Uint8Array(await file.arrayBuffer());
  if (/\.csv$/i.test(file.name)) {
    const text = new TextDecoder().decode(buf);
    return text.split(/\r?\n/).filter(Boolean).map(l => l.split(',').map(c => c.replace(/^"|"$/g, '')));
  }
  const zip = await unzip(buf);
  const ssXml = zip['xl/sharedStrings.xml'] || '';
  const strs = [...ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)]
    .map(m => unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('')));
  const sheetName = Object.keys(zip).find(k => /^xl\/worksheets\/sheet1\.xml$/.test(k)) ||
    Object.keys(zip).find(k => /^xl\/worksheets\//.test(k));
  const xml = zip[sheetName];
  const rows = [];
  for (const r of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
    const row = [];
    for (const c of r[1].matchAll(/<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
      const v = (c[3].match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      const inl = (c[3].match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
      row[colIndex(c[1])] = /t="s"/.test(c[2]) && v != null ? strs[+v] : unesc(inl ?? v ?? '');
    }
    for (let i = 0; i < row.length; i++) if (row[i] == null) row[i] = '';
    rows.push(row);
  }
  return rows;
}

/* ── header mapping ──────────────────────────────────────────────────────── */
const FIELD_HINTS = [
  ['id', /รหัสพนักงาน|employee.?id|emp.?code/i],
  ['name', /ชื่อ/i],
  ['position', /ตำแหน่ง|position/i],
  ['empStatus', /สถานะพนักงาน/i],
  ['area', /^พื้นที่/i],
  ['zone', /^เขต/i],
  ['branch', /สาขา$|^สาขา|branch|hub$/i],
  ['region', /ภูมิภาค|region/i],
  ['branchType', /ประเภทสาขา/i],
  ['dept', /แผนก|department/i],
  ['absentDays', /ขาดงาน/i],
  ['lateCount', /จำนวนครั้งที่สาย/i],
  ['lateMin', /เวลาที่สาย/i],
  ['earlyCount', /จำนวนครั้งที่ออกงานก่อน/i],
  ['earlyMin', /ออกงานก่อนเวลา\(/i],
];
const isDateHeader = (h) => /^\d{4}-\d{2}-\d{2}/.test(h) || /^\d{1,2}$/.test(String(h).trim()) ||
  /^\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?$/.test(String(h).trim());

// ไฟล์ attendance กับไฟล์ตารางกะที่ HCM ส่งออกมา ใช้รูปแบบวันที่คนละแบบ
// (2026-08-01 กับ 01/08/2026) ถ้าไม่ทำให้เป็นแบบเดียวกันก่อน จะจับคู่กะกับวันไม่ได้เลย
// วันที่ไทยเป็น วัน/เดือน/ปี — ยืนยันจากไฟล์จริงที่มีคอลัมน์ 31/08/2026 (31 เป็นเดือนไม่ได้)
export function normalizeDateKey(h) {
  const t = String(h).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(t)) return t.slice(0, 10);
  const m = t.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = '20' + y;
    return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }
  return t;   // เลขวันโดด ๆ (1, 2, 3…) ไม่มีเดือนให้อ้างอิง ปล่อยไว้ตามเดิม
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = rows[i].map(c => String(c || ''));
    if (r.some(c => /รหัสพนักงาน|employee.?id/i.test(c)) && r.filter(isDateHeader).length >= 3) return i;
  }
  return 1;
}

export function mapTable(rows) {
  const hIdx = findHeaderRow(rows);
  const header = rows[hIdx].map(c => String(c || '').trim());
  const fields = {}, dates = [];
  header.forEach((h, i) => {
    if (isDateHeader(h)) { dates.push({ i, key: normalizeDateKey(h) }); return; }
    for (const [f, re] of FIELD_HINTS) if (fields[f] === undefined && re.test(h)) { fields[f] = i; break; }
  });
  const records = rows.slice(hIdx + 1)
    .filter(r => String(r[fields.id ?? 0] || '').trim())
    .map(r => {
      const rec = { days: {} };
      for (const [f, i] of Object.entries(fields)) rec[f] = String(r[i] ?? '').trim();
      dates.forEach(d => { rec.days[d.key] = String(r[d.i] ?? '').trim(); });
      return rec;
    });
  return { dates: dates.map(d => d.key), records, headerRow: hIdx, columns: header };
}

/* ── classification ──────────────────────────────────────────────────────── */
export const LEAVE_TYPES = [
  { re: /ลาพักร้อน/, code: 'AL', label: 'ลาพักร้อน' },
  { re: /ลาป่วย/, code: 'SL', label: 'ลาป่วย' },
  { re: /ลากิจ/, code: 'PL', label: 'ลากิจ' },
  { re: /ลาไม่รับค่าจ้าง|ไม่รับค่าจ้าง/, code: 'UL', label: 'ลาไม่รับค่าจ้าง' },
  { re: /ลาคลอด/, code: 'ML', label: 'ลาคลอดบุตร' },
  { re: /ลาอุปสมบท|ลาบวช/, code: 'OL', label: 'ลาอุปสมบท' },
];

const toMin = (s) => { const m = /(\d{1,2}):(\d{2})/.exec(s || ''); return m ? +m[1] * 60 + +m[2] : null; };
export const fmtMin = (m) => `${String(Math.floor(m / 60) % 24).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;

export function parseShift(text) {
  if (!text) return null;
  const t = String(text).trim();
  if (!t || /วันหยุด|OFF|-$/i.test(t) && !/\d/.test(t)) return null;
  const m = t.match(/(\d{1,2}[:.]\d{2})\s*[-–]\s*(\d{1,2}[:.]\d{2})/);
  if (!m) return null;
  let s = toMin(m[1].replace('.', ':')), e = toMin(m[2].replace('.', ':'));
  if (e <= s) e += 1440;
  return { name: t.replace(m[0], '').trim() || 'กะ', start: s, end: e, raw: t };
}

// → { color: 'red'|'yellow'|'gray'|'gray-half'|'ok', color2, code, tip, late, early, leave, pending }
export function classify(cellText, shiftText, cfg, dateKey) {
  const t = String(cellText || '').trim();
  const shift = parseShift(shiftText);
  const grace = cfg.grace ?? 5;
  const earlyGrace = cfg.earlyGrace ?? 5;
  const isPH = (cfg.holidays || []).includes(dateKey);
  const pending = /รอการอนุมัติ/.test(t);
  const leave = LEAVE_TYPES.find(l => l.re.test(t));
  const half = /ครึ่งวัน/.test(t) ? (/บ่าย/.test(t) ? 'บ่าย' : 'เช้า') : null;
  const scan = t.match(/in\s*:\s*([^\s]+)\s+out\s*:\s*([^\s]+)/i);
  const inRaw = scan ? scan[1] : '', outRaw = scan ? scan[2] : '';
  const inM = /^\d{1,2}:\d{2}/.test(inRaw) ? toMin(inRaw) : null;
  const outM = /^\d{1,2}:\d{2}/.test(outRaw) ? toMin(outRaw) : null;
  // แสกนมาข้างเดียว (มีแค่ in หรือ out) — ตัดสินตรงนี้ก่อนแยกสาขา เพราะสาขา NoS/PH
  // คืนค่าออกไปก่อนจะถึงจุดที่ตรวจ INC ทำให้ช่องพวกนั้นรอดกฎ "แสกนไม่ครบ = ขาดงาน" ไปได้
  const inc = !!scan && (inM == null) !== (outM == null);
  const R = { code: '', tip: t || '—', late: 0, early: 0, leave: null, pending, inc, shift: shift ? shift.raw : '', in: inM, out: outM };

  // กฎแปลงรหัสที่แอดมินเพิ่มเอง — มาก่อนกฎมาตรฐาน เพื่อให้เขียนทับได้
  // ใช้รองรับคำใหม่ ๆ ที่ HCM ส่งมาโดยที่ระบบยังไม่รู้จัก เช่น "ทำงานนอกสถานที่"
  if (t) {
    const cu = (cfg.customCodes || []).find(c => c && c.on !== false && c.match && t.includes(c.match));
    if (cu) return { ...R, color: cu.color || 'gray', code: cu.code || '',
      tip: (cu.label || cu.match) + (pending ? ' (รอการอนุมัติ)' : ''),
      customPresent: cu.present !== false };
  }

  if (isPH) return { ...R, color: 'gray', code: 'PH', tip: 'วันหยุดนักขัตฤกษ์' };

  if (/วันหยุด/.test(t) && !scan) {
    if (shift && cfg.flagShiftOnHoliday !== false)
      return { ...R, color: 'red', code: 'OFF!', tip: `ตั้งกะ ${shift.raw} ในวันหยุด แต่ไม่มีการแสกน` };
    return { ...R, color: 'gray', code: 'OFF', tip: 'วันหยุดประจำสัปดาห์' };
  }

  if (leave && !half) {
    return { ...R, color: 'gray', code: leave.code + (pending ? '?' : ''), leave: leave.code,
      tip: leave.label + (pending ? ' (รอการอนุมัติ)' : '') };
  }

  if (!t) return shift
    ? { ...R, color: 'red', code: 'NS', tip: `มีกะ ${shift.raw} แต่ไม่มีข้อมูลการแสกน` }
    : { ...R, color: 'gray', code: '', tip: 'ไม่มีข้อมูล' };

  // half-day leave — base gray-half, then judge the worked half
  const base = half ? { color: 'gray-half', code: (leave ? leave.code : 'L') + '½', leave: leave ? leave.code : null,
    tip: `${leave ? leave.label : 'ลา'}ครึ่งวัน${half}${pending ? ' (รอการอนุมัติ)' : ''}` } : null;

  if (/ขาดงาน/.test(t) && !scan)
    return { ...R, color: 'red', code: 'AB', tip: 'ขาดงาน' };

  if (!shift && cfg.flagNoShift !== false && (scan || /ขาดงาน/.test(t)))
    return { ...R, color: 'red', code: 'NoS', tip: 'ไม่มีกะทำงานที่กำหนดไว้' };

  if (scan) {
    if (inM == null && outM == null) {
      if (base) return { ...R, ...base, tip: base.tip + ' · ไม่มีการแสกนในครึ่งวันที่เหลือ', color2: 'red' };
      return { ...R, color: 'red', code: /ขาดงาน/.test(t) ? 'AB' : 'NS', tip: 'ไม่มีการแสกนหน้า' };
    }
    if (inM == null || outM == null) {
      const c = { ...R, color: 'red', code: 'INC', tip: `แสกนไม่ครบ (${inM == null ? 'ไม่มี in' : 'ไม่มี out'})` };
      return base ? { ...R, ...base, color: 'gray-half', color2: 'red', code: base.code, tip: base.tip + ' · แสกนไม่ครบ' } : c;
    }
    let late = 0, early = 0;
    if (shift) {
      let i = inM, o = outM;
      if (i < shift.start - 720) i += 1440;
      if (o < i) o += 1440;
      // half-day leave: judge only the half actually worked
      const mid = Math.round((shift.start + shift.end) / 2);
      const expIn = half === 'เช้า' ? mid : shift.start;
      const expOut = half === 'บ่าย' ? mid : shift.end;
      late = Math.max(0, i - expIn);
      early = Math.max(0, expOut - o);
    }
    const flags = [];
    if (late > grace) flags.push(`สาย ${late} น.`);
    if (early > earlyGrace) flags.push(`ออกก่อน ${early} น.`);
    const timeTip = `เข้า ${fmtMin(inM)} · ออก ${fmtMin(outM)}${shift ? ` · กะ ${shift.raw}` : ''}`;
    if (flags.length) {
      const code = late > grace ? `L${late}` : `E${early}`;
      if (base) return { ...R, ...base, color: 'gray-half', color2: 'yellow', code: base.code + '·' + code, late, early, tip: `${base.tip} · ${flags.join(' · ')}` };
      return { ...R, color: 'yellow', code, late, early, tip: `${flags.join(' · ')} — ${timeTip}` };
    }
    if (base) return { ...R, ...base, tip: `${base.tip} · ${timeTip}` };
    return { ...R, color: 'ok', code: '', tip: timeTip };
  }

  if (base) return { ...R, ...base };
  return { ...R, color: 'gray', code: '', tip: t };
}

/* ── HCM sheet code (identical order to the source sheet's IFS ladder) ──── */
export function hcmCode(text) {
  const t = String(text || '');
  if (!t.trim()) return '';
  if (/[023456789]/.test(t)) return '1';              // any real clock digit → present
  if (/วันหยุด\s*PH/.test(t)) return 'PH';
  if (/ตั้งค่าทำงานในวันหยุด/.test(t)) return 'PH';
  if (/วันหยุด/.test(t)) return 'OFF';
  if (/ลากิจ/.test(t)) return 'PL';
  if (/ลาพักร้อน/.test(t)) return 'AL';
  if (/ลาป่วย/.test(t)) return 'SL';
  if (/ลาไม่รับค่าจ้าง/.test(t)) return 'UL';
  if (/ลาคลอดบุตร/.test(t)) return 'ML';
  if (/ขาดงาน/.test(t)) return 'AB';
  return '';
}
export const HCM_CODES = ['OFF', 'SL', 'PL', 'AB', 'AL', 'UL', 'PH', 'ML'];

/* ── "มาทำงาน" สำหรับคิด Attendance Rate ─────────────────────────────────
   แสกนไม่ครบ (มีแค่ in หรือ out) ไม่นับเป็นมาทำงาน — คิดเท่ากับขาดงาน
   แต่ยังโชว์รหัส INC และสีแดงตามเดิม
   ถ้าพนักงานยื่นแสกนย้อนหลังจนอนุมัติแล้ว ไฟล์รอบใหม่จาก HCM จะมีเวลาครบ
   พออัปโหลดทับ สถานะจะกลับมาเป็นมาทำงานเองโดยไม่ต้องแก้อะไร               */
export const isIncompleteScan = (c) => c.inc === true || c.code === 'INC' || c.color2 === 'red';
export const isPresent = (c) => c.hcm === '1' && !isIncompleteScan(c);

/* ── build the model ─────────────────────────────────────────────────────── */
export function buildModel(attRows, schedRows, cfg) {
  const att = mapTable(attRows);
  const sched = schedRows ? mapTable(schedRows) : null;
  const shiftById = {};
  const defaultShift = {};
  if (sched) sched.records.forEach(r => {
    shiftById[r.id] = r.days;
    // กะประจำตัว = กะที่คนนั้นขึ้นบ่อยที่สุดในไฟล์ตารางกะ ใช้เป็นค่าตั้งต้นของวันที่ไฟล์ยังไม่ครอบคลุม
    // เพราะแอดมินอัปตารางกะเฉพาะตอนมีพนักงานใหม่หรือเปลี่ยนกะ ไม่ได้อัปทุกเดือน
    const freq = {};
    Object.values(r.days).forEach(v => {
      const s = String(v || '').trim();
      if (s && parseShift(s)) freq[s] = (freq[s] || 0) + 1;
    });
    const top = Object.entries(freq).sort((a, b) => b[1] - a[1])[0];
    if (top) defaultShift[r.id] = top[0];
  });
  // ไฟล์ที่ดึงกลางเดือนจะมีคอลัมน์ของวันที่ยังไม่ถึงติดมาด้วย ว่างทั้งคอลัมน์ —
  // ถ้านับรวมจะไปบวมอยู่ในตัวหารของ Attendance Rate แล้วอัตราตกผิด ๆ
  // จึงตัดทิ้งตั้งแต่วันสุดท้ายที่มีข้อมูลจริงเป็นต้นไป
  // (วันว่างที่แทรกอยู่กลางช่วงยังเก็บไว้ เพราะนั่นคือข้อมูลขาดจริงที่ต้องเห็น)
  let lastData = -1;
  att.dates.forEach((d, i) => { if (att.records.some(r => (r.days[d] || '').trim())) lastData = i; });
  // วันสุดท้ายอาจเป็น "วันที่กะยังไม่จบ" — คนสแกนเข้าแล้วแต่ยังไม่ถึงเวลาเลิกงาน
  // ทั้งวันจึงไม่มีใครสแกนครบสักคน ถ้านับเข้าไปจะกลายเป็นขาดงานยกแผนก
  // ตัดออกจนกว่าจะมีคนสแกนครบอย่างน้อย 1 คน แล้วรอบอัปวันถัดไปจะดึงกลับมาเอง
  const hasFullScan = (i) => {
    const d = att.dates[i];
    return att.records.some(r => /in\s*:\s*\d{1,2}:\d{2}\s+out\s*:\s*\d{1,2}:\d{2}/i.test(String(r.days[d] || '')));
  };
  let lastClosed = lastData;
  if (cfg.dropOpenDay !== false && lastClosed >= 0 && !hasFullScan(lastClosed)) lastClosed--;
  const dates = att.dates.slice(0, lastClosed + 1);
  const openDay = lastClosed < lastData ? att.dates[lastData] : null;
  const employees = att.records.map(r => {
    const shifts = shiftById[r.id] || {};
    const cells = dates.map(d => {
      const raw = r.days[d];
      let sh = shifts[d];
      // วันที่ไฟล์ตารางกะไม่ครอบคลุม → ใช้กะประจำตัวแทน แต่ข้ามวันที่ระบบลงว่าหยุด/ลาไว้แล้ว
      // ไม่งั้นจะไปติดธง OFF! ("ตั้งกะในวันหยุดแต่ไม่แสกน") ทั้งที่เป็นวันหยุดจริง
      if (!String(sh || '').trim() && defaultShift[r.id] && !/วันหยุด|ลา/.test(String(raw || '')))
        sh = defaultShift[r.id];
      const c = classify(raw, sh, cfg, d);
      // แถวที่แอดมินเพิ่มเองเป็นคนกำหนดเองว่านับเป็นมาทำงานไหม — เก็บลง hcm เพื่อให้
      // isPresent และการบันทึก/อ่านกลับใช้ช่องเดิม ไม่ต้องเปลี่ยนรูปแบบไฟล์ที่เก็บไว้
      c.hcm = c.customPresent === undefined ? hcmCode(raw) : (c.customPresent ? '1' : '');
      return c;
    });
    const stat = { absent: 0, late: 0, lateMin: 0, maxLate: 0, early: 0, earlyMin: 0, leave: 0, incomplete: 0, off: 0, worked: 0, red: 0, present: 0, codes: {} };
    HCM_CODES.forEach(k => stat.codes[k] = 0);
    cells.forEach(c => { if (isPresent(c)) stat.present++; else if (stat.codes[c.hcm] !== undefined) stat.codes[c.hcm]++; });
    cells.forEach(c => {
      if (c.color === 'red') { stat.red++; if (c.code === 'AB' || c.code === 'NS') stat.absent++; if (c.code === 'INC') stat.incomplete++; }
      if (c.color === 'yellow' || c.color2 === 'yellow') { if (c.late > (cfg.grace ?? 5)) { stat.late++; stat.lateMin += c.late; if (c.late > stat.maxLate) stat.maxLate = c.late; } if (c.early > (cfg.earlyGrace ?? 5)) { stat.early++; stat.earlyMin += c.early; } }
      if (c.leave) stat.leave++;
      if (c.code === 'OFF') stat.off++;
      if (c.in != null && c.out != null) stat.worked++;
    });
    return { ...r, shifts, cells, stat };
  });
  return { dates, employees, columns: att.columns, hasSchedule: !!sched, openDay };
}

export function hubRollup(employees) {
  const map = new Map();
  employees.forEach(e => {
    const k = e.position || '—';
    if (!map.has(k)) map.set(k, { branch: k, region: '', type: '', headcount: 0, absent: 0, late: 0, leave: 0, lateMin: 0, incomplete: 0, worked: 0, slots: 0, present: 0, personDays: 0 });
    const h = map.get(k);
    h.headcount++; h.absent += e.stat.absent; h.late += e.stat.late; h.leave += e.stat.leave;
    h.lateMin += e.stat.lateMin; h.incomplete += e.stat.incomplete; h.worked += e.stat.worked;
    h.slots += e.cells.filter(c => c.color !== 'gray' && c.code !== 'OFF').length;
    h.present += e.cells.filter(isPresent).length;
    h.personDays += e.cells.filter(c => c.color !== 'out').length;
  });
  return [...map.values()].map(h => ({ ...h, attRate: attRate(h), onTime: h.slots ? Math.round((1 - (h.absent + h.late + h.incomplete) / h.slots) * 1000) / 10 : 100 }))
    .sort((a, b) => a.onTime - b.onTime);
}

/* ── demo data (built from the real HCM vocabulary) ──────────────────────── */
const POSITIONS = ['Hub Staff', 'Hub Staff Leader', 'Hub Supervisor', 'Hub Forklift Driver', 'Hub Quality Control Specialist', 'Hub Quality Control Officer', 'Hub Admin Officer', 'Hub Admin Supervisor', 'Hub Area Manager', 'Hub Deputy Manager', 'Mini CS Officer', 'Industrial Engineering Maintenance', 'Hub Standardization Officer'];
const BRANCHES = [['21 BPL_BHUB-บางพลี', 'BKK', 'B-Hub']];
const FIRST_M = ['นาย สมชาย', 'นาย ธนากร', 'นาย พงษ์ศักดิ์', 'นาย วีระพล', 'นาย อนุชา', 'นาย จีรพงษ์', 'นาย สายัณห์', 'นาย นุกูล'];
const FIRST_F = ['นางสาว ดวงใจ', 'นางสาว เปมิกา', 'นางสาว วนิดา', 'นางสาว ณัฐกานต์', 'นางสาว รุ่งธิวา', 'นางสาว สุภาพร', 'นางสาว กมลชนก', 'นาง อรพิน'];
const LAST = ['จินดาศรี', 'จำปา', 'เข็มเพชร', 'สระโร', 'ฉิมพลีสวรรค์', 'จันทร์สิงห์', 'แจ้งไพร', 'แบบทอง', 'ศรีสุวรรณ', 'บุญมาก', 'ทองอินทร์', 'พรมมา', 'แก้วมณี', 'สุขใจ'];
const SHIFTS = [['กะเช้า', '09:00-18:00'], ['กะบ่าย', '13:00-22:00'], ['กะดึก', '21:00-06:00'], ['กะเช้าตรู่', '06:00-15:00']];

function rng(seed) { return () => (seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296; }

export function demoData(count = 640, year = 2026, month = 8) {
  const r = rng(20260803);
  const days = new Date(year, month, 0).getDate();
  const dates = Array.from({ length: days }, (_, i) => `${year}-${String(month).padStart(2, '0')}-${String(i + 1).padStart(2, '0')}`);
  const head = ['รหัสพนักงาน', 'ชื่อ-สกุล', 'ตำแหน่ง', 'สถานะพนักงาน', 'สาขา', 'ภูมิภาค', 'ประเภทสาขา', 'จำนวนวันที่ขาดงาน', 'จำนวนครั้งที่สาย', 'เวลาที่สาย(นาที)', 'จำนวนครั้งที่ออกงานก่อนเวลา', 'ออกงานก่อนเวลา(นาที)'];
  const att = [['คำอธิบายการแสดงผลของสี: สีแดง: ขาดงาน/ไม่มีการแสกนหน้า/ไม่มีกะทำงาน/แสกนไม่ครบ; สีเหลือง: มาสายหรือออกงานก่อนเวลา; สีเทา: การลาหยุด/วันหยุด/PH'], [...head, ...dates]];
  const sch = [['ตารางกะการทำงาน'], ['รหัสพนักงาน', 'ชื่อ', 'แผนก', 'ตำแหน่ง', 'สาขา', ...dates]];
  const pad = (n) => String(n).padStart(2, '0');
  for (let i = 0; i < count; i++) {
    const b = BRANCHES[Math.floor(r() * BRANCHES.length)];
    const pos = POSITIONS[Math.floor(r() * (r() < 0.6 ? 4 : POSITIONS.length))];
    const name = (r() < 0.5 ? FIRST_M[Math.floor(r() * 8)] : FIRST_F[Math.floor(r() * 8)]) + ' ' + LAST[Math.floor(r() * LAST.length)];
    const id = 39000 + Math.floor(r() * 60000);
    const sh = SHIFTS[Math.floor(r() * (r() < 0.55 ? 1 : SHIFTS.length))];
    const discipline = r();           // per-person reliability
    const offA = Math.floor(r() * 7), offB = (offA + 3) % 7;
    const aRow = [id, name, pos, 'พนักงานปัจจุบัน', b[0], b[1], b[2], 0, 0, 0, 0, 0];
    const sRow = [id, name, 'Hub Operation', pos, b[0]];
    const [sHH, sMM] = sh[1].split('-')[0].split(':').map(Number);
    const [eHH, eMM] = sh[1].split('-')[1].split(':').map(Number);
    let absent = 0, lateN = 0, lateM = 0, earlyN = 0, earlyM = 0;
    dates.forEach((d, di) => {
      const dow = new Date(year, month - 1, di + 1).getDay();
      const isOff = dow === offA || dow === offB;
      const p = r();
      if (isOff) {
        if (p < 0.06) { sRow.push(sh[0] + sh[1]); aRow.push('วันหยุด'); }        // shift set on an off day
        else { sRow.push('วันหยุด'); aRow.push('วันหยุด'); }
        return;
      }
      sRow.push(sh[0] + sh[1]);
      const q = r();
      if (q < 0.030 + (1 - discipline) * 0.05) { aRow.push('ขาดงาน'); absent++; return; }
      if (q < 0.075) { const L = LEAVE_TYPES[Math.floor(r() * 4)]; aRow.push(`${L.label}1วัน${r() < 0.25 ? 'รอการอนุมัติ' : ''}`); return; }
      if (q < 0.095) {
        const startT = sHH * 60 + sMM, endT = (eHH * 60 + eMM > startT ? eHH * 60 + eMM : eHH * 60 + eMM + 1440), midT = Math.round((startT + endT) / 2);
        const morningOff = r() < 0.5;   // leave in the morning → works the afternoon half
        const inT = (morningOff ? midT : startT) + Math.floor(r() * 14) - 4;
        const outT = (morningOff ? endT : midT) + Math.floor(r() * 12) - 6;
        aRow.push(`in:${fmtMin((inT + 1440) % 1440)} out:${fmtMin((outT + 1440) % 1440)}  ลาพักร้อนครึ่งวัน${morningOff ? 'เช้า' : 'บ่าย'}`);
        return;
      }
      if (q < 0.125) { aRow.push(`in:${pad(sHH)}:${pad(Math.floor(r() * 30))} out:--`); return; }
      if (q < 0.14) { aRow.push('in:-- out:--'); return; }
      const late = r() < 0.16 + (1 - discipline) * 0.25 ? Math.floor(r() * 55) + 3 : Math.floor(r() * 8) - 6;
      const early = r() < 0.09 ? Math.floor(r() * 40) + 3 : -Math.floor(r() * 30);
      const inT = sHH * 60 + sMM + late, outT = eHH * 60 + eMM - early;
      if (late > 5) { lateN++; lateM += late; }
      if (early > 5) { earlyN++; earlyM += early; }
      aRow.push(`in:${fmtMin((inT + 1440) % 1440)} out:${fmtMin((outT + 1440) % 1440)}`);
    });
    aRow[7] = absent; aRow[8] = lateN; aRow[9] = lateM; aRow[10] = earlyN; aRow[11] = earlyM;
    att.push(aRow.map(String));
    sch.push(sRow.map(String));
  }
  return { att, sch };
}

/* ── exporters ───────────────────────────────────────────────────────────── */
export const CELL_HEX = { red: 'EC3013', yellow: 'E8A400', gray: 'D7D3D3', 'gray-half': 'EAE7E7', ok: 'FFFFFF' };

export function toCSV(header, rows) {
  const q = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  return '\uFEFF' + [header, ...rows].map(r => r.map(q).join(',')).join('\r\n');
}

export function toSpreadsheetML(header, rows, colorRows, meta) {
  const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const styles = Object.entries(CELL_HEX).map(([k, hex]) =>
    `<Style ss:ID="c_${k.replace('-', '_')}"><Interior ss:Color="#${hex}" ss:Pattern="Solid"/><Font ss:Color="${k === 'red' ? '#FFFFFF' : '#201E1D'}" ss:Size="9"/><Alignment ss:Horizontal="Center"/></Style>`).join('');
  const body = rows.map((r, ri) => '<Row>' + r.map((v, ci) => {
    const color = colorRows[ri] && colorRows[ri][ci];
    return `<Cell${color ? ` ss:StyleID="c_${color.replace('-', '_')}"` : ''}><Data ss:Type="String">${esc(v)}</Data></Cell>`;
  }).join('') + '</Row>').join('');
  return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles><Style ss:ID="hd"><Font ss:Bold="1" ss:Size="9"/><Interior ss:Color="#201E1D" ss:Pattern="Solid"/><Font ss:Color="#FFFFFF" ss:Bold="1" ss:Size="9"/></Style>${styles}</Styles>
<Worksheet ss:Name="${esc(meta || 'Attendance')}"><Table><Row>${header.map(h => `<Cell ss:StyleID="hd"><Data ss:Type="String">${esc(h)}</Data></Cell>`).join('')}</Row>${body}</Table></Worksheet></Workbook>`;
}

export function download(name, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a = document.createElement('a');
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

/* ── shift buckets ───────────────────────────────────────────────────────── */
export const BUCKETS = [['morning', 'กะเช้า'], ['afternoon', 'กะบ่าย'], ['night', 'กะดึก'], ['none', 'ไม่ได้ตั้งกะ']];
export function shiftBucket(raw) {
  if (!raw) return 'none';
  const m = String(raw).match(/(\d{1,2})[:.](\d{2})/);
  if (!m) return 'none';
  const h = +m[1];
  if (h >= 4 && h < 12) return 'morning';
  if (h >= 12 && h < 18) return 'afternoon';
  return 'night';
}

/* ── monthly snapshot for the archive ────────────────────────────────────── */
const blankAgg = () => ({ slots: 0, ok: 0, absent: 0, late: 0, lateMin: 0, early: 0, leave: 0, inc: 0, heads: 0, present: 0, personDays: 0 });
const rate = (a) => a.slots ? Math.round(a.ok / a.slots * 1000) / 10 : 100;
export const attRate = (a) => a.personDays ? Math.round(a.present / a.personDays * 1000) / 10 : 0;

export function snapshot(model, cfg, source) {
  const g = cfg.grace ?? 5, eg = cfg.earlyGrace ?? 5;
  const month = (model.dates[0] || '').slice(0, 7);
  const total = blankAgg();
  const byDay = model.dates.map(d => ({ date: d, ...blankAgg() }));
  const byShift = {}; BUCKETS.forEach(([k]) => byShift[k] = blankAgg());
  const byBranch = new Map();
  const warn = [];
  model.employees.forEach(e => {
    const b = e.position || '—';
    if (!byBranch.has(b)) byBranch.set(b, { branch: b, region: '', ...blankAgg() });
    const B = byBranch.get(b); B.heads++;
    total.heads++;
    e.cells.forEach((c, i) => {
      // Attendance Rate (as in the source sheet): present ÷ every employee-day, holidays included
      total.personDays++; byDay[i].personDays++; B.personDays++;
      const bkt0 = byShift[shiftBucket(c.shift)]; bkt0.personDays++;
      if (isPresent(c)) { total.present++; byDay[i].present++; B.present++; bkt0.present++; }
      if (c.code === 'OFF' || c.code === 'PH') return;
      const bucket = byShift[shiftBucket(c.shift)];
      const D = byDay[i];
      const hit = (A) => {
        A.slots++;
        if (c.color === 'ok' || c.leave) A.ok++;
        if (c.leave) A.leave++;
        if (c.code === 'AB' || c.code === 'NS' || c.code === 'NoS' || c.code === 'OFF!') A.absent++;
        if (c.code === 'INC') A.inc++;
        if (c.late > g) { A.late++; A.lateMin += c.late; }
        if (c.early > eg) A.early++;
      };
      hit(total); hit(D); hit(bucket); hit(B);
    });
    if (e.stat.absent >= (cfg.warnAbsent ?? 3) || e.stat.late >= (cfg.warnLate ?? 5))
      warn.push({ id: e.id, name: e.name, branch: e.branch, position: e.position, absent: e.stat.absent, late: e.stat.late, lateMin: e.stat.lateMin });
  });
  return {
    month, savedAt: new Date().toISOString(), source: source || 'upload',
    days: model.dates.length, employees: model.employees.length,
    total: { ...total, onTime: rate(total), attRate: attRate(total) },
    byDay: byDay.map(d => ({ ...d, onTime: rate(d), attRate: attRate(d) })),
    byShift: Object.fromEntries(Object.entries(byShift).map(([k, v]) => [k, { ...v, onTime: rate(v), attRate: attRate(v) }])),
    byBranch: [...byBranch.values()].map(b => ({ ...b, onTime: rate(b), attRate: attRate(b) })).sort((a, b) => a.attRate - b.attRate),
    warn: warn.sort((a, b) => (b.absent * 10 + b.late) - (a.absent * 10 + a.late)).slice(0, 20),
  };
}

/* Synthetic multi-year history so the archive has something to read on day one. */
export function demoArchive(endYear, endMonth, months) {
  const r = rng(778899);
  const out = [];
  for (let k = months; k >= 1; k--) {
    const dt = new Date(endYear, endMonth - 1 - k, 1);
    const y = dt.getFullYear(), mo = dt.getMonth() + 1;
    const days = new Date(y, mo, 0).getDate();
    const season = 1 + Math.sin((mo - 3) / 12 * Math.PI * 2) * 0.22;
    const drift = 1 + (months - k) * 0.004;
    const heads = 560 + Math.floor(r() * 120);
    const mk = (scale) => {
      const slots = Math.round(heads * days * 0.72 * scale);
      const absent = Math.round(slots * 0.031 * season / drift);
      const late = Math.round(slots * 0.14 * season / drift);
      const early = Math.round(slots * 0.035 * season);
      const leave = Math.round(slots * 0.052);
      const inc = Math.round(slots * 0.021 * season);
      const ok = slots - absent - late - early - inc;
      const personDays = Math.round(heads * days * scale);
      const present = Math.round(personDays * (0.55 + r() * 0.06) / (season > 1 ? season : 1));
      return { slots, ok, absent, late, lateMin: late * (9 + Math.round(r() * 8)), early, leave, inc, heads: Math.round(heads * scale), personDays, present,
        onTime: Math.round(ok / slots * 1000) / 10, attRate: Math.round(present / personDays * 1000) / 10 };
    };
    const total = mk(1);
    const byDay = Array.from({ length: days }, (_, i) => {
      const d = `${y}-${String(mo).padStart(2, '0')}-${String(i + 1).padStart(2, '0')}`;
      const dow = new Date(y, mo - 1, i + 1).getDay();
      const f = (dow === 1 ? 1.25 : dow === 5 ? 1.18 : 1) * (0.85 + r() * 0.3);
      const slots = Math.round(total.slots / days);
      const absent = Math.round(slots * 0.031 * f), late = Math.round(slots * 0.14 * f);
      const inc = Math.round(slots * 0.02 * f), early = Math.round(slots * 0.035);
      const ok = slots - absent - late - early - inc;
      const personDays = Math.round(total.personDays / days);
      const present = Math.round(personDays * (dow === 0 || dow === 6 ? 0.42 : 0.62) * (0.94 + r() * 0.12));
      return { date: d, slots, ok, absent, late, lateMin: late * 11, early, leave: Math.round(slots * 0.05), inc, heads: 0, personDays, present,
        onTime: Math.round(ok / slots * 1000) / 10, attRate: Math.round(present / personDays * 1000) / 10 };
    });
    out.push({
      month: `${y}-${String(mo).padStart(2, '0')}`, savedAt: new Date(y, mo, 1).toISOString(), source: 'demo',
      days, employees: heads, total, byDay,
      byShift: { morning: mk(0.52), afternoon: mk(0.24), night: mk(0.2), none: mk(0.04) },
      byBranch: POSITIONS.slice(0, 8).map((p) => { const a = mk(1 / 8); return { branch: p, region: '', ...a }; }).sort((a, b) => a.onTime - b.onTime),
      warn: [],
    });
  }
  return out;
}

/* ── multi-month dataset store (localStorage; uploading a new month never
   overwrites the previous ones — each month is its own record) ─────────── */
const DS_INDEX = 'hcm-datasets-v2';
const DS_KEY = (m) => 'hcm-ds-v2:' + m;
const COLORS = ['red', 'yellow', 'gray', 'gray-half', 'ok', 'out'];

export function serializeModel(model, meta) {
  const shifts = [];
  const shiftIdx = (raw) => { const s = raw || ''; let i = shifts.indexOf(s); if (i < 0) { shifts.push(s); i = shifts.length - 1; } return i; };
  const employees = model.employees.map(e => ({
    i: e.id, n: e.name, p: e.position || '', b: e.branch || '', r: e.region || '', t: e.branchType || '',
    c: e.cells.map(c => [c.code || '', COLORS.indexOf(c.color), c.color2 ? COLORS.indexOf(c.color2) : -1,
      shiftIdx(c.shift), c.in ?? -1, c.out ?? -1, c.late || 0, c.early || 0, c.leave || '', c.hcm || '']),
  }));
  return { v: 2, month: (model.dates[0] || '').slice(0, 7), dates: model.dates, shifts, employees,
    savedAt: new Date().toISOString(), ...(meta || {}) };
}

export function deserializeModel(d) {
  const dates = d.dates;
  const employees = d.employees.map(e => {
    const cells = e.c.map(a => {
      const inM = a[4] < 0 ? null : a[4], outM = a[5] < 0 ? null : a[5];
      const shift = d.shifts[a[3]] || '';
      const code = a[0], color = COLORS[a[1]] || 'ok', color2 = a[2] >= 0 ? COLORS[a[2]] : undefined;
      const times = inM != null ? `เข้า ${fmtMin(inM)} · ออก ${outM != null ? fmtMin(outM) : '--'}` : '';
      const bits = [];
      if (a[6]) bits.push(`สาย ${a[6]} น.`);
      if (a[7]) bits.push(`ออกก่อน ${a[7]} น.`);
      if (a[8]) bits.push('ลา ' + a[8]);
      if (code === 'OFF') bits.push('วันหยุดประจำสัปดาห์');
      if (code === 'PH') bits.push('วันหยุดนักขัตฤกษ์');
      if (code === 'AB' || code === 'NS') bits.push('ขาดงาน/ไม่มีการแสกน');
      if (code === 'INC') bits.push('แสกนไม่ครบ');
      return { code, color, color2, shift, in: inM, out: outM, late: a[6], early: a[7], leave: a[8] || null, hcm: a[9],
        inc: (inM == null) !== (outM == null),   // แสกนข้างเดียว — ดูจาก in/out ที่เก็บไว้ ไม่ต้องเปลี่ยนรูปแบบไฟล์
        tip: [bits.join(' · '), times, shift ? 'กะ ' + shift : ''].filter(Boolean).join(' — ') || '—' };
    });
    const stat = { absent: 0, late: 0, lateMin: 0, maxLate: 0, early: 0, earlyMin: 0, leave: 0, incomplete: 0, off: 0, worked: 0, red: 0, present: 0, codes: {} };
    HCM_CODES.forEach(k => stat.codes[k] = 0);
    cells.forEach(c => {
      if (c.color === 'red') { stat.red++; if (c.code === 'AB' || c.code === 'NS') stat.absent++; if (c.code === 'INC') stat.incomplete++; }
      if (c.late) { stat.late++; stat.lateMin += c.late; if (c.late > stat.maxLate) stat.maxLate = c.late; }
      if (c.early) { stat.early++; stat.earlyMin += c.early; }
      if (c.leave) stat.leave++;
      if (c.code === 'OFF') stat.off++;
      if (c.in != null && c.out != null) stat.worked++;
      if (isPresent(c)) stat.present++; else if (stat.codes[c.hcm] !== undefined) stat.codes[c.hcm]++;
    });
    return { id: e.i, name: e.n, position: e.p, branch: e.b, region: e.r, branchType: e.t, shifts: {}, cells, stat };
  });
  return { dates, employees, columns: [], hasSchedule: true, month: d.month, savedAt: d.savedAt, source: d.source };
}

export function listDatasets() {
  try { return JSON.parse(localStorage.getItem(DS_INDEX) || '[]'); } catch (e) { return []; }
}
export function saveDataset(model, meta) {
  const d = serializeModel(model, meta);
  if (!d.month) return null;
  localStorage.setItem(DS_KEY(d.month), JSON.stringify(d));
  const idx = listDatasets().filter(x => x.month !== d.month);
  idx.push({ month: d.month, employees: d.employees.length, days: d.dates.length, savedAt: d.savedAt, source: (meta && meta.source) || 'upload' });
  idx.sort((a, b) => a.month.localeCompare(b.month));
  while (idx.length > 15) { const drop = idx.shift(); localStorage.removeItem(DS_KEY(drop.month)); }
  localStorage.setItem(DS_INDEX, JSON.stringify(idx));
  return d.month;
}
export function loadDataset(month) {
  try { const raw = localStorage.getItem(DS_KEY(month)); return raw ? deserializeModel(JSON.parse(raw)) : null; } catch (e) { return null; }
}
export function deleteDataset(month) {
  localStorage.removeItem(DS_KEY(month));
  localStorage.setItem(DS_INDEX, JSON.stringify(listDatasets().filter(x => x.month !== month)));
}

/* ── counting window: accumulate per-employee stats across stored months ── */
export function statsAcross(months, from, to) {
  const out = new Map();
  months.forEach(mo => {
    const m = loadDataset(mo);
    if (!m) return;
    m.employees.forEach(e => {
      if (!out.has(e.id)) out.set(e.id, { id: e.id, name: e.name, position: e.position, absent: 0, late: 0, lateMin: 0, maxLate: 0, early: 0, incomplete: 0, leave: 0, present: 0, personDays: 0, red: 0 });
      const A = out.get(e.id);
      e.cells.forEach((c, i) => {
        const d = m.dates[i];
        if (!d || d < from || d > to) return;
        A.personDays++;
        if (isPresent(c)) A.present++;
        if (c.color === 'red') A.red++;
        if (c.code === 'AB' || c.code === 'NS' || c.code === 'NoS' || c.code === 'OFF!') A.absent++;
        if (c.code === 'INC') A.incomplete++;
        if (c.late) { A.late++; A.lateMin += c.late; if (c.late > A.maxLate) A.maxLate = c.late; }
        if (c.leave) A.leave++;
      });
    });
  });
  return out;
}

/* ── กฎหนังสือเตือน — แอดมินตั้งเองได้ ─────────────────────────────────
   แต่ละกฎคือ "ตัวชี้วัด ≥ ค่า → ใบเตือนระดับนี้" เช่น สาย ≥ 5 ครั้ง, ขาด ≥ 3 วัน,
   หรือสายครั้งเดียวเกิน 120 นาที ก็ออกใบเตือนทันที
   คนที่เข้าหลายกฎ จะได้ระดับของกฎที่ความรุนแรง (sev) สูงสุด                */
/* ── ตารางแปลงรหัส ─────────────────────────────────────────────────────
   BUILTIN_CODES = รหัสที่ระบบตัดสินเอง (แก้คำอธิบายได้ แต่ตรรกะอยู่ใน classify)
   ส่วนแถวที่แอดมินเพิ่มเองจะมี match = คำที่ให้ไปค้นในช่อง แล้วบังคับรหัส/สี/นับเป็นมาทำงาน */
export const CELL_COLORS = [
  ['ok', 'เขียว — มาทำงาน'],
  ['red', 'แดง — มีปัญหา'],
  ['yellow', 'เหลือง — สาย/ออกก่อน'],
  ['gray', 'เทา — ลา/วันหยุด'],
  ['gray-half', 'เทาอ่อน — ครึ่งวัน'],
];
export const BUILTIN_CODES = [
  { code: '(ว่าง)', label: 'มาทำงานปกติ — สแกนครบ ตรงเวลา', color: 'ok' },
  { code: 'AB', label: 'ขาดงาน', color: 'red' },
  { code: 'NS', label: 'ไม่มีการแสกนหน้า', color: 'red' },
  { code: 'INC', label: 'แสกนไม่ครบ (มีแค่ in หรือ out) — คิดเป็นขาดงานในอัตราการมาทำงาน', color: 'red' },
  { code: 'NoS', label: 'ไม่มีกะทำงาน', color: 'red' },
  { code: 'OFF!', label: 'ตั้งกะในวันหยุดแต่ไม่แสกน', color: 'red' },
  { code: 'L15', label: 'มาสาย 15 นาที', color: 'yellow' },
  { code: 'E20', label: 'ออกก่อนเวลา 20 นาที', color: 'yellow' },
  { code: 'AL / SL / PL / UL', label: 'ลาพักร้อน / ลาป่วย / ลากิจ / ลาไม่รับค่าจ้าง', color: 'gray' },
  { code: 'AL½', label: 'ลาครึ่งวัน (เช้า/บ่าย)', color: 'gray-half' },
  { code: 'OFF', label: 'วันหยุดประจำสัปดาห์', color: 'gray' },
  { code: 'PH', label: 'วันหยุดนักขัตฤกษ์', color: 'gray' },
  { code: '?', label: 'รอการอนุมัติ', color: 'gray' },
];

// คำในไฟล์ที่ระบบยังอ่านไม่ออก — เอาไว้เสนอให้แอดมินเพิ่มเป็นกฎแปลงรหัส
const KNOWN_RE = /วันหยุด|ขาดงาน|^in:|รอการอนุมัติ|ลาพักร้อน|ลาป่วย|ลากิจ|ลาไม่รับค่าจ้าง|ลาคลอด|ลาอุปสมบท|ลาบวช/;
export function unknownPhrases(attRows, customCodes) {
  const t = mapTable(attRows);
  const tally = new Map();
  t.records.forEach(r => t.dates.forEach(d => {
    const raw = String(r.days[d] || '').trim();
    if (!raw) return;
    // ตัดส่วนที่เป็นเวลาสแกนออก เหลือเฉพาะข้อความต่อท้าย
    const rest = raw.replace(/in\s*:\s*\S+\s+out\s*:\s*\S+/i, '').replace(/\d{1,2}:\d{2}/g, '').trim();
    if (!rest || KNOWN_RE.test(rest)) return;
    if ((customCodes || []).some(c => c.match && raw.includes(c.match))) return;
    tally.set(rest, (tally.get(rest) || 0) + 1);
  }));
  return [...tally.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([text, n]) => ({ text, n }));
}

export const WARN_METRICS = [
  ['absent', 'ขาดงาน', 'วัน'],
  ['late', 'มาสาย', 'ครั้ง'],
  ['lateMin', 'นาทีสายรวม', 'นาที'],
  ['maxLate', 'สายครั้งเดียวเกิน', 'นาที'],
  ['incomplete', 'แสกนไม่ครบ', 'ครั้ง'],
  ['early', 'ออกก่อนเวลา', 'ครั้ง'],
];
export const DEFAULT_WARN_RULES = [
  { id: 'w1', metric: 'absent', value: 3, level: 'เตือนครั้งที่ 1', sev: 1, on: true },
  { id: 'w2', metric: 'late', value: 5, level: 'เตือนครั้งที่ 1', sev: 1, on: true },
  { id: 'w3', metric: 'maxLate', value: 120, level: 'เตือนครั้งที่ 1', sev: 1, on: true },
  { id: 'w4', metric: 'absent', value: 6, level: 'เตือนครั้งที่ 2', sev: 2, on: true },
  { id: 'w5', metric: 'late', value: 10, level: 'เตือนครั้งที่ 2', sev: 2, on: true },
];
export const metricLabel = (k) => (WARN_METRICS.find(m => m[0] === k) || [k, k, ''])[1];
export const metricUnit = (k) => (WARN_METRICS.find(m => m[0] === k) || [k, k, ''])[2];
export const ruleText = (r) => `${metricLabel(r.metric)} ≥ ${r.value} ${metricUnit(r.metric)} → ${r.level}`;

// คืนกฎที่รุนแรงที่สุดที่พนักงานคนนี้เข้าเกณฑ์ (null = ยังไม่ถึงเกณฑ์ใด)
export function matchWarnRules(stat, rules) {
  let best = null;
  (rules || []).forEach(r => {
    if (!r.on) return;
    const v = Number(stat[r.metric] || 0);
    if (v < Number(r.value)) return;
    if (!best || Number(r.sev) > Number(best.sev)) best = r;
  });
  return best;
}

/* ── configurable shift groups (admin can rename / add) ─────────────────── */
export const DEFAULT_SHIFT_GROUPS = [
  { key: 'morning', name: 'กะเช้า', start: '09:00', end: '18:00' },
  { key: 'afternoon', name: 'กะบ่าย', start: '13:00', end: '22:00' },
  { key: 'night', name: 'กะดึก', start: '21:00', end: '06:00' },
];
const hourOf = (t) => { const m = String(t || '').match(/(\d{1,2})[:.](\d{2})/); return m ? +m[1] + (+m[2]) / 60 : null; };
export function normalizeGroups(groups) {
  return (groups || []).map(g => g.start != null ? g
    : { key: g.key, name: g.name, start: String(g.from ?? 0).padStart(2, '0') + ':00', end: String(g.to ?? 0).padStart(2, '0') + ':00' });
}
// จับกะจาก "เวลาเริ่มงาน" ในไฟล์ตารางกะ เทียบกับเวลาเริ่มงานของแต่ละกะที่ตั้งไว้ (ใกล้สุดภายใน 3 ชม.)
export function bucketWith(raw, groups) {
  const h = hourOf(raw);
  if (h == null) return 'none';
  let best = null, bestD = 99;
  normalizeGroups(groups).forEach(g => {
    const gh = hourOf(g.start);
    if (gh == null) return;
    const d = Math.min(Math.abs(h - gh), 24 - Math.abs(h - gh));
    if (d < bestD) { bestD = d; best = g.key; }
  });
  return best && bestD <= 3 ? best : 'other';
}
