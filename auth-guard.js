/* ── ด่านตรวจการล็อกอิน ────────────────────────────────────────────────
   ใส่ไว้ใน <head> ของทั้งสองหน้า ทำงานก่อนแอปจะเริ่มวาด

   data-require="admin" → ต้องเป็นแอดมินเท่านั้น (หน้า Attendance Admin)
   data-require="any"   → ล็อกอินอะไรก็ได้        (หน้า Attendance Viewer)

   ข้อควรรู้ตามตรง: นี่คือด่านกันคนทั่วไป ไม่ใช่การป้องกันระดับเซิร์ฟเวอร์
   คนที่รู้เรื่องเทคนิคแก้ค่าในเบราว์เซอร์ให้ผ่านด่านนี้ได้
   ตัวที่กันจริงคือ Apps Script ซึ่งตรวจรหัสทุกครั้งที่ขอข้อมูล
   และส่งกลับมาเฉพาะกะที่คนนั้นมีสิทธิ์เห็น                                */
(function () {
  var SESSION_KEY = 'hcm-session-v1';
  var LOGIN_PAGE = './index.html';

  var need = (document.currentScript && document.currentScript.getAttribute('data-require')) || 'any';

  var s = null;
  try { s = JSON.parse(localStorage.getItem(SESSION_KEY) || 'null'); } catch (e) { s = null; }

  if (!s || !s.me || !s.me.user) { location.replace(LOGIN_PAGE); return; }
  if (need === 'admin' && !s.me.isAdmin) { location.replace('./Attendance%20Viewer.dc.html'); return; }

  window.HCM_SESSION = s;

  // แถบเล็ก ๆ บอกว่าใครใช้อยู่ พร้อมปุ่มออกจากระบบ — ไม่ยุ่งกับโครงหน้าเดิม
  function bar() {
    if (!document.body || document.getElementById('hcm-who')) return;
    var me = s.me;
    var box = document.createElement('div');
    box.id = 'hcm-who';
    box.setAttribute('class', 'noprint');
    box.style.cssText = 'position:fixed;right:12px;bottom:12px;z-index:99999;display:flex;align-items:center;gap:10px;'
      + 'background:#201e1d;color:#f3f2f2;padding:7px 10px;font-family:Archivo,"Noto Sans Thai",system-ui,sans-serif;'
      + 'font-size:12px;box-shadow:0 3px 10px rgba(45,43,43,.28);max-width:min(92vw,420px);';
    var who = document.createElement('span');
    who.style.cssText = 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;';
    who.textContent = (me.name || me.user) + ' · ' + (me.isAdmin ? 'แอดมิน (เห็นทุกกะ)' : ('ดูแล ' + (me.shiftText || '—')));
    var out = document.createElement('button');
    out.type = 'button';
    out.textContent = 'ออกจากระบบ';
    out.style.cssText = 'flex:none;border:1px solid #f3f2f2;background:transparent;color:#f3f2f2;cursor:pointer;'
      + 'font-family:inherit;font-size:11px;padding:4px 8px;';
    out.onclick = function () {
      try { localStorage.removeItem(SESSION_KEY); } catch (e) {}
      location.replace(LOGIN_PAGE);
    };
    box.appendChild(who); box.appendChild(out);
    document.body.appendChild(box);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bar);
  else bar();
})();
