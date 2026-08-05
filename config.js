/* ── ตั้งค่าเชื่อมต่อส่วนกลาง ────────────────────────────────────────────
   ใส่ Web app URL ของ Apps Script ไว้ที่นี่ที่เดียว ทุกคนที่เปิดเว็บจะใช้ค่านี้
   ไม่ต้องให้แต่ละคนมาวาง URL เอง

   เปลี่ยน URL เมื่อไหร่ (เช่น deploy ใหม่แล้วได้ URL ใหม่) แก้บรรทัดเดียวนี้
   แล้ว push ขึ้น GitHub ทุกเครื่องจะใช้ค่าใหม่ทันทีที่เปิดเว็บ

   URL นี้ไม่ใช่ความลับ — ทุก action ต้องล็อกอินก่อนเสมอ รู้ URL เฉย ๆ
   ทำอะไรไม่ได้ แต่ก็ไม่ควรเผยแพร่ออกไปโดยไม่จำเป็น                        */
window.HCM_CONFIG = {
  endpoint: 'https://script.google.com/macros/s/AKfycbxh-sHgwznKpAmA_vFejYw05WWqMKYcd5DBAW9Ob_xONWp9YVsw62gtdF_ehKx2cfdw/exec',   // เช่น 'https://script.google.com/macros/s/AKfy..../exec'
};
