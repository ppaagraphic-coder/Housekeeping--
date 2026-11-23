/* วิธีติดตั้ง
 1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางโค้ดนี้
 2) แก้ PIN ใน setup() แล้วรัน setup() 1 ครั้ง (อนุญาตสิทธิ์)
 3) ทำให้ใช้งานได้ > Web app > Execute as: Me, Who has access: Anyone > คัดลอก URL
 4) นำ URL ไปใส่ตัวแปร API ใน index.html แล้วอัปโหลดขึ้น GitHub Pages
 หากแก้โค้ดภายหลัง ต้อง Deploy เวอร์ชันใหม่ PIN เก็บเฉพาะฝั่งนี้ */
const P = PropertiesService.getScriptProperties();
function setup() { P.setProperties({ PIN_ADMIN: '111111', PIN_HEAD: '222222' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('Docs') || ss.insertSheet('Docs').appendRow(['id', 'type', 'date', 'json', 'role', 'ts', 'signatures']);
}
// อ่านเฉพาะคอลัมน์ 1-6 (ไม่อ่านลายเซ็น) เพื่อให้โหลดเร็ว
function list_(s) {
  const n = s.getLastRow() - 1;
  return (n > 0 ? s.getRange(2, 1, n, 6).getValues() : []).map(x => ({ id: x[0], type: x[1], ts: x[5], d: JSON.parse(x[3]) }));
}
function doPost(e) {
  const r = JSON.parse(e.postData.contents);
  if (r.a === 'ping') return out({ ok: true });
  const role = r.pin === P.getProperty('PIN_ADMIN') ? 'admin' : r.pin === P.getProperty('PIN_HEAD') ? 'head' : '';
  if (!role) return out({ ok: false, err: 'PIN ไม่ถูกต้อง' });
  const s = sheet();
  if (r.a === 'login') return out({ ok: true, role, docs: list_(s) });
  if (r.a === 'list') return out({ ok: true, docs: list_(s) });
  if (r.a === 'save') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const n = s.getLastRow() - 1;
      const ids = n > 0 ? s.getRange(2, 1, n, 1).getValues().flat() : [];
      if (ids.includes(r.id)) return out({ ok: false, err: 'เอกสารนี้ถูกบันทึกและล็อกแล้ว' });
      const ts = new Date().toISOString(); // เวลาจากเซิร์ฟเวอร์
      s.appendRow([r.id, r.type, "'" + r.date, JSON.stringify(r.d), role, ts, JSON.stringify(r.sg || {})]);
      return out({ ok: true, ts });
    } finally { lock.releaseLock(); }
  }
  return out({ ok: false, err: 'คำสั่งไม่ถูกต้อง' });
}
