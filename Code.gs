/* วิธีติดตั้ง
 1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางโค้ดนี้
 2) แก้ PIN ใน setup() แล้วรัน setup() 1 ครั้ง (อนุญาตสิทธิ์)
 3) ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > Web app
    Execute as: Me, Who has access: Anyone > คัดลอก URL
 4) นำ URL ไปใส่ตัวแปร API ในไฟล์ index.html แล้วอัปโหลดขึ้น GitHub Pages
 PIN เก็บเฉพาะฝั่งนี้ ไม่อยู่ในหน้าเว็บ */
const P = PropertiesService.getScriptProperties();
function setup() { P.setProperties({ PIN_ADMIN: '111111', PIN_HEAD: '222222' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('Docs') || ss.insertSheet('Docs').appendRow(['id', 'type', 'date', 'json', 'role', 'ts']);
}
function doPost(e) {
  const r = JSON.parse(e.postData.contents);
  const role = r.pin === P.getProperty('PIN_ADMIN') ? 'admin' : r.pin === P.getProperty('PIN_HEAD') ? 'head' : '';
  if (!role) return out({ ok: false, err: 'PIN ไม่ถูกต้อง' });
  if (r.a === 'login') return out({ ok: true, role });
  const s = sheet(), lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (r.a === 'list') {
      const v = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 6).getValues() : [];
      return out({ ok: true, docs: v.map(x => ({ id: x[0], type: x[1], by: x[4], ts: x[5], d: JSON.parse(x[3]) })) });
    }
    if (r.a === 'save') {
      const ids = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().flat() : [];
      if (ids.includes(r.id)) return out({ ok: false, err: 'เอกสารนี้ถูกล็อกแล้ว แก้ไขไม่ได้' });
      const ts = new Date().toISOString(); // เวลาจากเซิร์ฟเวอร์
      s.appendRow([r.id, r.type, "'" + r.date, JSON.stringify(r.d), role, ts]);
      return out({ ok: true, ts });
    }
  } finally { lock.releaseLock(); }
  return out({ ok: false, err: 'คำสั่งไม่ถูกต้อง' });
}
