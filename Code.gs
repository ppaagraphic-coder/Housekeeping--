// วางในชีตที่จะใช้เก็บข้อมูล: Extensions > Apps Script > วางไฟล์นี้ > Deploy > Web app
// (Execute as: Me, Who has access: Anyone) แล้วนำ URL ไปใส่ API ใน index.html
// เมื่อแก้ไฟล์นี้ ต้อง Deploy > Manage deployments > ✎ > Version: New version > Deploy
const PINS = { '1111': 'hk', '2222': 'store', '9999': 'admin' }; // เปลี่ยน PIN ก่อนใช้งานจริง
const COLLS = ['cfg', 'logs', 'sends', 'rets', 'cases', 'manual', 'closes'];

function sh(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(['id', 'updated', 'json']); }
  return s;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// โหลดข้อมูล: since = เวลา (ms) ที่ซิงก์ล่าสุด → ส่งเฉพาะแถวที่ใหม่กว่า; ไม่ส่ง since หรือเคยมีการล้างข้อมูล → ส่งทั้งหมด (full)
function load(pin, since, check) {
  const role = PINS[pin];
  if (!role) return out({ err: 'pin' });
  if (check) return out({ role }); // ตรวจ PIN อย่างเดียว (เร็ว ไม่อ่านชีต)
  const rs = +PropertiesService.getScriptProperties().getProperty('resetAt') || 0;
  since = +since || 0;
  const full = !since || since < rs;
  if (full) since = 0;
  const ts = Date.now() - 3000; // เผื่อเวลาเหลื่อม; ฝั่งเว็บรวมตาม id จึงไม่ซ้ำ
  const data = {};
  COLLS.forEach(c => {
    if (role !== 'admin' && (c === 'manual' || c === 'closes')) { data[c] = []; return; } // ค่าใช้จ่าย/ปิดยอด เฉพาะแอดมิน
    const s = sh(c), n = s.getLastRow();
    if (n < 2) { data[c] = []; return; }
    data[c] = s.getRange(2, 1, n - 1, 3).getValues()
      .filter(r => !since || new Date(r[1]).getTime() > since)
      .map(r => JSON.parse(r[2]));
  });
  return out({ role, data, full, ts });
}

function doGet(e) { return load(e.parameter.pin, e.parameter.since, e.parameter.check); } // เผื่อเวอร์ชันเก่า

// บันทึกหลายรายการในครั้งเดียว (upsert ตาม id)
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.load || b.check) return load(b.pin, b.since, b.check); // โหลดด้วย POST: PIN ไม่อยู่ใน URL/log
  if (!PINS[b.pin]) return out({ err: 'pin' });
  if (b.reset) { // ล้างข้อมูล (เฉพาะแอดมิน): 'data' = ทุกอย่างยกเว้นราคา/สต็อก(cfg), 'all' = ทั้งหมด
    if (PINS[b.pin] !== 'admin') return out({ err: 'role' });
    const R = LockService.getScriptLock();
    R.waitLock(20000);
    try {
      COLLS.forEach(c => {
        if (b.reset !== 'all' && c === 'cfg') return;
        const s = sh(c), n = s.getLastRow();
        if (n > 1) s.deleteRows(2, n - 1);
      });
      PropertiesService.getScriptProperties().setProperty('resetAt', String(Date.now())); // ให้เครื่องอื่นโหลดใหม่ทั้งหมด
    } finally { R.releaseLock(); }
    return out({ ok: 1, reset: b.reset });
  }
  const L = LockService.getScriptLock();
  L.waitLock(20000);
  try {
    const by = {};
    b.ops.forEach(o => (by[o.c] = by[o.c] || []).push(o.r));
    for (const c in by) {
      if (COLLS.indexOf(c) < 0) continue;
      const s = sh(c), n = s.getLastRow();
      const ids = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
      const add = [];
      by[c].forEach(r => {
        const row = [r.id, new Date(), JSON.stringify(r)], i = ids.indexOf(String(r.id));
        if (i < 0) { ids.push(String(r.id)); add.push(row); }
        else s.getRange(i + 2, 1, 1, 3).setValues([row]);
      });
      if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, 3).setValues(add);
    }
  } finally { L.releaseLock(); }
  return out({ ok: 1 });
}
