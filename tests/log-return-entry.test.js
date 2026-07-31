const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const backendLf = backend.replace(/\r\n/g, '\n');

// ── Backend: dispatch + admin gate ────────────────────────────────────────────
assert(backend.includes("action === 'returnLogEntry'"), 'ต้อง dispatch returnLogEntry');
assert((backend.match(/action === 'returnLogEntry'/g) || []).length === 2,
  'ต้อง dispatch ทั้ง doGet และ doPost เหมือน deleteLogEntry');
assert(backend.includes('function returnLogEntry(payload)'));
assert(backend.includes('function returnLogEntryUnlocked(payload)'));
// คืนรายการแก้ยอดสต็อกจริง — ต้องเป็น Admin เท่านั้น
assert(/function returnLogEntryUnlocked[\s\S]{0,200}requireAdminUser/.test(backendLf),
  'returnLogEntryUnlocked ต้องเช็ค requireAdminUser');
// ต้องจับ lock เหมือน processTransaction กันคืนซ้อนกัน
assert(/function returnLogEntry\(payload\)[\s\S]{0,220}LockService\.getScriptLock/.test(backendLf));

// ── Backend: ต้องไม่ไปสร้าง Purchase History / เลขประจำชิ้น ตอนคืน ──────────────
// คืนของที่เบิก → กลายเป็น Input ซึ่งปกติจะ sync purchase history (ทั้งที่ไม่ได้ซื้อจริง)
// คืนของที่รับเข้า → กลายเป็น Output ซึ่งปกติจะออกเลขประจำชิ้นใหม่ (ไม่ควร)
assert(backend.includes('if (signedQty > 0 && !payload.skipPurchaseHistory) {'));
assert(backend.includes('if (signedQty < 0 && !payload.skipPartTags) {'));
assert(backend.includes('skipPurchaseHistory: true'));
assert(backend.includes('skipPartTags: true'));

// ── Backend: ต้องหาชีทสต็อกต้นทางเอง (ชีท Log ไม่เก็บชื่อชีทไว้) ─────────────────
assert(backend.includes('function findStockSheetNameForLogItem(item)'));
assert(!/function findStockSheetNameForLogItem[\s\S]{0,900}ensureSheetWithTemplate/.test(backendLf),
  'ห้ามใช้ ensureSheetWithTemplate ตอนค้นหาชีท เพราะจะสร้างชีทขยะถ้าไม่มี');
assert(/function findStockSheetNameForLogItem[\s\S]{0,300}getSheetByName/.test(backendLf));
assert(backend.includes('จึงคืนสต็อกอัตโนมัติไม่ได้'), 'หาชีทไม่เจอต้องแจ้ง ไม่ใช่คืนเงียบๆ');

// ── Backend: ทิศทางการกลับรายการ + กันคืนซ้ำ ─────────────────────────────────
// qty ใน Log เป็นค่าที่มีเครื่องหมาย: Output ติดลบ → คืนต้องเป็น Input
assert(backend.includes("var reverseType = signedQty < 0 ? 'Input' : 'Output';"));
assert(backend.includes("var LOG_RETURN_MARKER = 'RETURN_OF:';"));
assert(backend.includes('รายการนี้ถูกคืนไปแล้ว'), 'ต้องกันคืนซ้ำ');
assert(backend.includes('แถวนี้เป็นรายการคืนอยู่แล้ว'), 'ต้องกันคืนแถว Return เอง');
// ไม่ลบแถวเดิม — เก็บไว้เป็นหลักฐาน (ต่างจาก deleteLogEntry)
assert(!/function returnLogEntryUnlocked[\s\S]*?\n}/.test(backendLf.slice(backendLf.indexOf('function returnLogEntryUnlocked'))) ||
  !/function returnLogEntryUnlocked[\s\S]{0,4000}deleteRow/.test(backendLf),
  'returnLogEntry ต้องไม่ลบแถวเดิม');
// race guard เหมือน deleteLogEntry
assert(/function returnLogEntryUnlocked[\s\S]{0,2500}ข้อมูล Log มีการเปลี่ยนแปลงตั้งแต่โหลดหน้านี้/.test(backendLf));

// ── race guard ต้องเทียบ timestamp แบบ normalize แล้วเท่านั้น ──────────────────
// เซลล์ Timestamp เป็น Date จริง (Sheets แปลงให้) → ส่งออกเป็น ISO → เทียบสตริงตรงๆ
// กับ Date.toString() ของ Apps Script ไม่มีวันตรง ทำให้คืน/ลบไม่ได้เลยทั้งที่แถวถูกต้อง
assert(backend.includes('function normalizeLogTimestamp(value)'));
assert(backend.includes('function logTimestampsMatch(expected, actual)'));
assert(backend.includes("Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss')"));
// ทั้ง return และ delete ต้องใช้ตัวเทียบตัวเดียวกัน (delete มีบั๊กเดียวกันซ่อนอยู่)
assert((backend.match(/logTimestampsMatch\(payload\.timestamp/g) || []).length === 2,
  'ทั้ง returnLogEntry และ deleteLogEntry ต้องใช้ logTimestampsMatch');
assert(!/expectedTs && tsIdx !== undefined && String\(row\[tsIdx\]/.test(backend),
  'deleteLogEntry ต้องไม่เทียบสตริงดิบแบบเดิมแล้ว');
// สต็อกไม่พอตอนคืนขา Input ต้องได้ error ที่อ่านรู้เรื่อง
assert(backend.includes('คืนรายการนี้ไม่ได้: ต้องหักสต็อกออก'));

// ── Frontend: ปุ่ม + สถานะ ───────────────────────────────────────────────────
assert(html.includes("var LOG_RETURN_MARKER = 'RETURN_OF:';"));
assert(html.includes('data-log-return='), 'ต้องมีปุ่มคืนในตาราง Log');
assert(html.includes("action: 'returnLogEntry'"));
// ปุ่มคืนเป็นของ Admin เท่านั้น ส่วนปุ่มลบยังใช้ delete_logs เหมือนเดิม
assert(html.includes('var canReturnLog = isAdminUser();'));
assert(html.includes("var canDeleteLog = hasPermission('delete_logs');"));
// คอลัมน์ต้องโชว์ถ้ามีสิทธิ์ใดสิทธิ์หนึ่ง ไม่ใช่ผูกกับ delete อย่างเดียว
assert(html.includes('var showActionCol = canDeleteLog || canReturnLog;'));
assert(html.includes("colCount = showActionCol ? 13 : 12;"));
// แถวที่คืนแล้ว/แถว Return ต้องไม่มีปุ่มให้กดซ้ำ
assert(html.includes('✓ คืนแล้ว'));
assert(html.includes('↩ Return'));
// คืนแล้วสต็อกเปลี่ยนจริง ต้องโหลดข้อมูลอะไหล่ใหม่ ไม่ใช่รีเฟรชแค่ log
assert(html.includes('loadPartsData({ skipCache: true });'));
// คืนแล้วเลขประจำชิ้นถูกลบให้ใช้ซ้ำได้ — ต้องบอกผู้ใช้ด้วยว่าเลขไหนถูกคืน
assert(html.includes('var removedTags = (res && res.removed_tags) || [];'));
assert(html.includes("' · คืนเลข ' + removedTags.join(', ') + ' ให้ใช้ซ้ำได้'"));
assert(html.includes('ถ้ารายการนี้เคยออกเลขประจำชิ้นไว้ เลขนั้นจะถูกลบเพื่อให้กลับมาใช้ซ้ำได้'),
  'confirm ต้องบอกล่วงหน้าว่าเลขจะถูกลบ');

// ── Frontend: logic ของ index กันคืนซ้ำ (โหลดมารันจริง) ───────────────────────
const sandbox = {};
const htmlLf = html.replace(/\r\n/g, '\n');
['normalizeLogTimestampKey', 'logRowReturnKey', 'buildLogReturnedIndex', 'isLogRowReversal'].forEach(function (name) {
  const re = new RegExp('^    function[ ]+' + name + '\\([\\s\\S]*?\\n    }$', 'm');
  const match = htmlLf.match(re);
  assert(match, 'cannot extract ' + name);
  sandbox[name] = new Function('normalizeLogTimestampKey',
    "var LOG_RETURN_MARKER = 'RETURN_OF:';\n" + match[0] + '\nreturn ' + name + ';')(sandbox.normalizeLogTimestampKey);
});

// ── หัวใจของบั๊ก: เซลล์ Timestamp เป็น Date จริง → API ส่งกลับมาเป็น ISO ─────────
// ทั้งสองรูปแบบต้อง normalize ได้ค่าเดียวกัน ไม่งั้น guard ฝั่ง server จะตีว่าข้อมูลเปลี่ยน
// และป้าย "คืนแล้ว" ฝั่ง client จะไม่ขึ้น
assert.strictEqual(sandbox.normalizeLogTimestampKey('2026-07-30T11:13:17.000Z'), '2026-07-30 18:13:17',
  'ISO (UTC) ต้องแปลงเป็นเวลาไทย');
assert.strictEqual(sandbox.normalizeLogTimestampKey('2026-07-30 18:13:17'), '2026-07-30 18:13:17',
  'สตริงเวลาไทยต้องได้ค่าเดิม');
assert.strictEqual(sandbox.normalizeLogTimestampKey('2026-07-30T11:13:17.000Z'),
  sandbox.normalizeLogTimestampKey(new Date('2026-07-30T11:13:17.000Z')),
  'Date object กับ ISO string ต้องได้ค่าเดียวกัน');
assert.strictEqual(sandbox.normalizeLogTimestampKey(''), '');
assert.strictEqual(sandbox.normalizeLogTimestampKey(null), '');
assert.strictEqual(sandbox.normalizeLogTimestampKey('ไม่ใช่วันที่'), 'ไม่ใช่วันที่', 'ค่าที่ parse ไม่ได้ต้องคืนสตริงเดิม');

assert.strictEqual(sandbox.logRowReturnKey({ timestamp: '2026-07-30T11:13:17.000Z', partName: ' MT PIN ' }),
  '2026-07-30 18:13:17|MT PIN', 'key ต้อง normalize เวลา + trim ชื่อ');

// แถวจาก API มาเป็น ISO — marker ที่ backend เขียนไว้เป็นเวลาไทย ต้อง match กันให้ได้
const rows = [
  { no: 1, timestamp: '2026-07-30T11:13:17.000Z', partName: 'MT PIN', reason: 'Trial', reasonRemark: '' },
  { no: 2, timestamp: '2026-07-30T11:20:00.000Z', partName: 'MT PIN', reason: 'Return', reasonRemark: 'RETURN_OF:2026-07-30 18:13:17|MT PIN | คืนรายการโดย Admin' },
  { no: 3, timestamp: '2026-07-30T10:30:30.000Z', partName: 'SESAME', reason: 'N/A', reasonRemark: '' }
];
const idx = sandbox.buildLogReturnedIndex(rows);
assert.strictEqual(idx[sandbox.logRowReturnKey(rows[0])], true,
  'แถวที่ถูกคืนต้องถูก mark แม้ timestamp ที่ได้จาก API เป็น ISO แต่ marker เป็นเวลาไทย');
assert.strictEqual(idx[sandbox.logRowReturnKey(rows[2])], undefined, 'แถวที่ยังไม่ถูกคืนต้องไม่ถูก mark');
assert.strictEqual(sandbox.buildLogReturnedIndex([])['x'], undefined);
assert.strictEqual(sandbox.buildLogReturnedIndex(null) && Object.keys(sandbox.buildLogReturnedIndex(null)).length, 0,
  'rows เป็น null ต้องไม่ throw');

assert.strictEqual(sandbox.isLogRowReversal(rows[1]), true);
assert.strictEqual(sandbox.isLogRowReversal(rows[0]), false);
assert.strictEqual(sandbox.isLogRowReversal(null), false, 'null ต้องไม่ throw');

// marker ต้องตรงกันทั้งสองฝั่ง ไม่งั้นหน้าเว็บจะ mark ไม่เจอ
const backendMarker = /var LOG_RETURN_MARKER = '([^']+)'/.exec(backend)[1];
const htmlMarker = /var LOG_RETURN_MARKER = '([^']+)'/.exec(html)[1];
assert.strictEqual(backendMarker, htmlMarker, 'marker ฝั่ง backend และ frontend ต้องเป็นค่าเดียวกัน');

console.log('Log return-entry checks passed');
