// Inbox แท็บ "เปิด PR เกินงบ (Bypass)" — เดิมขึ้นทุกครั้งที่ Bypass ใน 30 วัน แม้ PR อนุมัติ/ตีกลับ/ยกเลิกไปแล้ว
// → ค้างอยู่ในกล่อง ปิดไม่ได้ (และ timestamp เป็น Date ทำให้ตัด 30 วันไม่ทำงาน)
// ตอนนี้แสดงเฉพาะ PR ที่ยัง PENDING — อนุมัติหรือตีกลับแล้วหายเอง
const { html, grabFn, assert } = require('./_export-extract');
const fs = require('fs');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

const filter = new Function(grabFn('filterOpenBudgetBypasses', backend, 0) + '\nreturn filterOpenBudgetBypasses;')();
const now = new Date('2026-11-13T12:00:00+07:00').getTime();
const cutoff = now - 30 * 86400000;
const row = function(pr, ts, extra) {
  return Object.assign({ action: 'BUDGET_BYPASS', pr_id: pr, timestamp: ts, ts_ms: new Date(ts.replace(' ', 'T') + '+07:00').getTime(), actor: 'Nun', line: 'H9', detail: 'd' }, extra || {});
};
const rows = [
  row('PR-2610-003', '2026-10-21 09:40:37'),
  row('PR-2610-010', '2026-10-30 15:48:33'),
  row('PR-2610-009', '2026-10-30 15:48:26'),
  row('PR-2611-001', '2026-11-10 10:00:00'),
  row('PR-2611-001', '2026-11-11 10:00:00', { detail: 'latest' }),
  row('PR-2609-001', '2026-09-01 10:00:00'),
  row('PR-2611-002', '2026-11-12 10:00:00', { action: 'CREATE' })
];
const status = { 'PR-2610-003': 'APPROVED', 'PR-2610-010': 'REJECTED', 'PR-2610-009': 'CANCELLED',
  'PR-2611-001': 'PENDING', 'PR-2609-001': 'PENDING', 'PR-2611-002': 'PENDING' };
const out = filter(rows, status, cutoff);
assert.deepStrictEqual(out.map(function(b) { return b.pr_id; }), ['PR-2611-001'], 'เฉพาะ PR ที่ยังรออนุมัติ ในช่วงเวลา');
assert.strictEqual(out[0].detail, 'latest', 'ใบละรายการเดียว (ล่าสุด)');
// timestamp แบบ Date (ts_ms) ต้องตัดช่วงเวลาได้จริง
const old = filter([{ action: 'BUDGET_BYPASS', pr_id: 'X', timestamp: 'Wed Sep 01 2026', ts_ms: new Date('2026-09-01T00:00:00+07:00').getTime() }], { X: 'PENDING' }, cutoff);
assert.strictEqual(old.length, 0, 'เกิน 30 วันถูกตัดแม้ timestamp เป็น Date');

const list = grabFn('listBudgetBypasses', backend, 0);
assert(list.includes("Utilities.formatDate(tsVal, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss')"), 'แสดงเวลาแบบอ่านง่าย');
assert(list.includes('getPrHeaderSheet()'), 'อ่านสถานะ PR จาก header');
assert(html.includes('✅ ไม่มี PR เกินงบที่รออนุมัติ'));
console.log('inbox-bypass-pending-only: OK');
