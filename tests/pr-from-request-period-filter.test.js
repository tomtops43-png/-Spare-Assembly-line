// โหมด "PR ขอซื้ออะไหล่" — เดิมดึงคำขอที่ยังเปิดอยู่ทุกเดือนมาทั้งหมด รายการเยอะเกิน
// ตอนนี้ดึงเฉพาะคำขอในรอบเดือนที่เลือก (ติ๊ก "รวมคำขอค้างจากเดือนก่อนๆ" ถึงจะได้ทั้งหมด)
const { html, grabFn, assert } = require('./_export-extract');

const filter = new Function(
  "var PR_FROM_REQUEST_OPEN_STATUSES = { 'Pending': true, 'Approved': true, 'On Hold': true };\n" +
  grabFn('filterOrderRequestsForPr') + '\nreturn filterOrderRequestsForPr;')();
const rows = [
  { request_id: 'A', status: 'Pending', requested_date: '2026-11-07 09:28:36' },
  { request_id: 'B', status: 'Approved', requested_date: '2026-11-01 08:00:00' },
  { request_id: 'C', status: 'Approved', requested_date: '2026-10-26 10:59:49' },
  { request_id: 'D', status: 'Converted to PR', requested_date: '2026-11-05 10:00:00' },
  { request_id: 'E', status: 'On Hold', requested_date: '2026-11-10 10:00:00' },
  { request_id: 'F', status: 'Rejected', requested_date: '2026-11-10 10:00:00' }
];
const ids = function(r) { return r.rows.map(function(x) { return x.request_id; }); };

let r = filter(rows, '2026-11', false);
assert.deepStrictEqual(ids(r), ['A', 'B', 'E'], 'เฉพาะเดือนที่เลือก + สถานะที่ยังไม่เปิด PR');
assert.strictEqual(r.backlog, 1, 'บอกจำนวนคำขอค้างจากเดือนอื่น');
r = filter(rows, '2026-11', true);
assert.deepStrictEqual(ids(r), ['A', 'B', 'C', 'E'], 'ติ๊กรวมค้าง = ทุกเดือน');
assert.strictEqual(r.backlog, 0);
assert.deepStrictEqual(ids(filter(rows, '', false)), ['A', 'B', 'C', 'E'], 'ไม่ได้เลือกเดือน = ทุกเดือน');

// UI + การผูกใช้งาน
assert(html.includes('id="prReqIncludeBacklog"'), 'มีช่องติ๊กรวมคำขอค้าง');
const loader = grabFn('loadOrderRequestPrCandidates');
assert(loader.includes('filterOrderRequestsForPr(rows, periodVal'), 'โหลดรายการผ่านตัวกรองเดือน');
assert(/On Hold'\) c\.selected = false/.test(loader), 'คำขอที่พักไว้ไม่ติ๊กอัตโนมัติ');
// เปิดจากหน้า "ขอซื้อ" (เลือกรายการมาเอง) ต้องไม่โดนกรองเดือน
assert(/if \(preset && preset\.length\) \{[\s\S]{0,300}idSet\[String\(r\.request_id\)\]/.test(loader));

console.log('pr-from-request-period-filter: OK');
