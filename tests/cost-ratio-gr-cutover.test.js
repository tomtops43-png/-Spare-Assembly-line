// กราฟ % ต้นทุนหลังเปิดระบบ PR → รับของ: ก่อนวันเริ่มระบบนับยอดสั่งซื้อ (Purchase History) แบบเดิม
// ตั้งแต่วันเริ่มระบบนับยอดรับของจริง (GR) แทน — ห้ามนับซ้ำ และรับโดยไม่มี PR ไม่ใช่ค่าใช้จ่าย
const { html, grabFn, assert } = require('./_export-extract');

const api = new Function(
  'function fccN(v) { var n = Number(String(v == null ? "" : v).replace(/,/g, "")); return isFinite(n) ? n : 0; }\n' +
  'function fccLineOf(l) { return String(l || "").trim(); }\n' +
  ['fccMonthKey', 'fccPurchaseSpendRows', 'fccDayKey', 'fccGrSpendRows', 'fccSpendRows'].map(function(n) { return grabFn(n); }).join('\n') +
  '\nreturn { spend: fccSpendRows, day: fccDayKey };'
)();

const hist = [
  { line: 'H9', requested_date: '2026-09-20', qty_ordered: 2, unit_price: 100, status: 'Received' },          // ก่อนเริ่มระบบ = นับ
  { line: 'H9', requested_date: '2026-10-03 08:00:00', qty_ordered: 1, unit_price: 999, status: 'Requested' }, // ก่อนเริ่มระบบ (เดือนเดียวกัน) = นับ
  { line: 'H9', requested_date: '2026-10-07 09:00:00', qty_ordered: 5, unit_price: 50, status: 'Requested' },  // หลังเริ่มระบบ = ไม่นับ (มาจาก GR แทน)
  { line: 'H9', requested_date: '2026-09-25', qty_ordered: 9, unit_price: 10, status: 'Cancelled' }            // ยกเลิก = ไม่นับ
];
const gr = {
  cutoverAt: '2026-10-05 10:00:00',
  receipts: [
    { gr_id: 'GR-1', received_at: '2026-10-08 09:00:00', pr_id: 'PR-2610-001', line: 'H9', qty: 4, unit_price: 50, amount: 200, status: 'POSTED' },
    { gr_id: 'GR-2', received_at: '2026-11-02 09:00:00', pr_id: 'PR-2610-001', line: 'H9', qty: 2, unit_price: 50, amount: '', status: 'POSTED' },
    { gr_id: 'GR-3', received_at: '2026-10-09 09:00:00', pr_id: '', no_pr_reason: 'ของแถม', line: 'H9', qty: 3, amount: 0, status: 'POSTED' },
    { gr_id: 'GR-4', received_at: '2026-10-09 10:00:00', pr_id: 'PR-2610-002', line: 'H9', qty: 1, amount: 70, status: 'REVERSED' }
  ]
};

function byMonth(rows) {
  const out = {};
  rows.forEach(function(r) { out[r.month] = (out[r.month] || 0) + r.amount; });
  return out;
}

// ยังไม่เปิดระบบ (cutover ว่าง) = ยอดสั่งซื้อเดิมทุกแถว ไม่สนใจ GR
assert.deepStrictEqual(byMonth(api.spend(hist, { cutoverAt: '', receipts: gr.receipts })), { '2026-09': 200, '2026-10': 999 + 250 });
assert.deepStrictEqual(byMonth(api.spend(hist, null)), { '2026-09': 200, '2026-10': 999 + 250 }, 'ยังโหลด GR ไม่ได้ = แบบเดิม');

// เปิดระบบแล้ว
const rows = api.spend(hist, gr);
assert.deepStrictEqual(byMonth(rows), { '2026-09': 200, '2026-10': 999 + 200, '2026-11': 100 },
  'ต.ค. = PR ก่อนวันเริ่ม 999 + GR 200 (ไม่นับ PR หลังวันเริ่ม 250 ซ้ำ) / พ.ย. = GR ที่ไม่มี amount คิด qty × ราคา');
assert(!rows.some(function(r) { return r.qty === 3; }), 'รับโดยไม่มี PR ต้องไม่ถูกนับ และไม่ขึ้นเตือนว่าไม่มีราคา');
assert(!rows.some(function(r) { return r.amount === 70; }), 'GR ที่ถูกคืนรายการต้องไม่ถูกนับ');

// วันที่ทั้งแบบสตริงและ Date ต้องเทียบกันได้ระดับวัน (เวลาไทย)
assert.strictEqual(api.day('2026-10-05 10:00:00'), '2026-10-05');
assert.strictEqual(api.day(new Date('2026-10-04T18:30:00Z')), '2026-10-05', 'Date ตอนตีหนึ่งครึ่งเวลาไทย ต้องเป็นวันที่ 5');
assert.strictEqual(api.day(''), '');

// Dashboard ต้องโหลด GR + PR ค้างรับ และกราฟต้องแสดงยอดผูกพันแยก
const ensure = grabFn('fccEnsureAsync');
assert(ensure.includes("action: 'getGrLog'") && ensure.includes("action: 'getOpenPrLines'"), 'Dashboard โหลด GR และ PR ค้างรับ');
const render = grabFn('fccRenderCostRatio');
assert(render.includes('ยอดผูกพัน (PR อนุมัติแล้ว รอของเข้า)'), 'กราฟแสดงยอดผูกพันแยกจากรายจ่าย');
assert(grabFn('ptAfterReceive').includes('fccInvalidateAsyncCache()'), 'รับของแล้วต้องล้างแคช Dashboard');

console.log('Cost ratio GR cutover checks passed');
