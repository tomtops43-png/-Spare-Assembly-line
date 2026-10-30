// แผงงบ Spare part ในหน้า PR ต้องตรงกับการ์ด Cost Ratio:
// 1) "ใช้" = ยอดสั่งซื้อ/รับจริง (fccSpendRows) + ของสิ้นเปลือง — ไม่ใช่ค่าเบิก (เบิกจริง × ราคา)
// 2) "PR ใบนี้" = เฉพาะรายการที่จะถูกส่งจริง (ตามแท็บไลน์ที่เปิด) ไม่ใช่ทุกรายการที่ติ๊กไว้ในทุกไลน์
const { html, grabFn, assert } = require('./_export-extract');

const compute = grabFn('prComputeBudget');
assert(compute.includes('fccSpendRows(A.hist, A.gr)'), 'งบ PR ต้องนับรายจ่ายด้วย fccSpendRows แบบเดียวกับการ์ด');
assert(!compute.includes('fccExpenseValidRows'), 'งบ PR ห้ามนับค่าเบิก — การ์ดไม่นับค่าเบิกในงบ');
assert(compute.includes('fccMiscExpenseRows(A.misc)'), 'ของสิ้นเปลืองยังนับในงบเหมือนการ์ด');
assert(grabFn('fccRenderCostRatio').includes('fccSpendRows(hist, grInfo)'), 'การ์ดใช้ fccSpendRows เป็นฐานเดียวกัน');

// รันจริง: ติ๊กไว้ 3 ไลน์ แต่เปิดแท็บ H9 อยู่ → PR ใบนี้ต้องมีแค่ H9 (ตรงกับที่ prSubmitCreatePr ส่ง)
const api = new Function(
  'var prCandidateState = [' +
  '{ line: "H9", selected: true, suggestedQty: 2, unitPrice: 100, name: "a" },' +
  '{ line: "H9", selected: false, suggestedQty: 9, unitPrice: 100, name: "b" },' +
  '{ line: "Coil Winding", selected: true, suggestedQty: 1000, unitPrice: 500, name: "c" },' +
  '{ line: "Lug&Screw", selected: true, suggestedQty: 1000, unitPrice: 700, name: "d" }];\n' +
  'var prLineTabState = "H9";\n' +
  'var document = { getElementById: function() { return null; } };\n' +
  'function fccN(v) { var n = Number(v); return isFinite(n) ? n : 0; }\n' +
  'function fccLineOf(l) { return String(l || "").trim(); }\n' +
  ['getFilteredPrCandidates', 'prSelectedAmountByLine'].map(n => grabFn(n)).join('\n') +
  '\nreturn { byLine: prSelectedAmountByLine, setTab: function(t) { prLineTabState = t; } };'
)();
assert.deepStrictEqual(api.byLine(), { H9: 200 }, 'แท็บ H9 = ยอดเฉพาะ H9 ที่ติ๊กไว้');
api.setTab('all');
assert.deepStrictEqual(api.byLine(), { H9: 200, 'Coil Winding': 500000, 'Lug&Screw': 700000 }, 'แท็บทั้งหมด = ทุกไลน์ที่ติ๊ก');

// prSubmitCreatePr ต้องใช้ชุดเดียวกัน
assert(/function prSubmitCreatePr\(\)\s*\{\s*var selected = getFilteredPrCandidates\(true\);/.test(html), 'ส่ง PR ใช้รายการชุดเดียวกับแผงงบ');

console.log('PR budget matches cost ratio checks passed');

// ---- หน้าตรวจ + แก้ยอด PR: คำนวณงบสดจากรายการในใบ ไม่ใช้สแนปช็อต ----
const appr = new Function(
  'function fccN(v) { var n = Number(v); return isFinite(n) ? n : 0; }\n' +
  'function fccLineOf(l) { return String(l || "").trim(); }\n' +
  'function prBudgetMonthKey() { return "2099-01"; }\n' +
  ['prPendingUsageMap', 'prApprovalAmountByLine', 'prApprovalMonthKey'].map(n => grabFn(n)).join('\n') +
  '\nreturn { pending: prPendingUsageMap, amount: prApprovalAmountByLine, month: prApprovalMonthKey };'
)();
const state = { header: { pr_id: 'PR-2610-009', line: 'H9', created_at: '2026-10-28 10:00:00' },
  lines: [{ line_no: 1, qty_approved: 2, unit_price: 100 }, { line_no: 2, qty_approved: 1, unit_price: 50 }] };
assert.deepStrictEqual(appr.amount(state), { H9: 250 }, 'PR ใบนี้ = จำนวนอนุมัติ × ราคา ลงไลน์ของหัว PR');
assert.deepStrictEqual(appr.amount(state, { 1: 5 }), { H9: 550 }, 'แก้จำนวนในช่องแล้วยอดตามทันที');
assert.strictEqual(appr.month(state), '2026-10', 'ใช้งบของเดือนที่สร้าง PR');
const pendingList = [
  { pr_id: 'PR-2610-009', line: 'H9', month: '2026-10', amount: 250 },
  { pr_id: 'PR-2610-010', line: 'H9', month: '2026-10', amount: 40 }
];
assert.deepStrictEqual(appr.pending(pendingList), { '2026-10': { H9: 290 } });
assert.deepStrictEqual(appr.pending(pendingList, 'PR-2610-009'), { '2026-10': { H9: 40 } }, 'ใบที่กำลังตรวจต้องไม่ถูกนับซ้ำใน "รอ"');
assert(grabFn('prComputeBudget').includes('prPendingUsageMap(arr[1], excludePrId)'));
assert(grabFn('renderPrApprovalBudget').includes('prComputeBudget(monthKey, prId)'));
console.log('PR approval live budget checks passed');
