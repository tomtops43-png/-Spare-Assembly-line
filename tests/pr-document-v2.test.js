const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

// ใบ PR แบบใหม่ (ภาษาอังกฤษ): โลโก้ + รายการพร้อมรูป/สต็อก/Min/Max/ราคา + งบ Spare part รายเดือน & รายไตรมาส
function grab(name) {
  const start = html.indexOf(name);
  assert(start > -1, 'cannot find ' + name);
  let i = html.indexOf('{', start), depth = 0;
  for (; i < html.length; i += 1) { if (html[i] === '{') depth += 1; else if (html[i] === '}' && --depth === 0) break; }
  return html.slice(start, i + 1);
}
const consts = html.match(/var PR_DOC_LOGO_PATH = [^\n]+\n/)[0] + html.match(/var PR_DOC_EN_MONTHS = [^\n]+\n/)[0];
const api = new Function('partsData', 'dashboardAllPartsData',
  'var window = { location: { href: "https://example.com/app/index.html" } };\n' +
  'var document = { getElementById: function() { return null; } };\n' +
  'function canonicalLineName(l) { return String(l || "-").trim() || "-"; }\n' +
  grab('function firstDefined(values, fallbackValue)') + '\n' + consts +
  grab('function prDocLogoUrl()') + '\n' + grab('function prDocStockStatus(item)') + '\n' +
  grab('function buildDocumentGradeReport(selected, options)') + '\n' +
  grab('function prDocBudgetRows(B, prByLine)') + '\n' + grab('function prDocBudgetHistory(B, lines, monthKey, prByLine)') + '\n' + grab('function prBudgetDataWithoutPr(A, excludePrId)') + '\n' +
  'return { build: buildDocumentGradeReport, status: prDocStockStatus, rows: prDocBudgetRows, history: prDocBudgetHistory, without: prBudgetDataWithoutPr, logo: prDocLogoUrl };'
);
const master = [
  { no: 'P1', line: 'H9', name: 'Sensor', model: 'E3Z', stock: 0, min: 2, max: 6 },
  { no: 'P2', line: 'H9', name: 'Belt', model: 'B-10', stock: 1, min: 3, max: 8 },
  { no: 'P3', line: 'H9', name: 'Bearing', model: '6202', stock: 10, min: 2, max: 10, image_main_url: 'https://drive.google.com/file/d/ABCDEFGHIJKLMN/view' }
];
const fn = api([], master);

assert.strictEqual(fn.logo(), 'https://example.com/app/images/Logo%20JR.jpg', 'โลโก้ใช้ URL เต็ม (เอกสารเขียนลง about:blank)');
assert.strictEqual(fn.status({ stockKnown: true, stock: 0, min: 2 }), 'out');
assert.strictEqual(fn.status({ stockKnown: true, stock: 1, min: 3 }), 'low');
assert.strictEqual(fn.status({ stockKnown: true, stock: 5, min: 3 }), 'ok');
assert.strictEqual(fn.status({ stockKnown: false, stock: 0, min: 3 }), 'unknown', 'ไม่รู้สต็อก ห้ามขึ้นว่าหมด');

const items = [
  { no: 'P1', line: 'H9', name: 'Sensor', model: 'E3Z', suggestedQty: 4, unit_price: 1500 },
  { no: 'P2', line: 'H9', name: 'Belt', model: 'B-10', suggestedQty: 5, unit_price: 200 },
  { no: 'P3', line: 'H9', name: 'Bearing', model: '6202', suggestedQty: 2, unit_price: '' }
];
const budgetDoc = {
  month: '2026-10', monthLabel: 'Oct 2026', quarterLabel: 'Q4 2026', quarterRange: 'Oct–Dec',
  rows: [{ line: 'H9', targetPct: 5, m: { budget: 20000, used: 5000, pr: 7000 }, q: { budget: 20000, used: 5000, pr: 7000 } }]
};
const doc = fn.build(items, { prNumber: 'PR-TEST-1', budgetDoc: budgetDoc, statusOverride: 'APPROVED', approvedByOverride: 'Boss' });
assert(doc.includes('class="doc-logo" src="https://example.com/app/images/Logo%20JR.jpg"'), 'มีโลโก้มุมบนซ้าย');
assert(doc.includes('PURCHASE REQUEST') && doc.includes('PR-TEST-1'));
['Picture', 'Stock Status', 'On Hand', 'Min', 'Max', 'Order Qty', 'Unit Price (THB)', 'Amount (THB)'].forEach(function(h) {
  assert(doc.includes('<th>' + h + '</th>'), 'ตารางรายการต้องมีคอลัมน์ ' + h);
});
assert(doc.includes('Out of Stock: <b>1</b>') && doc.includes('Below Min: <b>1</b>') && doc.includes('In Stock: <b>1</b>'), 'สรุปสถานะสต็อก');
assert(doc.includes('https://drive.google.com/thumbnail?id=ABCDEFGHIJKLMN&amp;sz=w400'), 'รูปจาก master ขึ้นในตาราง');
assert(doc.includes('exceeds Max'), 'Bearing 10 + 2 > Max 10 ต้องเตือน');
assert(doc.includes('7,000.00'), 'ยอดรวม (ไม่รวม TBC) = 6,000 + 1,000');
assert(doc.includes('excludes 1 TBC item(s)'));
assert(doc.includes('WITHIN BUDGET — OK TO PURCHASE') && doc.includes('THB 8,000.00'), 'คงเหลือหลัง PR = 20,000 − 5,000 − 7,000');
assert(doc.includes('Monthly — Oct 2026') && doc.includes('Quarterly — Q4 2026'), 'แสดงทั้งรายเดือนและรายไตรมาส');
assert(doc.includes('&#10004; APPROVED') && doc.includes('Boss'));
// ส่วนเก่าต้องหายไปหมด
['Annex', 'Drawing', 'Executive Summary', 'Critical Level', 'ใบขออนุมัติ'].forEach(function(t) { assert(!doc.includes(t), 'ต้องไม่มี ' + t + ' แล้ว'); });

const over = fn.build(items, { budgetDoc: Object.assign({}, budgetDoc, { rows: [{ line: 'H9', targetPct: 5, m: { budget: 10000, used: 5000, pr: 7000 }, q: { budget: 10000, used: 5000, pr: 7000 } }] }) });
assert(over.includes('OVER BUDGET') && over.includes('THB 2,000.00'), 'เกินงบ 2,000');
assert(over.includes('PENDING APPROVAL'));
assert(fn.build(items, {}).includes('Budget data was not available'), 'ไม่มีข้อมูลงบ ยังออกเอกสารได้');

const salesDoc = fn.build(items, { budgetDoc: Object.assign({}, budgetDoc, { rows: [{ line: 'H9', targetPct: 5, m: { sales: 400000, budget: 20000, used: 5000, pr: 7000 }, q: { sales: 1200000, budget: 60000, used: 5000, pr: 7000 } }] }) });
assert(salesDoc.includes('<th>Production Value (Sales)</th><th>Budget</th>'), 'คอลัมน์มูลค่าผลิตอยู่หน้างบ');
assert(salesDoc.includes('400,000.00') && salesDoc.includes('1,200,000.00'), 'แสดงมูลค่าผลิตเดือน + ไตรมาส');
assert(grab('function prComputeBudget(monthKey, excludePrId)').includes('sales: qSales'));

// เกินงบรายเดือน แต่ไตรมาสยังไม่เกิน → ต้องบอกให้ชัด (กล่องสีเหลือง)
const warn = fn.build(items, { budgetDoc: Object.assign({}, budgetDoc, { rows: [{ line: 'H9', targetPct: 5, m: { budget: 10000, used: 5000, pr: 7000 }, q: { budget: 30000, used: 5000, pr: 7000 } }] }) });
assert(warn.includes('OVER MONTHLY BUDGET — STILL WITHIN QUARTERLY BUDGET') && warn.includes('THB 18,000.00'), 'เกินเดือน 2,000 แต่ไตรมาสเหลือ 18,000');
assert(over.includes('OVER BUDGET — MONTHLY &amp; QUARTERLY'), 'เกินทั้งเดือนและไตรมาส');
assert(!warn.includes('MONTHLY &amp; QUARTERLY'), 'ไตรมาสไม่เกิน ห้ามขึ้นว่าเกินทั้งคู่');

// กราฟควบคุมงบ: 6 เดือน (ไตรมาสก่อน + ไตรมาสนี้) แถบสรุปไตรมาส + เส้นงบ + PR ใบนี้
const hB = { byLine: { H9: { history: [{ month: '2026-07', budget: 100, used: 80 }, { month: '2026-10', budget: 100, used: 30 }, { month: '2026-11', budget: 100, used: 90 }] } } };
const slots = fn.history(hB, ['H9'], '2026-11', { H9: 40 });
assert.deepStrictEqual(slots.map(function(x) { return x.month; }), ['2026-07', '2026-08', '2026-09', '2026-10', '2026-11', '2026-12']);
assert.strictEqual(slots[4].pr, 40, 'ยอดใบนี้อยู่ที่เดือนที่ขอ');
assert.strictEqual(slots[3].pr, 0);
assert(slots[5].future && !slots[4].future && slots[4].current);
assert.strictEqual(fn.history(hB, ['H9'], '2026-01', {})[0].month, '2025-10', 'ข้ามปีได้');
const chartDoc = fn.build(items, { budgetDoc: Object.assign({}, budgetDoc, { month: '2026-11', history: slots, rows: [{ line: 'H9', targetPct: 5, m: { budget: 100, used: 90, pr: 40 }, q: { budget: 200, used: 120, pr: 40 } }] }) });
assert(chartDoc.includes('<svg class="trend"'), 'มีกราฟ SVG');
assert(chartDoc.includes('Q4 2026 (to date) — 80% of budget · WITHIN'), 'แถบไตรมาสบอก % และสถานะ (30+90+40)/200');
assert(chartDoc.includes('Q3 2026 — 80% of budget · WITHIN'));
assert(chartDoc.includes('130 · 130%'), 'แท่งเดือน Nov = 90 + 40 เทียบงบ 100');
assert(chartDoc.includes('▲ this PR') && chartDoc.includes('upcoming'));
assert(chartDoc.includes('OVER MONTHLY BUDGET — STILL WITHIN QUARTERLY BUDGET'));

// งบต่อไลน์ของ PR ใบนี้
const B = { byLine: { H9: { sales: 2000, budget: 100, used: 40, targetPct: 5, quarter: { budget: 300, used: 90, sales: 6000 } }, 'Arc Chute': { budget: 50, used: 0, targetPct: 4 } } };
assert(grab('function prComputeBudget(monthKey, excludePrId)').includes('byLine[ln].history = history;'), 'เก็บประวัติรายเดือนไว้วาดกราฟ');
assert.deepStrictEqual(fn.rows(B, { H9: 30 }), [{ line: 'H9', targetPct: 5, m: { sales: 2000, budget: 100, used: 40, pr: 30 }, q: { sales: 6000, budget: 300, used: 90, pr: 30 } }], 'มีมูลค่าผลิต (ยอดขาย) ทั้งรายเดือน/ไตรมาส');
assert.strictEqual(fn.rows(B, {}).length, 2, 'ไม่มียอด PR = โชว์ทุกไลน์ที่มีงบ');

// ใบที่อนุมัติแล้วต้องไม่ถูกนับซ้ำ (ค้างรับ/รับของของใบนี้ตัดออก)
const A = { openPr: [{ pr_id: 'PR-1' }, { pr_id: 'PR-2' }], gr: { cutoverAt: 'x', receipts: [{ pr_id: 'PR-1' }, { pr_id: 'PR-3' }] }, hist: [] };
const cut = fn.without(A, 'PR-1');
assert.deepStrictEqual(cut.openPr, [{ pr_id: 'PR-2' }]);
assert.deepStrictEqual(cut.gr.receipts, [{ pr_id: 'PR-3' }]);
assert.strictEqual(A.openPr.length, 2, 'ไม่แก้ข้อมูลแคชเดิม');
assert.strictEqual(fn.without(A, ''), A);

// ผูกเข้ากับจุดปริ้นจริง
assert(grab('function prComputeBudget(monthKey, excludePrId)').includes('prBudgetDataWithoutPr(arr[0], excludePrId)'));
assert(grab('function prComputeBudget(monthKey, excludePrId)').includes('byLine[ln].quarter = { key: qKey, budget: qBudget, used: qUsed, sales: qSales };'));
assert(html.includes('prDocAttachBudget(edited.opts, edited.items, prDocMonthFromDate(header.created_at), header.pr_id)'), 'ปริ้นจาก Inbox คำนวณงบก่อน');
assert(html.includes("prDocAttachBudget(opts, selected, prBudgetMonthKey(), '')"), 'Preview คำนวณงบก่อน');
assert(!html.includes('id="prIncludeAnnexD"'), 'ตัวเลือก Drawing/Annex ถูกเอาออก');

console.log('pr-document-v2: ok');
