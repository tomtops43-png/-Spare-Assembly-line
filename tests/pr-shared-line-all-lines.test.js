// PR ที่เลือกไลน์ "ส่วนกลาง" — ใบ PR ต้องขึ้น Assembly Line (รวมทุกไลน์) ไม่ใช่ H9
// ต้นเหตุ: (1) ตอนส่ง PR เอาค่าช่อง Line ที่ซ่อนอยู่ (ค้าง H9) มาเป็นหัว PR (2) canonicalLineName('ส่วนกลาง') → mapLine → H9
const { html, grabFn, assert } = require('./_export-extract');

const consts = html.match(/var PR_DOC_ALL_LINES = [^\n]+/)[0] + '\n' + html.match(/var PR_DOC_REAL_LINES = [^\n]+/)[0];
const src = ['isSharedLineName', 'prHeaderLineForSubmit', 'prDocLineName', 'prBudgetWithAllLines', 'prDocBudgetRows', 'fccN']
  .map(function(n) { return grabFn(n); }).join('\n');
const api = new Function('canonicalLineName', consts + '\n' + src +
  '\nreturn { header: prHeaderLineForSubmit, name: prDocLineName, withAll: prBudgetWithAllLines, rows: prDocBudgetRows, ALL: PR_DOC_ALL_LINES };')(
  function(v) { return /h9/i.test(v) ? 'H9' : v === 'Arc Chute' ? 'Arc Chute' : 'H9'; });

assert.strictEqual(api.ALL, 'Assembly Line (All lines)');
// หัว PR: เคสในภาพ — โหมดคำขอซื้อ ช่อง Line ค้าง H9 แต่รายการเป็นส่วนกลาง
assert.strictEqual(api.header('from_request', 'H9', [{ line: 'ส่วนกลาง' }]), 'ส่วนกลาง');
assert.strictEqual(api.header('from_request', 'H9', [{ line: 'Arc Chute' }]), 'Arc Chute', 'โหมดอื่นไม่ใช้ช่อง Line ที่ซ่อนอยู่');
assert.strictEqual(api.header('from_request', 'all', [{ line: 'H9' }, { line: 'ส่วนกลาง' }]), 'ส่วนกลาง');
assert.strictEqual(api.header('line_specific', 'Coil Winding', [{ line: 'H9' }]), 'Coil Winding', 'โหมดเลือกไลน์ยังใช้ไลน์ที่เลือก');
assert.strictEqual(api.header('line_specific', 'all', [{ line: 'H9' }]), 'H9');
assert(html.includes("var headerLine = prHeaderLineForSubmit(String((prMode && prMode.value) || ''), prLineFilter ? prLineFilter.value : '', selected);"));

// ชื่อไลน์ในใบ PR
assert.strictEqual(api.name('ส่วนกลาง'), 'Assembly Line (All lines)');
assert.strictEqual(api.name('Assembly Line (All lines)'), 'Assembly Line (All lines)');
assert.strictEqual(api.name('H9'), 'H9');
assert(html.includes("var itemLine = prDocLineName(c.line || '-');"));
assert(html.includes("var lineLabel = isLineSpecific ? prDocLineName(selectedLineNow) : (lines.join(' / ') || '-');"));

// งบ: ส่วนกลาง = รวมทุกไลน์
const B = { byLine: {
  'H9': { sales: 1000, base: 30, budget: 30, used: 10, spent: 10, pending: 0, carryIn: 0, quarter: { key: 'Q4', budget: 90, used: 20, sales: 3000 }, history: [{ month: '2026-10', budget: 30, used: 5 }, { month: '2026-11', budget: 30, used: 10 }] },
  'Arc Chute': { sales: 2000, base: 40, budget: 40, used: 15, spent: 15, pending: 0, carryIn: 0, quarter: { key: 'Q4', budget: 120, used: 30, sales: 6000 }, history: [{ month: '2026-11', budget: 40, used: 15 }] },
  'ส่วนกลาง': { sales: 0, base: 0, budget: 0, used: 999, quarter: {}, history: [] }
} };
const B2 = api.withAll(B);
const all = B2.byLine['Assembly Line (All lines)'];
assert.strictEqual(all.budget, 70); assert.strictEqual(all.used, 25); assert.strictEqual(all.sales, 3000);
assert.strictEqual(all.quarter.budget, 210); assert.strictEqual(all.quarter.used, 50);
assert.strictEqual(all.targetPct, 2.33);
assert.deepStrictEqual(all.history, [{ month: '2026-10', budget: 30, used: 5 }, { month: '2026-11', budget: 70, used: 25 }]);
assert(!B.byLine['Assembly Line (All lines)'], 'ไม่แก้ object เดิม');
const rows = api.rows(B2, { 'Assembly Line (All lines)': 690 });
assert.strictEqual(rows.length, 1);
assert.strictEqual(rows[0].line, 'Assembly Line (All lines)');
assert.strictEqual(rows[0].m.budget, 70); assert.strictEqual(rows[0].m.pr, 690);
// ยอด PR 0 → fallback ไม่โผล่แถวรวม
assert(api.rows(B2, {}).every(function(r) { return r.line !== 'Assembly Line (All lines)'; }));
assert(html.includes("var ln = isSharedLineName(c.line) ? PR_DOC_ALL_LINES : fccLineOf(canonicalLineName(c.line || '-'));"));
assert(html.includes('history: prDocBudgetHistory(B, prByLine[PR_DOC_ALL_LINES] > 0 ? [PR_DOC_ALL_LINES] :'));

console.log('pr-shared-line-all-lines: OK');
