const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];

function extract(sig) {
  const start = script.indexOf('    function ' + sig);
  assert(start > -1, 'cannot find ' + sig);
  const end = script.indexOf('\n    }', start);
  return script.slice(start, end + 6);
}

// การ์ด Cost Ratio ดูได้ทั้งรายเดือนและรายไตรมาส (กราฟ + ตารางงบ)
const api = new Function(
  "var FCC_TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];\n" +
  extract('fccMonthKey(raw)') + '\n' + extract('fccMonthLabel(m)') + '\n' +
  extract('fccQuarterKey(monthKey)') + '\n' + extract('fccPeriodLabel(key)') + '\n' +
  extract('fccQuarterRangeLabel(key)') + '\n' + extract('fccRollupByQuarter(map)') + '\n' +
  'return { qk: fccQuarterKey, label: fccPeriodLabel, range: fccQuarterRangeLabel, rollup: fccRollupByQuarter };'
)();

assert.strictEqual(api.qk('2026-01'), '2026-Q1');
assert.strictEqual(api.qk('2026-03'), '2026-Q1');
assert.strictEqual(api.qk('2026-07'), '2026-Q3');
assert.strictEqual(api.qk('2026-12'), '2026-Q4');
assert.strictEqual(api.label('2026-Q3'), 'Q3 2026');
assert.strictEqual(api.label('2026-10'), 'ต.ค. 2026', 'key เดือนยังได้ชื่อเดือนแบบเดิม');
assert.strictEqual(api.range('2026-Q3'), 'ก.ค.–ก.ย.');
assert.deepStrictEqual(
  api.rollup({ '2026-07': 100, '2026-08': 50, '2026-09': 25, '2026-10': 10, '2025-12': 1 }),
  { '2026-Q3': 175, '2026-Q4': 10, '2025-Q4': 1 },
  'ต้องรวมยอดของ 3 เดือนในไตรมาสเดียวกัน'
);
assert(['2025-Q4', '2026-Q1', '2026-Q3'].slice().sort().join() === '2025-Q4,2026-Q1,2026-Q3', 'key ไตรมาสเรียงตามตัวอักษรได้ถูกลำดับ');

assert(html.includes('<select id="fccCostRatioPeriod"'), 'ต้องมีตัวเลือกรายเดือน/รายไตรมาส');
assert(html.includes('<option value="quarter">รายไตรมาส</option>'));
const render = extract('fccRenderCostRatio()');
['spendByMonth', 'valueByMonth', 'qtyByMonth', 'budgetByMonth', 'issueByMonth', 'miscByMonth', 'committedByMonth'].forEach(function(k) {
  assert(render.includes(k + ' = fccRollupByQuarter(' + k + ')'), 'รายไตรมาสต้องรวม ' + k);
});
assert(render.includes('label: fccPeriodLabel(m)'), 'ป้ายแกนกราฟต้องรองรับไตรมาส');
assert(render.includes("(isQuarter ? 'ไตรมาส' : 'เดือน')"), 'หัวตารางต้องเปลี่ยนตามมุมมอง');
assert(script.includes("fccCostRatioPeriodEl.addEventListener('change'"), 'เปลี่ยนมุมมองแล้วต้อง render ใหม่');

console.log('Cost ratio quarterly view checks passed');
