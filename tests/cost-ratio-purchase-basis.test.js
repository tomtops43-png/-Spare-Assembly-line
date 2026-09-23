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

// กราฟ % ต้นทุนคิดจากยอดสั่งซื้อ (PR) ใน Purchase History — ยอดเบิกมีกราฟแยกของตัวเองแล้ว
const api = new Function(
  'function fccN(v) { var n = Number(String(v == null ? "" : v).replace(/,/g, "")); return isFinite(n) ? n : 0; }\n' +
  'function fccLineOf(l) { return String(l || "").trim(); }\n' +
  extract('fccMonthKey(raw)') + '\n' +
  extract('fccPurchaseSpendRows(hist)') + '\nreturn { rows: fccPurchaseSpendRows };'
)();

const rows = api.rows([
  { line: 'Coil Winding', requested_date: '2026-09-22 17:42:17', qty_ordered: 12, unit_price: 7670, total_amount: 92040, status: 'Requested' },
  { line: 'Coil Winding', requested_date: '2026-09-22 17:21:59', qty_ordered: 10, unit_price: 3158.75, total_amount: 31587.5, status: 'Ordered' },
  { line: 'Coil Winding', requested_date: '2026-09-19', qty_ordered: 2, unit_price: '', total_amount: '', status: 'Requested', model: 'X-1' },
  { line: 'H9', requested_date: '2026-09-10', qty_ordered: 5, unit_price: 100, total_amount: '', status: 'Received' },
  { line: 'H9', requested_date: '2026-09-11', qty_ordered: 3, unit_price: 100, total_amount: 300, status: 'Cancelled' }
]);

assert.strictEqual(rows.length, 4, 'ใบที่ยกเลิกต้องไม่ถูกนับ');
assert(rows.every(function(r) { return r.month === '2026-09'; }), 'ต้องนับตามเดือนที่ขอซื้อ');
const coil = rows.filter(function(r) { return r.line === 'Coil Winding'; }).reduce(function(s, r) { return s + r.amount; }, 0);
assert.strictEqual(coil, 123627.5, 'ยอดสั่งซื้อของ Coil Winding ต้องเข้ากราฟ (92,040 + 31,587.50)');
assert.strictEqual(rows.filter(function(r) { return r.line === 'H9'; })[0].amount, 500, 'ไม่มี total_amount ต้องคำนวณ qty × ราคาเอง');
const unpriced = rows.filter(function(r) { return r.amount === 0; });
assert.strictEqual(unpriced.length, 1, 'รายการไม่มีราคาต้องแยกไว้เตือน ไม่รวมในยอด');
assert.strictEqual(unpriced[0].qty, 2);

const render = extract('fccRenderCostRatio()');
assert(render.includes('fccPurchaseSpendRows(hist)'), 'กราฟ % ต้นทุนต้องใช้ยอดสั่งซื้อ');
assert(!render.includes('fccExpenseValidRows('), 'กราฟ % ต้นทุนต้องไม่คิดจากยอดเบิกแล้ว');
assert(html.includes('<th style="padding:6px 10px;text-align:right;">ยอดสั่งซื้อ (PR)</th>'), 'หัวคอลัมน์ต้องบอกว่าเป็นยอดสั่งซื้อ');

console.log('Cost ratio purchase-basis checks passed');
