const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const start = html.indexOf('    function findPdfColumnStarts');
const end = html.indexOf('    function parsePurchaseHistoryPdf', start);
assert(start >= 0 && end > start);
const context = { console, Date, Number, String, Array, Math, Object, Promise };
vm.createContext(context);
vm.runInContext(html.slice(start, end), context);
const token = (text, x) => ({ text, x, width: 10 });
const line = (text, items) => ({ text, items: items || [] });
const splitHeader = [
  line('No. Category Item / Line Qty Critical Unit Total Drawing Reason', [
    token('No.', 20), token('Category', 55), token('Item /', 125), token('Line', 285), token('Qty', 325),
    token('Critical', 365), token('Unit', 430), token('Total', 485), token('Drawing', 545), token('Reason', 605)
  ]),
  line('Model Level Price Amount', [token('Model', 125), token('Level', 365), token('Price', 430), token('Amount', 485)])
];
const firstRow = [
  // Category token intentionally precedes the row number in PDF text order.
  line('Mechanical 1 Bearing Lug&Screw 4 High 100', [token('Mechanical', 55), token('1', 20), token('Bearing', 125), token('Lug&Screw', 285), token('4', 325), token('High', 365), token('100', 430)]),
  line('NSK / 6201', [token('NSK / 6201', 125)])
];
const secondRowOnNextPage = [
  line('Electrical 2 Sensor Lug&Screw 2 Critical TBC', [token('Electrical', 55), token('2', 20), token('Sensor', 125), token('Lug&Screw', 285), token('2', 325), token('Critical', 365), token('TBC', 430)]),
  line('Omron / E2E-X5', [token('Omron / E2E-X5', 125)])
];
const state = {};
const meta = { file_name: 'PR Report Lug&Screw.pdf', file_hash: 'hash' };
const header = { line: 'Lug&Screw', request_period: '2026-05', requested_date: '2026-05-30 00:00:00', prepared_by: 'Wanchai', has_pr_header: true };
const rows1 = context.parsePdfTableRows([line('6. Detailed Requested Item List by Line'), line('10.1 Lug&Screw'), line('10.1.1 Confirmed Price Items'), ...splitHeader, ...firstRow], meta, header, state);
const rows2 = context.parsePdfTableRows(secondRowOnNextPage, meta, header, state);
const rows = rows1.concat(rows2);
assert.strictEqual(rows.length, 2);
assert.strictEqual(rows[0].part_name, 'Bearing');
assert.strictEqual(rows[0].brand, 'NSK');
assert.strictEqual(rows[0].model, '6201');
assert.strictEqual(rows[0].qty_ordered, 4);
assert.strictEqual(rows[0].unit_price, 100);
assert.strictEqual(rows[0].line, 'Lug&Screw');
assert.strictEqual(rows[1].part_name, 'Sensor');
assert.strictEqual(rows[1].model, 'E2E-X5');
assert.strictEqual(rows[1].price_status, 'TBC');
assert.strictEqual(state.headerFound, true);
assert.strictEqual(state.rowNumbersFound, 2);
assert(state.sectionsFound >= 1);
const urgentState = {};
const urgentRows = context.parsePdfTableRows([line('5. Top Urgent Items'), ...splitHeader, ...firstRow], meta, header, urgentState);
assert.strictEqual(urgentRows.length, 0);
console.log('System-exported PR PDF parser regression checks passed');
