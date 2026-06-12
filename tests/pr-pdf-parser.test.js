const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const start = html.indexOf('    function normalizePdfHeaderToken');
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
  line('Mechanical 1 Bearing Lug&Screw 4 High 100', [token('Mechanical', 45), token('1', 20), token('Bearing', 90), token('Lug&Screw', 245), token('4', 315), token('High', 350), token('100', 415)]),
  line('NSK / 6201', [token('NSK / 6201', 90)])
];
const secondRowOnNextPage = [
  line('Electrical 2 Sensor Lug&Screw 2 Critical TBC', [token('Electrical', 45), token('2', 20), token('Sensor', 90), token('Lug&Screw', 245), token('2', 315), token('Critical', 350), token('TBC', 415)]),
  line('Omron / E2E-X5', [token('Omron / E2E-X5', 90)])
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


// H9 line-specific exports omit the Line column and can place left-aligned item
// text well before the centered Item / Model header. The raw-line fallback must
// recover both the item name and Brand / Model without treating Category as item.
const h9HeaderLines = [
  line('No. Category Item / Qty Critical Unit Total Drawing Reason', [
    token('No.', 20), token('Category', 70), token('Item /', 210), token('Qty', 400),
    token('Critical', 440), token('Unit', 500), token('Total', 555), token('Drawing', 610), token('Reason', 670)
  ]),
  line('Model Level Price Amount', [token('Model', 210), token('Level', 440), token('Price', 500), token('Amount', 555)])
];
const h9State = {};
const h9Header = { line: 'H9', request_period: '2026-05', requested_date: '2026-05-30 00:00:00', prepared_by: 'Wanchai', has_pr_header: true };
const h9Rows = context.parsePdfTableRows([
  line('6. Confirmed Price Items'),
  ...h9HeaderLines,
  line('Sensor 1 Fiber Unit Reflective 2 High 1900', [token('Sensor', 45), token('1', 20), token('Fiber Unit Reflective', 105), token('2', 390), token('High', 430), token('1900', 490)]),
  line('KEYENCE / FU-67V', [token('KEYENCE / FU-67V', 105)])
], { file_name: 'PR Report H9.pdf', file_hash: 'h9' }, h9Header, h9State);
assert.strictEqual(h9Rows.length, 1);
assert.strictEqual(h9Rows[0].part_name, 'Fiber Unit Reflective');
assert.strictEqual(h9Rows[0].brand, 'KEYENCE');
assert.strictEqual(h9Rows[0].model, 'FU-67V');
assert.strictEqual(h9Rows[0].line, 'H9');
assert.strictEqual(h9Rows[0].qty_ordered, 2);
assert.strictEqual(h9Rows[0].unit_price, 1900);



// Some browser/PDF combinations emit compact header labels and one physical
// header word per extracted line. These are still the application's own H9 report.
const compactH9State = {};
const compactH9Rows = context.parsePdfTableRows([
  line('7. Price Confirmation Required / TBC Items'),
  line('No.', [token('No.', 20)]),
  line('Category', [token('Category', 70)]),
  line('Item/Model', [token('Item/Model', 210)]),
  line('Qty', [token('Qty', 400)]),
  line('CriticalLevel', [token('CriticalLevel', 440)]),
  line('UnitPrice', [token('UnitPrice', 500)]),
  line('TotalAmount', [token('TotalAmount', 555)]),
  line('Drawing', [token('Drawing', 610)]),
  line('Reason', [token('Reason', 670)]),
  line('Sensor 1 Photoelectric Sensor 1 Critical TBC', [token('Sensor', 45), token('1', 20), token('Photoelectric Sensor', 105), token('1', 390), token('Critical', 430), token('TBC', 490)]),
  line('KEYENCE / PZ-G51N', [token('KEYENCE / PZ-G51N', 105)])
], { file_name: 'PR Report H9 compact.pdf', file_hash: 'h9compact' }, h9Header, compactH9State);
assert.strictEqual(compactH9State.headerFound, true);
assert.strictEqual(compactH9Rows.length, 1);
assert.strictEqual(compactH9Rows[0].part_name, 'Photoelectric Sensor');
assert.strictEqual(compactH9Rows[0].brand, 'KEYENCE');
assert.strictEqual(compactH9Rows[0].model, 'PZ-G51N');
assert.strictEqual(compactH9Rows[0].qty_ordered, 1);
assert.strictEqual(compactH9Rows[0].price_status, 'TBC');

console.log('System-exported PR PDF parser regression checks passed');
