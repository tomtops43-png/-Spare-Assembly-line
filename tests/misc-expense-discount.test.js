// ค่าใช้จ่ายสิ้นเปลือง: บิลที่มีส่วนลด/คูปองท้ายบิล
// ตัวอย่างจริง (Commercial Invoice 261107MUAJ0JU9): 2×174 = 348 − คูปอง 55 + ค่าจัดส่ง 36 = 329
const { html, backend, grabFn, assert } = require('./_export-extract');

const DISCOUNT = 'ส่วนลด';
function build(fields, rows) {
  const src = ['mxRowAmount', 'mxIsVatInclusiveVendor', 'mxVatInclusiveMode', 'mxExtraCharges', 'mxCollectItems'].map(function(n) { return grabFn(n); }).join('\n');
  return new Function('mxEl', 'mxEditMode', 'mxRowState', 'mxIsBlankRow', 'MISC_EXPENSE_DISCOUNT_CATEGORY',
    src + '\nreturn { extra: mxExtraCharges, collect: mxCollectItems, amount: mxRowAmount };')(
    function(id) { return id in fields ? (typeof fields[id] === 'object' ? fields[id] : { value: fields[id] }) : null; },
    function() { return false; },
    rows || [],
    function() { return false; },
    DISCOUNT);
}

const rows = [{ category: 'น็อต/สกรู/ฮาร์ดแวร์', item_name: 'V-shaped Torsion Spring 10pcs', qty: 2, unit: 'แพ็ค', unit_price: 174 }];
const fields = { miscExpenseVendor: 'Shuangpadong General Store', miscExpenseShipping: '36', miscExpenseDiscount: '55', miscExpenseVatEnabled: { checked: false }, miscExpenseVatPercent: '7' };
const api = build(fields, rows);

const ex = api.extra(348);
assert.strictEqual(ex.discount, 55);
assert.strictEqual(ex.subtotal + ex.vatAmount, 329, 'ยอดรวมทั้งบิล = ยอดจ่ายจริงตาม invoice');

const items = api.collect();
assert.deepStrictEqual(items.map(function(it) { return [it.category, it.total_amount]; }),
  [['น็อต/สกรู/ฮาร์ดแวร์', 348], ['ค่าจัดส่ง', 36], [DISCOUNT, -55]], 'ส่วนลดส่งเป็นแถวติดลบ');
assert.strictEqual(items.reduce(function(s, it) { return s + it.total_amount; }, 0), 329);

// พิมพ์ -55 ก็ได้ผลเหมือนกัน / ไม่มีส่วนลด = ไม่มีแถวส่วนลด
assert.strictEqual(build(Object.assign({}, fields, { miscExpenseDiscount: '-55' }), rows).extra(348).discount, 55);
assert(!build(Object.assign({}, fields, { miscExpenseDiscount: '' }), rows).collect().some(function(it) { return it.category === DISCOUNT; }));

// VAT คิดหลังหักส่วนลด: (348 + 36 − 55) × 7% = 23.03
const vat = build(Object.assign({}, fields, { miscExpenseVatEnabled: { checked: true } }), rows).extra(348);
assert.strictEqual(vat.vatAmount, 23.03);

// แถวหมวดส่วนลด (แก้ไขทีละแถว) คิดเป็นลบเสมอ
assert.strictEqual(api.amount({ category: DISCOUNT, qty: 1, unit_price: -55 }), -55);
assert.strictEqual(api.amount({ category: DISCOUNT, qty: 1, unit_price: 55 }), -55);
assert.strictEqual(api.amount({ category: 'อื่นๆ', qty: 1, unit_price: -55 }), 0, 'หมวดอื่นห้ามติดลบเหมือนเดิม');

// UI + draft
assert(html.includes('id="miscExpenseDiscount"'));
assert(grabFn('mxFormSnapshot').includes("discount: val('miscExpenseDiscount')"));
assert(grabFn('submitMiscExpenseForm').includes('ส่วนลดมากกว่ายอดรวมของบิลไม่ได้'));

// Backend: รับหมวดส่วนลด เก็บเป็นลบ หมวดอื่นยังห้ามติดลบ บิลรวมห้ามติดลบ
const cats = backend.match(/var MISC_EXPENSE_CATEGORIES = (\[[^\]]*\]);/)[1];
assert(cats.includes("'" + DISCOUNT + "'"));
const fieldsSrc = grabFn('buildMiscExpenseFields', backend, 0) + '\n' + grabFn('normalizeMiscExpenseCategory', backend, 0) +
  '\n' + grabFn('normalizeMiscExpenseLine', backend, 0);
const buildFields = new Function('MISC_EXPENSE_CATEGORIES', 'MISC_EXPENSE_DISCOUNT_CATEGORY', 'MISC_EXPENSE_LINES', 'MISC_EXPENSE_SHARED_LINE', 'normalizeMiscExpenseDate',
  fieldsSrc + '\nreturn buildMiscExpenseFields;')(eval(cats), DISCOUNT, ['H9', 'ส่วนกลาง'], 'ส่วนกลาง', function(d) { return d; });
const base = { date: '2026-11-07', line: 'H9', item_name: 'ส่วนลด/คูปอง', qty: 1 };
assert.strictEqual(buildFields(Object.assign({ category: DISCOUNT, total_amount: 55 }, base)).total, -55);
assert.strictEqual(buildFields(Object.assign({ category: DISCOUNT, total_amount: -55 }, base)).total, -55);
assert.throws(function() { buildFields(Object.assign({ category: 'อื่นๆ', total_amount: -55 }, base)); }, /ไม่ติดลบ/);
assert(grabFn('addMiscExpenseBatch', backend, 0).includes('ส่วนลดมากกว่ายอดรวมของบิลไม่ได้'));

console.log('misc-expense-discount: OK');
