// บิลฮาร์ดแวร์เฮาส์: ราคาสินค้ารวม VAT แล้ว → ถอด VAT ออกแบบท้ายบิล ยอดจ่ายไม่เพิ่ม
// ตัวอย่างจริง CS4691029-0433: 2×134 + 1×105 + 10×21 = 583 → รวมยอดขาย 544.86 + VAT 7% 38.14
const { html, grabFn, assert } = require('./_export-extract');

function build(fields, editMode, rows) {
  const src = ['mxIsVatInclusiveVendor', 'mxVatInclusiveMode', 'mxExtraCharges', 'mxCollectItems'].map(function(n) { return grabFn(n); }).join('\n');
  return new Function('mxEl', 'mxEditMode', 'mxRowState', 'mxIsBlankRow', 'mxRowAmount', src + '\nreturn { extra: mxExtraCharges, collect: mxCollectItems, isInc: mxIsVatInclusiveVendor };')(
    function(id) { return id in fields ? (typeof fields[id] === 'object' ? fields[id] : { value: fields[id] }) : null; },
    function() { return !!editMode; },
    rows || [],
    function() { return false; },
    function(r) { return Math.round(r.qty * r.unit_price * 100) / 100; });
}
const billRows = [
  { category: 'กล่องเครื่องมือ', item_name: 'กล่องตัดหัก 4 ช่อง CK-324', qty: 2, unit: 'ชิ้น', unit_price: 134 },
  { category: 'เครื่องมือช่าง', item_name: 'คีมปากเฉียง 8"', qty: 1, unit: 'ชิ้น', unit_price: 105 },
  { category: 'เครื่องมือช่าง', item_name: 'ดอกไขควง PH2*50mm', qty: 10, unit: 'ชุด', unit_price: 21 }
];
const hwFields = { miscExpenseVendor: 'HARDWAREHOUSE (ฮาร์ดแวร์เฮาส์)', miscExpenseShipping: '', miscExpenseVatEnabled: { checked: true }, miscExpenseVatPercent: '7' };

const api = build(hwFields, false, billRows);
['HARDWAREHOUSE', 'Hardware House', 'ฮาร์ดแวร์เฮาส์ สาขาบ่อวิน', 'ฮาร์ดแวร์เฮ้าส์'].forEach(function(v) { assert(api.isInc(v), v); });
['ไทยวัสดุ', 'โฮมโปร', ''].forEach(function(v) { assert(!api.isInc(v), v); });

const ex = api.extra(583);
assert.strictEqual(ex.inclusive, true);
assert.strictEqual(ex.subtotal, 544.86, 'รวมยอดขาย ตรงบิล');
assert.strictEqual(ex.vatAmount, 38.14, 'VAT 7% ตรงบิล');
assert.strictEqual(Math.round((ex.subtotal + ex.vatAmount) * 100) / 100, 583, 'ยอดเงินรวมทั้งสิ้น = ยอดจ่ายจริง');

// รายการที่ส่งขึ้นเซิร์ฟเวอร์ = ราคาตามบิลเป๊ะ ไม่มีแถว VAT (VAT อยู่ในราคาแล้ว ห้ามนับซ้ำ) → รวม 583
const items = api.collect();
assert.deepStrictEqual(items.map(function(it) { return it.total_amount; }), [268, 105, 210]);
assert(!items.some(function(it) { return it.category === 'ภาษีมูลค่าเพิ่ม (VAT)'; }), 'ห้ามเพิ่มแถว VAT ซ้ำ');

// ใบเบิกจ่าย: รายการคงราคาตามบิล ท้ายใบถอด VAT แบบท้ายบิล
const voucherSrc = ['mxIsVatInclusiveVendor', 'mxBahtText', 'mxVoucherData'].map(function(n) { return grabFn(n); }).join('\n');
const mxVoucherData = new Function('MX_VOUCHER_LINE_DEPT', voucherSrc + '\nreturn mxVoucherData;')({});
const saved = items.map(function(it) { return Object.assign({ vendor: 'ฮาร์ดแวร์เฮาส์', receipt_no: 'CS4691029-0433', line: 'H9' }, it); });
const d = mxVoucherData(saved);
assert.deepStrictEqual(d.items.map(function(it) { return [it.unitPrice, it.amount]; }), [[134, 268], [105, 105], [21, 210]], 'ราคาต่อหน่วย/จำนวนเงินตามบิล');
assert.strictEqual(d.vatInclusive, true);
assert.strictEqual(d.subtotal, 544.86, 'รวมไม่รวมภาษี');
assert.strictEqual(d.vat, 38.14, 'Vat 7%');
assert.strictEqual(d.total, 583, 'ยอดรวม');
assert.strictEqual(d.totalText, 'ห้าร้อยแปดสิบสามบาทถ้วน');
// ร้านอื่นที่ไม่มีแถว VAT = ไม่มี VAT เหมือนเดิม
const plain = mxVoucherData(saved.map(function(r) { return Object.assign({}, r, { vendor: 'ร้านไทยวัสดุ' }); }));
assert.strictEqual(plain.vat, 0);
assert.strictEqual(plain.total, 583);
// แก้ตัวเลขในหน้าต่างพิมพ์ก็ต้องถอด VAT แบบเดียวกัน
assert(grabFn('mxVoucherHtml').includes('data-vat-inclusive="1"'));
assert(grabFn('mxVoucherEditorRuntime').includes("sheet.getAttribute('data-vat-inclusive') === '1'"));

// ร้านอื่นยังคิด VAT บวกเพิ่มแบบเดิม
const other = build(Object.assign({}, hwFields, { miscExpenseVendor: 'ร้านไทยวัสดุ' }), false, billRows).extra(583);
assert.strictEqual(other.inclusive, false);
assert.strictEqual(other.vatAmount, 40.81);
// แก้ไขทีละแถวไม่แตะสูตรนี้
assert.strictEqual(build(hwFields, true, billRows).extra(583).inclusive, false);

// UI: บอกผู้ใช้ + ติ๊ก VAT ให้เองเมื่อพิมพ์ชื่อร้าน
assert(html.includes('id="miscExpenseVatHint"') && html.includes('id="miscExpenseSubtotalLabel"'));
assert(grabFn('mxRecalcTotals').includes("'รวมยอดขาย (ก่อน VAT)'"));

console.log('misc-expense-vat-inclusive-vendor: OK');
