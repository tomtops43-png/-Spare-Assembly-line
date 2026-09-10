// โหมดลูกค้า — ไฟล์ที่ส่งให้ลูกค้าต้องไม่มีราคาต้นทุน ผู้ขาย หรือค่าใช้จ่าย
// เจตนา: ต้อง "ตัดคอลัมน์ออกจากไฟล์จริง" ไม่ใช่ซ่อนด้วย CSS หรือปล่อยให้ผู้ใช้จำไม่ติ๊ก
// เพราะไฟล์ที่ออกไปแล้วเรียกคืนไม่ได้ ถ้าตัวเลขต้นทุนหลุดไปมือลูกค้าก็จบ
const { html, grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false };',
    grabVar('XP_DATASETS'),
    grabFn('xpDatasetByKey'),
    'function isAdminUser() { return __isAdmin; }',
    'var __isAdmin = true;',
    grabFn('xpVisibleDatasets')
  ]),
  '{ xpTable: xpTable, xpTableFromDump: xpTableFromDump, xpIsSensitiveHeader: xpIsSensitiveHeader, xpVisibleDatasets: xpVisibleDatasets, setCustomer: function(v) { xpFilters.customerMode = v; }, setAdmin: function(v) { __isAdmin = v; } }'
);

// ── หัวคอลัมน์ที่ต้องถือว่าเป็นข้อมูลราคา ─────────────────────────────
[
  'ราคาต่อหน่วย', 'มูลค่าคงคลัง', 'ยอดเงิน', 'ผู้ขาย', 'ต้นทุน', 'งบประมาณ',
  'Unit Price', 'Total Amount', 'Vendor', 'Supplier', 'unit_price', 'total_amount'
].forEach(function(h) {
  assert.strictEqual(mod.xpIsSensitiveHeader(h), true, 'ต้องถือว่า "' + h + '" เป็นข้อมูลราคา');
});
// ต้องไม่ไปตัดคอลัมน์ที่ไม่เกี่ยวข้องทิ้ง (จำนวนชิ้น/ชื่อ/ไลน์ ยังต้องอยู่ในไฟล์ลูกค้า)
['ชื่อรายการ', 'คงเหลือ', 'Min', 'Max', 'ไลน์', 'เครื่องจักร', 'Qty Ordered', 'สถานะ'].forEach(function(h) {
  assert.strictEqual(mod.xpIsSensitiveHeader(h), false, '"' + h + '" ไม่ควรถูกตัดในโหมดลูกค้า');
});

// ── xpTable ต้องตัดคอลัมน์ที่ทำเครื่องหมาย sensitive ────────────────
const cols = [
  { label: 'ชื่อรายการ', key: 'name' },
  { label: 'คงเหลือ', key: 'stock', type: 'num' },
  { label: 'ราคาต่อหน่วย', key: 'price', type: 'money', sensitive: true },
  { label: 'ผู้ขาย', key: 'supplier', sensitive: true }
];
const rows = [{ name: 'น็อต M8', stock: 12, price: 4.5, supplier: 'ร้านไทยวัสดุ' }];

mod.setCustomer(false);
const full = mod.xpTable(cols, rows);
assert.deepStrictEqual(full.headers, ['ชื่อรายการ', 'คงเหลือ', 'ราคาต่อหน่วย', 'ผู้ขาย']);
assert.deepStrictEqual(full.rows[0], ['น็อต M8', 12, 4.5, 'ร้านไทยวัสดุ']);
assert.deepStrictEqual(full.numeric, [false, true, true, false]);
assert.deepStrictEqual(full.money, [false, false, true, false], 'ช่องเงินต้องถูกทำเครื่องหมายไว้ให้ Excel ใส่ทศนิยม 2 ตำแหน่ง');

mod.setCustomer(true);
const cust = mod.xpTable(cols, rows);
assert.deepStrictEqual(cust.headers, ['ชื่อรายการ', 'คงเหลือ'], 'โหมดลูกค้าต้องเหลือแค่คอลัมน์ที่ไม่ใช่ราคา');
assert.deepStrictEqual(cust.rows[0], ['น็อต M8', 12], 'ค่าราคาต้องไม่อยู่ในแถวเลย ไม่ใช่แค่ซ่อนหัว');
assert.strictEqual(cust.rows[0].length, 2, 'จำนวนช่องในแถวต้องเท่ากับจำนวนหัวคอลัมน์');

// ── ตารางจากชีตดิบ (ไม่มีสเปกคอลัมน์) ต้องตัดตามชื่อหัวคอลัมน์ ───────
const dump = {
  headers: ['History ID', 'Part Name', 'Qty Ordered', 'Unit Price', 'Total Amount', 'Status'],
  rows: [['PH-1', 'เบรกเกอร์', 5, 250, 1250, 'Received']]
};
mod.setCustomer(false);
assert.strictEqual(mod.xpTableFromDump(dump).headers.length, 6);
mod.setCustomer(true);
const dumpCust = mod.xpTableFromDump(dump);
assert.deepStrictEqual(dumpCust.headers, ['History ID', 'Part Name', 'Qty Ordered', 'Status'], 'ชีตดิบต้องถูกตัดคอลัมน์ราคาออกด้วย');
assert.deepStrictEqual(dumpCust.rows[0], ['PH-1', 'เบรกเกอร์', 5, 'Received']);

// เดาชนิดคอลัมน์เป็นตัวเลขได้ถูก (Excel จะจัดขวา + ใส่ comma ให้)
assert.strictEqual(dumpCust.numeric[2], true, 'Qty Ordered เป็นตัวเลข');
assert.strictEqual(dumpCust.numeric[0], false, 'History ID เป็นข้อความ ไม่ใช่ตัวเลข');

// ── ชุดข้อมูลที่ทั้งชุดเป็นเรื่องเงิน ต้องหายไปทั้งชุดในโหมดลูกค้า ──
mod.setCustomer(false);
mod.setAdmin(true);
const normalKeys = mod.xpVisibleDatasets().map(function(d) { return d.key; });
assert(normalKeys.indexOf('miscExpenses') > -1, 'โหมดปกติต้องเห็นชุดค่าใช้จ่ายสิ้นเปลือง');
assert(normalKeys.indexOf('productionCostConfig') > -1);

mod.setCustomer(true);
const custKeys = mod.xpVisibleDatasets().map(function(d) { return d.key; });
assert(custKeys.indexOf('miscExpenses') === -1, 'ค่าใช้จ่ายสิ้นเปลืองต้องไม่โผล่ในโหมดลูกค้า — ซ่อนคอลัมน์ไม่พอเพราะทั้งชุดคือเรื่องเงิน');
assert(custKeys.indexOf('productionCostConfig') === -1, 'ราคาต่อชิ้น + เป้า % ต้องไม่โผล่ในโหมดลูกค้า');
assert(custKeys.indexOf('stockMaster') > -1, 'Master อะไหล่ยังต้องเห็นได้ (ตัดแค่คอลัมน์ราคา)');

// ── ชุด adminOnly ต้องหายไปเมื่อไม่ใช่ Admin ─────────────────────────
mod.setCustomer(false);
mod.setAdmin(false);
const nonAdminKeys = mod.xpVisibleDatasets().map(function(d) { return d.key; });
assert(nonAdminKeys.indexOf('userRoster') === -1, 'คนที่ไม่ใช่ Admin ต้องไม่เห็นชุดข้อมูลผู้ใช้');
assert(nonAdminKeys.indexOf('exportLog') === -1, 'คนที่ไม่ใช่ Admin ต้องไม่เห็นทะเบียนการส่งออก');
assert(nonAdminKeys.indexOf('logs') > -1, 'ชุดข้อมูลทั่วไปยังต้องเห็นได้');

// ── รายงานที่เป็นเรื่องราคาล้วน ต้องถูกซ่อนในโหมดลูกค้า ──────────────
const reportsSrc = grabVar('XP_REPORTS');
assert(/hiddenInCustomerMode: true/.test(reportsSrc), 'ต้องมีรายงานที่ถูกซ่อนในโหมดลูกค้า (Price Variance)');
const gridSrc = grabFn('xpRenderReportGrid');
assert(/hiddenInCustomerMode/.test(gridSrc), 'ตารางรายงานต้องเคารพ hiddenInCustomerMode');

// ── หน้าปกไฟล์ต้องระบุชัดว่าไฟล์นี้เป็นโหมดลูกค้า ───────────────────
const coverSrc = grabFn('xpBuildCover');
assert(/โหมดลูกค้า/.test(coverSrc), 'หน้าปกต้องบอกว่าไฟล์นี้ตัดข้อมูลราคาออกไปแล้ว ไม่งั้นคนรับไฟล์เข้าใจผิดว่าระบบไม่มีข้อมูลราคา');

// ── ไฟล์แนบ: ใบเสนอราคา/ใบเสร็จ ต้องไม่ติดไปในไฟล์ลูกค้า ────────────
const attachSrc = grabFn('xpBuildAttachmentSheet');
assert(/customerMode && pair\[0\] === 'ใบเสนอราคา'/.test(attachSrc), 'ลิงก์ใบเสนอราคาต้องไม่ติดไปในโหมดลูกค้า');
assert(/if \(!xpFilters\.customerMode\) \{[\s\S]*receipt_url/.test(attachSrc), 'ลิงก์บิล/ใบเสร็จต้องไม่ติดไปในโหมดลูกค้า');

console.log('export-customer-mode: OK');
