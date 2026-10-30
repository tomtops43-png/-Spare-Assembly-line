// แก้ราคา/ข้อมูลอะไหล่แล้วกดบันทึก → การ์ดต้องเปลี่ยนทันที ไม่ใช่โชว์ราคาเดิมจนกว่าเซิร์ฟเวอร์จะตอบ
// และต้องแพตช์แคชด้วย ไม่งั้นโหลดซ้ำหลังบันทึกเสร็จจะหยิบแคชเก่ามาโชว์ราคาเดิมกลับมา
const { html, grabFn, grabVar, assert } = require('./_export-extract');

const stored = {};
const localStorage = { setItem: (k, v) => { stored[k] = v; } };
let applied = null;

const api = new Function('localStorage', 'onApply',
  'var partsData = [];\nvar stockSheetMemoryCache = {};\n' +
  'function safeNum(v) { var n = Number(v); return isFinite(n) ? n : 0; }\n' +
  'function normalizeDriveImageUrl(v) { return v; }\n' +
  'function applyLoadedPartsData(d) { onApply(d); }\n' +
  grabVar('LOCAL_EDIT_FIELDS') + ';\n' +
  ['firstDefined', 'normalizeRecord', 'getItemIdentityKey', 'applyLocalItemEdit'].map(n => grabFn(n)).join('\n') +
  '\nreturn { normalizeRecord: normalizeRecord, applyLocalItemEdit: applyLocalItemEdit,' +
  ' set: function(p, c) { partsData = p; stockSheetMemoryCache = c; }, parts: function() { return partsData; } };'
)(localStorage, d => { applied = d; });

const raw = { no: 110, name: 'Air Cylinder', model: 'MGPM16-50Z', brand: 'SMC', line: 'H9', __sourceSheet: 'Main List Stock',
  stock: 0, min: 2, max: 5, unit: 'PCS', unit_price: 3660, currency: 'THB', image_main_url: 'img' };
const other = { no: 111, name: 'Valve', model: 'V1', __sourceSheet: 'Main List Stock', stock: 3, unit_price: 50 };
const item = api.normalizeRecord(raw, 0);
const cacheItems = [api.normalizeRecord(raw, 0), api.normalizeRecord(other, 1)];
api.set([item, api.normalizeRecord(other, 1)], { 'spare_parts_cache::H9::H9::ALL': { version: 1, updatedAt: 1, items: cacheItems } });

const payload = { no: '110', name: 'Air Cylinder', model: 'MGPM16-50Z', line: 'H9', stock: '0', min: '2', max: '5',
  unit: 'PCS', unit_price: '3437.25', currency: 'THB', supplier: '-', image_main_url: 'img' };
assert.strictEqual(api.applyLocalItemEdit(item, payload), 1, 'ต้องเจอรายการที่แก้ 1 ตัว');

const now = api.parts()[0];
assert.strictEqual(Number(now.unit_price), 3437.25, 'ราคาบนการ์ดต้องเป็นราคาใหม่ทันที');
assert.strictEqual(now.__sourceSheet, 'Main List Stock', 'ต้องคงชีตต้นทางเดิม');
assert.strictEqual(now.image_main_url, 'img');
assert.strictEqual(api.parts()[1].unit_price, 50, 'รายการอื่นต้องไม่ถูกแตะ');
assert(applied && applied.length === 2, 'ต้องวาดการ์ดใหม่ทันที');

const cached = JSON.parse(stored['spare_parts_cache::H9::H9::ALL']);
assert.strictEqual(Number(cached.items[0].unit_price), 3437.25, 'แคชต้องเป็นราคาใหม่ — โหลดซ้ำแล้วไม่เด้งกลับราคาเดิม');
assert.strictEqual(cached.items[1].unit_price, 50);

// ไม่มีรายการเดิม (เพิ่มใหม่/หาไม่เจอ) = ไม่ทำอะไร
applied = null;
assert.strictEqual(api.applyLocalItemEdit({}, payload), 0);
assert.strictEqual(applied, null);

// ฟอร์มแก้ไขต้องเรียกแพตช์ก่อนเข้าคิว และบันทึกพลาดต้องดึงข้อมูลจริงกลับมา
const submit = grabFn('submitEditModal');
assert(submit.indexOf('applyLocalItemEdit(oldItem, payload)') < submit.indexOf('enqueueUploadJob('), 'แพตช์ก่อนเข้าคิวอัปโหลด');
assert(submit.includes('loadPartsData({ skipCache: true })'), 'บันทึกไม่สำเร็จต้องโหลดข้อมูลจริงกลับมาแทนค่าที่แสดงไว้');

console.log('Edit item instant update checks passed');
