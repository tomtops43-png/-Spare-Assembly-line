// เติมราคาที่หายใน Purchase History ด้วยราคาปัจจุบันจากทะเบียนอะไหล่
const fs = require('fs');
const assert = require('assert');
const vm = require('vm');
const html = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

// indent = ระยะย่อหน้าของ function (index.html ย่อ 4, Backend.gs ไม่ย่อ) ใช้หาปีกกาปิดตัวจริง
function grab(src, name, indent) {
  const pad = indent === undefined ? '    ' : indent;
  const re = new RegExp('^' + pad + 'function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n' + pad + '\\}', 'm');
  const m = src.match(re);
  assert(m, 'ต้องดึง ' + name + ' ออกมาได้');
  return m[0];
}

// ── ปุ่มและการผูกอีเวนต์ ────────────────────────────────────────────────────
assert(html.includes('id="purchaseHistoryFillPriceBtn"'), 'ต้องมีปุ่มเติมราคาที่หาย');
assert(!/id="purchaseHistoryFillPriceBtn"[^>]*pt-legacy-hide/.test(html), 'ปุ่มต้องไม่ถูกซ่อนตอนเปิดระบบ PR ใหม่ — ราคาย้อนหลังยังต้องเติมได้');
assert(/purchaseHistoryFillPriceBtn\.addEventListener\('click', function\(\) \{ fillPurchaseHistoryMissingPrices\(\); \}\)/.test(html), 'ต้องผูกปุ่มกับฟังก์ชันเติมราคา');

// ── ตรรกะจับคู่ (รันจริง) ──────────────────────────────────────────────────
const ctx = {};
vm.createContext(ctx);
['normPhText', 'normPhModel', 'isMeaningfulPhModel', 'isMeaningfulPhBrand', 'purchaseHistoryMatchesItem',
  'isPurchaseHistoryPriceMissing', 'phMasterUnitPrice', 'buildPurchaseHistoryPriceFills'].forEach(function(name) {
  vm.runInContext(grab(html, name), ctx);
});

const parts = [
  { line: 'Coil Winding', no: '5', name: 'Bearing', model: '6204-ZZ', brand: 'NSK', unit_price: '120', currency: 'THB' },
  { line: 'Coil Winding', no: '6', name: 'Spindle', model: 'MF-DRWG-CWM-PF220-SPINDLE-20A', brand: 'YHM', unit_price: '1,500.50' },
  { line: 'H9', no: '', name: 'Bearing', model: '6204-ZZ', brand: 'NSK', unit_price: '999' },
  { line: 'Coil Winding', no: '7', name: 'Belt', model: '-', brand: '-', unit_price: '' },
  { line: 'Lug&Screw', no: '1', name: 'Cutter', model: 'CT-1', unit_price: '50' },
  { line: 'Lug&Screw', no: '2', name: 'Cutter', model: 'CT-1', unit_price: '70' }
];
const history = [
  { history_id: 'A', line: 'Coil Winding', part_id: '', part_name: 'Bearing', model: '6204 ZZ', unit_price: '' },
  { history_id: 'B', line: 'Coil Winding', part_id: '', part_name: 'Retractable Spindle', model: 'MF-DRWG-CWM-PF220- SPINDLE-20A', unit_price: '' },
  { history_id: 'C', line: 'Coil Winding', part_id: '', part_name: 'Bearing', model: '6204-ZZ', unit_price: '80' },
  { history_id: 'D', line: 'Coil Winding', part_id: '', part_name: '-', model: '-', unit_price: '' },
  { history_id: 'E', line: 'Coil Winding', part_id: '', part_name: 'Belt', model: '-', unit_price: '' },
  { history_id: 'F', line: 'Lug&Screw', part_id: '', part_name: 'Cutter', model: 'CT-1', unit_price: '' }
];
const plan = ctx.buildPurchaseHistoryPriceFills(history, parts);
const byId = {};
plan.fills.forEach(function(f) { byId[f.history_id] = f; });
assert.strictEqual(byId.A && byId.A.unit_price, 120, 'ต้องได้ราคาของ Line เดียวกัน ไม่ใช่ของ H9');
assert.strictEqual(byId.A.currency, 'THB');
assert.strictEqual(byId.B && byId.B.unit_price, 1500.5, 'รุ่นที่เว้นวรรคเพี้ยนต้องยังจับคู่ได้ และอ่านราคามีคอมมาได้');
assert(!byId.C, 'รายการที่มีราคาแล้วห้ามทับ');
assert(!byId.D, 'ไม่มีชื่อ/รุ่นให้เทียบ = ไม่เดา');
assert(!byId.E, 'อะไหล่ใน master ที่ยังไม่มีราคา เติมไม่ได้');
assert(!byId.F, 'เจอหลายตัวที่ราคาไม่เท่ากัน = ไม่เดา');
assert.strictEqual(plan.ambiguous, 1);
assert.strictEqual(plan.noMatch, 2);

// ── โหลด master สด และกันเลข NO ซ้ำข้ามชีต ────────────────────────────────
const loader = grab(html, 'loadPurchaseHistoryPriceMaster');
assert(loader.includes('fetchSheetRows(sheetName, { lite: true })'), 'ต้องดึงราคาสดจากชีต ไม่ใช่แคชเก่า');
assert(loader.includes('line: lineKey'), 'ต้องติด line ตามไลน์ของชีต');
assert(loader.includes("no: sheets.length > 1 ? '' : row.no"), 'ไลน์หลายชีตต้องไม่จับคู่ด้วยเลข NO');

const run = grab(html, 'fillPurchaseHistoryMissingPrices');
assert(run.includes("action: 'fillPurchaseHistoryPrices'"), 'ต้องส่งเป็นชุดเดียว');
assert(run.includes('showConfirmDialog('), 'ต้องให้ยืนยันก่อนบันทึก');
assert(run.includes('loadPurchaseHistory(true)'), 'บันทึกแล้วต้องโหลดตารางใหม่');
assert(html.includes("item.price_status === 'Master'"), 'ราคาที่เติมต้องมีป้ายบอกว่าเป็นราคาปัจจุบัน');

// ── Backend ────────────────────────────────────────────────────────────────
assert(backend.includes("if (action === 'fillPurchaseHistoryPrices') return respond(fillPurchaseHistoryPrices(body), e);"), 'ต้องมี route ใหม่');
const fill = grab(backend, 'fillPurchaseHistoryPrices', '');
assert(fill.includes('requirePurchaseHistoryEditor('), 'เฉพาะ Admin/Engineer');
assert(/values\[i\]\[12\][\s\S]*?\.trim\(\)\) continue;/.test(fill), 'แถวที่มีราคาแล้วห้ามทับ');
assert(fill.includes('calculatePurchaseHistoryTotal(values[i][10], fill.price)'), 'ต้องคำนวณ Total Amount ใหม่');
assert(fill.includes("values[i][28] = 'Master'"), 'ต้องบันทึก Price Status ว่ามาจาก master');
assert(fill.includes("'PRICE_FILL'"), 'ต้องเก็บ audit');

console.log('purchase-history-fill-prices OK');
