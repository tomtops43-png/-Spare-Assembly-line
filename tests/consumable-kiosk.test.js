const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(src, re, label) {
  const m = src.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── Backend: เบิกทั้งชุด ของไม่พอแม้ตัวเดียวต้องไม่ตัดอะไรเลย ─────────────────────
const plan = new Function(grab(backend, /^function planConsumableWithdraw\(itemsById, requested\) \{[\s\S]*?\n\}/m, 'planConsumableWithdraw') + '\nreturn planConsumableWithdraw;')();
const stock = {
  CS0001: { item_id: 'CS0001', name: 'ฟิตติ้ง PC 6-01', qty: 10, unit: 'ตัว', active: true },
  CS0002: { item_id: 'CS0002', name: 'น็อต M8', qty: 3, unit: 'ตัว', active: true },
  CS0009: { item_id: 'CS0009', name: 'ของเลิกใช้', qty: 50, unit: 'ตัว', active: false }
};
const ok = plan(stock, [{ item_id: 'CS0001', qty: 4 }, { item_id: 'CS0002', qty: 1 }, { item_id: 'CS0001', qty: 2 }]);
assert.strictEqual(ok.lines.length, 2, 'รายการซ้ำในตะกร้าต้องรวมเป็นแถวเดียว');
assert.strictEqual(ok.lines[0].qty, 6);
assert.strictEqual(ok.lines[0].balance, 4, 'ยอดคงเหลือหลังเบิก = 10 - 6');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 1 }, { item_id: 'CS0002', qty: 5 }]), /ของไม่พอเบิก: น็อต M8/, 'ของไม่พอต้องโยน error ทั้งชุด');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 6 }, { item_id: 'CS0001', qty: 6 }]), /ของไม่พอเบิก/, 'รวมยอดซ้ำแล้วเกินต้องถูกจับได้');
assert.throws(() => plan(stock, [{ item_id: 'CS0009', qty: 1 }]), /ไม่พบรายการ/, 'ของที่ปิดใช้งานต้องเบิกไม่ได้');
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: 0 }]), /จำนวนต้องมากกว่า 0/);
assert.throws(() => plan(stock, [{ item_id: 'CS0001', qty: -2 }]), /จำนวนต้องมากกว่า 0/);

// ── เลขใบเบิกวิ่งใหม่ทุกวัน และไม่สับสนกับวันอื่น ─────────────────────────────────
const nextTicket = new Function(grab(backend, /^function nextConsumableTicketNo\(prefix, existingTickets\) \{[\s\S]*?\n\}/m, 'nextConsumableTicketNo') + '\nreturn nextConsumableTicketNo;')();
assert.strictEqual(nextTicket('CS-261029', []), 'CS-261029-001');
assert.strictEqual(nextTicket('CS-261029', ['CS-261028-009', 'CS-261029-002', '', 'CS-261029-011', 'CS-261029-003']), 'CS-261029-012');

// ── Backend: สิทธิ์ + เขียนใต้ Lock + routing ────────────────────────────────────
assert(/function withdrawConsumables\(payload\) \{\n  var user = requireWarehouseWriter\(/.test(backend), 'เบิกต้องผ่านด่าน requireWarehouseWriter (บัญชีดูอย่างเดียวเบิกไม่ได้)');
['restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(fn) {
  assert(new RegExp('function ' + fn + '\\(payload\\) \\{\\n  var user = requireConsumableKeeper\\(').test(backend), fn + ' ต้องเป็น Admin/Leader เท่านั้น');
});
['withdrawConsumables', 'restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(fn) {
  const body = grab(backend, new RegExp('^function ' + fn + '\\(payload\\) \\{[\\s\\S]*?\\n\\}', 'm'), fn);
  assert(body.includes('LockService.getScriptLock()') && body.includes('lock.releaseLock()'), fn + ' ต้องเขียนใต้ Lock กันเบิกชนกันแล้วยอดเพี้ยน');
});
assert(/function adjustConsumableStock[\s\S]{0,600}if \(!reason\) throw/.test(backend), 'นับปรับยอดต้องบังคับเหตุผล');
['getConsumables', 'withdrawConsumables', 'restockConsumables', 'adjustConsumableStock', 'upsertConsumable', 'seedConsumables'].forEach(function(action) {
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(e.parameter), e);"), action + ' ต้องเรียกผ่าน GET ได้');
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(body), e);"), action + ' ต้องเรียกผ่าน POST ได้');
});
assert(backend.includes("SPARE_APP_CONFIG.consumableSheetName = SPARE_APP_CONFIG.consumableSheetName || 'Consumables'"));
assert(backend.includes("SPARE_APP_CONFIG.consumableLogSheetName = SPARE_APP_CONFIG.consumableLogSheetName || 'ConsumableLog'"));
const kit = new Function(grab(backend, /^var CONSUMABLE_CATEGORIES = [^\n]*$/m, 'categories') + '\n' + grab(backend, /^var CONSUMABLE_STARTER_KIT = \[[\s\S]*?\n\];/m, 'starter kit') + '\nreturn { cats: CONSUMABLE_CATEGORIES, kit: CONSUMABLE_STARTER_KIT };')();
kit.kit.forEach(function(row) { assert(kit.cats.indexOf(row[1]) > -1, 'ชุดเริ่มต้นต้องใช้หมวดที่มีจริง: ' + row[1]); });

// ── หน้าเว็บ: ตัวช่วยคำนวณ ─────────────────────────────────────────────────────
const fe = new Function([
  grab(htmlLf, /^ {4}function csStockLevel\(item\) \{[\s\S]*?\n {4}\}/m, 'csStockLevel'),
  grab(htmlLf, /^ {4}function csForecastDaysLeft\(qty, issuedInWindow, windowDays\) \{[\s\S]*?\n {4}\}/m, 'csForecastDaysLeft'),
  grab(htmlLf, /^ {4}function csReorderQty\(item\) \{[\s\S]*?\n {4}\}/m, 'csReorderQty'),
  grab(htmlLf, /^ {4}function csGroupHistory\(logs\) \{[\s\S]*?\n {4}\}/m, 'csGroupHistory'),
  grab(htmlLf, /^ {4}function csMyFavorites\(logs, me, limit\) \{[\s\S]*?\n {4}\}/m, 'csMyFavorites')
].join('\n') + '\nreturn { csStockLevel, csForecastDaysLeft, csReorderQty, csGroupHistory, csMyFavorites };')();

assert.strictEqual(fe.csStockLevel({ qty: 0, min: 5, max: 20 }).level, 'out');
assert.strictEqual(fe.csStockLevel({ qty: 5, min: 5, max: 20 }).level, 'low', 'ถึง Min พอดีถือว่าใกล้หมด');
assert.strictEqual(fe.csStockLevel({ qty: 6, min: 5, max: 20 }).level, 'ok');
assert.strictEqual(fe.csStockLevel({ qty: 10, min: 5, max: 20 }).pct, 50);
assert.strictEqual(fe.csStockLevel({ qty: 3, min: 0, max: 0 }).level, 'ok', 'ไม่ตั้ง Min = ไม่เตือน');

assert.strictEqual(fe.csForecastDaysLeft(30, 30, 30), 30, 'ใช้วันละ 1 เหลือ 30 = อีก 30 วัน');
assert.strictEqual(fe.csForecastDaysLeft(10, 60, 30), 5);
assert.strictEqual(fe.csForecastDaysLeft(10, 0, 30), null, 'ไม่มีประวัติเบิกต้องไม่เดา');
assert.strictEqual(fe.csForecastDaysLeft(0, 10, 30), 0);

assert.strictEqual(fe.csReorderQty({ qty: 12, min: 20, max: 100 }), 88, 'สั่งให้เติมถึง Max');
assert.strictEqual(fe.csReorderQty({ qty: 3, min: 5, max: 0 }), 7, 'ไม่มี Max ใช้ Min×2');

const groups = fe.csGroupHistory([
  { txn_id: 'a', ticket_no: 'CS-1', type: 'ISSUE' }, { txn_id: 'b', ticket_no: 'CS-1', type: 'ISSUE' },
  { txn_id: 'c', ticket_no: '', type: 'RESTOCK' }, { txn_id: 'd', ticket_no: 'CS-2', type: 'ISSUE' }
]);
assert.deepStrictEqual(groups.map(g => g.key + ':' + g.logs.length), ['CS-1:2', 'c:1', 'CS-2:1'], 'ใบเบิกเดียวกันต้องรวมเป็นใบเดียว ส่วนเติม/ปรับยอดแสดงทีละรายการ');

assert.deepStrictEqual(fe.csMyFavorites([
  { type: 'ISSUE', item_id: 'A', requester: 'me' }, { type: 'ISSUE', item_id: 'B', recorded_by: 'me' }, { type: 'ISSUE', item_id: 'B', requester: 'me' },
  { type: 'ISSUE', item_id: 'C', requester: 'other', recorded_by: 'other' }, { type: 'RESTOCK', item_id: 'D', recorded_by: 'me' }
], 'me', 8), ['B', 'A'], 'เบิกบ่อยของฉันต้องนับเฉพาะของตัวเองและเฉพาะการเบิก');

// ── หน้าเว็บ: เมนู Sidebar + หน้า + สลับหน้า ──────────────────────────────────────
const sideAt = htmlLf.indexOf('id="csSideBtn"');
assert(sideAt > htmlLf.indexOf('id="lineTabs"') && sideAt < htmlLf.indexOf('id="lineMenuFooterAvatar"'), 'ปุ่มเบิกของสิ้นเปลืองต้องอยู่ใน Sidebar ใต้รายชื่อไลน์ เหนือข้อมูลผู้ใช้');
assert(htmlLf.includes('id="consumablePage"') && htmlLf.includes('id="csCheckoutModal"') && htmlLf.includes('id="csSuccessModal"'));
assert(htmlLf.includes("'consumables': { btn: 'csSideBtn'"), 'ต้องเปิดเป็นแท็บแบบ browser ได้');
assert((htmlLf.match(/^ {6}csHidePage\(\);$/gm) || []).length >= 13, 'ทุก switch page ต้องซ่อนหน้าเบิกของสิ้นเปลือง');
assert(/function switchToMiscExpensePage\(\)[\s\S]{0,800}csHidePage\(\);/.test(htmlLf), 'หน้าค่าใช้จ่ายสิ้นเปลืองต้องซ่อนหน้านี้ด้วย');
assert(/function switchToConsumablePage\(\)[\s\S]{0,700}mxHidePage\(\);[\s\S]{0,400}setTabStyles\('consumables'\)/.test(htmlLf));
assert(/function renderAuthUserBox\(\)[\s\S]{0,400}if \(!currentUser\) \{[\s\S]{0,120}csSyncSideEntry\(\);/.test(htmlLf), 'ออกจากระบบต้องล้างตะกร้าและซ่อนปุ่ม');
assert(/function csCanWithdraw\(\) \{ return canWriteWarehouse\(\); \}/.test(htmlLf), 'ฝั่งเว็บต้องสะท้อนด่าน requireWarehouseWriter');
assert(htmlLf.includes(".cs-cartbar.hidden") || htmlLf.includes('#consumablePage .hidden'), 'คลาส cs-* ที่ตั้ง display เองต้องยังซ่อนด้วย .hidden ได้');

// ── นำเข้า CSV: ชุดฟิตติ้งลมต้องแปลงครบ 24 รายการ เดาหมวด/ไอคอนถูก และทยอยเพิ่มซ้ำได้ ─────
const imp = new Function([
  grab(htmlLf, /^ {4}var CS_CATEGORY_ICONS = \{[\s\S]*?\n {4}\};/m, 'CS_CATEGORY_ICONS'),
  'var CS_CATEGORIES = Object.keys(CS_CATEGORY_ICONS);',
  grab(htmlLf, /^ {4}var CS_PRESET_PNEUMATIC_FITTINGS = \[[\s\S]*?\]\.join\('\\n'\);/m, 'preset'),
  grab(htmlLf, /^ {4}var CS_IMPORT_GUESS = \[[\s\S]*?\n {4}\];/m, 'CS_IMPORT_GUESS'),
  grab(htmlLf, /^ {4}function csSplitCsvLine\(line, sep\) \{[\s\S]*?\n {4}\}/m, 'csSplitCsvLine'),
  grab(htmlLf, /^ {4}function csParseImportCsv\(text\) \{[\s\S]*?\n {4}\}/m, 'csParseImportCsv')
].join('\n') + '\nreturn { preset: CS_PRESET_PNEUMATIC_FITTINGS, parse: csParseImportCsv };')();
const presetRows = imp.parse(imp.preset);
assert.strictEqual(presetRows.errors.length, 0);
assert.strictEqual(presetRows.rows.length, 24, 'ชุดฟิตติ้งลมต้องครบ 24 รายการ');
assert(presetRows.rows.every(r => r.category === 'ฟิตติ้งลม/สายลม' && r.unit === 'ตัว'), 'ฟิตติ้งทุกตัวต้องเข้าหมวดฟิตติ้งลม');
const byName = {};
presetRows.rows.forEach(r => { byName[r.name] = r; });
assert.strictEqual(byName['ข้อต่อตรง ø6'].min, '20', 'min_stock ต้องไปเป็น Min');
assert.strictEqual(byName['ข้อต่อตรง ø4 x M5'].spec, 'ข้อต่อตรง เกลียวนอก', 'หมวดย่อยในไฟล์เก็บไว้เป็นสเปค');
assert.strictEqual(byName['ข้อลด ø8-ø6'].icon, '🔻');
assert.strictEqual(byName['ข้องอ 90 ø6 x 1/4'].icon, '↪️');
assert.strictEqual(byName['สามทาง ø8'].icon, '🔱');
assert.strictEqual(byName['ตัวอุด ø6'].icon, '⛔');
// ก๊อปจาก Excel = แยกด้วย tab, มีเครื่องหมายคำพูด/คอมมาในช่อง, หมวดของระบบตรงๆ
const tsv = imp.parse('name\tcategory\tspec\tmin\tqty\nน็อต M10\tน็อต/สกรู/แหวน\t"M10, ชุบซิงค์"\t50\t120\n\tอื่นๆ\t\t\t');
assert.strictEqual(tsv.rows.length, 1);
assert.deepStrictEqual([tsv.rows[0].category, tsv.rows[0].spec, tsv.rows[0].qty], ['น็อต/สกรู/แหวน', 'M10, ชุบซิงค์', '120']);
assert.strictEqual(tsv.errors.length, 1, 'แถวไม่มีชื่อต้องถูกรายงาน');
assert(/ไม่พบคอลัมน์ name/.test(imp.parse('foo,bar\n1,2').errors[0]));

const importFn = grab(backend, /^function importConsumables\(payload\) \{[\s\S]*?\n\}/m, 'importConsumables');
assert(importFn.includes('requireConsumableKeeper') && importFn.includes('LockService.getScriptLock()'), 'นำเข้าต้อง Admin/Leader และเขียนใต้ Lock');
assert(importFn.includes('if (seen[key]) { skipped.push('), 'ชื่อ+สเปคซ้ำของเดิมต้องข้าม (ทยอยเพิ่มไฟล์เดิมซ้ำได้)');
assert(backend.includes("if (action === 'importConsumables') return respond(importConsumables(body), e);"));
assert(backend.includes("if (action === 'importConsumables') return respond(importConsumables(e.parameter), e);"));

// ── แก้ไข/เติม/นับ จากการ์ดหน้าเบิก + รูปสินค้า ─────────────────────────────────
const shopSrc = grab(htmlLf, /^ {4}function renderCsShop\(\) \{[\s\S]*?\n {4}\}/m, 'renderCsShop');
assert(shopSrc.includes('var canKeep = csCanKeep();') && /var admin = !canKeep \? '' :/.test(shopSrc), 'ปุ่มจัดการบนการ์ดต้องโผล่เฉพาะ Admin/Leader');
['data-cs-edit', 'data-cs-restock', 'data-cs-adjust'].forEach(function(attr) {
  assert(shopSrc.includes(attr + '="'), 'การ์ดหน้าเบิกต้องมีปุ่ม ' + attr);
});
assert(shopSrc.includes('action + admin +'), 'ปุ่มจัดการต้องอยู่บนการ์ดแม้ของหมด (เติมได้ทันที)');
assert(shopSrc.includes('csVisual(it, 300)'), 'การ์ดต้องโชว์รูปแทนไอคอนเมื่อมีรูป');
assert(htmlLf.includes('id="csItemImageFile"') && htmlLf.includes("action: 'uploadConsumableImage'"), 'ฟอร์มแก้ไขต้องแนบรูปได้');
assert(htmlLf.includes("image_url: csEl('csItemImage').value"), 'บันทึกรายการต้องส่งลิงก์รูปไปด้วย');
assert(/compressImageToDataUrl\(file, \{ maxDim: 900/.test(htmlLf), 'ต้องย่อรูปก่อนอัปโหลด');

assert(backend.includes("'Created By', 'Created At', 'Image URL'];"), 'คอลัมน์รูปต้องต่อท้าย (ห้ามแทรกกลาง เพราะอ่านตามตำแหน่ง)');
assert(backend.includes("image_url: String(row[16] || '')"));
assert(/function uploadConsumableImage\(payload\) \{\n  var user = requireConsumableKeeper\(/.test(backend), 'อัปโหลดรูปต้อง Admin/Leader');
assert(backend.includes("if (action === 'uploadConsumableImage') return respond(uploadConsumableImage(body), e);"));
// ทุกที่ที่สร้างแถวใหม่ต้องเขียนครบทุกคอลัมน์ (setValues ช่วงกว้าง CONSUMABLE_HEADERS.length)
const headerCount = new Function(grab(backend, /^var CONSUMABLE_HEADERS = [^\n]*$/m, 'headers') + '\nreturn CONSUMABLE_HEADERS.length;')();
assert.strictEqual(headerCount, 17);
assert(backend.includes("timestamp, user.username || '', timestamp, ''];"), 'ชุดเริ่มต้นต้องมีช่องรูป (ว่าง)');
assert(backend.includes('timestamp, user.username || \'\', timestamp, f.imageUrl];'), 'นำเข้าต้องมีช่องรูป');
assert(backend.includes('existing.created_by, existing.created_at, f.imageUrl'), 'แก้ไขต้องเขียนช่องรูป');

// ── ฟิตติ้งลม: แยกชนิดย่อยจากชื่อ+สเปคอัตโนมัติ ──────────────────────────────────
const fit = new Function([
  grab(htmlLf, /^ {4}var CS_FITTING_CATEGORY = [^\n]*$/m, 'CS_FITTING_CATEGORY'),
  grab(htmlLf, /^ {4}var CS_FITTING_TYPES = \[[\s\S]*?\n {4}\];/m, 'CS_FITTING_TYPES'),
  grab(htmlLf, /^ {4}function csFittingType\(item\) \{[\s\S]*?\n {4}\}/m, 'csFittingType'),
  grab(htmlLf, /^ {4}function csFittingConn\(item\) \{[\s\S]*?\n {4}\}/m, 'csFittingConn')
].join('\n') + '\nreturn { type: csFittingType, conn: csFittingConn };')();
const F = (name, spec) => ({ name: name, spec: spec || '', category: 'ฟิตติ้งลม/สายลม' });
[
  [F('ข้อต่อตรง ø6', 'ข้อต่อตรง สาย-สาย'), 'straight', 'tube'],
  [F('ข้อต่อตรง ø4 x M5', 'ข้อต่อตรง เกลียวนอก'), 'straight', 'thread'],
  [F('ข้อลด ø8-ø6', 'ข้อลด สาย-สาย'), 'reducer', 'tube'],
  [F('ข้องอ 90 ø6 x 1/4', 'ข้องอ เกลียวนอก'), 'elbow', 'thread'],
  [F('สามทาง ø6', 'สามทาง สาย-สาย'), 'tee-t', 'tube'],
  [F('สามทาง T ø8'), 'tee-t', 'tube'],
  [F('สามทาง Y ø6'), 'tee-y', 'tube'],
  [F('ข้อต่อ Y-Union', 'PY 8'), 'tee-y', 'tube'],
  [F('ตัวอุด ø6', 'ตัวอุด'), 'plug', 'tube'],
  [F('ฟิตติ้งตรง เสียบสาย', 'PC 6-01'), 'straight', 'tube'],
  [F('ฟิตติ้งงอ 90°', 'PL 6-01'), 'elbow', 'tube'],
  [F('ข้อต่อสามทาง', 'PE 6'), 'tee-t', 'tube'],
  [F('สายลม PU', 'Ø6 mm'), 'tube', ''],
  [F('อะไรสักอย่าง'), 'other', 'tube']
].forEach(function(c) {
  assert.strictEqual(fit.type(c[0]), c[1], 'ชนิดของ "' + c[0].name + ' ' + c[0].spec + '" ต้องเป็น ' + c[1]);
  assert.strictEqual(fit.conn(c[0]), c[2], 'แบบการต่อของ "' + c[0].name + ' ' + c[0].spec + '" ต้องเป็น ' + (c[2] || '(ไม่มี)'));
});
assert.strictEqual(fit.type({ name: 'สามทาง', category: 'อื่นๆ' }), '', 'หมวดอื่นต้องไม่ถูกแยกชนิด');
assert(htmlLf.includes('id="csSubChips"'), 'ต้องมีแถบชิปชนิดย่อย');
assert(/if \(\(a = t\.getAttribute\('data-cs-cat'\)\)\) \{ csCategory = a; csSubType = 'all'; csConn = 'all';/.test(htmlLf), 'เปลี่ยนหมวดต้องรีเซ็ตตัวกรองย่อย');
