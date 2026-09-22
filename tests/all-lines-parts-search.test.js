const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── หน้า "อะไหล่ทุกไลน์" ต้องมีอยู่จริงและเข้าถึงได้จากแถบเมนู ────────────────
assert(htmlLf.includes('<button id="tabAllParts"'), 'ต้องมีปุ่มแท็บ ทุกไลน์');
assert(htmlLf.includes('<section id="allPartsPage"'), 'ต้องมี section ของหน้ารวมทุกไลน์');
['allPartsSearch', 'allPartsLineFilter', 'allPartsCategoryFilter', 'allPartsStatusFilter',
  'allPartsLineChips', 'allPartsTable', 'allPartsSummary', 'allPartsMoreBtn', 'allPartsRefreshBtn']
  .forEach(function(id) {
    assert(htmlLf.includes('id="' + id + '"'), 'ต้องมี #' + id);
  });
assert(/tabBtn\.addEventListener\('click', switchToAllPartsPage\)/.test(htmlLf), 'ปุ่มแท็บต้องผูกกับ switchToAllPartsPage');

// ── ทุกหน้าที่สลับไป ต้องซ่อนหน้านี้ด้วย ────────────────────────────────────
// ถ้าลืมจุดใดจุดหนึ่ง หน้ารวมทุกไลน์จะค้างซ้อนอยู่ใต้หน้าอื่น
// นับนอกตัว switchToAllPartsPage เอง — หน้านี้เรียก xpHidePage แต่ไม่ต้องซ่อนตัวเอง
const outside = htmlLf.replace(/^ {4}function switchToAllPartsPage\(\) \{[\s\S]*?\n {4}\}/m, '');
const xpCalls = (outside.match(/^ {6}xpHidePage\(\);$/gm) || []).length;
const apCalls = (outside.match(/^ {6}apHidePage\(\);$/gm) || []).length;
assert(xpCalls > 0 && apCalls === xpCalls, 'apHidePage ต้องถูกเรียกครบทุกจุดที่เรียก xpHidePage (' + apCalls + '/' + xpCalls + ')');
assert(/function apHidePage\(\)/.test(htmlLf), 'ต้องมี apHidePage');
// ตัวหน้าเองต้องซ่อนหน้าอื่นครบ ไม่งั้นกดเข้ามาแล้วหน้าเดิมยังค้างอยู่
const switcher = grab(/^ {4}function switchToAllPartsPage\(\) \{[\s\S]*?\n {4}\}/m, 'switchToAllPartsPage');
['stockPage', 'dashboardPage', 'txnPage', 'requestOrderPage', 'logPage', 'prReportPage',
  'purchaseHistoryPage', 'managePage', 'adminPage'].forEach(function(id) {
  assert(switcher.includes(id + '.classList.add(\'hidden\')'), 'switchToAllPartsPage ต้องซ่อน ' + id);
});
assert(switcher.includes('mxHidePage();') && switcher.includes('xpHidePage();'), 'ต้องซ่อนหน้าค่าใช้จ่าย/Export ด้วย');
assert(switcher.includes("setTabStyles('all-parts')"), 'ต้องไฮไลต์แท็บของตัวเอง');

// แท็บต้องเข้าระบบไฮไลต์/สิทธิ์เหมือนแท็บอื่น
assert(/if \(active === 'all-parts' && tabAllPartsEl\) tabAllPartsEl\.className = tabActiveClass;/.test(htmlLf),
  'setTabStyles ต้องรู้จัก all-parts');
assert(/tabAllPartsEl\.classList\.toggle\('hidden', !hasPermission\('view'\)\)/.test(htmlLf),
  'แท็บต้องซ่อนเมื่อไม่มีสิทธิ์ view');
assert(/allPartsPage'\)\.classList\.contains\('hidden'\)\) setTabStyles\('all-parts'\)/.test(htmlLf),
  'ย่อ/ขยายจอแล้วไฮไลต์แท็บต้องไม่หาย');

// ── ค้นหาต้องหน่วงคีย์ + รูปในลิสต์ต้อง lazy (ยามเดียวกับ ui-performance-guards) ──
assert(htmlLf.includes('var debouncedApRender = debounce('), 'ช่องค้นหาต้องหน่วงคีย์');
assert(/searchEl\.addEventListener\('input', debouncedApRender\)/.test(htmlLf), 'ช่องค้นหาต้องใช้ตัวที่หน่วงแล้ว');

// ── โหลดข้อมูล: ต้องใช้แคชก้อนเดียวกับหน้า Stock ไม่ยิงซ้ำโดยไม่จำเป็น ──────────
const loader = grab(/^ {4}function apLoadAllLines\(force\) \{[\s\S]*?\n {4}\}/m, 'apLoadAllLines');
assert(loader.includes("loadLineCache(lineKey, lineKey + '::ALL')"), 'ต้องอ่านแคชรายไลน์ตัวเดียวกับหน้า Stock');
assert(loader.includes("saveLineCache(lineKey, rows, lineKey + '::ALL')"), 'โหลดเสร็จต้องเขียนกลับแคชเดิม ให้หน้า Stock ใช้ต่อได้');
assert(loader.includes('lineKey === currentLine && partsData.length'), 'ไลน์ที่เปิดอยู่ต้องใช้ข้อมูลในมือ ไม่ยิงซ้ำ');
assert(/chain = chain\.then/.test(loader), 'ต้องยิงทีละไลน์ ไม่ยิงพร้อมกันทุกไลน์');
assert(loader.includes('{ lite: true'), 'ต้องโหลดแบบ lite เหมือนหน้า Stock');

// ── กระโดดไปไลน์ของอะไหล่ ต้องล้างฟิลเตอร์ที่ค้างอยู่ก่อน ───────────────────
// ไม่ล้าง = กดแล้ว "ไม่เจอ" ทั้งที่ของอยู่จริง เพราะฟิลเตอร์เดิมค้างจากการใช้งานก่อนหน้า
const opener = grab(/^ {4}function apOpenInLine\(lineKey, no, name, sheetName\) \{[\s\S]*?\n {4}\}/m, 'apOpenInLine');
assert(opener.includes('apResetStockFilters();'), 'ต้องล้างฟิลเตอร์หน้า Stock ก่อน');
assert(opener.includes('switchToStockPage();'), 'ต้องพาไปหน้า Stock');
assert(opener.includes('loadPartsData();'), 'สลับไลน์แล้วต้องโหลดข้อมูลไลน์ใหม่');
// ช่องค้นหาหน้า Stock เทียบกับ ชื่อ/รุ่น/แบรนด์ เท่านั้น — ส่ง No. ไปจะหาไม่เจอ
assert(/var keyword = String\(name \|\| ''\)\.trim\(\);/.test(opener), 'ต้องส่ง "ชื่อ" ไปค้น ไม่ใช่ No.');
const reset = grab(/^ {4}function apResetStockFilters\(\) \{[\s\S]*?\n {4}\}/m, 'apResetStockFilters');
['categoryFilter', 'brandFilter', 'statusFilter', 'subSheetFilter'].forEach(function(id) {
  assert(reset.includes("'" + id + "'"), 'ต้องล้าง ' + id);
});
assert(reset.includes("tableStatusMode = 'all'"), 'ต้องล้างโหมดสถานะที่กดจากการ์ดสถิติ');
assert(reset.includes('noPriceOnlyState = false') && reset.includes('newArrivalOnlyState = false'),
  'ต้องล้างโหมดกรองพิเศษ (ยังไม่มีราคา / ของใหม่)');

// ── รันของจริง: รวมพูล / ค้นหา / ฟิลเตอร์ / แบ่งหน้า ────────────────────────
const src = [
  grab(/^ {4}function safeNum\(value\) \{[\s\S]*?\n {4}\}/m, 'safeNum'),
  grab(/^ {4}function escHtml\(str\) \{[\s\S]*?\n {4}\}/m, 'escHtml'),
  grab(/^ {4}function getStockStatus\(item\) \{[\s\S]*?\n {4}\}/m, 'getStockStatus'),
  grab(/^ {4}function getStatusTone\(status\) \{[\s\S]*?\n {4}\}/m, 'getStatusTone'),
  grab(/^ {4}function getItemSourceSheet\(item\) \{[\s\S]*?\n {4}\}/m, 'getItemSourceSheet'),
  grab(/^ {4}var LINE_SHEET_OPTIONS = \{[\s\S]*?\n {4}\};/m, 'LINE_SHEET_OPTIONS'),
  grab(/^ {4}var AP_PAGE_SIZE = 60;[\s\S]*?\n(?= {4}var debouncedApRender)/m, 'all-parts module')
].join('\n');

function fakeInput(value) { return { value: value }; }
function fakeSelect(value) { return { value: value, innerHTML: '' }; }
const nodes = {
  allPartsSearch: fakeInput(''),
  allPartsLineFilter: fakeSelect('all'),
  allPartsCategoryFilter: fakeSelect('all'),
  allPartsStatusFilter: fakeSelect('all'),
  allPartsLineChips: { innerHTML: '' },
  allPartsTable: { innerHTML: '' },
  allPartsSummary: { textContent: '' },
  allPartsMoreBtn: { hidden: false, classList: { toggle: function(c, on) { nodes.allPartsMoreBtn.hidden = on; } } }
};
const api = new Function('document', 'console', src +
  '\nreturn { setByLine: function(v) { apByLine = v; }, apply: apApplyPool, render: apRender,' +
  ' more: function() { apVisibleCount += AP_PAGE_SIZE; }, reset: function() { apVisibleCount = AP_PAGE_SIZE; },' +
  ' pool: function() { return apPool; }, pageSize: AP_PAGE_SIZE };')(
  { getElementById: function(id) { return nodes[id] || null; } },
  { warn: function() {}, log: function() {} }
);

function part(no, name, extra) {
  return Object.assign({ no: no, name: name, model: 'M-' + no, brand: 'ACME', category: 'General',
    location: 'A-01', stock: 5, min: 2, __sourceSheet: 'Main List Stock' }, extra || {});
}

api.setByLine({
  'H9': [part('1001', 'Bearing 6203'), part('1002', 'Oil Seal', { stock: 0 })],
  'Arc Chute': [part('2001', 'Arc Plate', { __sourceSheet: 'Arc chut', location: 'B-07' }),
    part('2001', 'Arc Plate', { __sourceSheet: 'Arc chut', location: 'B-07' })] // ตัวซ้ำ ต้องถูกตัด
});
api.apply();
assert.strictEqual(api.pool().length, 3, 'รายการซ้ำในไลน์เดียวกันต้องถูกตัดออก');
assert(nodes.allPartsTable.innerHTML.includes('Bearing 6203'), 'ต้องเห็นอะไหล่ของ H9');
assert(nodes.allPartsTable.innerHTML.includes('Arc Plate'), 'ต้องเห็นอะไหล่ของ Arc Chute ในหน้าเดียวกัน');
assert(nodes.allPartsTable.innerHTML.includes('Arc chut'), 'ต้องบอกชีตย่อยที่ของอยู่จริง');

// ปุ่มกระโดดต้องพกข้อมูลครบพอให้เปิดถูกไลน์/ถูกชีต
const btn = /data-ap-open="1" data-ap-line="([^"]*)" data-ap-no="([^"]*)" data-ap-name="([^"]*)" data-ap-sheet="([^"]*)"/
  .exec(nodes.allPartsTable.innerHTML.replace(/\s+/g, ' '));
assert(btn, 'ทุกแถวต้องมีปุ่มเปิดในไลน์พร้อม data ครบ');
assert(btn[1] && btn[2] && btn[3], 'ปุ่มต้องรู้ line/no/name');

// รูปในลิสต์ต้อง lazy (ยามเดียวกับ ui-performance-guards)
api.setByLine({ 'H9': [part('1003', 'Belt', { image_main_url: 'https://x/y.jpg' })] });
api.apply();
assert(/<img[^>]*loading="lazy"[^>]*decoding="async"/.test(nodes.allPartsTable.innerHTML), 'รูปในลิสต์ต้อง lazy + decode async');

// ── ค้นด้วย No. ต้องเจอ (หน้า Stock เดิมค้นได้แค่ ชื่อ/รุ่น/แบรนด์) ──────────
api.setByLine({
  'H9': [part('1001', 'Bearing 6203'), part('1002', 'Oil Seal', { stock: 0 })],
  'Arc Chute': [part('2001', 'Arc Plate', { __sourceSheet: 'Arc chut', location: 'B-07' })]
});
api.apply();
nodes.allPartsSearch.value = '2001';
api.render();
assert(nodes.allPartsTable.innerHTML.includes('Arc Plate'), 'ค้นด้วยเลขที่อะไหล่ต้องเจอ');
assert(!nodes.allPartsTable.innerHTML.includes('Bearing 6203'), 'ตัวที่ไม่ตรงต้องไม่โผล่');
nodes.allPartsSearch.value = 'b-07';
api.render();
assert(nodes.allPartsTable.innerHTML.includes('Arc Plate'), 'ค้นด้วยที่เก็บต้องเจอ (ไม่สนตัวพิมพ์)');
nodes.allPartsSearch.value = '';

// ── ฟิลเตอร์ไลน์ / สถานะ ────────────────────────────────────────────────────
nodes.allPartsLineFilter.value = 'Arc Chute';
api.render();
assert(!nodes.allPartsTable.innerHTML.includes('Bearing 6203'), 'เลือกไลน์แล้วต้องเหลือเฉพาะไลน์นั้น');
nodes.allPartsLineFilter.value = 'all';
nodes.allPartsStatusFilter.value = 'out';
api.render();
assert(nodes.allPartsTable.innerHTML.includes('Oil Seal'), 'กรองของหมดสต็อกต้องเห็นตัวที่ stock = 0');
assert(!nodes.allPartsTable.innerHTML.includes('Arc Plate'), 'ของที่ยังมีสต็อกต้องไม่ติดมาด้วย');
nodes.allPartsStatusFilter.value = 'all';

// ── แบ่งหน้า: ของเยอะต้องไม่วาดทีเดียวหมด ──────────────────────────────────
const many = [];
for (let i = 0; i < api.pageSize + 25; i += 1) many.push(part('90' + i, 'Part ' + i));
api.setByLine({ 'H9': many });
api.reset();
api.apply();
assert.strictEqual((nodes.allPartsTable.innerHTML.match(/<tr /g) || []).length, api.pageSize,
  'ต้องวาดแค่หน้าแรก ไม่เทมาทั้งหมด');
assert.strictEqual(nodes.allPartsMoreBtn.hidden, false, 'ยังมีของเหลือ ต้องโชว์ปุ่มโหลดเพิ่ม');
api.more();
api.render();
assert.strictEqual((nodes.allPartsTable.innerHTML.match(/<tr /g) || []).length, many.length, 'กดโหลดเพิ่มแล้วต้องครบ');
assert.strictEqual(nodes.allPartsMoreBtn.hidden, true, 'ครบแล้วต้องซ่อนปุ่มโหลดเพิ่ม');

console.log('All-lines parts search checks passed');
