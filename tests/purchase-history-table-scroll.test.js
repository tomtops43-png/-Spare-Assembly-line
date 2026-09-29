const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// อาการเดิม: ตาราง Purchase History ยาว 20 แถว scrollbar แนวนอนอยู่ใต้แถวสุดท้าย
// จะเลื่อนดูคอลัมน์ขวาต้องเลื่อนหน้าลงไปล่างสุดก่อน
// แก้: ตารางอยู่ในกล่องสูงไม่เกินจอ (scrollbar อยู่ในจอเสมอ) + แถบเลื่อนด้านบน + หัวตาราง/คอลัมน์ Action ค้าง

// ---- 1. โครงสร้าง markup ----
const page = html.slice(html.indexOf('<section id="purchaseHistoryPage"'), html.indexOf('<section id="miscExpensePage"'));
['purchaseHistoryScrollToolbar', 'purchaseHistoryScrollLeft', 'purchaseHistoryScrollTrack',
  'purchaseHistoryScrollThumb', 'purchaseHistoryScrollRight', 'purchaseHistoryScrollShell', 'purchaseHistoryScroll'].forEach(function(id) {
  assert(page.indexOf('id="' + id + '"') > -1, 'มี #' + id + ' ในหน้า Purchase History');
});
assert(page.indexOf('id="purchaseHistoryScrollToolbar"') < page.indexOf('id="purchaseHistoryScroll"'),
  'แถบเลื่อนต้องอยู่เหนือตาราง');
assert(/<div id="purchaseHistoryScroll" class="ph-scroll">\s*<table class="ph-table/.test(page), 'ตารางอยู่ในกล่องเลื่อน .ph-scroll');
assert(/<th class="[^"]*ph-sticky-end[^"]*">Action<\/th>/.test(page), 'หัวคอลัมน์ Action ค้างขวา');

// ---- 2. CSS: กล่องสูงไม่เกินจอ + หัวตารางค้าง ----
// Tailwind ที่ใช้เป็น v2 static — arbitrary value อย่าง max-h-[70vh] ใช้ไม่ได้ ต้องเป็น CSS ตรงๆ
assert(/\.ph-scroll \{[^}]*max-height: calc\(100vh - \d+px\)[^}]*overflow: auto/.test(html), 'กล่องเลื่อนสูงไม่เกินจอ เลื่อนได้ทั้ง 2 แกน');
assert(/\.ph-table thead th \{[^}]*position: sticky; top: 0/.test(html), 'หัวตารางค้างบนในกล่อง');
assert(/\.ph-table \.ph-sticky-end \{ position: sticky; right: 0; \}/.test(html), 'คอลัมน์ Action ค้างขวา');
assert(/\.ph-table tbody td \{[^}]*background: #fff/.test(html), 'td มีพื้นทึบ — คอลัมน์ค้างไม่โปร่งเห็นข้อมูลที่เลื่อนผ่าน');

// ---- 3. JS: แถว render คอลัมน์ Action ค้าง + sync แถบเลื่อน ----
assert(/'<td class="ph-sticky-end px-3 py-3">' \+ actions \+ '<\/td><\/tr>'/.test(html), 'แถวใส่ ph-sticky-end ที่ช่อง Action');
const renderFn = html.slice(html.indexOf('function renderPurchaseHistory()'), html.indexOf('function loadPurchaseHistory('));
assert(renderFn.indexOf('updatePurchaseHistoryScrollUi();') > -1, 'render แล้วอัปเดตแถบเลื่อน');
assert(/purchaseHistoryScroll\.addEventListener\('scroll', updatePurchaseHistoryScrollUi/.test(html), 'เลื่อนตาราง → แถบเลื่อนตาม');
assert(/new ResizeObserver\(/.test(html), 'แท็บโผล่ (เคย render ตอนกว้าง 0) → แถบเลื่อนคำนวณใหม่');

// ---- 4. ตรรกะคำนวณ thumb ----
const fnSrc = html.slice(html.indexOf('function updatePurchaseHistoryScrollUi()'), html.indexOf('function scrollPurchaseHistoryBy('));
function makeEl() { return { classList: { toggle: function(c, on) { this[c] = !!on; } }, style: {} }; }
const scroller = { scrollLeft: 0, scrollWidth: 2000, clientWidth: 1000 };
const thumb = makeEl(), shell = makeEl(), toolbar = makeEl(), leftBtn = {}, rightBtn = {};
const run = new Function('purchaseHistoryScroll', 'purchaseHistoryScrollThumb', 'purchaseHistoryScrollToolbar',
  'purchaseHistoryScrollShell', 'purchaseHistoryScrollLeftBtn', 'purchaseHistoryScrollRightBtn', 'purchaseHistoryScrollTrack',
  fnSrc + '; updatePurchaseHistoryScrollUi();');
const track = { clientWidth: 400 };
run(scroller, thumb, toolbar, shell, leftBtn, rightBtn, track);
assert.strictEqual(thumb.style.width, '200px', 'thumb กว้างตามสัดส่วนที่มองเห็น');
assert.strictEqual(thumb.style.transform, 'translateX(0px)');
assert.strictEqual(leftBtn.disabled, true, 'อยู่ซ้ายสุด → ปุ่มซ้ายกดไม่ได้');
assert.strictEqual(rightBtn.disabled, false);
assert.strictEqual(shell.classList['ph-can-right'], true);
scroller.scrollLeft = 1000;
run(scroller, thumb, toolbar, shell, leftBtn, rightBtn, track);
assert.strictEqual(thumb.style.transform, 'translateX(200px)', 'ขวาสุด → thumb ชิดขวา');
assert.strictEqual(rightBtn.disabled, true);
assert.strictEqual(shell.classList['ph-can-left'], true);
scroller.scrollWidth = 1000; scroller.scrollLeft = 0;
run(scroller, thumb, toolbar, shell, leftBtn, rightBtn, track);
assert.strictEqual(toolbar.classList['ph-no-overflow'], true, 'ตารางไม่ล้น → แถบเลื่อนจางลง');

console.log('PASS purchase-history-table-scroll');
