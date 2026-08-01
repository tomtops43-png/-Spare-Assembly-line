const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// ยามกันงานหนักหลุดกลับเข้ามาใน hot path (จอกระตุก/พิมพ์หน่วง)
// ทุกข้อในนี้เคยเป็นปัญหาจริงที่วัดได้ ไม่ใช่การเดา

function slice(startMarker, endMarker, source, label) {
  const start = source.indexOf(startMarker);
  assert(start > -1, 'can find ' + label);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert(end > start, 'can slice ' + label);
  return source.slice(start, end);
}

// ---- 1. resize ต้องถูกจำกัดรอบ ----
// เดิม resize ยิง applyFilters(false) ทุก event = กรอง partsData ทั้งก้อน + วาดกริดใหม่
// หลายสิบครั้ง/วินาที (มือถือ address bar ซ่อน/โผล่ตอนเลื่อนจอก็นับเป็น resize)
assert(html.includes('function rafThrottle(fn)'), 'rafThrottle helper exists');
assert(!/window\.addEventListener\('resize', function\(\)/.test(html), 'no raw unthrottled resize handler remains');
const resizeRelayout = slice('var handleResizeRelayout = rafThrottle(', 'window.addEventListener(\'resize\', handleResizeRelayout)', html, 'resize relayout handler');
assert(resizeRelayout.includes('pageSize === lastResizePageSize'), 'resize skips re-render when pageSize did not change');
assert(resizeRelayout.includes('isDesktop === lastResizeIsDesktop'), 'resize also tracks the 1024 breakpoint (sort control swaps there)');

// ---- 2. applyFilters ต้องไม่คิดงานของฟิลเตอร์ที่ปิดอยู่ ----
const applyFilters = slice('function applyFilters(resetPage)', '\n    function syncCoilFilterVisibility', html, 'applyFilters');
// สตริงค้นหาต้องสร้างเฉพาะตอนมีคำค้นจริง — indexOf('') คืน 0 เสมอ คิดตอนไม่ได้พิมพ์ = เสียเปล่า
assert(applyFilters.includes('hasKeyword && getItemSearchText(item)'), 'search text is only built when a keyword is present');
assert(!/\[item\.name, item\.model, item\.brand, item\.coil_size\]\.join/.test(applyFilters), 'applyFilters no longer builds the search string inline per item');
assert(applyFilters.includes('needSourceSheet && getItemSourceSheet(item)'), 'source sheet only resolved when that filter is active');
assert(applyFilters.includes('if (needStatus)'), 'stock status only computed when the status filter is active');
// เช็คถูกๆ ต้องมาก่อนเช็คแพงๆ เพื่อให้ตัดจบเร็ว
const catIdx = applyFilters.indexOf("category !== 'all'");
const searchIdx = applyFilters.indexOf('hasKeyword && getItemSearchText');
assert(catIdx > -1 && searchIdx > catIdx, 'cheap equality checks run before the expensive search check');

// ---- 3. แคชสตริงค้นหา ----
assert(html.includes('function getItemSearchText(item)'), 'memoized search-text helper exists');
assert(html.includes('item.__searchText === undefined'), 'search text is memoized per item');

// ---- 4. ช่องค้นหาที่วาดใหม่ทั้งลิสต์ต้องหน่วงคีย์ ----
assert(html.includes('var debouncedLogFilterRender = debounce(runLogFilterRender, 250)'), 'log search is debounced');
assert(/logFilterSearch'\) el\.addEventListener\('input', function\(\) \{ setter\(el\.value\); debouncedLogFilterRender\(\); \}\)/.test(html), 'log search input uses the debounced renderer');
// dropdown/ปุ่มกดทีเดียวจบ ต้องตอบสนองทันที ไม่ต้องหน่วง
assert(/el\.addEventListener\('change', function\(\) \{ setter\(el\.value\); runLogFilterRender\(\); \}\)/.test(html), 'log dropdowns still render immediately (no debounce)');
assert(html.includes("debounce(renderTxnResultList, 200)"), 'transaction search is debounced');
assert(html.includes("debounce(applyUserFiltersAndRender, 200)"), 'admin user search is debounced');
assert(html.includes("debounce(scRenderTable, 200)"), 'stock-count search is debounced');
assert(html.includes('var debouncedRenderPrCandidates = debounce(renderPrCandidates, 250)'), 'PR candidate search is debounced');
assert(html.includes("id === 'prCandidateSearch' ? debouncedRenderPrCandidates : renderPrCandidates"), 'only the PR text input is debounced, not its dropdowns');
assert(html.includes('p.__roSearchText === undefined'), 'request-order all-lines pool search is memoized');

// ---- 5. ไม่มี log/timer หนักค้างใน hot path ----
assert(!html.includes("console.time('filterLine')"), 'filterLine timer removed from the filter hot path');
assert(!html.includes("console.time('renderCards')"), 'renderCards timer removed from the render hot path');
// อันนี้ dump partsData ทั้งก้อน — DevTools ค้าง reference ทำให้ทั้งชุดไม่ถูก GC
assert(!html.includes("console.log('[item data after save/load]', partsData)"), 'full partsData dump removed (was pinning the dataset in memory)');
assert(!html.includes("console.log('btn exists'"), 'leftover debug logs removed from the badge updater');
assert(!html.includes("console.log('[render detail item]', item)"), 'full item dump removed from detail render');

// ---- 6. รูปในลิสต์ต้อง lazy ----
// เอาเฉพาะ <img> ที่ถูกสร้างเป็นสตริงใน JS ตอน render ลิสต์ (ต่อสตริงด้วย ' + ) เท่านั้น
// ไม่รวมรูป preview เดี่ยวๆ ที่อยู่ใน HTML ตรงๆ (เช่น #eImageMainPreview) — พวกนั้นเห็นทันที
// ที่เปิด modal อยู่แล้ว ใส่ lazy ไปก็ไม่ได้อะไร มีแต่จะทำให้ขึ้นช้ากว่าเดิม
// รูปของแดชบอร์ด FCC ไม่ได้เขียน attribute ตรงๆ แต่ยืมมาจาก fccImgAttrs() — เช็คที่ตัว helper แทน
assert(/function fccImgAttrs[\s\S]{0,600}loading="lazy" decoding="async"/.test(html), 'fccImgAttrs supplies lazy + async decoding to dashboard images');

// ดูทีละบรรทัด เพราะ regex ที่จับแค่ '<img ... จะโดนตัดที่ quote ตัวแรก มองไม่เห็น attribute ที่เหลือ
const imgLines = html.split(/\r?\n/).filter(l => l.includes("'<img"));
assert(imgLines.length >= 8, 'found the JS-built list images to check, got ' + imgLines.length);
let checked = 0;
imgLines.forEach(l => {
  // ยกเว้นตัวที่ต่อ attribute จาก helper (ตรวจแยกไปแล้วด้านบน)
  if (/class="fcc-(q|spare)-img"/.test(l)) return;
  // ยกเว้นรูปในเอกสารปริ้น/PDF — ตรงนั้นต้องการให้รูปโหลด+decode เสร็จก่อนสั่งปริ้น ใส่ async
  // แล้วเสี่ยงได้ PDF ช่องรูปว่าง (ดู waitForReportImages) สังเกตจากใช้ esc() ซึ่งเป็น helper
  // ของเอกสารปริ้นโดยเฉพาะ ส่วนที่อื่นในแอปใช้ escHtml()
  if (/drawing-full-image/.test(l) || /\besc\(/.test(l)) return;
  checked += 1;
  assert(/loading="lazy"/.test(l), 'JS-built list image is lazy-loaded: ' + l.trim().slice(0, 90));
  assert(/decoding="async"/.test(l), 'JS-built list image decodes off the main thread: ' + l.trim().slice(0, 90));
});
assert(checked >= 7, 'actually checked the in-app list images, got ' + checked);
// กันเผลอใส่ decoding="async" ให้รูปในเอกสารปริ้นในอนาคต (เคยเผลอมาแล้วตอนทำ PR นี้)
const annexImg = (html.match(/var imgHtml = c\.image \?[^\n]*/) || [''])[0];
assert(annexImg && !/decoding="async"/.test(annexImg), 'PDF export images must not decode asynchronously');

console.log('ui-performance-guards: all assertions passed');
