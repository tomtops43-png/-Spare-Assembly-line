const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

function extract(sig) {
  const start = script.indexOf('    function ' + sig);
  assert(start > -1, 'cannot find ' + sig);
  const end = script.indexOf('\n    }', start);
  assert(end > -1, 'cannot find end of ' + sig);
  return script.slice(start, end + 6);
}

// ── ค่าคงที่ในโค้ดต้องตรงกับ markup จริง ไม่งั้นคำนวณคอลัมน์ผิดแล้วช่องโหว่กลับมา ──
const gridTag = html.match(/<div id="desktopCardGrid"[^>]*>/)[0];
assert(gridTag.indexOf('minmax(220px,1fr)') > -1,
  'markup ต้องยังเป็น minmax(220px,1fr) ให้ตรงกับ STOCK_GRID_MIN_COL');
assert(/class="[^"]*\bgap-3\b/.test(gridTag), 'markup ต้องยังเป็น gap-3 ให้ตรงกับ STOCK_GRID_GAP (12px)');
assert(script.indexOf('var STOCK_GRID_MIN_COL = 220;') > -1);
assert(script.indexOf('var STOCK_GRID_GAP = 12;') > -1);

const mobileTag = html.match(/<div id="mobileCardList"[^>]*>/)[0];
assert(/\bgrid-cols-2\b/.test(mobileTag), 'กริดมือถือต้องยังเป็น 2 คอลัมน์');
assert(script.indexOf('var STOCK_GRID_MOBILE_COLS = 2;') > -1);

// ── รันตัวคำนวณจริง ──────────────────────────────────────────────────────────
const bundle = [
  '    var STOCK_GRID_MIN_COL = 220;',
  '    var STOCK_GRID_GAP = 12;',
  '    var STOCK_GRID_MOBILE_COLS = 2;',
  '    var pageSize = 10;',
  '    var window = { innerWidth: 0 };',
  '    var document = { getElementById: function() { return DOC_GRID; } };',
  '    var DOC_GRID = null;',
  extract('getStockGridColumnCount()'),
  extract('updateResponsivePageSize()')
].join('\n');
const api = new Function(bundle +
  '\nreturn function(innerWidth, gridWidth) {' +
  '  window.innerWidth = innerWidth;' +
  '  DOC_GRID = gridWidth ? { clientWidth: gridWidth, parentNode: null } : null;' +
  '  var cols = getStockGridColumnCount();' +
  '  updateResponsivePageSize();' +
  '  return { cols: cols, pageSize: pageSize };' +
  '};')();

// สูตรเดียวกับที่เบราว์เซอร์ใช้ตัดคอลัมน์ auto-fit
function expectedCols(width) {
  return Math.max(1, Math.floor((width + 12) / (220 + 12)));
}

// ── หัวใจของบั๊ก: pageSize ต้องหารลงตัวกับจำนวนคอลัมน์เสมอ ───────────────────
// ไม่งั้นแถวสุดท้ายจะมีช่องโหว่ทั้งที่ยังมีของหน้าถัดไป
const widths = [1024, 1100, 1280, 1366, 1440, 1536, 1600, 1680, 1920, 2200, 2560, 3440];
widths.forEach(function(w) {
  const gridWidth = Math.round(w * 0.94); // หักขอบ/padding คร่าวๆ
  const r = api(w, gridWidth);
  assert.strictEqual(r.cols, expectedCols(gridWidth), 'จอ ' + w + ' ต้องได้ ' + expectedCols(gridWidth) + ' คอลัมน์');
  assert.strictEqual(r.pageSize % r.cols, 0,
    'จอ ' + w + 'px: pageSize ' + r.pageSize + ' ต้องหารด้วย ' + r.cols + ' คอลัมน์ลงตัว (แถวสุดท้ายห้ามมีช่องโหว่)');
  assert(r.pageSize > 0, 'pageSize ต้องมากกว่า 0');
});

// เคสจริงจากที่ผู้ใช้เจอ: 5 คอลัมน์ ต้องไม่ได้ 24 อีกแล้ว
const case1366 = api(1366, Math.round(1366 * 0.94));
assert.strictEqual(case1366.cols, 5);
assert.strictEqual(case1366.pageSize, 25, 'จอ 5 คอลัมน์ต้องได้ 25 ใบ (5 แถวเต็ม) ไม่ใช่ 24');

// จอกว้าง 7 คอลัมน์ เดิมได้ 30 (เหลือ 2 ใบแถวสุดท้าย) ต้องไม่เกิดอีก
const case1920 = api(1920, Math.round(1920 * 0.94));
assert.strictEqual(case1920.cols, 7);
assert.strictEqual(case1920.pageSize % 7, 0, 'จอ 7 คอลัมน์ต้องหารลงตัว');

// ── มือถือ: 2 คอลัมน์คงที่ ────────────────────────────────────────────────────
[360, 414, 768, 1023].forEach(function(w) {
  const r = api(w, 0);
  assert.strictEqual(r.cols, 2, 'จอเล็กต้องใช้กริดมือถือ 2 คอลัมน์');
  assert.strictEqual(r.pageSize % 2, 0, 'จอเล็ก pageSize ต้องหาร 2 ลงตัว');
});

// ── ความหนาแน่นต้องใกล้เคียงของเดิม ไม่ใช่โดดไปหน้าละ 100 ใบ ─────────────────
widths.forEach(function(w) {
  const r = api(w, Math.round(w * 0.94));
  const target = w < 768 ? 10 : (w < 1200 ? 12 : (w < 1600 ? 24 : 30));
  assert(r.pageSize <= target + r.cols && r.pageSize >= Math.max(r.cols, target - r.cols),
    'จอ ' + w + ': pageSize ' + r.pageSize + ' ต้องใกล้เคียงเป้า ' + target);
});

// ── กันตอน grid ยังไม่ถูกวาด (clientWidth = 0) ต้องไม่พังเป็น 0 คอลัมน์ ───────
const noGrid = api(1440, 0);
assert(noGrid.cols >= 1, 'ตอนยังไม่มีกริดต้องยังคืนอย่างน้อย 1 คอลัมน์');
assert.strictEqual(noGrid.pageSize % noGrid.cols, 0);

console.log('Stock grid page size checks passed');
