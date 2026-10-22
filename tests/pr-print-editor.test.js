const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// ---- ปริ้นใบ PR จาก Inbox: เปิดหน้าต่างแก้ไขก่อน แล้วค่อย build รายงานจากค่าที่แก้ ----
const openFn = html.slice(html.indexOf('function openApprovedPrPrintWindow('), html.indexOf('function openPrPrintEditor('));
assert(openFn.includes('openPrPrintEditor(header, items, opts);'), 'ปุ่มปริ้นเปิดหน้าต่างแก้ไขก่อน');
assert(!openFn.includes('.print()'), 'ไม่สั่งปริ้นทันทีโดยข้ามหน้าแก้ไข');
assert(/function printPrDocument\(header, items, opts\) \{\s*var html = buildDocumentGradeReport\(items, opts\);/.test(html), 'ปริ้นจากค่าที่แก้แล้ว');

// collectPrPrintEdits: อ่านค่าจากฟอร์ม → items/opts ใหม่ (ไม่แก้ของเดิม) และข้ามแถวที่ลบ
function grab(name) {
  const start = html.indexOf(name);
  let i = html.indexOf('{', start), depth = 0;
  for (; i < html.length; i += 1) { if (html[i] === '{') depth += 1; else if (html[i] === '}' && --depth === 0) break; }
  return html.slice(start, i + 1);
}
const itemFields = html.match(/var PR_PRINT_EDIT_ITEM_FIELDS = \[[\s\S]*?\];/)[0];
const collect = new Function(itemFields + '\n' + grab('function collectPrPrintEdits(') + '; return collectPrPrintEdits;')();

function input(v) { return { value: v }; }
function fakeRoot(head, rows) {
  return {
    querySelectorAll: function(sel) {
      assert.strictEqual(sel, '[data-pe-head]');
      return Object.keys(head).map(function(k) { return { value: head[k], getAttribute: function() { return k; } }; });
    },
    querySelector: function(sel) {
      const m = sel.match(/data-pe-row="(\d+)"/);
      const row = m && rows[m[1]];
      if (!row) return null;
      return { querySelector: function(s) { const f = s.match(/data-pe-f="(\w+)"/)[1]; return row[f] !== undefined ? input(row[f]) : null; } };
    }
  };
}
const items = [
  { name: 'Inlet Center Tube', suggestedQty: 10, unit_price: 3158.75, purpose: '' },
  { name: 'MT PIN', suggestedQty: 12, unit_price: 7670, purpose: '' }
];
const opts = { preparedByOverride: 'Admin', prNumber: 'PR-1' };
const out = collect(
  fakeRoot({ preparedByOverride: ' Wanchai ' }, { 0: { name: 'Inlet Tube 25/32A', suggestedQty: '5', unit_price: '3000', purpose: 'เปลี่ยนด่วน' } }),
  items, opts, { 1: true }
);
assert.strictEqual(out.opts.preparedByOverride, 'Wanchai');
assert.strictEqual(out.opts.prNumber, 'PR-1', 'ค่าที่ไม่ได้อยู่ในฟอร์มยังอยู่');
assert.strictEqual(out.items.length, 1, 'แถวที่กดลบไม่ถูกพิมพ์');
assert.strictEqual(out.items[0].name, 'Inlet Tube 25/32A');
assert.strictEqual(out.items[0].suggestedQty, 5);
assert.strictEqual(out.items[0].unit_price, 3000);
assert.strictEqual(out.items[0].unitPrice, 3000);
assert.strictEqual(out.items[0].reason, 'เปลี่ยนด่วน');
assert.strictEqual(items[0].suggestedQty, 10, 'ไม่แก้ข้อมูลต้นฉบับ');
assert.strictEqual(opts.preparedByOverride, 'Admin', 'ไม่แก้ opts ต้นฉบับ');

console.log('pr-print-editor: ok');
