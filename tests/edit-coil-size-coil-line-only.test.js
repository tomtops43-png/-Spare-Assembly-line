const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

function grab(src, name) {
  const m = src.match(new RegExp('^ {0,4}function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n {0,4}\\}', 'm'));
  assert(m, 'ต้องดึง ' + name + ' ออกมาได้');
  return m[0];
}

// ── ฟอร์มแก้ไข: ช่อง Coil Size ขึ้นเฉพาะไลน์ Coil Winding ───────────────────────────
assert(htmlLf.includes('<div id="eCoilSizeWrap">'), 'ช่อง Coil Size ในฟอร์มแก้ไขต้องมีกล่องให้ซ่อนได้');

function run(sheet, line) {
  const els = {
    eSheet: { value: sheet },
    eLine: { value: line },
    eCoilSizeWrap: { hidden: false, classList: { toggle: function(c, on) { els.eCoilSizeWrap.hidden = on; } } }
  };
  const fn = new Function('document', 'SHEET_LINE_MAP', 'canonicalLineName', [
    grab(htmlLf, 'editFormLineKey'),
    grab(htmlLf, 'syncEditCoilSizeVisibility'),
    'syncEditCoilSizeVisibility();'
  ].join('\n'));
  fn({ getElementById: function(id) { return els[id] || null; } },
    { 'Coil Winding': 'Coil Winding', 'Standard Spare part': 'Arc Chute', 'Main List Stock': 'H9', 'Lug&Screw': 'Lug&Screw' },
    function(v) { return /coil/i.test(String(v)) ? 'Coil Winding' : String(v || ''); });
  return els.eCoilSizeWrap.hidden;
}
assert.strictEqual(run('Coil Winding', 'Coil Winding'), false, 'ไลน์ Coil ต้องเห็นช่อง Coil Size');
assert.strictEqual(run('Standard Spare part', 'Arc Chute'), true, 'Arc Chute ห้ามเห็น');
assert.strictEqual(run('Main List Stock', 'H9'), true, 'H9 ห้ามเห็น');
assert.strictEqual(run('Lug&Screw', 'Lug&Screw'), true, 'Lug&Screw ห้ามเห็น');

// เรียกทุกครั้งที่เปิดฟอร์มแก้ไข และเมื่อแก้ช่อง Line
assert(/fillLocationSelect\(document\.getElementById\('eLocation'\)[^\n]*\n\s*syncEditCoilSizeVisibility\(\);/.test(htmlLf));
assert(htmlLf.includes("eLineEl.addEventListener('change', syncEditCoilSizeVisibility);"));
// ซ่อนแล้วส่งค่าว่าง ต้องไม่ลบ Coil Size เดิมในชีต
assert(backend.includes("if (key === 'coil_size' && String(nextValue || '').trim() === '') continue;"));
console.log('edit-coil-size-coil-line-only: ok');
