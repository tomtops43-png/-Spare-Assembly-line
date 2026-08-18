const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

// ── ดึง bindBackdropClose ตัวจริงออกมารันจริง ────────────────────────────────
const src = script.match(/function bindBackdropClose\(element, onClose, isBackdrop\)[\s\S]*?\n    \}/);
assert(src, 'ต้องมีฟังก์ชัน bindBackdropClose');
const bindBackdropClose = new Function(src[0] + '\nreturn bindBackdropClose;')();

// element ปลอมแบบง่ายๆ พอให้ addEventListener/dispatch ทำงาน
function makeEl(name) {
  return {
    name: name,
    handlers: {},
    addEventListener: function(type, fn) { (this.handlers[type] = this.handlers[type] || []).push(fn); },
    removeEventListener: function(type, fn) {
      (this.handlers[type] || []).some(function(h, i, arr) { if (h === fn) { arr.splice(i, 1); return true; } return false; });
    },
    fire: function(type, target) { (this.handlers[type] || []).forEach(function(fn) { fn({ type: type, target: target }); }); }
  };
}

// จำลอง: กดเมาส์ที่ down แล้วปล่อยที่ up — เบราว์เซอร์ยิง click ที่บรรพบุรุษร่วม (clickTarget)
function drag(modal, down, up, clickTarget) {
  modal.fire('mousedown', down);
  modal.fire('mouseup', up);
  modal.fire('click', clickTarget);
}

function setup() {
  const modal = makeEl('modal');       // ตัว overlay = พื้นหลัง
  const input = makeEl('input');       // ช่องกรอกในฟอร์ม
  const state = { closed: 0 };
  bindBackdropClose(modal, function() { state.closed += 1; });
  return { modal: modal, input: input, state: state };
}

// ── เคสของบั๊ก: ลากเลือกข้อความในช่องกรอก แล้วปล่อยเมาส์นอกกรอบ ───────────────
// เบราว์เซอร์ตั้ง target ของ click เป็นตัว modal เอง (บรรพบุรุษร่วม) — ต้องไม่ปิด
var t = setup();
drag(t.modal, t.input, t.modal, t.modal);
assert.strictEqual(t.state.closed, 0, 'ลากเลือกข้อความจากในฟอร์มออกไปปล่อยนอกกรอบ ต้องไม่ปิด modal');

// ── คลิกพื้นหลังจริงๆ ต้องยังปิดได้เหมือนเดิม ────────────────────────────────
t = setup();
drag(t.modal, t.modal, t.modal, t.modal);
assert.strictEqual(t.state.closed, 1, 'คลิกพื้นหลังต้องปิด modal');

// ── ลากจากพื้นหลังเข้ามาปล่อยในฟอร์ม ต้องไม่ปิด ──────────────────────────────
t = setup();
drag(t.modal, t.modal, t.input, t.modal);
assert.strictEqual(t.state.closed, 0, 'ลากจากพื้นหลังเข้ามาในฟอร์ม ต้องไม่ปิด modal');

// ── คลิกในฟอร์มปกติ ต้องไม่ปิด ───────────────────────────────────────────────
t = setup();
drag(t.modal, t.input, t.input, t.input);
assert.strictEqual(t.state.closed, 0, 'คลิกในฟอร์มต้องไม่ปิด modal');

// ── คลิกพื้นหลังซ้ำ 2 ครั้ง ต้องนับครบ (ธงถูกรีเซ็ตถูกต้อง) ───────────────────
t = setup();
drag(t.modal, t.modal, t.modal, t.modal);
drag(t.modal, t.modal, t.modal, t.modal);
assert.strictEqual(t.state.closed, 2, 'ธง pressed/released ต้องรีเซ็ตหลังทุกครั้ง');

// ── รองรับ matcher เอง (imageViewer ปิดได้ทั้งตัว viewer และกรอบใน) ───────────
var viewer = makeEl('viewer');
var inner = makeEl('inner');
var img = makeEl('img');
var viewerClosed = 0;
bindBackdropClose(viewer, function() { viewerClosed += 1; }, function(target) { return target === viewer || target === inner; });
drag(viewer, inner, inner, inner);
assert.strictEqual(viewerClosed, 1, 'matcher ที่ส่งเข้ามาต้องถูกใช้');
drag(viewer, img, viewer, viewer);
assert.strictEqual(viewerClosed, 1, 'ลากจากรูปออกไปปล่อยนอกกรอบ ต้องไม่ปิด viewer');

// ── ห้ามมี modal ไหนกลับไปใช้ท่าเดิม (เช็ค target แค่ในอีเวนต์ click) ─────────
const naive = script.match(/if \(e\.target === [A-Za-z0-9_$]+\)/g) || [];
assert.deepStrictEqual(naive, [],
  'ห้ามเช็ค e.target ในอีเวนต์ click ตรงๆ — ต้องใช้ bindBackdropClose ไม่งั้นบั๊กลากเลือกข้อความจะกลับมา');

console.log('Modal backdrop drag-close checks passed');
