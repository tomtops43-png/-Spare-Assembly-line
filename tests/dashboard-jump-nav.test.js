const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── แถบข้ามไปยังการ์ด/กราฟบน Dashboard ───────────────────────────────────────
assert(htmlLf.includes('<nav id="fccJump"'), 'ต้องมีแถบข้ามไปยังส่วนต่างๆ');
assert(htmlLf.includes('id="fccJumpList"'), 'ต้องมีที่สำหรับใส่ปุ่มข้าม');

// แถบต้องอยู่ "นอก" .fcc-canvas — canvas มี overflow:hidden ซึ่งทำให้ position:sticky
// ของลูกไปยึดกับ canvas แทน viewport แล้วแถบจะเลื่อนหายไปกับหน้า ไม่ค้างอยู่ด้านบน
const dashStart = htmlLf.indexOf('<section id="dashboardPage"');
assert(dashStart > -1, 'ต้องหา section dashboard เจอ');
const jumpAt = htmlLf.indexOf('<nav id="fccJump"', dashStart);
const canvasAt = htmlLf.indexOf('<div class="fcc-canvas">', dashStart);
assert(jumpAt > dashStart && canvasAt > dashStart, 'ทั้งคู่ต้องอยู่ใน section dashboard');
assert(jumpAt < canvasAt, 'แถบข้ามต้องอยู่ก่อน/นอก .fcc-canvas ไม่งั้น sticky ใช้ไม่ได้');

// CSS ที่ทำให้แถบใช้งานได้จริง
assert(/\.fcc-jump\{[^}]*position:sticky/.test(htmlLf.replace(/\s+/g, ' ').replace(/ \{/g, '{')) ||
  /\.fcc-jump\{[\s\S]{0,200}position:sticky/.test(htmlLf), 'แถบต้องเป็น sticky');
assert(/scroll-margin-top:\s*\d+px/.test(htmlLf), 'ต้องกันแถบ sticky บังหัวการ์ดตอนกดข้าม');

// ── ทุกปุ่มต้องชี้ไป id ที่มีอยู่จริงในหน้า ─────────────────────────────────
// พิมพ์ id ผิดตัวเดียว = ปุ่มนั้นกดแล้วไม่ไปไหน และไม่มีอะไรฟ้อง
const sectionsSrc = grab(/^ {4}var FCC_SECTIONS = \[[\s\S]*?\n {4}\];/m, 'FCC_SECTIONS');
const FCC_SECTIONS = new Function(sectionsSrc + '\nreturn FCC_SECTIONS;')();
assert(FCC_SECTIONS.length >= 10, 'ต้องมีรายการข้ามครบทุกโซนหลัก');
FCC_SECTIONS.forEach(function(s) {
  assert(s.id && s.label && s.icon, 'ทุกรายการต้องมี id/label/icon');
  assert(htmlLf.includes('id="' + s.id + '"'), 'ไม่พบการ์ด id="' + s.id + '" ที่ปุ่มข้ามชี้ไป');
});
const ids = FCC_SECTIONS.map(function(s) { return s.id; });
assert.strictEqual(new Set(ids).size, ids.length, 'id ต้องไม่ซ้ำกัน');

// การ์ดใหม่ที่เพิ่งเพิ่มเข้าหน้าต้องถูกใส่ในลิสต์ด้วย ไม่งั้นแถบจะตกรุ่นเงียบๆ
['fccMachineCard', 'fccMiscCard', 'fccRepairCard', 'fccCostRatioCard'].forEach(function(id) {
  assert(ids.indexOf(id) > -1, 'ลิสต์ต้องมี ' + id);
});

// ต้องใช้ handler เดิมของ FCC (scroll:<id>) ไม่ผูก listener ใหม่รายปุ่ม
assert(/data-fcc-act="scroll:' \+ s\.id \+ '"/.test(htmlLf), 'ปุ่มต้องใช้ act "scroll:<id>" ตัวเดิม');
assert(/fccInited = true;\n\s*fccBuildJump\(\);/.test(htmlLf), 'fccInit ต้องสร้างแถบข้าม');

// ── ไฮไลต์ตามตำแหน่งที่กำลังดูอยู่ ──────────────────────────────────────────
const src = [
  grab(/^ {4}function fccEsc\(v\) \{[\s\S]*?\n {4}\}/m, 'fccEsc'),
  sectionsSrc,
  grab(/^ {4}var fccJumpBuilt = false;$/m, 'fccJumpBuilt'),
  grab(/^ {4}var fccJumpRaf = 0;$/m, 'fccJumpRaf'),
  grab(/^ {4}var fccJumpActiveId = '';$/m, 'fccJumpActiveId'),
  grab(/^ {4}function fccBuildJump\(\) \{[\s\S]*?\n {4}\}/m, 'fccBuildJump'),
  grab(/^ {4}function fccQueueJumpSync\(\) \{[\s\S]*?\n {4}\}/m, 'fccQueueJumpSync'),
  grab(/^ {4}function fccSyncJumpActive\(\) \{[\s\S]*?\n {4}\}/m, 'fccSyncJumpActive')
].join('\n');

// ตำแหน่งจำลอง: การ์ดที่ผ่านใต้แถบไปแล้ว (top <= 140) ตัวล่างสุดคือตัวที่กำลังดู
let tops = {};
function fakeCard(id) {
  const cls = new Set();
  return {
    id: id,
    classList: {
      contains: (c) => cls.has(c),
      toggle: (c, on) => { on ? cls.add(c) : cls.delete(c); return on; },
      has: (c) => cls.has(c)
    },
    getAttribute: (k) => (k === 'data-jump' ? id : null),
    getBoundingClientRect: () => ({ top: tops[id] === undefined ? 9999 : tops[id], left: 0, right: 100 })
  };
}
const chips = [];
const jumpList = {
  innerHTML: '',
  scrollLeft: 0,
  get children() { return chips; },
  getBoundingClientRect: () => ({ left: 0, right: 1000 })
};
const page = { classList: { contains: () => false } };
const nodes = { fccJumpList: jumpList, dashboardPage: page };
FCC_SECTIONS.forEach(function(s) { nodes[s.id] = fakeCard(s.id); });

const api = new Function('document', 'window', 'requestAnimationFrame',
  src + '\nreturn { build: fccBuildJump, sync: fccSyncJumpActive };')(
  { getElementById: (id) => nodes[id] || null },
  { addEventListener: function() {} },
  function() {}
);

api.build();
assert(jumpList.innerHTML.includes('data-fcc-act="scroll:' + FCC_SECTIONS[0].id + '"'),
  'ต้องสร้างปุ่มจาก FCC_SECTIONS');
assert.strictEqual((jumpList.innerHTML.match(/class="fcc-jump-it"/g) || []).length, FCC_SECTIONS.length,
  'ต้องมีปุ่มครบทุกส่วน');

// จำลองปุ่มจริงให้ sync ใช้ (ของจริงเป็น DOM ที่เกิดจาก innerHTML)
FCC_SECTIONS.forEach(function(s) { chips.push(fakeCard(s.id)); });

// เลื่อนมาอยู่ที่การ์ดใบที่ 3 — ใบ 1-3 ผ่านแถบไปแล้ว ที่เหลือยังอยู่ล่างจอ
tops[FCC_SECTIONS[0].id] = -900;
tops[FCC_SECTIONS[1].id] = -400;
tops[FCC_SECTIONS[2].id] = 60;
api.sync();
assert(chips[2].classList.has('on'), 'ต้องไฮไลต์การ์ดที่กำลังดูอยู่');
assert(!chips[1].classList.has('on') && !chips[3].classList.has('on'), 'ต้องไฮไลต์ทีละอันเดียว');

// เลื่อนต่อไปการ์ดใบที่ 4
tops[FCC_SECTIONS[2].id] = -200;
tops[FCC_SECTIONS[3].id] = 100;
api.sync();
assert(chips[3].classList.has('on') && !chips[2].classList.has('on'), 'ไฮไลต์ต้องขยับตามการเลื่อน');

// อยู่บนสุดยังไม่เลื่อน — ต้องไฮไลต์ตัวแรกไว้ ไม่ใช่ปล่อยว่าง
tops = {};
api.sync();
assert(chips[0].classList.has('on'), 'ตอนอยู่บนสุดต้องไฮไลต์รายการแรก');

console.log('Dashboard jump navigation checks passed');
