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

// ── ต้องเลื่อนดูรายการที่ล้นออกนอกจอได้ ─────────────────────────────────────
// เมาส์ทั่วไปไม่มีล้อแนวนอน + เราซ่อน scrollbar ไว้ ถ้าไม่มีทางเลื่อนเลย
// รายการท้ายๆ จะกดไม่ได้ถาวร
assert(htmlLf.includes('id="fccJumpPrev"') && htmlLf.includes('id="fccJumpNext"'),
  'ต้องมีปุ่มเลื่อนซ้าย-ขวา');
assert(/list\.addEventListener\('wheel'/.test(htmlLf), 'ต้องแปลงล้อเมาส์แนวตั้งเป็นการเลื่อนแนวนอน');
assert(/kind === 'jumpscroll'/.test(htmlLf), 'ปุ่มลูกศรต้องมี handler');
assert(!/\.fcc-jump-list\{[^}]*scroll-behavior:smooth/.test(htmlLf.replace(/\n/g, '')),
  'ห้ามใส่ scroll-behavior:smooth ที่ตัว list — หมุนล้อแล้วจะหนืด');

// ── ไฮไลต์ตามตำแหน่งที่กำลังดูอยู่ ──────────────────────────────────────────
const src = [
  grab(/^ {4}function fccEsc\(v\) \{[\s\S]*?\n {4}\}/m, 'fccEsc'),
  sectionsSrc,
  grab(/^ {4}var fccJumpBuilt = false;$/m, 'fccJumpBuilt'),
  grab(/^ {4}var fccJumpRaf = 0;$/m, 'fccJumpRaf'),
  grab(/^ {4}var fccJumpActiveId = '';$/m, 'fccJumpActiveId'),
  grab(/^ {4}function fccBuildJump\(\) \{[\s\S]*?\n {4}\}/m, 'fccBuildJump'),
  grab(/^ {4}function fccSyncJumpArrows\(\) \{[\s\S]*?\n {4}\}/m, 'fccSyncJumpArrows'),
  grab(/^ {4}function fccScrollJumpList\(dir\) \{[\s\S]*?\n {4}\}/m, 'fccScrollJumpList'),
  grab(/^ {4}function fccQueueJumpSync\(\) \{[\s\S]*?\n {4}\}/m, 'fccQueueJumpSync'),
  grab(/^ {4}function fccSyncJumpActive\(\) \{[\s\S]*?\n {4}\}/m, 'fccSyncJumpActive')
].join('\n');

const CARD_H = 300;
let tops = {};
function fakeEl(id) {
  const cls = new Set();
  const attrs = {};
  return {
    id: id,
    classList: {
      toggle: (c, on) => { on ? cls.add(c) : cls.delete(c); return on; },
      has: (c) => cls.has(c),
      contains: (c) => cls.has(c)
    },
    setAttribute: (k, v) => { attrs[k] = v; },
    removeAttribute: (k) => { delete attrs[k]; },
    getAttribute: (k) => (k === 'data-jump' ? id : (k in attrs ? attrs[k] : null)),
    has: (k) => k in attrs,
    getBoundingClientRect: () => {
      const top = tops[id] === undefined ? 9999 : tops[id];
      return { top: top, bottom: top + CARD_H, left: 0, right: 100 };
    }
  };
}
const chips = [];
const jumpList = {
  innerHTML: '', scrollLeft: 0, clientWidth: 900, scrollWidth: 900,
  get children() { return chips; },
  addEventListener: function() {},
  scrollTo: function(o) { this.scrollLeft = o.left; },
  getBoundingClientRect: () => ({ left: 0, right: 900 })
};
const prevBtn = fakeEl('fccJumpPrev');
const nextBtn = fakeEl('fccJumpNext');
const page = { classList: { contains: () => false } };
const bar = { getBoundingClientRect: () => ({ bottom: 60 }) };
const nodes = { fccJumpList: jumpList, dashboardPage: page, fccJump: bar, fccJumpPrev: prevBtn, fccJumpNext: nextBtn };
FCC_SECTIONS.forEach(function(s) { nodes[s.id] = fakeEl(s.id); });

// resize ต้องผ่าน rafThrottle (ยามใน ui-performance-guards บังคับไว้) — ใช้ตัวจริงจาก index.html
const rafThrottleSrc = grab(/^ {4}function rafThrottle\(fn\) \{[\s\S]*?\n {4}\}/m, 'rafThrottle');
const api = new Function('document', 'window', 'requestAnimationFrame',
  rafThrottleSrc + '\n' + src +
  '\nreturn { build: fccBuildJump, sync: fccSyncJumpActive, arrows: fccSyncJumpArrows, scrollList: fccScrollJumpList };')(
  { getElementById: (id) => nodes[id] || null },
  { addEventListener: function() {}, innerHeight: 800 },
  function() {}
);

api.build();
assert(jumpList.innerHTML.includes('data-fcc-act="scroll:' + FCC_SECTIONS[0].id + '"'),
  'ต้องสร้างปุ่มจาก FCC_SECTIONS');
assert.strictEqual((jumpList.innerHTML.match(/class="fcc-jump-it"/g) || []).length, FCC_SECTIONS.length,
  'ต้องมีปุ่มครบทุกส่วน');
FCC_SECTIONS.forEach(function(s) { chips.push(fakeEl(s.id)); });

// ── การ์ดที่วางคู่กันในแถวเดียว ต้องติดสีทั้งคู่ ────────────────────────────
// เดิมไฮไลต์ใบเดียว ทำให้ Risk Heatmap / Top Intelligence ที่เห็นพร้อมกันบนจอ
// ติดสีแค่ใบเดียว ดูเหมือนระบบชี้ผิดใบ
tops = {};
tops[FCC_SECTIONS[2].id] = -900;      // เลื่อนผ่านไปแล้ว อยู่เหนือจอ
tops[FCC_SECTIONS[3].id] = 100;       // คู่ที่กำลังดู
tops[FCC_SECTIONS[4].id] = 100;
api.sync();
assert(chips[3].classList.has('on') && chips[4].classList.has('on'),
  'การ์ดที่อยู่แถวเดียวกันและเห็นพร้อมกัน ต้องติดสีทั้งคู่');
assert(!chips[2].classList.has('on'), 'การ์ดที่เลื่อนผ่านไปแล้วต้องไม่ติดสี');
assert(!chips[5].classList.has('on'), 'การ์ดที่ยังอยู่ล่างจอต้องไม่ติดสี');

// เลื่อนต่อ — ไฮไลต์ต้องขยับตาม
tops = {};
tops[FCC_SECTIONS[6].id] = 120;
api.sync();
assert(chips[6].classList.has('on'), 'ไฮไลต์ต้องขยับตามการเลื่อน');
assert(!chips[3].classList.has('on') && !chips[4].classList.has('on'), 'ของเดิมต้องถูกปลดสี');

// อยู่ระหว่างการ์ด (ไม่มีใบไหนอยู่ในช่วงอ่าน) — ต้องยึดใบล่าสุดที่ผ่านไปแล้ว ไม่ใช่ปล่อยว่าง
tops = {};
tops[FCC_SECTIONS[1].id] = -900;
api.sync();
assert(chips[1].classList.has('on'), 'ต้องมีตัวที่ active เสมอ ไม่ปล่อยว่าง');

// ── ปุ่มลูกศร: โผล่เมื่อรายการล้น และหรี่เมื่อเลื่อนสุดทาง ───────────────────
jumpList.scrollWidth = 900; jumpList.clientWidth = 900; jumpList.scrollLeft = 0;
api.arrows();
assert(!prevBtn.classList.has('show') && !nextBtn.classList.has('show'),
  'รายการไม่ล้น ต้องไม่โชว์ปุ่มลูกศร');

jumpList.scrollWidth = 1800; jumpList.scrollLeft = 0;
api.arrows();
assert(prevBtn.classList.has('show') && nextBtn.classList.has('show'), 'รายการล้นต้องโชว์ปุ่ม');
assert(prevBtn.has('disabled'), 'อยู่ซ้ายสุดแล้ว ปุ่มซ้ายต้องหรี่');
assert(!nextBtn.has('disabled'), 'ยังเลื่อนขวาได้ ปุ่มขวาต้องกดได้');

api.scrollList(1);
assert(jumpList.scrollLeft > 0, 'กดลูกศรขวาแล้วต้องเลื่อนจริง');
jumpList.scrollLeft = 900;   // สุดขวา
api.arrows();
assert(nextBtn.has('disabled') && !prevBtn.has('disabled'), 'สุดขวาแล้วปุ่มขวาต้องหรี่ ปุ่มซ้ายกดได้');

console.log('Dashboard jump navigation checks passed');
