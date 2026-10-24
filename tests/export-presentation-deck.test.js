// PowerPoint นำเสนอผู้บริหาร / ลูกค้า (Export Center → ระดับ 1)
// สิ่งที่เทสต์นี้กันไว้:
//  1) ตัวเลขในสไลด์มาจากตัวคำนวณของ Dashboard (xpDashboardData → fccCrunch*) ไม่ใช่สูตรใหม่
//  2) โหมดลูกค้า: ไม่มีเงิน ไม่มีเรื่องภายใน (คะแนนสุขภาพ, PR ค้าง, ของที่ต้องสั่ง, Insight) หลุดไปถึงลูกค้า
//  3) ภาษาอังกฤษ: ไม่มีอักษรไทยที่ระบบสร้างหลุดเข้าไฟล์ (ข้อมูลตัวอย่างเป็นอังกฤษล้วน)
//  4) ปุ่ม/การผูก event/ไลบรารีที่โหลด ต้องมีครบ
const { html, grabFn, grabVar, grabScalar, buildModule, baseHelpers, assert } = require('./_export-extract');

const THAI = /[฀-๿]/;

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", category: "", txnType: "", groupBy: "month", lang: "th" };',
    'var currentUser = { username: "manager01", role: "admin" };',
    grabFn('fccN'), grabFn('fccPrice'), grabFn('fccInt'), grabFn('fccMoney'),
    grabFn('fccLineOf'), grabFn('fccPartKey'),
    grabFn('fccCrunchParts'), grabFn('fccCrunchLogs'), grabFn('fccCrunchRequests'), grabFn('fccHealth'),
    grabFn('xpDashKpiRows'), grabFn('xpDashLineTable'), grabFn('xpDashCategoryTable'),
    grabFn('xpDeckMoney'), grabFn('xpDeckPct'), grabFn('xpDeckTopUsage'), grabFn('xpDeckTopValue'),
    grabFn('xpDeckDeadStock'), grabFn('xpDeckRunningOut'), grabFn('xpDeckActionQueue'),
    grabFn('xpDeckHighlights'), grabFn('xpDeckCut'), grabFn('xpBuildDeckModel')
  ]),
  '{ fccCrunchParts: fccCrunchParts, fccCrunchLogs: fccCrunchLogs, fccCrunchRequests: fccCrunchRequests, fccHealth: fccHealth, xpBuildDeckModel: xpBuildDeckModel, xpDeckActionQueue: xpDeckActionQueue, xpDeckMoney: xpDeckMoney, setLang: function(l) { xpFilters.lang = l; }, setCustomer: function(v) { xpFilters.customerMode = v; } }'
);

// ── ข้อมูลตัวอย่าง (อังกฤษล้วน — อักษรไทยที่หลุดในโหมดอังกฤษจึงมาจากโค้ด) ──
const today = new Date();
function stamp(daysAgo) {
  const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysAgo, 9, 0, 0);
  const p = function(n) { return ('0' + n).slice(-2); };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' 09:00:00';
}
const parts = [
  { name: 'Breaker', model: 'BK-10', line: 'Lug&Screw', category: 'Electrical', stock: 0, min: 10, unit_price: 100 },
  { name: 'Bolt M8', model: 'M8', line: 'Lug&Screw', category: 'Hardware', stock: 200, min: 50, unit_price: 2 },
  { name: 'Belt', model: 'B-77', line: 'H9', category: 'Mechanical', stock: 1, min: 5, unit_price: 500 },
  { name: 'Fuse', model: 'FZ-1', line: 'H9', category: 'Electrical', stock: 20, min: 0, unit_price: 30 },
  { name: 'Spring', model: 'SP-2', line: 'Arc Chute', category: 'Mechanical', stock: 3, min: 4, unit_price: 50 }
];
const logs = [
  { timestamp: stamp(0), type: 'Output', process: 'Lug&Screw', partName: 'Breaker', model: 'BK-10', qty: 6 },
  { timestamp: stamp(0), type: 'Input', process: 'H9', partName: 'Belt', model: 'B-77', qty: 4 },
  { timestamp: stamp(3), type: 'Output', process: 'H9', partName: 'Belt', model: 'B-77', qty: 3 },
  { timestamp: stamp(10), type: 'Output', process: 'Lug&Screw', partName: 'Bolt M8', model: 'M8', qty: 50 },
  { timestamp: stamp(12), type: 'Output', process: 'Arc Chute', partName: 'Spring', model: 'SP-2', qty: 5 }
];
const reqs = [
  { request_id: 'RQ1', line: 'H9', status: 'Pending', requested_date: stamp(20).slice(0, 10) },
  { request_id: 'RQ2', line: 'H9', status: 'Pending', requested_date: stamp(2).slice(0, 10) },
  { request_id: 'RQ3', line: 'Lug&Screw', status: 'Approved', requested_date: stamp(5).slice(0, 10) }
];
const P = mod.fccCrunchParts(parts);
const L = mod.fccCrunchLogs(logs, 'all');
const R = mod.fccCrunchRequests(reqs, 'all');
const S = { P: P, L: L, R: R, H: mod.fccHealth(P, R) };

function allText(model) {
  const out = [model.title, model.footer];
  model.slides.forEach(function(s) {
    ['title', 'sub', 'badge'].forEach(function(k) { if (s[k]) out.push(s[k]); });
    (s.meta || []).forEach(function(t) { out.push(t); });
    (s.items || []).forEach(function(it) { out.push(typeof it === 'string' ? it : it.text); });
    (s.cards || []).concat(s.stats || []).forEach(function(c) { out.push(c.label, c.value, c.unit, c.note); });
    if (s.gauge) out.push(s.gauge.grade);
    if (s.gauge) s.gauge.notes.forEach(function(n) { out.push(n); });
    if (s.table) {
      s.table.headers.forEach(function(h) { out.push(h); });
      s.table.rows.forEach(function(r) { r.forEach(function(c) { out.push(c); }); });
    }
    if (s.chart) {
      s.chart.labels.forEach(function(l) { out.push(l); });
      s.chart.series.forEach(function(se) { out.push(se.name); });
    }
  });
  return out.filter(function(t) { return t !== undefined && t !== null && t !== ''; }).map(String);
}
function kinds(model) { return model.slides.map(function(s) { return s.kind; }); }
function slideOf(model, kind) { return model.slides.filter(function(s) { return s.kind === kind; })[0]; }

// ── ฉบับผู้บริหาร (ไทย) ────────────────────────────────────────────────
mod.setLang('th');
mod.setCustomer(false);
const exec = mod.xpBuildDeckModel(S);
['cover', 'kpis', 'lines', 'movement', 'risk', 'table', 'requests', 'bullets', 'closing'].forEach(function(k) {
  assert(kinds(exec).indexOf(k) > -1, 'ฉบับผู้บริหารต้องมีสไลด์ชนิด ' + k);
});
assert.strictEqual(exec.slides[0].kind, 'cover', 'สไลด์แรกต้องเป็นปก');
assert.strictEqual(exec.slides[exec.slides.length - 1].kind, 'closing', 'สไลด์สุดท้ายต้องเป็นสรุป/สิ่งที่ต้องทำต่อ');

// ตัวเลขต้องมาจาก fccCrunch* ตรงๆ
const kp = slideOf(exec, 'kpis');
function cardVal(label) { const c = kp.cards.filter(function(x) { return x.label === label; })[0]; assert(c, 'ต้องมีการ์ด ' + label); return c.value; }
assert.strictEqual(cardVal('อะไหล่ทั้งหมด'), String(P.total));
assert.strictEqual(cardVal('หมดสต็อก'), String(P.out));
assert.strictEqual(cardVal('ต่ำกว่า Min'), String(P.low));
assert.strictEqual(cardVal('Need PO'), String(P.need));
assert.strictEqual(cardVal('มูลค่าคงคลัง'), mod.xpDeckMoney(P.value), 'มูลค่าคงคลังต้องย่อเป็น ฿K/฿M');
assert.strictEqual(kp.gauge.score, S.H.score, 'เกจต้องเป็นคะแนนสุขภาพตัวเดียวกับ Dashboard');
assert(!kp.cards.some(function(c) { return c.label === 'Factory Health Score'; }), 'คะแนนสุขภาพแสดงเป็นเกจแล้ว ไม่ต้องซ้ำเป็นการ์ด');

const mv = slideOf(exec, 'movement');
assert.strictEqual(mv.chart.labels.length, 14, 'กราฟเคลื่อนไหวต้องมี 14 วัน');
assert.deepStrictEqual(mv.chart.series[0].values, L.days.map(function(d) { return d.out; }), 'เส้นเบิกออกต้องตรงกับ Dashboard');
assert.deepStrictEqual(mv.chart.series[1].values, L.days.map(function(d) { return d.in; }), 'เส้นรับเข้าต้องตรงกับ Dashboard');

const ln = slideOf(exec, 'lines');
ln.chart.labels.forEach(function(k, i) {
  const sum = ln.chart.series.reduce(function(a, s) { return a + s.values[i]; }, 0);
  assert.strictEqual(sum, P.lines[k].total, 'กราฟซ้อนของไลน์ ' + k + ' ต้องรวมได้เท่าจำนวน SKU');
});

// คิวสั่งซื้อ: หมด → วิกฤต, ≤ ครึ่ง Min → เร่งด่วน, ต่ำกว่า Min → ควรเติม (เหมือน Action Queue)
const q = mod.xpDeckActionQueue(S);
assert.strictEqual(q.crit, 1, 'Breaker หมด = วิกฤต');
assert.strictEqual(q.high, 1, 'Belt 1/5 = เร่งด่วน');
assert.strictEqual(q.med, 1, 'Spring 3/4 = ควรเติม');
assert.strictEqual(q.rows[0].name, 'Breaker');

// ── ฉบับลูกค้า ────────────────────────────────────────────────────────
mod.setCustomer(true);
const cust = mod.xpBuildDeckModel(S);
['risk', 'requests', 'bullets'].forEach(function(k) {
  assert(kinds(cust).indexOf(k) === -1, 'ฉบับลูกค้าต้องไม่มีสไลด์ภายใน: ' + k);
});
assert.strictEqual(slideOf(cust, 'kpis').gauge, null, 'ฉบับลูกค้าต้องไม่มีคะแนนสุขภาพ (CRITICAL บนจอลูกค้าไม่ควรเห็น)');
const custText = allText(cust).join('\n');
assert(custText.indexOf('฿') === -1 && custText.indexOf('THB') === -1, 'ฉบับลูกค้าต้องไม่มีจำนวนเงิน');
['มูลค่า', 'หมดสต็อก', 'ต่ำกว่า Min', 'Need PO', 'PR', 'Dead Stock', 'ต้องสั่ง'].forEach(function(w) {
  assert(custText.indexOf(w) === -1, 'ฉบับลูกค้าต้องไม่มีคำว่า "' + w + '"');
});
assert(cust.slides[0].badge, 'ปกฉบับลูกค้าต้องมีป้ายบอกว่าเป็นฉบับลูกค้า');
assert(/ลูกค้า/.test(cust.footer), 'footer ทุกหน้าต้องบอกว่าเป็นฉบับลูกค้า');

// ── ภาษาอังกฤษ: ไม่มีอักษรไทยหลุด ─────────────────────────────────────
[false, true].forEach(function(cm) {
  mod.setLang('en');
  mod.setCustomer(cm);
  const m = mod.xpBuildDeckModel(S);
  const leaks = allText(m).filter(function(t) { return THAI.test(t); });
  assert.strictEqual(leaks.length, 0, (cm ? 'ฉบับลูกค้า' : 'ฉบับผู้บริหาร') + ' ภาษาอังกฤษยังมีภาษาไทยหลุด:\n  ' + leaks.join('\n  '));
  if (cm) assert(allText(m).join(' ').indexOf('THB') === -1, 'ฉบับลูกค้าภาษาอังกฤษก็ต้องไม่มีเงิน');
});
mod.setLang('en');
mod.setCustomer(false);
assert(/^THB /.test(mod.xpDeckMoney(1650)), 'ฉบับอังกฤษใช้ THB แทน ฿');
mod.setLang('th');
mod.setCustomer(false);

// ── ไม่มีข้อมูล log / PR ก็ต้องสร้างได้ ไม่ throw ──────────────────────
const empty = mod.xpBuildDeckModel({ P: mod.fccCrunchParts([]), L: mod.fccCrunchLogs([], 'all'), R: mod.fccCrunchRequests([], 'all'), H: mod.fccHealth(mod.fccCrunchParts([]), mod.fccCrunchRequests([], 'all')) });
assert(empty.slides.length >= 4, 'ข้อมูลว่างต้องยังได้สไลด์หลัก');
assert(slideOf(empty, 'requests').chart === null, 'ไม่มี PR ต้องไม่วาดกราฟว่าง');

// ── ตัวเรียกใช้: ต้องดึงข้อมูลชุดเดียวกับ Dashboard + บันทึกทะเบียน ──────
const runSrc = grabFn('xpRunPptDeck');
assert(/xpDashboardData\(\)/.test(runSrc), 'ต้องใช้ xpDashboardData() — ตัวเลขจะได้ตรงกับหน้า Dashboard');
assert(/xpLogExport\('presentation-deck'/.test(runSrc), 'ต้องบันทึกทะเบียนการส่งออก (ไล่ได้ว่าใครส่งอะไรให้ลูกค้า)');
assert(/\.pptx'/.test(runSrc), 'ไฟล์ต้องเป็น .pptx');
assert(/xpFileSuffix\(\)/.test(runSrc), 'ชื่อไฟล์ต้องบอก customer / EN เหมือนไฟล์อื่น');
assert(/xpSetBusy\(false\)/.test(runSrc), 'ต้องปลดสถานะ busy เสมอแม้ล้มเหลว');

// ไลบรารีโหลดตอนกดเท่านั้น จาก CDN ที่ระบุเวอร์ชันแน่นอน
assert(/var XP_PPTX_LIB_URL = 'https:\/\/cdnjs\.cloudflare\.com\/ajax\/libs\/pptxgenjs\/\d+\.\d+\.\d+\/pptxgen\.bundle\.js';/.test(html), 'ต้องโหลด PptxGenJS แบบระบุเวอร์ชัน');
assert(!/<script src="[^"]*pptxgen/.test(html), 'ห้ามโหลด PptxGenJS ตอนเปิดเว็บ (หนัก ~500KB) — โหลดตอนกดปุ่มเท่านั้น');
assert(/xpPptxLibPromise = null; reject/.test(grabFn('xpLoadPptxLib')), 'โหลดล้มเหลวต้องล้าง promise ให้กดลองใหม่ได้');

// ── UI ────────────────────────────────────────────────────────────────
assert(html.indexOf('id="xpPptBtn"') > -1, 'ต้องมีปุ่ม PowerPoint ในระดับ 1');
const initSrc = html.slice(html.indexOf('(function initExportCenter()'));
assert(/xpEl\('xpPptBtn'\);\s*if \(pptBtn\) pptBtn\.addEventListener\('click', xpRunPptDeck\)/.test(initSrc), 'ปุ่มต้องผูกกับ xpRunPptDeck');
assert(/xpPptBtn/.test(grabFn('xpSetBusy')), 'ปุ่มต้องถูกปิดระหว่างทำงาน กันกดซ้ำ');

console.log('export-presentation-deck: OK (' + exec.slides.length + ' สไลด์ผู้บริหาร / ' + cust.slides.length + ' สไลด์ลูกค้า)');
