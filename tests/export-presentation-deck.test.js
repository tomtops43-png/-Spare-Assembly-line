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
    grabFn('xpDeckHighlights'), grabFn('xpDeckCut'), grabFn('xpDeckMonthLabel'), grabFn('xpDeckMonthlyIssue'), grabFn('xpDeckRangeText'), grabFn('xpBuildDeckModel')
  ]),
  '{ fccCrunchParts: fccCrunchParts, fccCrunchLogs: fccCrunchLogs, fccCrunchRequests: fccCrunchRequests, fccHealth: fccHealth, xpBuildDeckModel: xpBuildDeckModel, xpDeckMonthlyIssue: xpDeckMonthlyIssue, setRange: function(f, t) { xpFilters.from = f; xpFilters.to = t; }, xpDeckActionQueue: xpDeckActionQueue, xpDeckMoney: xpDeckMoney, setLang: function(l) { xpFilters.lang = l; }, setCustomer: function(v) { xpFilters.customerMode = v; } }'
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
[['en', false], ['en', true], ['zh', false], ['zh', true]].forEach(function(pair) {
  const cm = pair[1];
  mod.setLang(pair[0]);
  mod.setCustomer(cm);
  const m = mod.xpBuildDeckModel(S);
  const leaks = allText(m).filter(function(t) { return THAI.test(t); });
  assert.strictEqual(leaks.length, 0, (cm ? 'ฉบับลูกค้า' : 'ฉบับผู้บริหาร') + ' ภาษา ' + pair[0] + ' ยังมีภาษาไทยหลุด:\n  ' + leaks.join('\n  '));
  if (pair[0] === 'zh') assert(/[\u4e00-\u9fff]/.test(m.slides[1].title), 'ฉบับจีนต้องเป็นภาษาจีนจริง ไม่ใช่ตกไปใช้อังกฤษ');
  if (cm) assert(allText(m).join(' ').indexOf('THB') === -1, 'ฉบับลูกค้าภาษาอังกฤษก็ต้องไม่มีเงิน');
});
mod.setLang('en');
mod.setCustomer(false);
assert(/^THB /.test(mod.xpDeckMoney(1650)), 'ฉบับอังกฤษใช้ THB แทน ฿');
mod.setLang('th');
mod.setCustomer(false);

// ── มูลค่าการเบิกรายเดือน ─────────────────────────────────────────────
// ตั้ง "วันนี้" ตายตัว เพื่อให้รู้แน่ว่า 12 เดือนล่าสุดคือเดือนไหน
const NOW = new Date(2026, 9, 24); // 24 ต.ค. 2026
const mMaster = [
  { name: 'Breaker', model: 'BK-10', unit_price: 100, no: '1', __sheet: 'Lug&Screw' },
  { name: 'Belt', model: 'B-77', unit_price: 500, no: '2', __sheet: 'H9', image_main_url: 'https://drive.google.com/file/d/BELT123/view' },
  { name: 'Mystery', model: 'X', unit_price: '', no: '3', __sheet: 'H9' }
];
const mLogs = [
  { timestamp: '2026-10-03 08:00:00', type: 'Output', process: 'Lug&Screw', partName: 'Breaker', model: 'BK-10', qty: 3, category: 'Electrical' },
  { timestamp: '2026-10-05 08:00:00', type: 'Output', process: 'H9', partName: 'Belt', model: 'B-77', qty: 2, category: 'Mechanical' },
  { timestamp: '2026-10-06 08:00:00', type: 'Output', process: 'H9', partName: 'Mystery', model: 'X', qty: 4, category: 'Mechanical' },
  { timestamp: '2026-10-07 08:00:00', type: 'Input', process: 'H9', partName: 'Belt', model: 'B-77', qty: 50 },
  { timestamp: '2026-09-10 08:00:00', type: 'Output', process: 'H9', partName: 'Belt', model: 'B-77', qty: 1, category: 'Mechanical' },
  { timestamp: '2025-10-31 08:00:00', type: 'Output', process: 'H9', partName: 'Belt', model: 'B-77', qty: 9 }, // พ้น 12 เดือนแล้ว
  { timestamp: '2025-11-01 08:00:00', type: 'Output', process: 'H9', partName: 'Belt', model: 'B-77', qty: 1 }
];
const mMisc = [{ month: '2026-10', date: '2026-10-02', line: 'H9', total_amount: 250 }, { month: '2024-01', total_amount: 999 }];
mod.setLang('th');
mod.setCustomer(false);
mod.setRange('', '');
const M = mod.xpDeckMonthlyIssue(mLogs, mMaster, mMisc, NOW);
assert.strictEqual(M.months.length, 12, 'ค่าเริ่มต้นต้องเป็น 12 เดือนล่าสุด');
assert.strictEqual(M.months[0], '2025-11');
assert.strictEqual(M.months[11], '2026-10', 'เดือนสุดท้ายต้องเป็นเดือนปัจจุบัน');
assert.strictEqual(M.byLine['Lug&Screw'][11], 300, 'Breaker 3 × 100');
assert.strictEqual(M.byLine['H9'][11], 1000, 'Belt 2 × 500 (Mystery ไม่มีราคา = 0, รับเข้าไม่นับ)');
assert.strictEqual(M.misc[11], 250, 'ค่าสิ้นเปลืองแยกชุด');
assert.strictEqual(M.total[11], 1550, 'รวม = เบิก 1,300 + สิ้นเปลือง 250');
assert.strictEqual(M.byLine['H9'][10], 500, 'ก.ย. Belt 1 × 500');
assert.strictEqual(M.byLine['H9'][0], 500, 'พ.ย. ปีก่อน ยังอยู่ใน 12 เดือน');
assert.strictEqual(M.total.reduce(function(a, b) { return a + b; }, 0), 2550, 'ต.ค. ปีก่อน (พ้น 12 เดือน) และ misc ปี 2024 ต้องไม่ถูกนับ');
assert.strictEqual(M.unpricedItems, 1, 'ต้องบอกว่ามีรายการไม่มีราคา');
assert.strictEqual(M.unpricedQty, 4);
assert.strictEqual(M.thisMonth, 1550);
assert.strictEqual(M.prevMonth, 500);
assert.strictEqual(M.ytd, 2050, 'สะสมปี 2026 = ก.ย. 500 + ต.ค. 1,550');
assert.strictEqual(M.peak.value, 1550);
assert.strictEqual(M.top[0].name, 'Belt', 'อันดับ 1 ของเดือนล่าสุดเรียงตามมูลค่า');
assert.strictEqual(M.top[0].value, 1000);
assert.deepStrictEqual(M.cats.map(function(c) { return c.name; }), ['Mechanical', 'Electrical']);

const withM = mod.xpBuildDeckModel(S, M);
const ms = slideOf(withM, 'monthly');
assert(ms, 'ฉบับผู้บริหารต้องมีสไลด์มูลค่าการเบิกรายเดือน');
assert(kinds(withM).indexOf('monthly') === kinds(withM).indexOf('movement') + 1, 'วางต่อจากสไลด์เบิก/รับ 14 วัน');
assert(slideOf(withM, 'monthTop'), 'ต้องมีสไลด์เดือนล่าสุดเบิกอะไรไปบ้าง');
assert.strictEqual(ms.chart.labels.length, 12);
assert(ms.chart.series.some(function(se) { return se.name === 'ค่าสิ้นเปลือง'; }), 'ค่าสิ้นเปลืองต้องเป็นอีกชุดในกราฟ');
assert(/ยังไม่มีราคา 1 รายการ/.test(ms.sub), 'ต้องเตือนเรื่องรายการที่ไม่มีราคา');
assert.strictEqual(ms.table.rows[ms.table.rows.length - 1][0], 'รวม', 'แถวสุดท้ายของตารางเป็นยอดรวม');

// อันดับตามมูลค่าเบิกทั้งช่วง: Belt (H9) 2+1+1 ชิ้น × 500 = 2,000 > Breaker 300, ไม่มีราคาไม่ติดอันดับ
assert.deepStrictEqual(M.rankAll.map(function(x) { return x.name; }), ['Belt', 'Breaker'], 'เรียงจากมูลค่าเบิกมากไปน้อย ไม่รวมรายการไม่มีราคา');
assert.strictEqual(M.rankAll[0].value, 2000);
assert.strictEqual(M.rankAll[0].count, 3, 'นับจำนวนครั้งที่เบิก');
assert.strictEqual(M.issueSum, 2300, 'มูลค่าเบิกรวมทั้งช่วง (ไม่รวมค่าสิ้นเปลือง)');
const rankSlide = withM.slides.filter(function(sl) { return sl.title === 'อันดับอะไหล่ตามมูลค่าเบิก'; })[0];
assert(rankSlide, 'ต้องมีสไลด์จัดอันดับตามมูลค่าเบิก');
assert(kinds(withM).indexOf('monthTop') === withM.slides.indexOf(rankSlide) + 1, 'วางระหว่างสไลด์รายเดือนกับสไลด์เดือนล่าสุด');
assert.strictEqual(rankSlide.table.rows[0][1], 'Belt');
assert.strictEqual(rankSlide.table.rows[rankSlide.table.rows.length - 1][8], '100.0%', 'แถวสุดท้าย % สะสมต้องครบ 100% เมื่อมีไม่ถึง 15 รายการ');
assert(/คิดเป็น 100%/.test(rankSlide.sub), 'หัวสไลด์ต้องบอกสัดส่วนของ 15 อันดับแรก (มีแค่ 2 รายการ = 100%)');

// ชื่อ/รุ่นในตารางต้องเต็ม ไม่ถูกตัดเป็น "…" (เคยตัดที่ 40 ตัวอักษร ผู้ใช้อ่านไม่ออกว่าอะไหล่ตัวไหน)
const longName = 'MT PIN CWM Coil Winding Machine LP-2 Die Holder Assembly Left';
const longModel = 'ME-ACCE-PIN-CWM-HRC-0042-LEFT';
const MLong = mod.xpDeckMonthlyIssue([{ timestamp: '2026-10-03 08:00:00', type: 'Output', process: 'Coil Winding', partName: longName, model: longModel, qty: 2 }],
  [{ name: longName, model: longModel, unit_price: 1000 }], [], NOW);
const longDeck = mod.xpBuildDeckModel(S, MLong);
const longText = allText(longDeck);
assert(longText.indexOf(longName) > -1, 'ชื่ออะไหล่ยาวต้องแสดงเต็มในตาราง');
assert(longText.indexOf(longModel) > -1, 'รุ่นยาวต้องแสดงเต็มในตาราง');
longDeck.slides.forEach(function(sl) {
  (sl.table ? sl.table.rows : []).forEach(function(r) {
    r.forEach(function(c) { assert(String(c).indexOf('…') === -1, 'เซลล์ตารางสไลด์ "' + sl.title + '" ถูกตัดเป็น …: ' + c); });
  });
});

// ช่วงข้อมูล: ไม่ได้เลือก = 12 เดือนล่าสุด แสดงทั้งบนปกและสไลด์รายเดือน
assert(withM.slides[0].meta.some(function(t) { return /ช่วงข้อมูล: 12 เดือนล่าสุด/.test(t); }), 'ปกต้องบอกช่วงข้อมูล');
assert(/^ช่วงข้อมูล: 12 เดือนล่าสุด/.test(ms.sub), 'สไลด์รายเดือนต้องบอกช่วงข้อมูล');
// ยอดรวมบนหัวแท่ง = ผลรวมทุกไลน์ + ค่าสิ้นเปลือง ของเดือนนั้น
assert.deepStrictEqual(ms.chart.totals.values, M.total.map(Math.round), 'ยอดรวมบนหัวแท่งต้องตรงกับแถวรวม');
ms.chart.labels.forEach(function(_, i) {
  const sum = ms.chart.series.reduce(function(a, se) { return a + se.values[i]; }, 0);
  assert.strictEqual(sum, ms.chart.totals.values[i], 'แท่งซ้อนเดือนที่ ' + i + ' รวมได้เท่ายอดรวม');
});
// ไลน์ที่มูลค่าเป็น 0 ทั้งช่วง (เบิกแต่ไม่มีราคา) ต้องไม่มีแถว/สีในกราฟ
const MZero = mod.xpDeckMonthlyIssue(mLogs.concat([{ timestamp: '2026-10-08 08:00:00', type: 'Output', process: 'H1', partName: 'Mystery', model: 'X', qty: 1 }]), mMaster, mMisc, NOW);
const zeroSlide = slideOf(mod.xpBuildDeckModel(S, MZero), 'monthly');
assert(zeroSlide.chart.series.every(function(se) { return se.name !== 'H1'; }), 'ไลน์ที่ไม่มีมูลค่าต้องไม่โผล่ในกราฟ/ตาราง');
// เลือกช่วงวันที่เอง: แสดงวันที่ตามที่เลือก
mod.setRange('2026-09-01', '2026-09-30');
const rangeDeck = mod.xpBuildDeckModel(S, mod.xpDeckMonthlyIssue(mLogs, mMaster, mMisc, NOW));
assert(rangeDeck.slides[0].meta.some(function(t) { return /ช่วงข้อมูล: 1 ก\.ย\. 2569 – 30 ก\.ย\. 2569/.test(t); }), 'ปกต้องบอกวันที่ที่เลือก: ' + rangeDeck.slides[0].meta.join(' | '));
mod.setRange('', '');

// ภาคผนวกรูปอะไหล่: อยู่ท้ายเล่ม (หลังสไลด์ปิด) ชื่อในตารางอันดับลิงก์ไปหาได้ และมีลิงก์กลับ
const gal = slideOf(withM, 'gallery');
assert(gal, 'ต้องมีสไลด์ภาคผนวกรูปอะไหล่');
assert.strictEqual(withM.slides[withM.slides.length - 1], gal, 'ภาคผนวกต้องอยู่หน้าท้ายสุด');
assert.strictEqual(withM.slides[withM.slides.length - 2].kind, 'closing', 'ภาคผนวกต่อจากสไลด์ปิด');
assert.strictEqual(gal.id, 'gallery');
assert.strictEqual(rankSlide.linkTo, 'gallery', 'ชื่อในตารางอันดับต้องผูกลิงก์ไปภาคผนวก');
assert.strictEqual(rankSlide.linkCol, 1, 'ลิงก์อยู่ที่คอลัมน์ชื่ออะไหล่');
assert.strictEqual(gal.backTo, rankSlide.id, 'ภาคผนวกต้องมีลิงก์กลับไปตารางอันดับ');
assert.deepStrictEqual(gal.items.map(function(it) { return it.name; }), rankSlide.table.rows.map(function(r) { return r[1]; }), 'รูปเรียงตามอันดับเดียวกับตาราง');
assert.strictEqual(gal.items[0].imageUrl, 'https://drive.google.com/file/d/BELT123/view', 'ลิงก์รูปมาจาก image_main_url ใน Master');
assert.strictEqual(gal.items[1].imageUrl, '', 'ไม่มีรูปใน Master = ว่าง (การ์ดขึ้น "ไม่มีรูป")');
assert(/ภาคผนวก/.test(gal.title));
const runSrc1 = grabFn('xpRunPptDeck');
assert(/xpDeckGalleryImages\(model\)/.test(runSrc1), 'ต้องดึงรูปก่อนวาดสไลด์');
const galSrc = grabFn('xpDeckGalleryImages');
assert(/xpFetchPartImages\(/.test(galSrc), 'ใช้ตัวดึงรูปชุดเดียวกับ Excel');
assert(/\.catch\(function\(err\)/.test(galSrc), 'ดึงรูปไม่ได้ต้องยังได้ไฟล์');
// เคยค้างจริง: ทางเซิร์ฟเวอร์ไม่ตอบ หน้าจอค้างที่ "กำลังดึงรูปอะไหล่..." ไม่รู้จบ — ต้องมีเพดานเวลา
assert(/Promise\.race\(\[fetching, timeout\]\)/.test(galSrc), 'ดึงรูปต้องมีเพดานเวลา ไม่งั้นหน้าจอค้างได้');
assert(/var XP_DECK_IMAGE_TIMEOUT_MS = \d+;/.test(require('fs').readFileSync('index.html', 'utf8')), 'ต้องกำหนดเพดานเวลาดึงรูป');
const renderSrc = grabFn('xpRenderDeck');
assert(/hyperlink = \{ slide: opt\.linkSlide/.test(renderSrc), 'ตารางต้องผูกลิงก์กระโดดไปหน้าภาคผนวก');
assert(/hyperlink: \{ slide: slideNo\[s\.backTo\]/.test(renderSrc), 'ภาคผนวกต้องมีลิงก์กลับ');

// ฉบับลูกค้า: เป็นตัวเลขเงิน ต้องไม่มี
mod.setCustomer(true);
const custM = mod.xpBuildDeckModel(S, M);
assert(kinds(custM).indexOf('monthly') === -1 && kinds(custM).indexOf('monthTop') === -1, 'ฉบับลูกค้าต้องไม่มีสไลด์มูลค่าเบิก');
mod.setCustomer(false);

// ภาษาอังกฤษ: ชื่อเดือนและข้อความต้องไม่มีไทย
mod.setLang('en');
const MEn = mod.xpDeckMonthlyIssue(mLogs, mMaster, mMisc, NOW);
const enText = allText(mod.xpBuildDeckModel(S, MEn)).filter(function(t) { return THAI.test(t); });
assert.strictEqual(enText.length, 0, 'สไลด์มูลค่ารายเดือนฉบับอังกฤษมีไทยหลุด:\n  ' + enText.join('\n  '));
mod.setLang('zh');
const zhText = allText(mod.xpBuildDeckModel(S, mod.xpDeckMonthlyIssue(mLogs, mMaster, mMisc, NOW))).filter(function(t) { return THAI.test(t); });
assert.strictEqual(zhText.length, 0, 'สไลด์มูลค่ารายเดือนฉบับจีนมีไทยหลุด:\n  ' + zhText.join('\n  '));
mod.setLang('th');

// เลือกช่วงวันที่ในหน้า Export = ใช้ช่วงนั้นแทน 12 เดือน
mod.setRange('2026-09-01', '2026-09-30');
const MR = mod.xpDeckMonthlyIssue(mLogs, mMaster, mMisc, NOW);
assert.deepStrictEqual(MR.months, ['2026-09'], 'ใช้เฉพาะเดือนในช่วงที่เลือก');
assert.strictEqual(MR.total[0], 500);
mod.setRange('', '');

// ไม่มีการเบิกเลย = ไม่ใส่สไลด์ (ไม่ใช่กราฟศูนย์ทั้งแถว)
const MEmpty = mod.xpDeckMonthlyIssue([], mMaster, [], NOW);
assert.strictEqual(MEmpty.hasData, false);
assert(kinds(mod.xpBuildDeckModel(S, MEmpty)).indexOf('monthly') === -1, 'ไม่มีข้อมูลต้องไม่มีสไลด์รายเดือน');

// ตัวเรียกใช้: ฉบับลูกค้าไม่ต้องดึง log/master เพิ่ม และค่าสิ้นเปลืองล้มต้องไม่ทำให้ทั้งไฟล์ล้ม
const runSrc0 = grabFn('xpRunPptDeck');
assert(/xpFilters\.customerMode\) return \[S, null\]/.test(runSrc0), 'ฉบับลูกค้าไม่ต้องคำนวณมูลค่าเบิก');
assert(/xpFetchMiscExpenses\(\)\.catch/.test(runSrc0), 'ดึงค่าสิ้นเปลืองไม่ได้ต้องไปต่อได้');
assert(/xpDeckMonthlyIssue\(/.test(runSrc0), 'ต้องคำนวณมูลค่ารายเดือนก่อนสร้างสไลด์');

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
// ไลบรารีต้องมีสำเนาในเว็บเราเองเป็นตัวแรก (เครือข่ายโรงงานบล็อก CDN — เคยกดแล้วโหลดไม่ขึ้นจริง)
const urls = buildModule([grabVar('XP_PPTX_LIB_URLS')], 'XP_PPTX_LIB_URLS');
assert.strictEqual(urls[0], 'vendor/pptxgen.bundle.js', 'ตัวแรกต้องเป็นไฟล์ในเว็บเราเอง');
const fs = require('fs');
assert(fs.existsSync(urls[0]), 'ต้องมีไฟล์ ' + urls[0] + ' อยู่ใน repo');
const vendorHead = fs.readFileSync(urls[0], 'utf8').slice(0, 80);
const ver = (vendorHead.match(/PptxGenJS (\d+\.\d+\.\d+)/) || [])[1];
assert(ver, 'ไฟล์ vendor ต้องเป็น PptxGenJS bundle');
urls.slice(1).forEach(function(u) {
  assert(/^https:\/\//.test(u) && u.indexOf('@' + ver + '/') > -1, 'CDN สำรองต้องเป็นเวอร์ชันเดียวกับไฟล์ vendor (' + ver + '): ' + u);
});
assert(/cp -r vendor _site\/vendor/.test(fs.readFileSync('.github/workflows/deploy-pages.yml', 'utf8')), 'deploy ต้อง copy โฟลเดอร์ vendor ขึ้นเว็บด้วย');
assert(/return tryAt\(i \+ 1\)/.test(grabFn('xpLoadPptxLib')), 'โหลดตัวแรกไม่ได้ต้องลองตัวถัดไป');
assert(!/<script src="[^"]*pptxgen/.test(html), 'ห้ามโหลด PptxGenJS ตอนเปิดเว็บ (หนัก ~500KB) — โหลดตอนกดปุ่มเท่านั้น');
assert(/xpPptxLibPromise = null; throw err/.test(grabFn('xpLoadPptxLib')), 'โหลดล้มเหลวต้องล้าง promise ให้กดลองใหม่ได้');

// ── UI ────────────────────────────────────────────────────────────────
assert(html.indexOf('id="xpPptBtn"') > -1, 'ต้องมีปุ่ม PowerPoint ในระดับ 1');
const initSrc = html.slice(html.indexOf('(function initExportCenter()'));
assert(/xpEl\('xpPptBtn'\);\s*if \(pptBtn\) pptBtn\.addEventListener\('click', xpRunPptDeck\)/.test(initSrc), 'ปุ่มต้องผูกกับ xpRunPptDeck');
assert(/xpPptBtn/.test(grabFn('xpSetBusy')), 'ปุ่มต้องถูกปิดระหว่างทำงาน กันกดซ้ำ');

console.log('export-presentation-deck: OK (' + exec.slides.length + ' สไลด์ผู้บริหาร / ' + cust.slides.length + ' สไลด์ลูกค้า)');
