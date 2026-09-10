// รายงานวิเคราะห์ 10 ตัวของ Export Center — คำนวณจากข้อมูลจริงเท่านั้น
// เทสต์นี้ป้อนข้อมูลตัวอย่างที่รู้คำตอบล่วงหน้า แล้วเช็คว่าเลขที่ออกมาตรง
// เจตนาสำคัญที่กันไว้: ห้ามเติมศูนย์แทน "ไม่รู้" และห้ามข้ามไลน์ตอนจับคู่อะไหล่
const { grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", groupBy: "month" };',
    grabVar('XP_DICTIONARY'),
    grabVar('XP_REPORTS'),
    grabFn('xpBuildKpis'),
    grabFn('xpKpiTable'),
    grabFn('xpBuildTrend'),
    grabFn('xpBuildLineSummary'),
    grabFn('xpBuildDictionary')
  ]),
  '{ XP_REPORTS: XP_REPORTS, xpBuildKpis: xpBuildKpis, xpBuildTrend: xpBuildTrend, xpBuildLineSummary: xpBuildLineSummary, xpBuildDictionary: xpBuildDictionary, setFilters: function(f) { Object.keys(f).forEach(function(k) { xpFilters[k] = f[k]; }); } }'
);

function report(key) {
  const r = mod.XP_REPORTS.filter(function(x) { return x.key === key; })[0];
  assert(r, 'ไม่พบรายงาน ' + key);
  return r;
}

assert.strictEqual(mod.XP_REPORTS.length, 10, 'ต้องมีรายงานวิเคราะห์ 10 ตัว');
mod.XP_REPORTS.forEach(function(r) {
  assert(r.key && r.name && r.desc, 'รายงานต้องมี key/name/desc');
  assert(Array.isArray(r.needs) && r.needs.length, r.key + ' ต้องระบุว่าใช้ข้อมูลชุดไหน');
  assert(typeof r.build === 'function', r.key + ' ต้องมีตัวคำนวณ');
});

// ── ข้อมูลตัวอย่าง ────────────────────────────────────────────────────
const master = [
  { __sheet: 'Lug&Screw', no: '1', name: 'เบรกเกอร์', model: 'BK-10', line: 'Lug&Screw', category: 'ไฟฟ้า', stock: 4, min: 10, max: 40, unit: 'PCS', unit_price: 100 },
  { __sheet: 'Lug&Screw', no: '2', name: 'น็อต M8', model: 'M8', line: 'Lug&Screw', category: 'ฮาร์ดแวร์', stock: 200, min: 50, max: 300, unit: 'PCS', unit_price: 2 },
  { __sheet: 'H9', no: '3', name: 'สายพาน', model: 'B-77', line: 'H9', category: 'กลไก', stock: 6, min: 2, max: 10, unit: 'PCS', unit_price: 500 },
  // ของที่ไม่เคยเบิก + มีของเหลือ = dead stock
  { __sheet: 'H9', no: '4', name: 'ฟิวส์เก่า', model: 'FZ-1', line: 'H9', category: 'ไฟฟ้า', stock: 20, min: 0, max: 0, unit: 'PCS', unit_price: 30 },
  // ของที่ยังไม่ลงราคา — ต้องไม่โผล่ใน ABC แต่ต้องยังนับจำนวนชิ้นได้
  { __sheet: 'H9', no: '5', name: 'ยางรอง', model: 'RB-2', line: 'H9', category: 'กลไก', stock: 5, min: 1, max: 8, unit: 'PCS', unit_price: '' }
];
const logs = [
  { timestamp: '2026-09-02 08:00:00', type: 'Output', process: 'Lug&Screw', category: 'ไฟฟ้า', partName: 'เบรกเกอร์', model: 'BK-10', qty: 6, by: 'somchai', machine: 'LS-10', stockBefore: 10, stockAfter: 4 },
  { timestamp: '2026-09-03 08:00:00', type: 'Output', process: 'Lug&Screw', category: 'ฮาร์ดแวร์', partName: 'น็อต M8', model: 'M8', qty: 50, by: 'somchai', machine: 'LS-10', stockBefore: 250, stockAfter: 200 },
  { timestamp: '2026-09-04 08:00:00', type: 'Output', process: 'H9', category: 'กลไก', partName: 'สายพาน', model: 'B-77', qty: 2, by: 'somsak', machine: 'H9-1', stockBefore: 8, stockAfter: 6 },
  { timestamp: '2026-09-05 08:00:00', type: 'Input', process: 'H9', category: 'กลไก', partName: 'สายพาน', model: 'B-77', qty: 4, by: 'admin', machine: '', stockBefore: 4, stockAfter: 8 },
  { timestamp: '2026-09-06 08:00:00', type: 'Output', process: 'H9', category: 'กลไก', partName: 'ยางรอง', model: 'RB-2', qty: 3, by: 'somsak', machine: 'H9-1', stockBefore: 8, stockAfter: 5 }
];

// ── Top 50 อะไหล่ที่เบิกเยอะสุด ────────────────────────────────────────
const top = report('topIssue').build({ logs: logs, master: master, failed: {} });
const topRows = top.table.rows;
const hName = top.table.headers.indexOf('ชื่ออะไหล่');
const hQty = top.table.headers.indexOf('ยอดเบิก (ชิ้น)');
const hVal = top.table.headers.indexOf('มูลค่ารวม');
assert.strictEqual(topRows.length, 4, 'ต้องมี 4 อะไหล่ที่มีการเบิก (ไม่นับรายการรับเข้า)');
// เรียงตามมูลค่า: เบรกเกอร์ 6×100=600, น็อต 50×2=100, สายพาน 2×500=1000, ยางรอง 3×0=0
assert.strictEqual(topRows[0][hName], 'สายพาน', 'อันดับ 1 ต้องเป็นตัวที่มูลค่าสูงสุด (1,000 บาท)');
assert.strictEqual(topRows[0][hVal], 1000);
assert.strictEqual(topRows[1][hName], 'เบรกเกอร์');
assert.strictEqual(topRows[1][hVal], 600);
// ของที่ยังไม่ลงราคา: มูลค่า 0 แต่จำนวนชิ้นต้องถูกต้อง ไม่ใช่หายไปจากรายงาน
const rubber = topRows.filter(function(r) { return r[hName] === 'ยางรอง'; })[0];
assert(rubber, 'อะไหล่ที่ยังไม่ลงราคาต้องยังอยู่ในรายงาน');
assert.strictEqual(rubber[hQty], 3, 'จำนวนชิ้นต้องถูกแม้ไม่มีราคา');
assert.strictEqual(rubber[hVal], 0);
// เครื่องจักรที่ใช้ต้องมาจาก Log จริง
const hMachine = top.table.headers.indexOf('เครื่องที่ใช้');
assert.strictEqual(topRows[1][hMachine], 'LS-10');

// ── Dead Stock ────────────────────────────────────────────────────────
const dead = report('deadStock').build({ master: master, logs: logs, failed: {} });
const dName = dead.table.headers.indexOf('ชื่ออะไหล่');
const dVal = dead.table.headers.indexOf('มูลค่าที่จม');
assert.strictEqual(dead.table.rows.length, 1, 'มีของเดียวที่มีของเหลือแต่ไม่เคยเบิก');
assert.strictEqual(dead.table.rows[0][dName], 'ฟิวส์เก่า');
assert.strictEqual(dead.table.rows[0][dVal], 600, '20 ชิ้น × 30 บาท = 600');

// ── ABC Analysis ──────────────────────────────────────────────────────
const abc = report('abc').build({ logs: logs, master: master, failed: {} });
const aGrade = abc.table.headers.indexOf('ชั้น');
const aName = abc.table.headers.indexOf('ชื่ออะไหล่');
assert.strictEqual(abc.table.rows.length, 3, 'ของที่ยังไม่ลงราคาต้องไม่ถูกจัดชั้น (จัดไม่ได้จริง)');
assert.strictEqual(abc.table.rows[0][aName], 'สายพาน');
assert.strictEqual(abc.table.rows[0][aGrade], 'A', 'ตัวที่กินมูลค่ามากสุดต้องเป็นชั้น A');
// รวม 1700 → สายพาน 1000 (58.8% สะสม) = A, เบรกเกอร์ 600 (94.1% สะสม) = B, น็อต 100 (100%) = C
assert.strictEqual(abc.table.rows[1][aGrade], 'B');
assert.strictEqual(abc.table.rows[2][aGrade], 'C');
assert(/ยังไม่ลงราคา/.test(abc.note), 'ต้องบอกผู้อ่านว่ามีของที่จัดชั้นไม่ได้กี่รายการ');

// ── Stock Turnover ────────────────────────────────────────────────────
const turn = report('turnover').build({ logs: logs, master: master, failed: {} });
const tName = turn.table.headers.indexOf('ชื่ออะไหล่');
const tTurns = turn.table.headers.indexOf('รอบหมุน (เท่า)');
const tVerdict = turn.table.headers.indexOf('ประเมิน');
const nut = turn.table.rows.filter(function(r) { return r[tName] === 'น็อต M8'; })[0];
assert.strictEqual(nut[tTurns], 0.25, 'เบิก 50 คงเหลือ 200 = 0.25 รอบ');
assert.strictEqual(nut[tVerdict], 'หมุนช้า (เก็บมากกว่าที่ใช้)');
const fuse = turn.table.rows.filter(function(r) { return r[tName] === 'ฟิวส์เก่า'; })[0];
assert.strictEqual(fuse[tVerdict], 'ไม่เคยเบิก');
assert(/ไม่ใช่ค่าเฉลี่ยคงคลัง/.test(turn.note), 'ต้องบอกตรงๆ ว่าตัวหารเป็นยอดคงเหลือปัจจุบัน ไม่ใช่ค่าเฉลี่ย — ระบบไม่มี snapshot รายวัน');

// ── Min/Max Health ────────────────────────────────────────────────────
const mm = report('minmax').build({ logs: logs, master: master, failed: {} });
const mName = mm.table.headers.indexOf('ชื่ออะไหล่');
const mHits = mm.table.headers.indexOf('ครั้งที่ตกใต้ Min');
const mLow = mm.table.headers.indexOf('ยอดต่ำสุดที่เคยเหลือ');
const breaker = mm.table.rows.filter(function(r) { return r[mName] === 'เบรกเกอร์'; })[0];
assert(breaker, 'เบรกเกอร์ต้องติดรายงานเพราะเบิกแล้วยอดเหลือ 4 ต่ำกว่า Min 10');
assert.strictEqual(breaker[mHits], 1, 'นับจาก Stock After ในชีต Log ซึ่งเป็นเหตุการณ์จริง');
assert.strictEqual(breaker[mLow], 4);
assert(/Stock After/.test(mm.note), 'ต้องบอกที่มาว่าใช้คอลัมน์ Stock After ไม่ใช่ประเมินจากยอดปัจจุบัน');

// ── ต้นทุนอะไหล่ต่อมูลค่าผลิต ──────────────────────────────────────────
const cost = report('costVsProduction').build({
  logs: logs, master: master,
  miscExpenses: [{ month: '2026-09', date: '2026-09-03', line: 'Lug&Screw', total_amount: 300 }],
  productionVolume: [{ month: '2026-09', line: 'Lug&Screw', actual_qty: 1000, production_value: 10000 }],
  costConfig: [{ line: 'Lug&Screw', target_pct: 12 }],
  failed: {}
});
const cLine = cost.table.headers.indexOf('ไลน์');
const cPct = cost.table.headers.indexOf('% จริง');
const cResult = cost.table.headers.indexOf('ผล');
const lsRow = cost.table.rows.filter(function(r) { return r[cLine] === 'Lug&Screw'; })[0];
// ค่าอะไหล่ = 600 + 100 = 700, สิ้นเปลือง 300 → 1000/10000 = 10%
assert.strictEqual(lsRow[cPct], 10, 'ค่าอะไหล่ 700 + สิ้นเปลือง 300 ÷ มูลค่าผลิต 10,000 = 10%');
assert.strictEqual(lsRow[cResult], 'อยู่ในเป้า', 'เป้า 12% จริง 10% = อยู่ในเป้า');
// ไลน์ที่ไม่มีมูลค่าผลิต ต้องไม่โชว์ % ปลอม
const h9Row = cost.table.rows.filter(function(r) { return r[cLine] === 'H9'; })[0];
assert.strictEqual(h9Row[cPct], '', 'ไม่มีมูลค่าผลิต = ต้องเว้นว่าง ไม่ใช่ 0%');
assert.strictEqual(h9Row[cResult], 'ไม่มีมูลค่าผลิตในเดือนนี้');

// ── Lead Time ─────────────────────────────────────────────────────────
const lead = report('leadTime').build({
  orderRequests: [
    { request_id: 'RQ1', line: 'H9', item_name: 'สายพาน', status: 'Approved', requested_date: '2026-09-01', approved_date: '2026-09-04', request_qty: 2 },
    { request_id: 'RQ2', line: 'H9', item_name: 'น็อต', status: 'Pending', requested_date: '2026-09-02', approved_date: '', request_qty: 5 }
  ],
  purchaseHistory: [
    { 'History ID': 'PH1', Line: 'H9', 'Part Name': 'สายพาน', Status: 'Received', 'Requested Date': '2026-09-01', 'Ordered Date': '2026-09-05', 'Received Date': '2026-09-15' }
  ],
  failed: {}
});
const leadMap = {};
lead.table.rows.forEach(function(r) { leadMap[r[0]] = r[1]; });
assert.strictEqual(leadMap['จำนวนคำขอที่วัดเวลารออนุมัติได้'], 1, 'คำขอที่ยังไม่อนุมัติต้องไม่ถูกนับ (ไม่เดาแทน)');
assert.strictEqual(leadMap['เวลารออนุมัติ — ค่ากลาง'], 3, '1 ก.ย. ถึง 4 ก.ย. = 3 วัน');
assert.strictEqual(leadMap['ขอ → รับของ — ค่ากลาง'], 14, '1 ก.ย. ถึง 15 ก.ย. = 14 วัน');
assert.strictEqual(leadMap['สั่ง → รับของ — ค่ากลาง'], 10);
assert(lead.extraSheets && lead.extraSheets.length, 'ต้องมีชีตรายรายการแนบไปด้วย');

// ── Price Variance ────────────────────────────────────────────────────
const pv = report('priceVariance').build({
  purchaseHistory: [
    { 'Part Name': 'เบรกเกอร์', 'Model / Part No.': 'BK-10', Line: 'Lug&Screw', 'Unit Price': 100, 'Qty Ordered': 5, 'Total Amount': 500, Brand: 'ABB' },
    { 'Part Name': 'เบรกเกอร์', 'Model / Part No.': 'BK-10', Line: 'Lug&Screw', 'Unit Price': 150, 'Qty Ordered': 2, 'Total Amount': 300, Brand: 'Schneider' },
    { 'Part Name': 'น็อต M8', 'Model / Part No.': 'M8', Line: 'Lug&Screw', 'Unit Price': 2, 'Qty Ordered': 100, 'Total Amount': 200, Brand: 'local' }
  ],
  failed: {}
});
assert.strictEqual(pv.table.rows.length, 1, 'ต้องนับเฉพาะของที่ซื้อ 2 ครั้งขึ้นไป');
const pMin = pv.table.headers.indexOf('ราคาต่ำสุด');
const pMax = pv.table.headers.indexOf('ราคาสูงสุด');
const pSpread = pv.table.headers.indexOf('ส่วนต่าง (%)');
assert.strictEqual(pv.table.rows[0][pMin], 100);
assert.strictEqual(pv.table.rows[0][pMax], 150);
assert.strictEqual(pv.table.rows[0][pSpread], 50, '(150-100)/100 = 50%');
assert(/สเปกต่างกัน/.test(pv.note), 'ต้องเตือนว่าส่วนต่างสูงอาจมาจากสเปกต่างกัน ไม่ใช่ซื้อแพงเสมอ');

// ── ความแม่นยำการนับสต็อก ─────────────────────────────────────────────
const acc = report('countAccuracy').build({
  stockCount: [
    { session_id: 'S1', month: '2026-09', line: 'H9', round_no: 1, total_items: 100, matched: 95, diff_count: 5, adjusted_count: 5, status: 'approved' },
    { session_id: 'S2', month: '2026-09', line: 'H9', round_no: 2, total_items: 100, matched: 99, diff_count: 1, adjusted_count: 1, status: 'approved' },
    { session_id: 'S3', month: '2026-09', line: 'Lug&Screw', round_no: 1, total_items: 50, matched: 40, diff_count: 10, adjusted_count: 0, status: 'pending_approval' }
  ],
  failed: {}
});
const acLine = acc.table.headers.indexOf('ไลน์');
const acPct = acc.table.headers.indexOf('ความแม่น (%)');
const h9acc = acc.table.rows.filter(function(r) { return r[acLine] === 'H9'; })[0];
assert.strictEqual(h9acc[acPct], 97, '(95+99)/200 = 97%');
const lsacc = acc.table.rows.filter(function(r) { return r[acLine] === 'Lug&Screw'; })[0];
assert.strictEqual(lsacc[acPct], 80);

// ── อายุอะไหล่จากทะเบียนชิ้น ────────────────────────────────────────────
const life = report('partLife').build({
  partTags: [
    { tag_no: 'T1', part_name: 'ใบมีด', model: 'BL-1', line: 'Lug&Screw', machine: 'LS-10', installed_at: '2026-01-01', removed_at: '2026-03-02' },
    { tag_no: 'T2', part_name: 'ใบมีด', model: 'BL-1', line: 'Lug&Screw', machine: 'LS-10', installed_at: '2026-03-02', removed_at: '2026-05-01' },
    // ชิ้นที่ยังอยู่ในเครื่อง (ไม่มีวันถอด) ต้องไม่ถูกนับ เพราะยังไม่รู้อายุจริง
    { tag_no: 'T3', part_name: 'ใบมีด', model: 'BL-1', line: 'Lug&Screw', machine: 'LS-10', installed_at: '2026-05-01', removed_at: '' }
  ],
  failed: {}
});
const lCount = life.table.headers.indexOf('จำนวนชิ้นที่วัดได้');
const lAvg = life.table.headers.indexOf('อายุเฉลี่ย (วัน)');
assert.strictEqual(life.table.rows.length, 1);
assert.strictEqual(life.table.rows[0][lCount], 2, 'ชิ้นที่ยังอยู่ในเครื่องต้องไม่ถูกนับ');
assert.strictEqual(life.table.rows[0][lAvg], 60, '(60 + 60) / 2 = 60 วัน');
assert(/ยังอยู่ในเครื่องไม่ถูกนับ/.test(life.note), 'ต้องบอกว่าไม่นับชิ้นที่ยังอยู่ในเครื่อง');

console.log('export-analytics-reports: OK (10 รายงาน)');
