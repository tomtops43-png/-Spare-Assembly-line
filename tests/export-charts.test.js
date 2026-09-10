// กราฟของ Export Center — ตารางเปล่าอ่านยาก ทุกรายงานจึงต้องมีกราฟที่วาดจากตารางจริง
// จุดที่กันไว้:
//  1) ทุกรายงานทั้ง 10 ตัวต้องมีสเปกกราฟ และต้องหาคอลัมน์ที่อ้างไว้เจอจริง
//     (ถ้าใครเปลี่ยนชื่อหัวคอลัมน์ กราฟจะเงียบหายไปโดยไม่มีใครรู้ — เทสต์นี้จับได้)
//  2) กราฟต้องใช้ตารางเดียวกับที่ลงไฟล์ ไม่มีชุดข้อมูลแยกให้เลขไม่ตรงกัน
//  3) โหมดลูกค้าตัดคอลัมน์ราคาออก กราฟต้องถอยไปใช้จำนวนชิ้น ไม่ใช่หายไปทั้งกราฟ
//  4) ค่าว่าง (คำนวณไม่ได้) ต้องไม่ถูกวาดเป็นแท่ง 0
//  5) Chart.js โหลดไม่ได้ต้องยังเห็นตารางได้ปกติ ไม่ทำให้รายงานพัง
const { html, grabFn, grabVar, buildModule, baseHelpers, assert } = require('./_export-extract');

const mod = buildModule(
  baseHelpers().concat([
    'var xpFilters = { customerMode: false, from: "", to: "", line: "", category: "", txnType: "", groupBy: "month", lang: "th" };',
    grabVar('XP_CHART_COLORS'),
    grabVar('XP_CHART_SPECS'),
    grabVar('XP_DICTIONARY'),
    grabVar('XP_REPORTS'),
    grabFn('xpColIdx'),
    grabFn('xpChartSeries'),
    grabFn('xpChartSeriesFromTable'),
    grabFn('xpChartConfig'),
    grabFn('xpChartHeight'),
    grabFn('xpBuildTrend'),
    grabFn('xpBuildLineSummary'),
    grabFn('xpSummaryChartSeries')
  ]),
  '{ XP_REPORTS: XP_REPORTS, XP_CHART_SPECS: XP_CHART_SPECS, xpChartSeries: xpChartSeries, xpChartConfig: xpChartConfig, xpChartHeight: xpChartHeight, xpBuildTrend: xpBuildTrend, xpBuildLineSummary: xpBuildLineSummary, xpSummaryChartSeries: xpSummaryChartSeries, setLang: function(l) { xpFilters.lang = l; }, setCustomer: function(v) { xpFilters.customerMode = v; } }'
);

// ── ข้อมูลตัวอย่าง (มีทั้งของที่มีราคาและไม่มีราคา, ทั้งเบิกและรับเข้า) ──
const master = [
  { __sheet: 'LugScrew', no: '1', name: 'เบรกเกอร์', model: 'BK-10', line: 'Lug&Screw', category: 'ไฟฟ้า', location: 'A-1', stock: 4, min: 10, max: 40, unit: 'PCS', unit_price: 100 },
  { __sheet: 'LugScrew', no: '2', name: 'น็อต M8', model: 'M8', line: 'Lug&Screw', category: 'ฮาร์ดแวร์', location: 'A-2', stock: 200, min: 50, max: 300, unit: 'PCS', unit_price: 2 },
  { __sheet: 'H9', no: '3', name: 'สายพาน', model: 'B-77', line: 'H9', category: 'กลไก', location: 'B-1', stock: 6, min: 2, max: 10, unit: 'PCS', unit_price: 500 },
  { __sheet: 'H9', no: '4', name: 'ฟิวส์เก่า', model: 'FZ-1', line: 'H9', category: 'ไฟฟ้า', location: 'B-2', stock: 20, min: 0, max: 0, unit: 'PCS', unit_price: 30 }
];
const logs = [
  { timestamp: '2026-08-02 08:00:00', type: 'Output', process: 'Lug&Screw', category: 'ไฟฟ้า', partName: 'เบรกเกอร์', model: 'BK-10', qty: 6, by: 'a', machine: 'LS-10', stockBefore: 10, stockAfter: 4 },
  { timestamp: '2026-09-03 08:00:00', type: 'Output', process: 'Lug&Screw', category: 'ฮาร์ดแวร์', partName: 'น็อต M8', model: 'M8', qty: 50, by: 'a', machine: 'LS-10', stockBefore: 250, stockAfter: 200 },
  { timestamp: '2026-09-04 08:00:00', type: 'Output', process: 'H9', category: 'กลไก', partName: 'สายพาน', model: 'B-77', qty: 2, by: 'b', machine: 'H9-1', stockBefore: 8, stockAfter: 6 },
  { timestamp: '2026-09-05 08:00:00', type: 'Input', process: 'H9', category: 'กลไก', partName: 'สายพาน', model: 'B-77', qty: 4, by: 'c', machine: '', stockBefore: 4, stockAfter: 8 }
];
const ctx = {
  master: master,
  logs: logs,
  orderRequests: [{ request_id: 'RQ1', line: 'H9', item_name: 'สายพาน', model: 'B-77', status: 'Approved', requested_date: '2026-09-01', approved_date: '2026-09-04', request_qty: 2 }],
  purchaseHistory: [
    { 'History ID': 'PH1', Line: 'H9', 'Part Name': 'สายพาน', 'Model / Part No.': 'B-77', Status: 'Received', 'Requested Date': '2026-09-01', 'Ordered Date': '2026-09-05', 'Received Date': '2026-09-09', 'Unit Price': 500, 'Qty Ordered': 2, 'Total Amount': 1000, Brand: 'Gates' },
    { 'History ID': 'PH2', Line: 'H9', 'Part Name': 'สายพาน', 'Model / Part No.': 'B-77', Status: 'Received', 'Requested Date': '2026-09-02', 'Ordered Date': '2026-09-06', 'Received Date': '2026-09-10', 'Unit Price': 650, 'Qty Ordered': 1, 'Total Amount': 650, Brand: 'Optibelt' }
  ],
  miscExpenses: [{ expense_id: 'EX1', month: '2026-09', date: '2026-09-03', line: 'Lug&Screw', category: 'ฮาร์ดแวร์', item_name: 'ประแจ', total_amount: 300 }],
  productionVolume: [
    { month: '2026-09', line: 'Lug&Screw', actual_qty: 1000, production_value: 10000 },
    { month: '2026-09', line: 'H9', actual_qty: 500, production_value: 0 }
  ],
  costConfig: [{ line: 'Lug&Screw', target_pct: 12 }],
  stockCount: [{ session_id: 'S1', month: '2026-09', line: 'H9', round_no: 1, total_items: 100, matched: 95, diff_count: 5, adjusted_count: 5, status: 'approved' }],
  partTags: [
    { tag_no: 'T1', part_name: 'ใบมีด', model: 'BL-1', line: 'Lug&Screw', machine: 'LS-10', installed_at: '2026-01-01', removed_at: '2026-03-02' },
    { tag_no: 'T2', part_name: 'ใบมีด', model: 'BL-1', line: 'Lug&Screw', machine: 'LS-10', installed_at: '2026-03-02', removed_at: '2026-05-01' }
  ],
  machines: [{ machine_id: 'MC-1', line: 'H9', machine_name: 'H9-1', active: true }],
  prBundle: { prHeaders: { headers: ['pr_id', 'status', 'total_amount'], rows: [['PR-1', 'PENDING', 5000]] }, prLines: { headers: [], rows: [] }, prAudit: { headers: [], rows: [] } },
  failed: {}
};

// ── ทุกรายงานต้องมีสเปกกราฟ และต้องหาคอลัมน์เจอจริง ─────────────────
assert.strictEqual(mod.XP_REPORTS.length, 10, 'ต้องมีรายงาน 10 ตัว');
const noChart = [];
mod.XP_REPORTS.forEach(function(rp) {
  const spec = mod.XP_CHART_SPECS[rp.key];
  assert(spec, 'รายงาน ' + rp.key + ' ยังไม่มีสเปกกราฟ — ตารางเปล่าอ่านยาก ทุกรายงานต้องมีกราฟ');
  assert(['bar', 'hbar', 'doughnut'].indexOf(spec.type) > -1, rp.key + ' ชนิดกราฟไม่รู้จัก: ' + spec.type);
  assert(Array.isArray(spec.labelCols) && spec.labelCols.length, rp.key + ' ต้องระบุคอลัมน์ป้าย');
  assert(Array.isArray(spec.valueCols) && spec.valueCols.length, rp.key + ' ต้องระบุคอลัมน์ค่า');

  const built = rp.build(ctx);
  const series = mod.xpChartSeries(rp.key, built.table);
  if (!series) { noChart.push(rp.key); return; }
  assert(series.items.length > 0, rp.key + ' ได้ชุดกราฟว่าง');
  assert(series.valueLabel, rp.key + ' ต้องรู้ว่ากราฟกำลังวัดอะไร (ไว้ทำ legend)');
  series.items.forEach(function(it) {
    assert(typeof it.value === 'number' && isFinite(it.value), rp.key + ' มีค่าที่ไม่ใช่ตัวเลขในกราฟ: ' + it.value);
    assert(typeof it.label === 'string' && it.label.length, rp.key + ' มีแท่งที่ไม่มีป้ายชื่อ');
  });
  if (spec.limit) assert(series.items.length <= spec.limit, rp.key + ' เกิน limit ที่ตั้งไว้');
  // config ต้องประกอบได้จริง
  const cfg = mod.xpChartConfig(series, { title: rp.name });
  assert(cfg.data.labels.length === series.items.length, rp.key + ' จำนวน label ไม่ตรงกับข้อมูล');
  assert(cfg.options.animation === false, rp.key + ' ต้องปิด animation (ตอน render เป็นรูปต้องได้ภาพเต็มเฟรมแรก)');
  assert(mod.xpChartHeight(series) >= 220, rp.key + ' ความสูงกราฟน้อยเกินไป');
});
// ข้อมูลตัวอย่างนี้ควรวาดกราฟได้แทบทุกรายงาน — ถ้าวาดไม่ได้เกินครึ่งคือคอลัมน์ที่อ้างไว้เพี้ยน
assert(noChart.length <= 2, 'รายงานที่วาดกราฟไม่ได้เยอะเกินไป (' + noChart.join(', ') + ') — ตรวจว่าชื่อคอลัมน์ในสเปกยังตรงกับตารางจริง');

// ── กราฟต้องมาจากตารางเดียวกับที่ลงไฟล์ (ตัวเลขต้องตรงกัน) ──────────
const topTable = mod.XP_REPORTS[0].build(ctx).table;
const topSeries = mod.xpChartSeries('topIssue', topTable);
const nameIdx = topTable.headers.indexOf('ชื่ออะไหล่');
const valIdx = topTable.headers.indexOf('มูลค่ารวม');
assert.strictEqual(topSeries.items[0].value, topTable.rows[0][valIdx], 'ค่าแท่งแรกต้องเท่ากับแถวแรกในตารางเป๊ะ');
assert(topSeries.items[0].label.indexOf(topTable.rows[0][nameIdx]) > -1, 'ป้ายแท่งต้องมาจากชื่อในตาราง');
// แท่งนอนต้องเรียงมากไปน้อย (คนอ่านกราฟคาดหวังแบบนี้)
for (let i = 1; i < topSeries.items.length; i += 1) {
  assert(topSeries.items[i - 1].value >= topSeries.items[i].value, 'กราฟแท่งนอนต้องเรียงจากมากไปน้อย');
}

// ── ABC ต้องรวมยอดตามชั้น ไม่ใช่วาดทีละอะไหล่ ────────────────────────
const abcSeries = mod.xpChartSeries('abc', mod.XP_REPORTS[2].build(ctx).table);
assert(abcSeries.items.length <= 3, 'ABC ต้องยุบเหลือ 3 ชั้น แต่ได้ ' + abcSeries.items.length + ' แท่ง');
assert(abcSeries.items.map(function(i) { return i.label; }).indexOf('A') > -1, 'ต้องมีชั้น A');
const abcCfg = mod.xpChartConfig(abcSeries, {});
assert.strictEqual(abcCfg.type, 'doughnut', 'ABC ควรเป็นกราฟวงกลม (สัดส่วนของทั้งหมด)');
assert(Array.isArray(abcCfg.data.datasets[0].backgroundColor), 'กราฟวงกลมต้องมีสีต่อชิ้น');

// ── เส้นเป้าต้องซ้อนบนกราฟต้นทุน (เห็นทันทีว่าเกินเป้าไหม) ───────────
let costReport = null;
mod.XP_REPORTS.forEach(function(r) { if (r.key === 'costVsProduction') costReport = r.build(ctx); });
const costSeries = mod.xpChartSeries('costVsProduction', costReport.table);
const costCfg = mod.xpChartConfig(costSeries, {});
assert.strictEqual(costCfg.data.datasets.length, 2, 'กราฟต้นทุนต้องมีทั้งแท่งจริงและเส้นเป้า');
assert.strictEqual(costCfg.data.datasets[1].type, 'line', 'เป้าต้องเป็นเส้น ไม่ใช่แท่ง');
assert(costCfg.data.datasets[1].borderDash, 'เส้นเป้าควรเป็นเส้นประให้ต่างจากข้อมูลจริง');
// เดือน/ไลน์ที่ไม่มีมูลค่าผลิต (% คำนวณไม่ได้ = ช่องว่าง) ต้องไม่ถูกวาดเป็นแท่ง 0
const pctIdx = costReport.table.headers.indexOf('% จริง');
const blankRows = costReport.table.rows.filter(function(r) { return r[pctIdx] === ''; }).length;
assert(blankRows > 0, 'ข้อมูลตัวอย่างควรมีแถวที่คำนวณ % ไม่ได้ เพื่อทดสอบเคสนี้');
assert.strictEqual(costSeries.items.length, costReport.table.rows.length - blankRows,
  'แถวที่คำนวณไม่ได้ต้องไม่กลายเป็นแท่ง 0 บนกราฟ — 0% กับ "ไม่รู้" คนละความหมาย');

// ── โหมดลูกค้า: ต้องถอยไปใช้จำนวนชิ้น ไม่ใช่กราฟหายไป ────────────────
mod.setCustomer(true);
const custTable = mod.XP_REPORTS[0].build(ctx).table;
assert.strictEqual(custTable.headers.indexOf('มูลค่ารวม'), -1, 'โหมดลูกค้าต้องไม่มีคอลัมน์มูลค่า');
const custSeries = mod.xpChartSeries('topIssue', custTable);
assert(custSeries, 'โหมดลูกค้าต้องยังมีกราฟ');
assert.strictEqual(custSeries.valueLabel, 'ยอดเบิก (ชิ้น)', 'ต้องถอยไปวัดเป็นจำนวนชิ้นแทนมูลค่า');
mod.setCustomer(false);

// ── ภาษาอังกฤษ: legend/ป้ายค่าต้องเป็นอังกฤษ ──────────────────────────
mod.setLang('en');
const enSeries = mod.xpChartSeries('topIssue', mod.XP_REPORTS[0].build(ctx).table);
assert(enSeries, 'โหมดอังกฤษต้องยังหาคอลัมน์เจอ (สเปกอ้างชื่อไทยแล้วแปลก่อนค้น)');
assert.strictEqual(enSeries.valueLabel, 'Total Value', 'legend ต้องเป็นอังกฤษ');
// ต้องสร้างตารางใหม่ในภาษาเดียวกัน — สเปกกราฟค้นคอลัมน์จากหัวตารางที่แปลแล้ว
// (ในแอปตารางกับกราฟถูกสร้างในรอบเดียวกันเสมอ และเปลี่ยนภาษาจะล้าง xpLastReport ทิ้ง)
let enCostReport = null;
mod.XP_REPORTS.forEach(function(r) { if (r.key === 'costVsProduction') enCostReport = r.build(ctx); });
const enCost = mod.xpChartConfig(mod.xpChartSeries('costVsProduction', enCostReport.table), {});
assert.strictEqual(enCost.data.datasets[1].label, 'Target (%)', 'ป้ายเส้นเป้าต้องเป็นอังกฤษ');
// ตารางที่สร้างไว้ภาษาอื่นต้องไม่ถูกวาดกราฟผิด ๆ — คืน null ดีกว่าได้กราฟที่คอลัมน์ไม่ตรง
assert.strictEqual(mod.xpChartSeries('costVsProduction', costReport.table), null,
  'ตารางภาษาไทยกับโหมดอังกฤษต้องหาคอลัมน์ไม่เจอ แล้วคืน null ไม่ใช่เดาคอลัมน์');
mod.setLang('th');

// ── ชุดกราฟสรุปของ Executive Workbook ────────────────────────────────
const trend = mod.xpBuildTrend(ctx);
const lineSummary = mod.xpBuildLineSummary(ctx);
const summary = mod.xpSummaryChartSeries(trend, lineSummary, mod.XP_REPORTS[0].build(ctx), costReport);
assert(summary.length >= 5, 'ชีตกราฟสรุปควรมีอย่างน้อย 5 กราฟ แต่ได้ ' + summary.length);
summary.forEach(function(c) {
  assert(c.title && !/undefined/.test(c.title), 'กราฟสรุปต้องมีหัวข้อ');
  assert(c.series && c.series.items.length, 'กราฟ "' + c.title + '" ไม่มีข้อมูล');
});
const titles = summary.map(function(c) { return c.title; });
assert(titles.indexOf('ยอดเบิกออกตามช่วงเวลา') > -1, 'ต้องมีกราฟเทรนด์ยอดเบิก');
assert(titles.indexOf('เปรียบเทียบยอดเบิกตามไลน์') > -1, 'ต้องมีกราฟเทียบไลน์');
// ตัวช่วยชั่วคราวต้องไม่ค้างอยู่ใน XP_CHART_SPECS
assert.strictEqual(mod.XP_CHART_SPECS.__tmp, undefined, 'xpChartSeriesFromTable ต้องลบสเปกชั่วคราวทิ้งทุกครั้ง');

// ── ต้องทนกรณีไม่มีข้อมูล / Chart.js โหลดไม่ได้ ──────────────────────
assert.strictEqual(mod.xpChartSeries('topIssue', { headers: [], rows: [] }), null, 'ตารางว่างต้องคืน null ไม่ throw');
assert.strictEqual(mod.xpChartSeries('ไม่มีรายงานนี้', { headers: ['a'], rows: [['x']] }), null, 'รายงานที่ไม่มีสเปกต้องคืน null');
const renderSrc = grabFn('xpRenderChartInto');
assert(/if \(!window\.Chart\)/.test(renderSrc), 'ต้องเช็คว่า Chart.js โหลดมาหรือยัง');
assert(/ดูข้อมูลจากตารางด้านล่างได้ตามปกติ/.test(renderSrc), 'Chart.js โหลดไม่ได้ต้องยังเห็นตารางได้ ไม่ทำให้รายงานพัง');
assert(/catch \(err\)/.test(renderSrc), 'วาดกราฟพลาดต้องไม่ทำให้ทั้งหน้าล้ม');
const pngSrc = grabFn('xpChartToPng');
assert(/left:-99999px/.test(pngSrc), 'ต้องแปะ canvas เข้า DOM ชั่วคราวแบบซ่อนไว้ ไม่งั้นบางเบราว์เซอร์ได้รูปเปล่า');
assert(/holder\.parentNode\.removeChild\(holder\)/.test(pngSrc), 'ต้องเก็บ canvas ชั่วคราวทิ้งทุกกรณี');
assert(/chart\.destroy\(\)/.test(pngSrc), 'ต้องทำลาย chart instance ทิ้ง กัน memory leak');

// ── หน้าจอ: กราฟต้องมาก่อนตาราง และมีปุ่มยุบตาราง ────────────────────
const previewSrc = grabFn('xpRenderReportPreview');
assert(previewSrc.indexOf('xpReportChart') < previewSrc.indexOf('xpReportTableWrap'),
  'กราฟต้องอยู่เหนือตาราง — คนอ่านต้องเห็นภาพรวมก่อนลงรายละเอียด');
assert(/xpDestroyLiveCharts\(\)/.test(previewSrc), 'เปิดรายงานใหม่ต้องทำลายกราฟเดิมก่อน');
assert(/xpToggleTableBtn/.test(previewSrc), 'ต้องมีปุ่มยุบ/แสดงตาราง');
assert(/xpChartSeries\(rp\.key, table\)/.test(previewSrc), 'กราฟต้องวาดจากตารางเต็ม ไม่ใช่แค่ 200 แถวที่โชว์บนจอ');
const afterSrc = grabFn('xpAfterFilterChange');
assert(/xpDestroyLiveCharts\(\)/.test(afterSrc), 'เปลี่ยนตัวกรองต้องล้างกราฟที่ค้างอยู่');

// ── Excel: ต้องฝังรูปกราฟได้ ─────────────────────────────────────────
const writeSrc = grabFn('xpWriteXlsxStyled');
assert(/wb\.addImage\(\{ base64: img\.dataUrl, extension: 'png' \}\)/.test(writeSrc), 'ต้องฝังรูปกราฟลง workbook');
assert(/ws\.addImage\(imgId/.test(writeSrc), 'ต้องวางรูปลงชีต');
assert(/spec\.images \|\| \[\]/.test(writeSrc), 'ชีตที่ไม่มีรูปต้องทำงานได้เหมือนเดิม');
const wbSrc = grabFn('xpRunWorkbook');
assert(/04 กราฟสรุป/.test(wbSrc), 'Executive Workbook ต้องมีชีตกราฟสรุป');
assert(/xpChartToPng/.test(wbSrc), 'ต้อง render กราฟเป็นรูปก่อนฝัง');
assert(/if \(chartImages\.length\)/.test(wbSrc), 'ถ้า render รูปไม่ได้เลย ต้องไม่ใส่ชีตกราฟเปล่า');
const reportDlSrc = grabFn('xpDownloadReport');
assert(/xpChartToPng/.test(reportDlSrc), 'ไฟล์รายงานระดับ 3 ก็ควรมีกราฟแนบไปด้วย');

// ── PDF: ต้องมีกราฟแท่งนอนในหน้า Top parts / Dead stock / สรุปไลน์ ──
const pdfSrc = grabFn('xpRunPdfReport');
assert((pdfSrc.match(/xpSvgHBarChart\(/g) || []).length >= 3, 'PDF ควรมีกราฟแท่งนอนอย่างน้อย 3 จุด');
assert(/xpSvgBarChart\(/.test(pdfSrc), 'PDF ต้องยังมีกราฟเทรนด์แบบแท่งตั้ง');
const svgSrc = grabFn('xpSvgHBarChart');
assert(/label\.slice\(0, 39\)/.test(svgSrc), 'ชื่อยาวต้องถูกตัดไม่ให้ทับแท่ง');
assert(!/canvas/.test(svgSrc), 'กราฟใน PDF ต้องเป็น SVG ไม่ใช่ canvas (บางเบราว์เซอร์พรินต์ canvas ออกมาเป็นช่องว่าง)');

// ── ไม่เพิ่ม library ใหม่ — Chart.js โหลดอยู่แล้วในหน้าเว็บ ───────────
assert(/chart\.js@4\.4\.3/.test(html), 'Chart.js ต้องยังถูกโหลดในหน้าเว็บ (กราฟบนจอพึ่งตัวนี้)');
assert.strictEqual((html.match(/cdn\.jsdelivr\.net\/npm\/chart\.js/g) || []).length, 1, 'ห้ามโหลด Chart.js ซ้ำซ้อน');

console.log('export-charts: OK (' + Object.keys(mod.XP_CHART_SPECS).length + ' สเปกกราฟ, ' + summary.length + ' กราฟสรุป)');
