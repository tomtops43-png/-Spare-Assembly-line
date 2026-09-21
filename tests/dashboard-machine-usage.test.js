const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── การ์ด "ประวัติการเบิกของเครื่องจักร" บน Dashboard ──────────────────────────
// ตอบคำถามว่า ไลน์ไหน เครื่องไหน เบิกอะไหล่อะไรไปกี่ครั้ง — อ่านจากคอลัมน์ Machine
// ของชีต Log ที่บังคับกรอกตอนเบิก ไม่ใช่ข้อมูลที่เดาเอง
assert(htmlLf.includes('id="fccMachineCard"'), 'ต้องมีการ์ดประวัติการเบิกของเครื่องจักร');
assert(htmlLf.includes('id="fccMachineTable"'), 'ต้องมีตารางเครื่องจักร');
assert(htmlLf.includes('id="fccMachineSummary"'), 'ต้องมีแถบสรุปหัวการ์ด');
assert(htmlLf.includes('id="fccMachineLineFilter"'), 'ต้องกรองตามไลน์ได้');
assert(htmlLf.includes('id="fccMachineRangeFilter"'), 'ต้องกรองตามช่วงเวลาได้');
assert(htmlLf.includes('id="fccMachineSortFilter"'), 'ต้องเลือกวิธีเรียงได้');
// ต้องคำนวณใหม่จาก log ที่โหลดไว้แล้ว ไม่ยิง API ซ้ำทุกครั้งที่เปลี่ยนตัวกรอง
assert(/\['fccMachineLineFilter', 'fccMachineRangeFilter', 'fccMachineSortFilter'\][\s\S]{0,220}addEventListener\('change', fccRenderMachineUsage\)/.test(htmlLf),
  'ตัวกรองทั้งสามต้องผูกกับ fccRenderMachineUsage');
assert(/fccRenderMiscExpense\(\);\n\s*fccRenderMachineUsage\(\);/.test(htmlLf),
  'renderDashboardPhase1 ต้องเรียก fccRenderMachineUsage หลังโหลด log เสร็จ');
// กางดูรายการอะไหล่ของแต่ละเครื่องผ่าน delegated handler ตัวเดิมของ FCC
assert(/kind === 'mexp'[\s\S]{0,160}fccRenderMachineUsage\(\)/.test(htmlLf),
  'ต้องมี action กางแถวเครื่องจักร');

// ── ตรรกะรวมยอด: ประกอบฟังก์ชันจริงจาก index.html มารันกับ log จำลอง ──────────
const src = [
  grab(/^ {4}function fccEsc\(v\) \{[\s\S]*?\n {4}\}/m, 'fccEsc'),
  grab(/^ {4}function fccN\(v\) \{.*\}$/m, 'fccN'),
  grab(/^ {4}function fccPrice\(v\) \{[\s\S]*?\n {4}\}/m, 'fccPrice'),
  grab(/^ {4}function fccInt\(n\) \{.*\}$/m, 'fccInt'),
  grab(/^ {4}function fccLineOf\(raw\) \{[\s\S]*?\n {4}\}/m, 'fccLineOf'),
  grab(/^ {4}function fccPartKey\(line, name, model\) \{[\s\S]*?\n {4}\}/m, 'fccPartKey'),
  grab(/^ {4}function fccRelTime\(ts\) \{[\s\S]*?\n {4}\}/m, 'fccRelTime'),
  grab(/^ {4}function fccMonthKey\(raw\) \{[\s\S]*?\n {4}\}/m, 'fccMonthKey'),
  grab(/^ {4}function fccExpenseValidRows\(logs\) \{[\s\S]*?\n {4}\}/m, 'fccExpenseValidRows'),
  grab(/^ {4}function fccMachineAnomalies\(rows\) \{[\s\S]*?\n {4}\}/m, 'fccMachineAnomalies'),
  grab(/^ {4}function fccRenderMachineUsage\(\) \{[\s\S]*?\n {4}\}/m, 'fccRenderMachineUsage')
].join('\n');

function makeEl() { return { innerHTML: '', textContent: '', value: '' }; }
const els = {
  fccMachineTable: makeEl(), fccMachineSummary: makeEl(), fccMachineAlert: makeEl(),
  fccMachineNote: makeEl(), fccMachineSub: makeEl(),
  fccMachineLineFilter: makeEl(), fccMachineRangeFilter: makeEl(), fccMachineSortFilter: makeEl()
};
els.fccMachineLineFilter.value = '__all__';
els.fccMachineRangeFilter.value = '0';   // ทั้งหมด — กันเทสต์พังเมื่อเวลาผ่านไป
els.fccMachineSortFilter.value = 'count';

const harness = new Function('document', 'partsData', 'dashboardAllPartsData', 'fccLast', 'formatThaiDateTime',
  'var fccMachineExp = {};\n' + src +
  '\nreturn { render: fccRenderMachineUsage, exp: fccMachineExp, valid: fccExpenseValidRows, anomalies: fccMachineAnomalies };');

const parts = [
  { line: 'H9', name: 'Bearing 6204', model: 'SKF-01', unit_price: 100 },
  { line: 'Coil Winding', name: 'End Mill', model: 'YHM-004', unit_price: 250 }
];
const now = Date.now();
const ts = (daysAgo) => new Date(now - daysAgo * 86400000).toISOString();
const logs = [
  { type: 'Output', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 2, timestamp: ts(1), machine: 'H9-01', by: 'Jang' },
  { type: 'Output', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 1, timestamp: ts(3), machine: 'H9-01', by: 'Jang' },
  { type: 'Output', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 5, timestamp: ts(2), machine: 'H9-02', by: 'Tom' },
  { type: 'Output', process: 'Coil Winding', partName: 'End Mill', model: 'YHM-004', qty: 1, timestamp: ts(4), machine: 'CW-01', by: 'Tom' },
  // เบิกเก่าก่อนมีช่องเลือกเครื่อง — ต้องไม่ถูกนับรวมมั่วเข้าเครื่องใดเครื่องหนึ่ง
  { type: 'Output', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 9, timestamp: ts(5), machine: '', by: 'Old' },
  // รับเข้า ไม่ใช่การเบิก — ห้ามนับ
  { type: 'Input', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 50, timestamp: ts(1), machine: 'H9-01', by: 'Jang' }
];
const fccLast = { Logs: logs };
const api = harness(
  { getElementById: (id) => els[id] || null },
  parts, parts, fccLast,
  (t) => new Date(t).toISOString().slice(0, 10)
);

// machine + ms ต้องไหลออกมาจากตัวแปลง log กลาง ไม่งั้นการ์ดนี้ไม่มีข้อมูลให้จัดกลุ่ม
const validRows = api.valid(logs);
assert.strictEqual(validRows.length, 5, 'ต้องนับเฉพาะรายการเบิก (Output)');
assert.strictEqual(validRows[0].machine, 'H9-01', 'ต้องอ่านชื่อเครื่องจาก log');
assert(validRows[0].ms > 0, 'ต้องมี timestamp ไว้หา "เบิกล่าสุด"');

api.render();
const table = els.fccMachineTable.innerHTML;
const summary = els.fccMachineSummary.innerHTML;
const note = els.fccMachineNote.innerHTML;

assert(table.includes('H9-01') && table.includes('H9-02') && table.includes('CW-01'),
  'ต้องแยกเป็นรายเครื่อง ไม่ยุบรวมทั้งไลน์');
assert(table.includes('รวม 3 เครื่อง'), 'ต้องสรุปจำนวนเครื่องที่มีการเบิก');
// H9-01 เบิก 2 ครั้ง (3 ชิ้น = 300฿), H9-02 เบิก 1 ครั้ง (5 ชิ้น = 500฿), CW-01 1 ครั้ง (1 ชิ้น = 250฿)
assert(/H9-01[\s\S]{0,400}>2<[\s\S]{0,200}>3<[\s\S]{0,200}฿300/.test(table),
  'H9-01 ต้องได้ 2 ครั้ง / 3 ชิ้น / ฿300');
assert(summary.includes('3 เครื่อง'), 'แถบสรุปต้องบอกจำนวนเครื่อง');
assert(summary.includes('4 ครั้ง'), 'แถบสรุปต้องรวมครั้งที่เบิกของทุกเครื่อง');
assert(summary.includes('H9-01'), 'แถบสรุปต้องชี้เครื่องที่เบิกบ่อยสุด');
// รายการที่ไม่ได้ระบุเครื่องต้องถูกแจ้ง ไม่ใช่หายเงียบจนตัวเลขดูน้อยกว่าจริง
assert(note.includes('1 ครั้ง') && note.includes('9 ชิ้น'), 'ต้องเตือนรายการที่ยังไม่ระบุเครื่อง');

// ── กางแถวแล้วต้องเห็นว่าเครื่องนั้นเบิก "อะไหล่อะไร กี่ครั้ง" ────────────────
assert(!table.includes('อะไหล่ที่เบิกใส่เครื่องนี้'), 'ค่าเริ่มต้นต้องยังไม่กาง');
api.exp['H9||H9-01'] = true;
api.render();
const expanded = els.fccMachineTable.innerHTML;
assert(expanded.includes('อะไหล่ที่เบิกใส่เครื่องนี้'), 'กางแล้วต้องเห็นตารางอะไหล่');
assert(expanded.includes('Bearing 6204 (SKF-01)'), 'ต้องบอกชื่ออะไหล่ + รุ่น');

// ── ตัวกรองไลน์ต้องตัดเครื่องของไลน์อื่นออกจริง ───────────────────────────────
els.fccMachineLineFilter.value = 'Coil Winding';
api.render();
const cwOnly = els.fccMachineTable.innerHTML;
assert(cwOnly.includes('CW-01'), 'กรองไลน์ Coil Winding ต้องเหลือเครื่องของไลน์นั้น');
assert(!cwOnly.includes('H9-01') && !cwOnly.includes('H9-02'), 'ต้องไม่มีเครื่องของไลน์อื่นปน');

// ── ไม่มีข้อมูลต้องบอกทางออก ไม่ใช่ปล่อย skeleton ค้าง ────────────────────────
fccLast.Logs = [];
api.render();
assert(els.fccMachineTable.innerHTML.includes('ยังไม่มีประวัติการเบิกที่ระบุเครื่อง'),
  'ไม่มีข้อมูลต้องขึ้นข้อความอธิบาย');

// ── เตือนเครื่องที่เบิกถี่ผิดปกติ (เทียบกับค่าปกติของเครื่องนั้นเอง) ──────────
// แต่ละเครื่องต้องเทียบกับตัวเอง ไม่ใช่เทียบข้ามเครื่อง เพราะเครื่องใหญ่/เล็ก
// กินอะไหล่คนละระดับกันอยู่แล้ว
function row(machine, daysAgo, line) {
  return { machine: machine, ms: now - daysAgo * 86400000, qty: 1, line: line || 'H9', partName: 'Bearing 6204', model: 'SKF-01' };
}
const anomalyRows = [];
// SPIKE-01: ฐาน 90 วันก่อนหน้า 3 ครั้ง (= ~1 ครั้ง/30 วัน) แต่ 30 วันล่าสุดพุ่งเป็น 6 → ~6 เท่า
[100, 70, 45].forEach(function(d) { anomalyRows.push(row('SPIKE-01', d)); });
[2, 5, 9, 14, 20, 27].forEach(function(d) { anomalyRows.push(row('SPIKE-01', d)); });
// STEADY-01: เบิกสม่ำเสมอ ฐาน 9 ครั้ง (~3/30วัน) ล่าสุด 3 ครั้ง → ปกติ ห้ามเตือน
[35, 45, 55, 65, 75, 85, 95, 105, 115].forEach(function(d) { anomalyRows.push(row('STEADY-01', d)); });
[4, 12, 25].forEach(function(d) { anomalyRows.push(row('STEADY-01', d)); });
// NEW-01: เพิ่งมีในระบบ ไม่มีฐานให้เทียบ → ห้ามเตือนว่าผิดปกติ (ยังไม่รู้ค่าปกติของมัน)
[1, 6, 11, 18].forEach(function(d) { anomalyRows.push(row('NEW-01', d)); });
// ONCE-01: ล่าสุดเบิกแค่ 2 ครั้ง น้อยเกินกว่าจะสรุป → ห้ามเตือน
[3, 8].forEach(function(d) { anomalyRows.push(row('ONCE-01', d)); });
[50, 60, 70].forEach(function(d) { anomalyRows.push(row('ONCE-01', d)); });

const found = api.anomalies(anomalyRows);
const names = found.map(function(a) { return a.machine; });
assert(names.indexOf('SPIKE-01') > -1, 'เครื่องที่เบิกพุ่งผิดปกติต้องถูกจับได้');
assert(names.indexOf('STEADY-01') === -1, 'เครื่องที่เบิกสม่ำเสมอต้องไม่ถูกเตือน');
assert(names.indexOf('NEW-01') === -1, 'เครื่องใหม่ที่ยังไม่มีฐานเทียบ ต้องไม่ถูกเหมาว่าผิดปกติ');
assert(names.indexOf('ONCE-01') === -1, 'จำนวนครั้งน้อยเกินไปต้องไม่เตือน (กันเตือนมั่วจนคนเลิกเชื่อ)');
const spike = found[0];
assert.strictEqual(spike.recent, 6, 'ต้องนับครั้งใน 30 วันล่าสุดให้ถูก');
assert(spike.ratio >= 3, 'ฐาน ~1 ครั้ง/30วัน เบิกจริง 6 ครั้ง ต้องได้อัตราส่วน ≥3 เท่า');
assert.strictEqual(spike.level, 'crit', '≥3 เท่าต้องเป็นระดับวิกฤต');
assert(spike.topPart.includes('Bearing 6204'), 'ต้องบอกด้วยว่าอะไหล่ตัวไหนที่เบิกถี่');

// ต้องขึ้นแถบเตือนบนการ์ด พร้อมตัวเลขที่ตรวจสอบย้อนได้ ไม่ใช่บอกแค่ "ผิดปกติ"
fccLast.Logs = anomalyRows.map(function(r) {
  return { type: 'Output', process: r.line, partName: r.partName, model: r.model, qty: r.qty,
    timestamp: new Date(r.ms).toISOString(), machine: r.machine };
});
els.fccMachineLineFilter.value = '__all__';
api.render();
const alertHtml = els.fccMachineAlert.innerHTML;
assert(alertHtml.includes('เบิกอะไหล่ถี่ผิดปกติ'), 'ต้องมีแถบเตือนบนการ์ด');
assert(alertHtml.includes('SPIKE-01'), 'แถบเตือนต้องระบุชื่อเครื่อง');
assert(alertHtml.includes('เบิก 6 ครั้งใน 30 วัน'), 'ต้องโชว์ตัวเลขจริงให้คนตรวจสอบย้อนได้');
assert(!alertHtml.includes('STEADY-01'), 'เครื่องปกติต้องไม่โผล่ในแถบเตือน');
assert(alertHtml.includes('data-fcc-act="mexp:'), 'กดที่แถบเตือนต้องกางดูรายการอะไหล่ของเครื่องนั้นได้');
assert(els.fccMachineTable.innerHTML.includes('🚨'), 'แถวในตารางของเครื่องผิดปกติต้องติดป้ายด้วย');

// ไม่มีเครื่องผิดปกติ = ต้องไม่เหลือแถบเตือนค้างจากรอบก่อน
fccLast.Logs = [{ type: 'Output', process: 'H9', partName: 'Bearing 6204', model: 'SKF-01', qty: 1, timestamp: ts(2), machine: 'CALM-01' }];
api.render();
assert.strictEqual(els.fccMachineAlert.innerHTML, '', 'ไม่มีเครื่องผิดปกติต้องไม่ขึ้นแถบเตือน');

console.log('Dashboard machine usage history checks passed');
