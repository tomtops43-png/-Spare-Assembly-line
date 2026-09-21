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
  grab(/^ {4}function fccRenderMachineUsage\(\) \{[\s\S]*?\n {4}\}/m, 'fccRenderMachineUsage')
].join('\n');

function makeEl() { return { innerHTML: '', textContent: '', value: '' }; }
const els = {
  fccMachineTable: makeEl(), fccMachineSummary: makeEl(),
  fccMachineNote: makeEl(), fccMachineSub: makeEl(),
  fccMachineLineFilter: makeEl(), fccMachineRangeFilter: makeEl(), fccMachineSortFilter: makeEl()
};
els.fccMachineLineFilter.value = '__all__';
els.fccMachineRangeFilter.value = '0';   // ทั้งหมด — กันเทสต์พังเมื่อเวลาผ่านไป
els.fccMachineSortFilter.value = 'count';

const harness = new Function('document', 'partsData', 'dashboardAllPartsData', 'fccLast', 'formatThaiDateTime',
  'var fccMachineExp = {};\n' + src + '\nreturn { render: fccRenderMachineUsage, exp: fccMachineExp, valid: fccExpenseValidRows };');

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

console.log('Dashboard machine usage history checks passed');
