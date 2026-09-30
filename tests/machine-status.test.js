// 🏭 สถานะเครื่องจักร — ทุกคนในไลน์เปิด/ปิดสถานะเครื่องเอง + บอกสาเหตุ + ผูกอะไหล่ที่รอ
// ฝั่ง backend รันของจริงบนชีตจำลอง (ดู _machine-status-backend.js)
// ฝั่งเว็บตัดฟังก์ชันคำนวณชั่วโมงเครื่องหยุด/เช็คอะไหล่มาแล้ว ออกมารันจริง + เช็คการเดินสายหน้าใหม่
const { html, backend, grabFn, grabVar, assert } = require('./_export-extract');
const { makeMachineStatusBackend } = require('./_machine-status-backend');

const MACHINES = [
  ['Machine ID', 'Line', 'Machine Name', 'Active', 'Created By', 'Created At', 'Updated By', 'Updated At'],
  ['MC-H9-1', 'H9', 'H9 Press 1', true, 'admin', '', 'admin', ''],
  ['MC-H9-2', 'H9', 'H9 Press 2', true, 'admin', '', 'admin', ''],
  ['MC-LS-4', 'Lug&Screw', 'Lug & Screw 4', true, 'admin', '', 'admin', ''],
  ['MC-LS-9', 'Lug&Screw', 'Lug & Screw 9 (ปลดแล้ว)', false, 'admin', '', 'admin', '']
];
const USERS = {
  h9: { username: 'somchai', role: 'user', line: 'H9' },
  ls: { username: 'somsak', role: 'leader', line: 'Lug&Screw' },
  floater: { username: 'wichai', role: 'user', line: '' },
  admin: { username: 'admin', role: 'admin', line: 'H9' },
  exec: { username: 'boss', role: 'user', line: '', viewOnly: true }
};
function fresh() {
  return makeMachineStatusBackend({ users: USERS, sheets: { Machines: MACHINES.map(function(r) { return r.slice(); }) } });
}
function throwsWith(fn, text, msg) {
  assert.throws(fn, function(err) { return String(err && err.message).indexOf(text) > -1; }, msg);
}
const bearing = { key: '12::Stock for MC::6204ZZ::Bearing::A-01', no: '12', name: 'Bearing', model: '6204ZZ', sheet: 'Stock for MC', qty_needed: 2, stock_at_save: 0 };

// ── บอร์ด: เครื่องที่ยังไม่เคยตั้งสถานะ = ทำงาน, เครื่องที่ปลดจากทะเบียนไม่โผล่ ───────────
let b = fresh();
let board = b.getMachineStatusBoard({ authToken: 'h9' });
assert.strictEqual(board.machines.length, 3, 'เครื่องที่ปิดใช้งานในทะเบียนต้องไม่โผล่บนบอร์ด');
assert(board.machines.every(function(m) { return m.status === 'running'; }), 'ยังไม่เคยตั้งสถานะ = ทำงาน');
assert.deepStrictEqual(board.machines.map(function(m) { return m.can_edit; }), [true, true, false],
  'คนไลน์ H9 แก้ได้เฉพาะเครื่อง H9');
assert.strictEqual(b.getMachineStatusBoard({ authToken: 'h9', line: 'Lug&Screw' }).machines.length, 1, 'กรองตามไลน์ได้');
assert(b.getMachineStatusBoard({ authToken: 'floater' }).machines.every(function(m) { return m.can_edit; }),
  'คนที่ไม่ได้ผูกไลน์ (ดูแลทุกไลน์) แก้ได้ทุกเครื่อง');
assert(b.getMachineStatusBoard({ authToken: 'admin' }).machines.every(function(m) { return m.can_edit; }), 'Admin แก้ได้ทุกไลน์');

// ── สิทธิ์ ──────────────────────────────────────────────────────────────
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-LS-4', status: 'stopped', reason_type: 'breakdown' }); },
  'เฉพาะเครื่องในไลน์ของคุณ', 'ข้ามไลน์ต้องถูกกัน (backend เช็คซ้ำ ไม่เชื่อปุ่มหน้าเว็บ)');
throwsWith(function() { b.updateMachineStatus({ authToken: 'exec', machine_id: 'MC-H9-1', status: 'stopped', reason_type: 'breakdown' }); },
  'ดูอย่างเดียว', 'บัญชีผู้บริหาร (ดูอย่างเดียว) เปลี่ยนสถานะไม่ได้');
throwsWith(function() { b.updateMachineStatus({ authToken: 'nobody', machine_id: 'MC-H9-1', status: 'running' }); }, 'เข้าสู่ระบบ');
throwsWith(function() { b.updateMachineStatus({ authToken: 'admin', machine_id: 'MC-LS-9', status: 'stopped', reason_type: 'breakdown' }); },
  'ปิดใช้งาน', 'เครื่องที่ปลดจากทะเบียนแล้วห้ามตั้งสถานะ');
throwsWith(function() { b.updateMachineStatus({ authToken: 'admin', machine_id: 'MC-XXX', status: 'stopped', reason_type: 'breakdown' }); }, 'ไม่พบเครื่องจักร');
assert.strictEqual(b.updateMachineStatus({ authToken: 'ls', machine_id: 'MC-LS-4', status: 'maintenance', reason_type: 'breakdown' }).status, 'success',
  'Leader ของไลน์เปลี่ยนเครื่องไลน์ตัวเองได้');

// ── ตรวจข้อมูล ─────────────────────────────────────────────────────────
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'broken' }); }, 'สถานะไม่ถูกต้อง');
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'stopped' }); }, 'เลือกสาเหตุ',
  'เครื่องไม่ได้ทำงานต้องบอกสาเหตุ');
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'stopped', reason_type: 'other', comment: '  ' }); },
  'คอมเมนต์', 'สาเหตุ "อื่นๆ" ต้องมีคอมเมนต์');
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'waiting_parts', parts_json: '[]' }); },
  'อย่างน้อย 1 รายการ', 'รออะไหล่ต้องระบุอะไหล่');
const tooMany = JSON.stringify(Array.from({ length: 16 }, function(_, i) { return { name: 'P' + i, manual: true }; }));
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'waiting_parts', parts_json: tooMany }); }, 'สูงสุด 15');

// ── วงจรเคส: หยุด → อัปเดตรายละเอียด → เปลี่ยนสาเหตุ → กลับมาทำงาน ──────────────
b = fresh();
const t0 = b.clock.now;
let res = b.updateMachineStatus({
  authToken: 'h9', machine_id: 'MC-H9-1', status: 'waiting_parts', reason_type: 'breakdown',
  parts_json: JSON.stringify([bearing, { name: 'ซีลยาง 40mm', manual: 'true', qty_needed: '0' }, { name: '   ' }]),
  comment: 'รอ supplier ส่ง', pr_ref: 'PR-001'
});
let m = res.machine;
assert.strictEqual(m.status, 'waiting_parts');
assert.strictEqual(m.reason_type, 'spare_part', 'รออะไหล่ = สาเหตุขาด Spare Part เสมอ ไม่ว่าส่งอะไรมา');
assert.strictEqual(m.parts.length, 2, 'แถวอะไหล่ที่ไม่มีชื่อต้องถูกตัดทิ้ง');
assert.strictEqual(m.parts[0].qty_needed, 2);
assert.strictEqual(m.parts[1].manual, true, 'manual ที่ส่งมาเป็นข้อความ "true" ต้องอ่านเป็น true');
assert.strictEqual(m.parts[1].qty_needed, 1, 'จำนวน 0/ไม่ถูกต้อง ปัดเป็น 1');
assert.strictEqual(m.since, '2026-10-01 08:00:00', 'Since เป็นเวลาไทย');
assert(/^MSC-/.test(m.case_id), 'หยุดครั้งแรกต้องเปิดเคสใหม่');
const caseId = m.case_id;
let log = b.getMachineStatusHistory({ authToken: 'h9' }).history;
assert.strictEqual(log.length, 1);
assert.strictEqual(log[0].from_status, 'running');
assert.strictEqual(log[0].to_status, 'waiting_parts');
assert.strictEqual(log[0].changed_by, 'somchai');

b.advanceMinutes(45);
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'waiting_parts', parts_json: JSON.stringify([bearing]), comment: 'supplier แจ้งส่งพรุ่งนี้' }).machine;
assert.strictEqual(m.since, '2026-10-01 08:00:00', 'แก้แค่รายละเอียดในสถานะเดิม ห้ามรีเซ็ตเวลาเริ่มหยุด');
assert.strictEqual(m.case_id, caseId, 'ยังเป็นเคสเดิม');
assert.strictEqual(m.pr_ref, '', 'ไม่ส่งเลข PR มา = ลบออก (ฟอร์มส่งค่าปัจจุบันมาทั้งชุดเสมอ)');
log = b.getMachineStatusHistory({ authToken: 'h9' }).history;
assert.strictEqual(log[0].from_status, 'waiting_parts');
assert.strictEqual(log[0].prev_duration_min, 0, 'สถานะไม่เปลี่ยน = ไม่นับเป็นช่วงที่จบ');

b.advanceMinutes(45);
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'maintenance', reason_type: 'spare_part', parts_json: JSON.stringify([bearing]) }).machine;
assert.strictEqual(m.case_id, caseId, 'เปลี่ยนจากรออะไหล่ → ซ่อมบำรุง ยังเป็นเคสเดียวกัน (เครื่องยังไม่ได้เดิน)');
assert.strictEqual(m.since, '2026-10-01 09:30:00', 'เปลี่ยนสถานะ = เริ่มนับเวลาของสถานะใหม่');
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9' }).history[0].prev_duration_min, 90, 'รออะไหล่มา 90 นาที');

b.advanceMinutes(30);
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'running', reason_type: 'breakdown', parts_json: JSON.stringify([bearing]), pr_ref: 'PR-9' }).machine;
assert.strictEqual(m.status, 'running');
assert.strictEqual(m.reason_type, '', 'กลับมาทำงาน = ล้างสาเหตุ');
assert.deepStrictEqual(m.parts, [], 'กลับมาทำงาน = ล้างอะไหล่ที่รอ');
assert.strictEqual(m.pr_ref, '');
assert.strictEqual(m.case_id, '', 'ปิดเคสแล้ว');
log = b.getMachineStatusHistory({ authToken: 'h9' }).history;
assert.strictEqual(log[0].case_id, caseId, 'log ตอนกลับมาทำงานต้องผูกเคสเดิม (ใช้คิดเวลาเฉลี่ยกว่าจะกลับมาเดิน)');
assert.strictEqual(log[0].prev_duration_min, 30);

b.advanceMinutes(60);
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'stopped', reason_type: 'no_job' }).machine;
assert(m.case_id && m.case_id !== caseId, 'หยุดรอบใหม่ = เคสใหม่');
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9' }).history[0].prev_duration_min, 60, 'นับเวลาที่เดินเครื่องได้ด้วย');

// ── ประวัติ: ใหม่สุดก่อน กรองเครื่อง/ไลน์ และตัดตามจำนวนวัน ────────────────────
b.updateMachineStatus({ authToken: 'ls', machine_id: 'MC-LS-4', status: 'stopped', reason_type: 'breakdown' });
const all = b.getMachineStatusHistory({ authToken: 'h9' }).history;
assert.strictEqual(all.length, 6);
assert.strictEqual(all[0].machine_id, 'MC-LS-4', 'ใหม่สุดขึ้นก่อน');
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9', machine_id: 'MC-H9-1' }).history.length, 5);
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9', line: 'Lug&Screw' }).history.length, 1);
b.advanceMinutes(3 * 24 * 60);
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9', days: 2 }).history.length, 0, 'เก่ากว่าช่วงวันที่ขอต้องไม่ส่งมา');
assert.strictEqual(b.getMachineStatusHistory({ authToken: 'h9', days: 7 }).history.length, 6);

// ── Sheets แปลงเวลาเป็น Date object เอง — ต้องอ่านกลับเป็นข้อความเวลาไทย ────────────
b.sheets.MachineStatus.rows[1][8] = new Date('2026-10-01T01:00:00Z');
const row = b.getMachineStatusBoard({ authToken: 'h9' }).machines.filter(function(x) { return x.machine_id === 'MC-H9-1'; })[0];
assert.strictEqual(row.since, '2026-10-01 08:00:00');


// ── รันได้บางส่วน (degraded): ยังผลิตได้ แต่รออะไหล่บางตัว / บอกผลกระทบ ────────────────
b = fresh();
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'degraded' }); },
  'อย่างน้อย 1 อย่าง', 'รันได้บางส่วนต้องบอกอะไหล่ที่รอ หรือผลกระทบ');
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'degraded', impact_json: JSON.stringify({ impacts: ['slow'], capacity_pct: 120 }) }); },
  '1–99%', 'เปอร์เซ็นต์กำลังผลิตต้องสมเหตุสมผล');
throwsWith(function() { b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'degraded', impact_json: JSON.stringify({ impacts: ['at_risk'] }) }); },
  'ระบุวันที่', 'เสี่ยงหยุดต้องมีวันที่ต้องเปลี่ยนให้ทัน');
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'degraded', reason_type: 'breakdown',
  impact_json: JSON.stringify({ impacts: ['slow', 'workaround', 'bogus', 'slow'], capacity_pct: '70', models: 'ไม่ได้ติ๊กจึงไม่เก็บ' }) }).machine;
assert.strictEqual(m.status, 'degraded');
assert.strictEqual(m.reason_type, '', 'รันได้บางส่วนไม่ใช่เครื่องหยุด จึงไม่เก็บสาเหตุที่หยุด');
assert.deepStrictEqual(m.impact, { impacts: ['slow', 'workaround'], capacity_pct: 70, models: '', risk_until: '' },
  'เก็บเฉพาะผลกระทบที่รู้จัก ไม่ซ้ำ และเก็บช่องเสริมเฉพาะที่ติ๊ก');
assert(/^MSC-/.test(m.case_id), 'รันได้บางส่วนเปิดเคสเหมือนเครื่องมีปัญหา');
const degCase = m.case_id;
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'degraded', parts_json: JSON.stringify([bearing]),
  impact_json: JSON.stringify({ impacts: ['at_risk', 'partial_models'], risk_until: '2026-10-05', models: 'รุ่น 16A' }) }).machine;
assert.strictEqual(m.reason_type, 'spare_part', 'มีอะไหล่ที่รอ = สาเหตุขาด Spare Part (ใช้กับสถิติอะไหล่)');
assert.strictEqual(m.impact.risk_until, '2026-10-05');
assert.strictEqual(m.impact.models, 'รุ่น 16A');
assert.strictEqual(m.impact.capacity_pct, '', 'ไม่ติ๊กกำลังผลิตลด = ไม่เก็บ %');
b.advanceMinutes(120);
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'stopped', reason_type: 'spare_part', parts_json: JSON.stringify([bearing]) }).machine;
assert.strictEqual(m.case_id, degCase, 'รันบางส่วนแล้วอาการหนักขึ้นจนหยุด = เคสเดียวกัน');
assert.deepStrictEqual(m.impact, {}, 'สถานะอื่นไม่เก็บผลกระทบของรันบางส่วน');
let dlog = b.getMachineStatusHistory({ authToken: 'h9', machine_id: 'MC-H9-2' }).history;
assert.strictEqual(dlog[0].from_status, 'degraded');
assert.strictEqual(dlog[0].prev_duration_min, 120);
assert.strictEqual(dlog[1].impact.risk_until, '2026-10-05', 'ประวัติเก็บผลกระทบไว้ด้วย');
m = b.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-2', status: 'running', impact_json: JSON.stringify({ impacts: ['slow'] }) }).machine;
assert.strictEqual(m.case_id, '');
assert.deepStrictEqual(m.impact, {}, 'กลับมาทำงานปกติ = ล้างผลกระทบ');

// ── ชีตจากเวอร์ชันก่อน (ยังไม่มีคอลัมน์ Impact JSON) ต้องเติมหัวให้เอง ข้อมูลเดิมอ่านได้ ──────
const oldStatusHead = ['Machine ID', 'Line', 'Machine Name', 'Status', 'Reason Type', 'Parts JSON', 'Comment', 'PR Ref', 'Since', 'Case ID', 'Updated By', 'Updated At'];
const legacy = makeMachineStatusBackend({ users: USERS, sheets: {
  Machines: MACHINES.map(function(r) { return r.slice(); }),
  MachineStatus: [oldStatusHead, ['MC-H9-1', 'H9', 'H9 Press 1', 'stopped', 'breakdown', '[]', 'เก่า', '', '2026-09-30 08:00:00', 'MSC-old', 'somchai', '2026-09-30 08:00:00']],
  MachineStatusLog: [['Log ID', 'Case ID', 'Machine ID', 'Line', 'Machine Name', 'From Status', 'To Status', 'Reason Type', 'Parts JSON', 'Comment', 'PR Ref', 'Changed By', 'Changed At', 'Prev Duration Min']]
} });
const legacyRow = legacy.getMachineStatusBoard({ authToken: 'h9' }).machines[0];
assert.strictEqual(legacyRow.status, 'stopped');
assert.deepStrictEqual(legacyRow.impact, {}, 'แถวเก่าไม่มีคอลัมน์ผลกระทบ = ว่าง');
assert.strictEqual(legacy.sheets.MachineStatus.rows[0][12], 'Impact JSON', 'เติมหัวคอลัมน์ใหม่ต่อท้าย');
legacy.updateMachineStatus({ authToken: 'h9', machine_id: 'MC-H9-1', status: 'degraded', impact_json: JSON.stringify({ impacts: ['workaround'] }) });
assert.strictEqual(legacy.sheets.MachineStatusLog.rows[0][14], 'Impact JSON');
assert.strictEqual(legacy.getMachineStatusBoard({ authToken: 'h9' }).machines[0].case_id, 'MSC-old', 'หยุด → รันบางส่วน ยังเป็นเคสเดิม');

// ── backend: route + export whitelist ──────────────────────────────────────
['getMachineStatusBoard', 'updateMachineStatus', 'getMachineStatusHistory'].forEach(function(action) {
  assert.strictEqual((backend.match(new RegExp("action === '" + action + "'", 'g')) || []).length, 2, action + ' ต้องมีทั้ง doGet และ doPost');
});
assert(/'MachineStatus', 'MachineStatusLog'/.test(backend), 'ชีตใหม่ต้องอยู่ใน whitelist ของ Export');

// ════════════════════ ฝั่งเว็บ ════════════════════
const web = new Function('partsData', [
  grabVar('MS_DOWN_STATUSES', html),
  grabVar('MS_ISSUE_STATUSES', html),
  grabFn('safeNum'),
  grabFn('getItemIdentityKey'),
  grabFn('findItemByIdentityKey'),
  grabFn('msParseTime'),
  grabFn('msDuration'),
  grabFn('msHours'),
  grabFn('msBuildSegments'),
  grabFn('msFindPartItem'),
  grabFn('msPartStock'),
  grabFn('msPartHasEnough'),
  grabFn('msMissingParts'),
  grabFn('msPartsReady'),
  grabFn('msLostCapacityMs'),
  grabFn('escHtml'),
  grabFn('msDateLabel'),
  grabFn('msTodayYmd'),
  grabFn('msImpactPillsHtml'),
  'return { msParseTime: msParseTime, msDuration: msDuration, msHours: msHours, msBuildSegments: msBuildSegments,' +
  ' msMissingParts: msMissingParts, msPartsReady: msPartsReady, msPartStock: msPartStock,' +
  ' msLostCapacityMs: msLostCapacityMs, msImpactPillsHtml: msImpactPillsHtml };'
].join('\n'));

const stockItem = { no: '12', subLine: 'Stock for MC', __sourceSheet: 'Stock for MC', model: '6204ZZ', name: 'Bearing', location: 'A-02', stock: 1 };
const w = web([stockItem]);

// เวลาเป็นเวลาไทยเสมอ ไม่ขึ้นกับ timezone ของเครื่องที่เปิด
assert.strictEqual(w.msParseTime('2026-10-01 08:00:00').toISOString(), '2026-10-01T01:00:00.000Z');
assert.strictEqual(w.msParseTime(''), null);
assert.strictEqual(w.msDuration(30 * 1000), 'ไม่ถึง 1 นาที');
assert.strictEqual(w.msDuration(12 * 60000), '12 นาที');
assert.strictEqual(w.msDuration(3 * 3600000), '3 ชม.');
assert.strictEqual(w.msDuration(3 * 3600000 + 20 * 60000), '3 ชม. 20 นาที');
assert.strictEqual(w.msDuration(52 * 3600000), '2 วัน 4 ชม.');
assert.strictEqual(w.msHours(90 * 60000), '1.5 ชม.');

// อะไหล่: location ย้ายแล้ว identity key ไม่ตรง ต้องถอยไปจับชื่อ+รุ่น+ชีต และใช้สต็อกสด
const waiting = { status: 'waiting_parts', parts: [Object.assign({}, bearing, { qty_needed: 2 })] };
assert.strictEqual(w.msPartStock(waiting.parts[0]).stock, 1, 'ใช้สต็อกสดจาก partsData ไม่ใช่ยอดตอนบันทึก');
assert.strictEqual(w.msMissingParts(waiting).length, 1, 'มี 1 ต้องใช้ 2 = ยังขาด');
assert.strictEqual(w.msPartsReady(waiting), false);
stockItem.stock = 2;
assert.strictEqual(w.msMissingParts(waiting).length, 0);
assert.strictEqual(w.msPartsReady(waiting), true, 'ของเข้าครบ = ขึ้นป้าย "อะไหล่มาแล้ว"');
const withManual = { status: 'waiting_parts', parts: waiting.parts.concat([{ name: 'ซีลยาง', manual: true, qty_needed: 1 }]) };
assert.strictEqual(w.msPartsReady(withManual), false, 'มีอะไหล่ที่พิมพ์เอง ระบบรู้ไม่ได้ว่ามาหรือยัง = ไม่ขึ้นป้าย');
assert.strictEqual(w.msMissingParts(withManual).length, 0, 'ตัวที่พิมพ์เองไม่นับเป็นของที่กดสั่งซื้อได้');
assert.strictEqual(w.msPartsReady({ status: 'running', parts: waiting.parts }), false);
const w2 = web([]);
assert.strictEqual(w2.msPartStock(bearing).stock, 0, 'ไม่มี partsData ของไลน์นี้ = ถอยไปใช้ยอด ณ ตอนบันทึก');
assert.strictEqual(w2.msPartStock({ name: 'X', stock_at_save: '' }), null);

// ช่วงเวลาเครื่องหยุด: ตัดตามหน้าต่างเวลา + ช่วงที่ยังไม่จบนับถึงตอนนี้
const H = 3600000;
const now = Date.parse('2026-10-10T01:00:00Z');
const winStart = now - 24 * H;
const hist = [ // ใหม่ → เก่า เหมือนที่ backend ส่งมา
  { machine_id: 'A', line: 'H9', machine_name: 'A', from_status: 'maintenance', to_status: 'running', case_id: 'C1', changed_at: '2026-10-10 06:00:00', parts: [] },
  { machine_id: 'A', line: 'H9', machine_name: 'A', from_status: 'waiting_parts', to_status: 'maintenance', reason_type: 'spare_part', case_id: 'C1', changed_at: '2026-10-10 04:00:00', parts: [] },
  { machine_id: 'A', line: 'H9', machine_name: 'A', from_status: 'running', to_status: 'waiting_parts', reason_type: 'spare_part', case_id: 'C1', changed_at: '2026-10-09 20:00:00', parts: [bearing] },
  { machine_id: 'B', line: 'H9', machine_name: 'B', from_status: 'stopped', to_status: 'stopped', reason_type: 'breakdown', case_id: 'C0', changed_at: '2026-10-10 07:00:00', parts: [] }
];
const machinesNow = [
  { machine_id: 'A', status: 'running' },
  { machine_id: 'B', status: 'stopped' },
  { machine_id: 'C', line: 'H9', machine_name: 'C', status: 'stopped', reason_type: 'no_job', since: '2026-09-01 08:00:00', parts: [] }
];
const segs = w.msBuildSegments(hist, machinesNow, winStart, now);
function hoursOf(id, status) {
  return segs.filter(function(s) { return s.machine_id === id && (!status || s.status === status); })
    .reduce(function(t, s) { return t + (s.end - s.start); }, 0) / H;
}
assert.strictEqual(hoursOf('A', 'waiting_parts'), 8, 'รออะไหล่ 20:00 → 04:00');
assert.strictEqual(hoursOf('A', 'maintenance'), 2, 'ซ่อม 04:00 → 06:00');
assert.strictEqual(hoursOf('A', 'running'), 0, 'ช่วงที่เดินเครื่องไม่นับ');
assert.strictEqual(hoursOf('B'), 24, 'หยุดมาก่อนช่วงที่เลือก: ตัดที่ต้นหน้าต่าง + ช่วงล่าสุดนับถึงตอนนี้');
assert.strictEqual(hoursOf('C'), 24, 'หยุดมานานแต่ไม่มี log ในช่วงนี้ ก็ต้องนับ (ตัดที่ต้นหน้าต่าง)');
assert.strictEqual(segs.filter(function(s) { return s.machine_id === 'A'; })[0].case_id, 'C1');


// รันได้บางส่วน: นับเป็นช่วงมีปัญหา (แยกจากชั่วโมงหยุด) + กำลังผลิตที่เสียไปคิดจาก % ที่กรอก
const degHist = [
  { machine_id: 'D', line: 'H9', machine_name: 'D', from_status: 'degraded', to_status: 'running', case_id: 'C9', changed_at: '2026-10-10 06:00:00', parts: [] },
  { machine_id: 'D', line: 'H9', machine_name: 'D', from_status: 'running', to_status: 'degraded', case_id: 'C9', changed_at: '2026-10-09 20:00:00', parts: [bearing],
    impact: { impacts: ['slow'], capacity_pct: 70 } }
];
const degSegs = w.msBuildSegments(degHist, [{ machine_id: 'D', status: 'running' }], winStart, now);
assert.strictEqual(degSegs.length, 1);
assert.strictEqual(degSegs[0].status, 'degraded');
assert.strictEqual((degSegs[0].end - degSegs[0].start) / H, 10);
assert.strictEqual(w.msLostCapacityMs(degSegs[0]) / H, 3, 'รัน 70% นาน 10 ชม. = เสียไป 3 ชม.เครื่อง');
assert.strictEqual(w.msLostCapacityMs(Object.assign({}, degSegs[0], { impact: { impacts: ['slow'] } })), 0, 'ไม่ใส่ % = ไม่เดาตัวเลข');
assert.strictEqual(w.msLostCapacityMs(Object.assign({}, degSegs[0], { impact: { impacts: ['workaround'], capacity_pct: 70 } })), 0, 'ไม่ได้ติ๊กกำลังผลิตลด = ไม่นับ');
assert.strictEqual(w.msLostCapacityMs(Object.assign({}, degSegs[0], { status: 'stopped' })), 0);
// ป้ายเสี่ยงหยุด: เลยกำหนดแล้วต้องขึ้นแดง
assert(w.msImpactPillsHtml({ impacts: ['at_risk'], risk_until: '2020-01-01' }).indexOf('เลยกำหนดเปลี่ยนแล้ว') > -1);
assert(w.msImpactPillsHtml({ impacts: ['at_risk'], risk_until: '2999-01-31' }).indexOf('ภายใน 31/01') > -1);
assert(w.msImpactPillsHtml({ impacts: ['slow'], capacity_pct: 70 }).indexOf('กำลังผลิต ~70%') > -1);
assert.strictEqual(w.msImpactPillsHtml({}), '');

// ── เดินสายหน้าใหม่ ────────────────────────────────────────────────────────
assert(html.includes('id="tabMachineStatus"'), 'มีแท็บเมนูของหน้าใหม่');
assert(html.includes('id="machineStatusPage"'), 'มีหน้าใหม่แยก');
assert(html.includes('<style id="msStyles">'), 'สีทั้งหมดต้องเป็น CSS จริง (Tailwind v2 static ไม่มี slate/rose/emerald)');
// ทุกฟังก์ชันสลับหน้าต้องซ่อนหน้านี้ ไม่งั้นสองหน้าซ้อนกัน
const switchers = html.match(/^ {4}function switchTo\w+Page\(\) \{[\s\S]*?\n {4}\}/gm) || [];
assert(switchers.length >= 12, 'หาฟังก์ชันสลับหน้าเจอครบ');
switchers.forEach(function(fn) {
  const name = fn.match(/function (\w+)/)[1];
  if (name === 'switchToMachineStatusPage') return;
  assert(fn.includes('msHidePage();'), name + ' ต้องซ่อนหน้าสถานะเครื่องจักร');
});
const own = switchers.filter(function(fn) { return fn.trim().indexOf('function switchToMachineStatusPage') === 0; })[0];
assert(own && own.includes('xpHidePage();') && own.includes('mxHidePage();') && own.includes("setTabStyles('machine-status')"),
  'หน้าใหม่ต้องซ่อนหน้าอื่นทั้งหมดเอง');
const tabStyles = grabFn('setTabStyles');
assert(tabStyles.includes("active === 'machine-status'"), 'แท็บ active ได้');
assert(tabStyles.includes("tabMachineStatusEl.classList.toggle('hidden', !hasPermission('view'))"), 'ซ่อนแท็บถ้าไม่มีสิทธิ์ดู');
assert(/loadLogData\(\);\n {10}msOnLineChanged\(\);/.test(html), 'สลับไลน์จากเมนูซ้าย = โหลดเครื่องของไลน์ใหม่');
assert(grabFn('applyAuthorizationUi').includes('msLoadBadge()'), 'badge เครื่องไม่ได้ทำงานโหลดหลัง login');
// ปุ่มสั่งซื้อต้องใช้ flow ติดดาว → PR เดิม (ผ่านงบ/อนุมัติเหมือนใบอื่น)
const buy = grabFn('msBuyParts');
assert(buy.includes('buildStarredEntry(') && buy.includes('openPrFromStarred()'), 'สั่งซื้อ = ติดดาวแล้วเปิด PR');
assert(buy.includes("hasPermission('view_logs')"), 'คนที่เข้าหน้า PR ไม่ได้ ติดดาวไว้ให้หัวหน้า');
// เปลี่ยนสถานะต้องเช็คสิทธิ์ทั้งจาก backend (can_edit) และสิทธิ์เขียนคลัง
assert(/function msCanEdit\(m\) \{ return !!\(m && m\.can_edit && canWriteWarehouse\(\)\); \}/.test(html));

assert(/var MS_RUN_STATUSES = \['running', 'degraded'\];/.test(html), 'รันได้บางส่วนนับเป็นเครื่องพร้อมรัน');
assert(grabFn('msUpdateBadgeFromList').includes('MS_DOWN_STATUSES.indexOf(m.status) > -1'), 'badge นับเฉพาะเครื่องที่ไม่ได้ผลิต (ไม่รวมรันบางส่วน)');

// ── style block ของ Dashboard ต้องปิดก่อนเปิดบล็อกถัดไป ────────────────────────────
// เดิม <style id="fccStyles"> ไม่มี </style> → แท็ก <style> ถัดไปกลายเป็นข้อความ CSS ขยะ
// กินกฎข้อแรกของบล็อกถัดไปทิ้ง (badge แจ้งเตือนเลยไม่มีพื้นแดง เห็นแต่ตัวเลขสีขาว)
// ดูเฉพาะใน <head> — ในสคริปต์มีสตริง '<style>' สำหรับหน้าพิมพ์ ซึ่งไม่ใช่แท็กจริง
const styleTags = html.slice(0, html.indexOf('</head>')).match(/<\/?style[^>]*>/g) || [];
let depth = 0;
styleTags.forEach(function(t) {
  depth += t[1] === '/' ? -1 : 1;
  assert(depth === 0 || depth === 1, 'แท็ก <style> ห้ามซ้อนกัน/ห้ามลืมปิด: ' + t);
});
assert.strictEqual(depth, 0, 'ทุก <style> ต้องมี </style>');
assert(/#machineStatusNavBadge \{ background-color:#f43f5e;/.test(html), 'badge เครื่องจักรต้องมีพื้นแดงจริง (Tailwind v2 ไม่มี bg-rose-500)');

// ── เลือกอะไหล่ได้หลายตัว: รายการไม่ปิดหลังกดเลือก และกดซ้ำ = เอาออก ─────────────────
const pickHandler = html.slice(html.indexOf("if (results) results.addEventListener('click'"), html.indexOf('// คลิกนอกช่องค้นหา'));
assert(pickHandler.includes('msTogglePickedItem(it);') && pickHandler.includes('msRenderPartResults();'), 'เลือกแล้ววาดรายการใหม่ ไม่ปิด');
assert(!pickHandler.includes("msEl('msPartSearch').value = ''"), 'ห้ามล้างคำค้นหลังเลือก (จะเลือกตัวถัดไปไม่ได้)');
assert(html.includes('if (!e.target.isConnected) return;'), 'แถวที่ถูกวาดใหม่ต้องไม่นับเป็นคลิกนอกรายการ');
assert(grabFn('msTogglePickedItem').includes('msEditState.parts.splice(at, 1)'), 'กดตัวที่เลือกแล้ว = เอาออก');

// ── ชิปกรองไลน์บนหน้า = ทางเดียวกับเมนูไลน์ซ้าย ───────────────────────────────────
assert(html.includes('id="msLineChips"'));
const sel = grabFn('msSelectLine');
['setCurrentLine(key)', 'renderLineTabs()', 'applyLineTheme()', 'loadPartsData()', 'msOnLineChanged()'].forEach(function(call) {
  assert(sel.includes(call), 'เลือกไลน์จากชิปต้องเรียก ' + call + ' เหมือนเมนูซ้าย');
});

console.log('machine status checks passed');
