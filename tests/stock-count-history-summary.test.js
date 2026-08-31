const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = htmlLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── เดือนของรอบนับต้องอ่านรู้เรื่อง ────────────────────────────────────────────
// เดือนถูกส่งขึ้นชีทเป็น '2026-08' (ค่าจาก input[type=month]) แต่ Google Sheets แปลงเป็น Date
// แล้วส่งกลับมาเป็น ISO UTC ที่ถอยไปเป็นสิ้นเดือนก่อนหน้า — เดิมเอามาโชว์ดิบ ๆ ประวัติจึงขึ้นว่า
// "2026-07-31T17:00:00.000Z" ทั้งที่เป็นรอบเดือนสิงหาคม อ่านไม่รู้เรื่องและเดือนผิดด้วย
const monthsConst = grab(/^ {4}var SC_THAI_MONTHS = \[[\s\S]*?\];/m, 'SC_THAI_MONTHS');
const fmtMonthSrc = grab(/^ {4}function scFmtMonth\(value\) \{[\s\S]*?\n {4}\}/m, 'scFmtMonth');
const scFmtMonth = new Function(monthsConst + '\n' + fmtMonthSrc + '\nreturn scFmtMonth;')();

// ค่าจริงที่หลุดมาโชว์บนหน้าจอ — ต้องอ่านที่โซนเวลาไทยถึงจะได้เดือนที่ถูก (17:00Z = 00:00 ไทยวันถัดไป)
assert.strictEqual(scFmtMonth('2026-07-31T17:00:00.000Z'), 'สิงหาคม 2569');
assert.strictEqual(scFmtMonth('2026-08'), 'สิงหาคม 2569', 'ค่าดิบจาก input[type=month] ต้องได้ผลเดียวกัน');
// ข้ามปี: 31 ธ.ค. 17:00Z = 1 ม.ค. เวลาไทย → ต้องเป็นรอบเดือนมกราคมของปีถัดไป
assert.strictEqual(scFmtMonth('2026-12-31T17:00:00.000Z'), 'มกราคม 2570');
assert.strictEqual(scFmtMonth('2026-1'), 'มกราคม 2569', 'เดือนเลขหลักเดียวต้องยังอ่านได้');
assert.strictEqual(scFmtMonth(''), '-');
assert.strictEqual(scFmtMonth(null), '-');
// ค่าที่ parse ไม่ได้ต้องคืนของเดิม ไม่ใช่ขึ้น Invalid Date / NaN
assert.strictEqual(scFmtMonth('ขยะ'), 'ขยะ');

// ── รวมแถวประวัติเป็นรอบนับ ────────────────────────────────────────────────────
// 1 รอบมีได้หลายคนนับ (นับซ้ำทาน) ถ้าโชว์ทีละแถวจะเห็นเดือนเดียวกันซ้ำหลายบรรทัด
// โดยไม่รู้ว่าเป็นรอบเดียวกัน — สรุปลงกลุ่มแล้วคนอ่านนึกว่านับ 2 รอบแยกกัน
const groupSrc = grab(/^ {4}function scGroupHistoryRows\(rows\) \{[\s\S]*?\n {4}\}/m, 'scGroupHistoryRows');
const scGroupHistoryRows = new Function(groupSrc + '\nreturn scGroupHistoryRows;')();

const rows = [
  { session_id: 'SC-1', count_group_id: 'SCG-202608-COIL-ALL', round_no: 1, created_by: 'ช่างเอ',
    month: '2026-07-31T17:00:00.000Z', line: 'Coil Winding', category: 'all',
    submitted_at: '2026-08-31 09:00:00', status: 'approved', is_pending: false,
    total_items: 74, matched: 62, diff_count: 12, adjusted_count: '', approved_by: 'Admin', approved_at: '2026-08-31 10:36:32' },
  { session_id: 'SC-2', count_group_id: 'SCG-202608-COIL-ALL', round_no: 2, created_by: 'Admin',
    month: '2026-07-31T17:00:00.000Z', line: 'Coil Winding', category: 'all',
    submitted_at: '2026-08-31 10:30:00', status: 'approved', is_pending: false,
    total_items: 74, matched: 62, diff_count: 12, adjusted_count: 12, approved_by: 'Admin', approved_at: '2026-08-31 10:36:32' },
  { session_id: 'SC-legacy', count_group_id: '', round_no: '', created_by: 'Admin',
    month: '2026-07-31T17:00:00.000Z', line: 'H9', category: 'all',
    submitted_at: '2026-08-31 13:24:41', status: 'approved', is_pending: false,
    total_items: 57, matched: 41, diff_count: 16, adjusted_count: 16, approved_by: 'Admin', approved_at: '2026-08-31 13:24:41' },
];
const groups = scGroupHistoryRows(rows);
assert.strictEqual(groups.length, 2, 'สองแถวที่อยู่รอบเดียวกันต้องยุบเป็นรายการเดียว');
assert.strictEqual(groups[0].rows.length, 2, 'รอบแรกต้องมีผู้นับ 2 คน');
assert.strictEqual(groups[0].rows[0].round_no, 1, 'ต้องเรียงตามลำดับรอบ');
assert.strictEqual(groups[0].total, 74);
assert.strictEqual(groups[0].pct, 84, '62/74 = 84%');
// adjusted_count ถูกประทับไว้ที่แถวเดียวของรอบ ไม่ใช่ทุกแถว ต้องกวาดหาทั้งกลุ่ม
assert.strictEqual(groups[0].adjusted, 12, 'ต้องเจอจำนวนที่ปรับแม้จะอยู่คนละแถวกับแถวล่าสุด');
// แถวเก่าที่ยังไม่มี count_group_id (ส่งก่อนมีฟีเจอร์นี้) ต้องยังแสดงได้ ไม่ถูกยุบรวมมั่ว
assert.strictEqual(groups[1].rows.length, 1);
assert.strictEqual(groups[1].line, 'H9');

// ── ข้อความสรุปสำหรับวางในกลุ่มแชท ─────────────────────────────────────────────
const lineLabelSrc = grab(/^ {4}function scLineLabel\(line\) \{[\s\S]*?\n {4}\}/m, 'scLineLabel');
const fmtTimeSrc = grab(/^ {4}function scFmtTime\(value\) \{[\s\S]*?\n {4}\}/m, 'scFmtTime');
const summarySrc = grab(/^ {4}function scHistorySummaryText\(g\) \{[\s\S]*?\n {4}\}/m, 'scHistorySummaryText');
const scHistorySummaryText = new Function(
  monthsConst + '\n' + fmtMonthSrc + '\n' + lineLabelSrc + '\n' + fmtTimeSrc + '\n' + summarySrc +
  '\nreturn scHistorySummaryText;')();

const text = scHistorySummaryText(groups[0]);
// ต้องอ่านรู้เรื่องโดยไม่ต้องเปิดระบบดูประกอบ — คนในกลุ่มไม่ได้เปิดหน้าจออยู่
assert(text.indexOf('สิงหาคม 2569') > -1, 'สรุปต้องบอกเดือนแบบอ่านรู้เรื่อง');
assert(text.indexOf('Coil Winding') > -1, 'สรุปต้องบอกไลน์');
assert(text.indexOf('รอบ 1: ช่างเอ') > -1 && text.indexOf('รอบ 2: Admin') > -1,
  'สรุปต้องบอกว่าใครนับรอบไหน');
assert(text.indexOf('74') > -1 && text.indexOf('62') > -1 && text.indexOf('12') > -1,
  'สรุปต้องมี ทั้งหมด/ตรง/ไม่ตรง ครบ');
assert(text.indexOf('ปรับยอดแล้ว: 12') > -1, 'สรุปต้องบอกจำนวนที่ปรับยอดไปจริง');
assert(text.indexOf('ความแม่นยำ: 84%') > -1, 'สรุปต้องมีความแม่นยำ');
assert(text.indexOf('อนุมัติโดย Admin') > -1, 'สรุปต้องบอกว่าใครอนุมัติ');
assert(text.indexOf('undefined') === -1 && text.indexOf('NaN') === -1 && text.indexOf('Invalid') === -1,
  'สรุปห้ามมี undefined/NaN/Invalid Date หลุดไปให้คนในกลุ่มเห็น');
// ต้องเป็นข้อความหลายบรรทัดจริง ไม่ใช่ก้อนเดียวยาว ๆ ที่วางในแชทแล้วอ่านไม่ออก
assert(text.split('\n').length >= 10, 'สรุปต้องแตกบรรทัดให้อ่านง่ายในแชท');

// รอบที่ยังไม่อนุมัติต้องบอกสถานะให้ชัด ไม่ใช่ปล่อยว่างให้เข้าใจผิดว่าปิดรอบแล้ว
const pendingGroup = scGroupHistoryRows([
  Object.assign({}, rows[0], { status: 'pending_approval', is_pending: true, approved_by: '', adjusted_count: '' })
])[0];
const pendingText = scHistorySummaryText(pendingGroup);
assert(pendingText.indexOf('รออนุมัติ') > -1, 'รอบที่ยังไม่อนุมัติต้องบอกว่ารออนุมัติ');
assert(pendingText.indexOf('อนุมัติโดย') === -1, 'ยังไม่อนุมัติต้องไม่ขึ้นบรรทัดผู้อนุมัติ');
assert(pendingText.indexOf('ปรับยอดแล้ว') === -1, 'ยังไม่อนุมัติต้องไม่ขึ้นว่าปรับยอดแล้ว');

// ── สรุปรวมทุกไลน์ของเดือนล่าสุด ───────────────────────────────────────────────
// ที่โพสต์ลงกลุ่มจริงคือภาพรวมทั้งโรงงาน ไม่ใช่ทีละไลน์
const monthSumSrc = grab(/^ {4}function scHistoryMonthSummaryText\(groups\) \{[\s\S]*?\n {4}\}/m, 'scHistoryMonthSummaryText');
const scHistoryMonthSummaryText = new Function(
  monthsConst + '\n' + fmtMonthSrc + '\n' + lineLabelSrc + '\n' + monthSumSrc +
  '\nreturn scHistoryMonthSummaryText;')();
const all = scHistoryMonthSummaryText(groups);
assert(all.indexOf('สิงหาคม 2569') > -1, 'สรุปรวมต้องบอกเดือน');
assert(all.indexOf('รวมทุกไลน์: 131') > -1, 'ต้องบวกยอดข้ามไลน์ให้ถูก (74 + 57)');
assert(all.indexOf('ตรงกับระบบ: 103') > -1, '62 + 41 = 103');
assert(all.indexOf('ไม่ตรง: 28') > -1, '12 + 16 = 28');
assert(all.indexOf('Coil Winding') > -1 && all.indexOf('H9') > -1, 'ต้องแตกรายไลน์ให้ดูด้วย');
assert(all.indexOf('ช่างเอ') > -1 && all.indexOf('Admin') > -1, 'ต้องรวมรายชื่อผู้นับ');
assert(all.indexOf('undefined') === -1 && all.indexOf('NaN') === -1,
  'สรุปรวมห้ามมี undefined/NaN หลุดไป');

// เดือนคนละเดือนต้องไม่ถูกบวกรวมกัน — สรุปเดือนนี้ต้องเป็นของเดือนนี้เท่านั้น
const mixed = scGroupHistoryRows([
  rows[0],
  Object.assign({}, rows[2], { session_id: 'SC-old', month: '2026-06-30T17:00:00.000Z', total_items: 999, matched: 999, diff_count: 0 })
]);
const mixedText = scHistoryMonthSummaryText(mixed);
assert(mixedText.indexOf('999') === -1, 'รอบของเดือนอื่นต้องไม่ถูกนับรวมในสรุปเดือนนี้');
assert(mixedText.indexOf('รวมทุกไลน์: 74') > -1, 'ต้องเหลือเฉพาะยอดของเดือนล่าสุด');

console.log('stock-count-history-summary: OK');
