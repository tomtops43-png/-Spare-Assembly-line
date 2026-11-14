// Inbox PR รออนุมัติ / อนุมัติแล้ว — เดิมเรียงตามลำดับในชีต ดูไม่ออกว่าใบไหนล่าสุด
// ตอนนี้เรียงใหม่สุดก่อน + แบ่งหัวข้อตามวัน (วันนี้ / เมื่อวาน / วันที่) และแสดงเวลาแบบอ่านง่าย
const { html, grabFn, assert } = require('./_export-extract');

const src = ['inboxParseDate', 'inboxDayKey', 'inboxDayLabel', 'inboxFmtDateTime', 'inboxGroupByDay']
  .map(function(n) { return grabFn(n, html); }).join('\n');
const consts = html.match(/var INBOX_TH_MONTHS = [^\n]+\n\s*var INBOX_TH_DAYS = [^\n]+/)[0];
const api = new Function(consts + '\n' + src + '\nreturn { inboxGroupByDay, inboxFmtDateTime, inboxDayLabel };')();

const now = new Date('2026-11-14T10:00:00+07:00').getTime();
const cards = [
  { pr_id: 'PR-2610-003', created_at: 'Wed Oct 21 2026 09:40:32 GMT+0700 (เวลาอินโดจีน)' },
  { pr_id: 'PR-2610-002', created_at: 'Wed Oct 21 2026 09:35:02 GMT+0700 (เวลาอินโดจีน)' },
  { pr_id: 'PR-2611-002', created_at: 'Wed Nov 11 2026 11:53:22 GMT+0700 (เวลาอินโดจีน)' },
  { pr_id: 'PR-202607-H9-389', created_at: '2026-07-08 13:03:44' },
  { pr_id: 'PR-2611-009', created_at: '2026-11-14 08:00:00' },
  { pr_id: 'PR-2611-008', created_at: '2026-11-13 23:30:00' },
  { pr_id: 'PR-X', created_at: '' }
];
const groups = api.inboxGroupByDay(cards, now);
assert.deepStrictEqual(groups.map(function(g) { return g.cards.map(function(c) { return c.pr_id; }); }), [
  ['PR-2611-009'], ['PR-2611-008'], ['PR-2611-002'], ['PR-2610-003', 'PR-2610-002'], ['PR-202607-H9-389'], ['PR-X']
], 'ใหม่สุดก่อน แบ่งกลุ่มตามวัน (เวลาไทย)');
assert(groups[0].label.startsWith('วันนี้'), groups[0].label);
assert(groups[1].label.startsWith('เมื่อวาน'), groups[1].label);
assert.strictEqual(groups[3].label, 'พุธ 21 ต.ค. 2026');
assert.strictEqual(groups[5].label, 'ไม่ระบุวันที่');
assert.strictEqual(api.inboxFmtDateTime('Wed Oct 21 2026 09:40:32 GMT+0700 (เวลาอินโดจีน)'), '21 ต.ค. 2026 09:40 น.');
assert(html.includes('inboxDayGroupsHtml(approved, true)'), 'กลุ่มอนุมัติแล้วใช้การแบ่งตามวัน');
assert(html.includes('inboxDayGroupsHtml(pending, false)'), 'กลุ่มรออนุมัติใช้การแบ่งตามวัน');
console.log('inbox-pr-sort-by-day: OK');
