// หน้า PR Report: ใช้ "รอบเดือน" อย่างเดียว — ช่วงวันที่ต้องเป็น 1 ถึงวันสุดท้ายของเดือนเสมอ
// บั๊กเดิม: เลือก November 2026 แต่ได้ 31-Oct-2026 ถึง 29-Nov-2026 (toISOString แปลงเป็น UTC เวลาไทยเลยเลื่อน 1 วัน)
const { html, grabFn, assert } = require('./_export-extract');

// ผู้ใช้เลือกวันที่เองไม่ได้แล้ว — เหลือแค่ช่องซ่อนให้โค้ดนับยอดเบิกอ่าน
assert(html.includes('<input id="prFrom" type="hidden" />') && html.includes('<input id="prTo" type="hidden" />'), 'ช่องวันที่ต้องถูกซ่อน');
assert(!/>\s*จากวันที่\s*\n\s*<input id="prFrom"/.test(html), 'ไม่มีป้าย "จากวันที่" ให้เลือกเองแล้ว');

const fn = grabFn('syncPrPeriodToDates');
assert(!fn.includes('toISOString'), 'ห้ามใช้ toISOString (เลื่อนวันตาม timezone)');
function run(period) {
  const els = { prPeriod: { value: period }, prFrom: { value: '' }, prTo: { value: '' }, prStep3: null };
  const doc = { getElementById: function(id) { return els[id]; } };
  new Function('document', 'renderPrBudgetPanel', fn + '\nsyncPrPeriodToDates();')(doc, function() {});
  return [els.prFrom.value, els.prTo.value];
}
assert.deepStrictEqual(run('2026-11'), ['2026-11-01', '2026-11-30'], 'พ.ย. ต้องเป็น 1–30');
assert.deepStrictEqual(run('2026-12'), ['2026-12-01', '2026-12-31'], 'ธ.ค. ข้ามปีต้องถูก');
assert.deepStrictEqual(run('2028-02'), ['2028-02-01', '2028-02-29'], 'ปีอธิกสุรทิน');
assert.deepStrictEqual(run('2026-01'), ['2026-01-01', '2026-01-31']);

// วันเริ่มต้องเป็นเที่ยงคืนเวลาเครื่อง ไม่ใช่ UTC
assert(grabFn('runPrCandidateSelection').includes("document.getElementById('prFrom').value + 'T00:00:00'"));
// เปิด Set ที่บันทึกไว้ → คิดช่วงวันที่จากรอบเดือนใหม่
assert(grabFn('loadPrSet').includes("setVal('prPeriod', set.period); syncPrPeriodToDates();"));

console.log('pr-period-month-only: OK');
