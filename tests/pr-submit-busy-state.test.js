// ปุ่ม "ส่งอนุมัติ (สร้าง PR)" ต้องบอกสถานะตั้งแต่กดจนส่งเสร็จ และกดซ้ำไม่ได้
// เดิมช่วงตรวจงบก่อนขึ้นกล่องยืนยันไม่มีอะไรบอกเลย → คนส่งกดซ้ำ และ server ออกเลข PR ใหม่ทุกครั้ง = PR ซ้ำ
const { html, grabFn, assert } = require('./_export-extract');

function el(id) {
  const cls = new Set();
  return { id, innerHTML: '<svg></svg> 📤 ส่งอนุมัติ (สร้าง PR)', className: '', disabled: false, attrs: {},
    setAttribute(k, v) { this.attrs[k] = v; },
    classList: { add(...c) { c.forEach((x) => cls.add(x)); }, remove(...c) { c.forEach((x) => cls.delete(x)); },
      toggle(c, on) { if (on) cls.add(c); else cls.delete(c); }, contains(c) { return cls.has(c); } } };
}
const dom = { prPrintBtn: el('prPrintBtn'), prSubmitStatus: el('prSubmitStatus'), prBudgetBypassBtn: el('prBudgetBypassBtn') };
dom.prSubmitStatus.innerHTML = '';
let lockCalls = 0;
const consts = ['var prSubmitBusy', 'var prPrintBtnIdleHtml', 'var PR_SUBMIT_SPINNER', 'var PR_SUBMIT_STATES']
  .map(function(k) { const i = html.indexOf(k); return html.slice(i, html.indexOf('\n    var', i + 5) > -1 && k !== 'var PR_SUBMIT_STATES' ? html.indexOf('\n', i) : html.indexOf('\n    };', i) + 7); })
  .join('\n');
const api = new Function('document', 'escHtml', 'prUpdateSubmitLock', consts + '\n' + grabFn('prSetSubmitState') +
  '\nreturn { set: prSetSubmitState, busy: function() { return prSubmitBusy; } };')(
  { getElementById: (id) => dom[id] || null }, (v) => String(v).replace(/</g, '&lt;'), () => { lockCalls++; });

const btn = dom.prPrintBtn, box = dom.prSubmitStatus;
api.set('checking');
assert(api.busy() && btn.disabled && dom.prBudgetBypassBtn.disabled, 'ตรวจงบ: ล็อกปุ่ม');
assert(btn.innerHTML.includes('animate-spin') && btn.innerHTML.includes('กำลังตรวจงบ'), 'ปุ่มมีสปินเนอร์ + ข้อความ');
assert(box.innerHTML.includes('กำลังตรวจสอบงบ') && !box.classList.contains('hidden'));
api.set('confirm');
assert(api.busy() && btn.disabled && btn.innerHTML.includes('รอยืนยัน') && box.classList.contains('hidden'));
api.set('sending');
assert(btn.innerHTML.includes('กำลังส่ง PR') && box.innerHTML.includes('ห้ามกดซ้ำ') && box.className.includes('violet'));
assert.strictEqual(btn.attrs['aria-busy'], 'true');
api.set('done', '✅ ส่ง PR-1 เข้า Inbox');
assert(api.busy() && btn.disabled && btn.innerHTML === '✅ ส่งแล้ว' && box.innerHTML.includes('PR-1') && box.className.includes('emerald'));
api.set('idle');
assert(!api.busy() && btn.innerHTML === '<svg></svg> 📤 ส่งอนุมัติ (สร้าง PR)', 'กลับเป็นปุ่มเดิม');
assert(!dom.prBudgetBypassBtn.disabled && lockCalls === 1, 'ปลดล็อกแล้วให้ prUpdateSubmitLock คิดล็อกงบต่อ');
api.set('error', '⚠️ ส่ง PR ไม่สำเร็จ: timeout');
assert(!api.busy() && box.innerHTML.includes('timeout') && box.className.includes('rose'), 'ไม่สำเร็จ: ปลดปุ่ม + แจ้งเหตุ');

// flow ใน prSubmitCreatePr
const fs0 = html.indexOf('function prSubmitCreatePr(){');
const flow = html.slice(fs0, html.indexOf("if (prPrintBtn) prPrintBtn.addEventListener('click'", fs0));
assert(fs0 > -1 && flow.length > 1000);
assert(flow.includes('if (prSubmitBusy) return;'), 'กันเรียกซ้ำระหว่างทำงาน');
assert(flow.indexOf("prSetSubmitState('checking');") < flow.indexOf('prRunBudgetGate(prByLine)'), 'ขึ้นสถานะก่อนตรวจงบ');
['prSetSubmitState(\'confirm\')', 'if (!ok) { prSetSubmitState(\'idle\'); return; }', "prSetSubmitState('sending');", "prSetSubmitState('done',", "prSetSubmitState('error',"].forEach(function(t) {
  assert(flow.includes(t), 'flow ต้องมี: ' + t);
});
assert(!flow.includes("prPrintBtn.textContent = '⏳ กำลังส่ง...'"), 'เลิกใช้แบบเดิม');
assert(html.includes('if (btn && !prSubmitBusy) {'), 'prUpdateSubmitLock ไม่ปลดปุ่มระหว่างส่ง');
assert(html.includes('id="prSubmitStatus"'));

console.log('pr-submit-busy-state: OK');
