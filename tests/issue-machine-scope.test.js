const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

function extract(sig) {
  const start = script.indexOf('    function ' + sig);
  assert(start > -1, 'cannot find ' + sig);
  const end = script.indexOf('\n    }', start);
  assert(end > -1, 'cannot find end of ' + sig);
  return script.slice(start, end + 6);
}

const api = new Function([
  extract('parseItemMachines(item)'),
  extract('getSharedMachineNames(items)')
].join('\n') + '\nreturn { parseItemMachines: parseItemMachines, getSharedMachineNames: getSharedMachineNames };')();

const M16 = ['CWM-02 (16A)', 'CWM-05 (16A)', 'CWM-10 (16A)', 'CWM-11 (16A)'];
const M32 = ['CWM-07 (32A)', 'CWM-08 (32A)', 'CWM-09 (32A)', 'CWM-12 (32A)', 'CWM-13 (32A)', 'CWM-14 (32A)'];

// ── แปลงคอลัมน์ Machines ของอะไหล่ ───────────────────────────────────────────
assert.deepStrictEqual(api.parseItemMachines({ machines: M16.join(', ') }), M16);
assert.deepStrictEqual(api.parseItemMachines({ machines: 'CWM-02 (16A) ,  CWM-05 (16A)' }), ['CWM-02 (16A)', 'CWM-05 (16A)'],
  'ต้อง trim ช่องว่างรอบชื่อออก ไม่งั้นจับคู่กับทะเบียนไม่ติด');
[undefined, null, '', '   ', ',,,'].forEach(function(v) {
  assert.deepStrictEqual(api.parseItemMachines({ machines: v }), [], 'ค่าว่างต้องได้ array ว่าง: ' + JSON.stringify(v));
});
assert.deepStrictEqual(api.parseItemMachines({}), [], 'อะไหล่ที่ไม่เคยผูกเครื่องต้องได้ array ว่าง');
assert.deepStrictEqual(api.parseItemMachines(null), []);

// ── ตะกร้า: เครื่องที่ใช้ได้ร่วมกันของทุกรายการ ──────────────────────────────
// ชิ้นเดียว → ได้เครื่องของชิ้นนั้น
assert.deepStrictEqual(api.getSharedMachineNames([{ machines: M16.join(', ') }]), M16);

// หลายชิ้นขนาดเดียวกัน → ยังได้ครบ
assert.deepStrictEqual(api.getSharedMachineNames([
  { machines: M16.join(', ') },
  { machines: M16.join(', ') }
]), M16, 'ของขนาดเดียวกันต้องเลือกได้ครบทุกเครื่อง');

// ชิ้นที่ยังไม่เคยผูกเครื่อง ต้องไม่ไปตัดตัวเลือกของชิ้นอื่น
assert.deepStrictEqual(api.getSharedMachineNames([
  { machines: M16.join(', ') },
  { machines: '' },
  {}
]), M16, 'ชิ้นที่ไม่ได้ผูกเครื่องต้องไม่ทำให้ตัวเลือกของชิ้นอื่นหาย');

// ตะกร้าที่ไม่มีชิ้นไหนผูกเครื่องเลย → ไม่จำกัด (คืน array ว่าง = แสดงทุกเครื่อง)
assert.deepStrictEqual(api.getSharedMachineNames([{ machines: '' }, {}]), []);
assert.deepStrictEqual(api.getSharedMachineNames([]), []);

// ── ของคนละขนาดปนกันในตะกร้า → ตัดกันไม่เหลือ = ไม่จำกัด (ถอยไปแสดงทุกเครื่อง) ──
// สำคัญ: ต้องไม่ทำให้ dropdown ว่างจนเบิกไม่ได้
assert.deepStrictEqual(api.getSharedMachineNames([
  { machines: M16.join(', ') },
  { machines: M32.join(', ') }
]), [], 'ของ 16A ปนกับ 32A ต้องถอยไปแสดงทุกเครื่อง ไม่ใช่ dropdown ว่าง');

// ทับกันบางส่วน → เหลือเฉพาะตัวที่ซ้ำกัน
assert.deepStrictEqual(api.getSharedMachineNames([
  { machines: 'CWM-02 (16A), CWM-05 (16A), CWM-10 (16A)' },
  { machines: 'CWM-05 (16A), CWM-10 (16A), CWM-11 (16A)' }
]), ['CWM-05 (16A)', 'CWM-10 (16A)'], 'ต้องเหลือเฉพาะเครื่องที่ทั้งสองชิ้นใส่ได้');

// ── ตัวกรองในตัว populateMachineSelect ───────────────────────────────────────
const popSrc = extract('populateMachineSelect(selectId, otherId, line, allowedNames)');
assert(popSrc.indexOf('allowList.indexOf(n) > -1') > -1, 'ต้องกรองรายชื่อตามที่อะไหล่ผูกไว้');
assert(popSrc.indexOf('var isFiltered = filtered.length > 0;') > -1,
  'กรองแล้วไม่เหลือสักตัว (เครื่องถูกลบ/เปลี่ยนชื่อ) ต้องถอยไปแสดงทั้งหมด ไม่ใช่ปล่อยว่าง');
assert(popSrc.indexOf('setMachineSelectScopeHint(') > -1, 'ต้องบอกผู้ใช้ว่ารายการถูกตัดให้เหลือเฉพาะที่ผูกไว้');
assert(popSrc.indexOf('MACHINE_SELECT_OTHER') === -1 || script.indexOf("'<option value=\"' + MACHINE_SELECT_OTHER") > -1,
  'ต้องยังมีตัวเลือก "อื่นๆ (พิมพ์เอง)" ไว้เป็นทางออกฉุกเฉิน');

// ทางออกฉุกเฉินต้องยังอยู่จริง — ตัดตัวเลือกแล้วห้ามไม่มีทางเบิกเครื่องนอกรายการ
assert(script.indexOf('✏️ อื่นๆ (พิมพ์เอง)') > -1, 'ต้องคงตัวเลือก "อื่นๆ (พิมพ์เอง)" ไว้');

// ── ต่อสายครบทั้ง 2 หน้าจอ ───────────────────────────────────────────────────
assert(script.indexOf("populateMachineSelect('quickIssueMachine', 'quickIssueMachineOther', line || currentLine, parseItemMachines(item))") > -1,
  'เบิกด่วนต้องส่งเครื่องที่อะไหล่ผูกไว้เข้าไป');
assert(script.indexOf("populateMachineSelect('issueCartMachine', 'issueCartMachineOther', currentLine, getSharedMachineNames(items))") > -1,
  'ตะกร้าต้องส่งเครื่องที่ใช้ได้ร่วมกันเข้าไป');

console.log('Issue machine scope checks passed');
