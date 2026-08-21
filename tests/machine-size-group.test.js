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

// COIL_SIZE_OPTIONS + canonicalCoilSize คือฐานของการจับคู่ ต้องดึงของจริงมาใช้
const optsLine = script.match(/var COIL_SIZE_OPTIONS = \[[^\]]*\];/)[0];
const bundle = [
  '    ' + optsLine,
  extract('canonicalCoilSize(value)'),
  extract('getMachineSizeTag(machine)'),
  extract('getMachineSizesForCoilSize(coilSize)'),
  extract('groupMachinesBySize(list)')
].join('\n');
const api = new Function(bundle +
  '\nreturn { COIL_SIZE_OPTIONS: COIL_SIZE_OPTIONS, getMachineSizeTag: getMachineSizeTag,' +
  ' getMachineSizesForCoilSize: getMachineSizesForCoilSize, groupMachinesBySize: groupMachinesBySize };')();

// ── เครื่องจักรจริงของไลน์ Coil Winding (16 ตัว ตามที่ลงทะเบียนไว้) ──────────
const CWM = [
  'CWM-01 (10A)', 'CWM-02 (16A)', 'CWM-03 (20A)', 'CWM-04 (10A)',
  'CWM-05 (16A)', 'CWM-06 (20A)', 'CWM-07 (32A)', 'CWM-08 (32A)',
  'CWM-09 (32A)', 'CWM-10 (16A)', 'CWM-11 (16A)', 'CWM-12 (32A)',
  'CWM-13 (32A)', 'CWM-14 (32A)', 'CWM-15 (20A)', 'CWM-16 (10A)'
].map(function(n) { return { machine_name: n }; });

// ── แกะขนาดออกจากชื่อเครื่อง ─────────────────────────────────────────────────
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'CWM-07 (32A)' }), '32A');
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'CWM-01 (10A)' }), '10A');
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'CWM-17 32A' }), '32A', 'ไม่มีวงเล็บก็ต้องแกะได้');
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'CWM-18 ( 16A )' }), '16A', 'มีช่องว่างในวงเล็บก็ต้องได้');
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'cwm-19 (20a)' }), '20A', 'ตัวพิมพ์เล็กต้องได้');
assert.strictEqual(api.getMachineSizeTag({ machine_name: 'Winding Machine 1' }), '', 'ไม่มีขนาดต้องคืนค่าว่าง');
assert.strictEqual(api.getMachineSizeTag({ machine_name: '' }), '');

// ── Coil Size ของอะไหล่ → ขนาดเครื่องที่เข้าได้ ──────────────────────────────
assert.deepStrictEqual(api.getMachineSizesForCoilSize('10A'), ['10A']);
assert.deepStrictEqual(api.getMachineSizesForCoilSize('16A'), ['16A']);
assert.deepStrictEqual(api.getMachineSizesForCoilSize('20A'), ['20A']);
// "25/32A" ต้องครอบทั้ง 25A และ 32A — ตอนนี้ลงทะเบียนไว้แค่ 32A แต่เผื่อวันหลังมี 25A
assert.deepStrictEqual(api.getMachineSizesForCoilSize('25/32A'), ['25A', '32A']);
// ค่าเก่าที่ไม่มี A ต้อง canonical ให้ตรงกันก่อน
assert.deepStrictEqual(api.getMachineSizesForCoilSize('25/32'), ['25A', '32A']);
assert.deepStrictEqual(api.getMachineSizesForCoilSize('16'), ['16A']);
// Common = ไม่จำกัดกลุ่ม, ไม่ระบุ = ไม่จำกัดกลุ่ม
assert.strictEqual(api.getMachineSizesForCoilSize('Common'), null, 'Common ต้องใช้ได้ทุกเครื่อง');
assert.strictEqual(api.getMachineSizesForCoilSize(''), null);
assert.strictEqual(api.getMachineSizesForCoilSize('-'), null);
assert.strictEqual(api.getMachineSizesForCoilSize('ไม่มีตัวเลข'), null, 'ค่าที่ไม่มีตัวเลขต้องไม่จำกัดกลุ่ม');

// ทุกค่าใน COIL_SIZE_OPTIONS ต้องแปลงได้ ไม่ระเบิด
api.COIL_SIZE_OPTIONS.forEach(function(opt) {
  const r = api.getMachineSizesForCoilSize(opt);
  assert(r === null || Array.isArray(r), 'ค่า ' + opt + ' ต้องคืน null หรือ array');
});

// ── จัดกลุ่มเครื่องจักรจริง ──────────────────────────────────────────────────
const groups = api.groupMachinesBySize(CWM);
assert.deepStrictEqual(groups.map(function(g) { return g.size; }), ['10A', '16A', '20A', '32A'],
  'ต้องเรียงกลุ่มตามตัวเลข 10A → 16A → 20A → 32A');
assert.deepStrictEqual(groups.map(function(g) { return g.machines.length; }), [3, 4, 3, 6],
  'จำนวนเครื่องแต่ละกลุ่มต้องตรงกับที่ลงทะเบียนไว้');

// กลุ่ม 32A ต้องได้ CWM-07,08,09,12,13,14 พอดี (ตรงกับที่เคยติ๊กมือไว้)
const g32 = groups.filter(function(g) { return g.size === '32A'; })[0];
assert.deepStrictEqual(g32.machines.map(function(m) { return m.machine_name; }),
  ['CWM-07 (32A)', 'CWM-08 (32A)', 'CWM-09 (32A)', 'CWM-12 (32A)', 'CWM-13 (32A)', 'CWM-14 (32A)'],
  'กลุ่ม 32A ต้องตรงกับที่เคยติ๊กมือไว้เป๊ะ');

// อะไหล่ 25/32A → ติ๊กได้ 6 ตัวพอดี
const sizes2532 = api.getMachineSizesForCoilSize('25/32A');
const picked = CWM.filter(function(m) { return sizes2532.indexOf(api.getMachineSizeTag(m)) > -1; });
assert.strictEqual(picked.length, 6, 'อะไหล่ 25/32A ต้องจับคู่ได้ 6 เครื่อง');

// ── เครื่องที่ไม่ได้แปะขนาด ต้องไม่หายไป ─────────────────────────────────────
const mixed = api.groupMachinesBySize(CWM.concat([{ machine_name: 'Winding Machine เก่า' }]));
const lastGroup = mixed[mixed.length - 1];
assert.strictEqual(lastGroup.size, '', 'กลุ่มไม่ระบุขนาดต้องอยู่ท้ายสุด');
assert.strictEqual(lastGroup.label, 'ไม่ระบุขนาด');
assert.strictEqual(lastGroup.machines.length, 1, 'เครื่องที่ไม่มีขนาดต้องยังอยู่ในรายการ ไม่ถูกกรองทิ้ง');
const totalShown = mixed.reduce(function(a, g) { return a + g.machines.length; }, 0);
assert.strictEqual(totalShown, 17, 'ทุกเครื่องต้องถูกแสดง ไม่มีตัวไหนหาย');

// ไลน์ที่ไม่มีขนาดเลย → กลุ่มเดียว ไม่ระบุขนาด (ตัว render จะ fallback เป็นรายการเรียบ)
const plain = api.groupMachinesBySize([{ machine_name: 'Machine A' }, { machine_name: 'Machine B' }]);
assert.strictEqual(plain.length, 1);
assert.strictEqual(plain[0].size, '');

// ── โค้ดฝั่ง UI ต้องต่อสายครบ ────────────────────────────────────────────────
const renderSrc = extract('renderEditMachinesCheckboxes(line, selectedNames)');
assert(renderSrc.indexOf('groupMachinesBySize(list)') > -1, 'ต้องจัดกลุ่มก่อนวาด');
assert(renderSrc.indexOf('getMachineSizesForCoilSize(coilSize)') > -1, 'ต้องรู้ว่าอะไหล่อยู่กลุ่มไหน');
assert(renderSrc.indexOf('data-machine-group-toggle') > -1, 'ต้องมีปุ่มเลือก/ล้างยกกลุ่ม');
assert(renderSrc.indexOf('ตรงกับอะไหล่') > -1, 'ต้องไฮไลต์กลุ่มที่ตรงกับ Coil Size');

// ติ๊กอัตโนมัติเฉพาะตอนยังไม่เคยผูกเครื่องไว้เลย — ห้ามทับของที่ผู้ใช้เลือกไว้แล้ว
assert(renderSrc.indexOf('if (!selected.length && matchSizes) {') > -1,
  'ต้องติ๊กอัตโนมัติเฉพาะอะไหล่ที่ยังไม่เคยผูกเครื่องไว้');

// เปลี่ยน Coil Size ต้องวาดกลุ่มใหม่ และต้องยกค่าที่ติ๊กไว้ไปด้วย
assert(/\['eCoilSize', 'eCoilSizeOther'\][\s\S]{0,400}renderEditMachinesCheckboxes\([\s\S]{0,80}getEditMachinesSelection\(\)\)/.test(script),
  'เปลี่ยน Coil Size ต้อง re-render โดยยกค่าที่ติ๊กไว้ไปด้วย');

console.log('Machine size group checks passed');
