const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

// ── ดึงตัวคำนวณ "ของใหม่" ออกมารันจริง ───────────────────────────────────────
function extract(signature) {
  const start = script.indexOf('    function ' + signature);
  assert(start > -1, 'cannot find ' + signature);
  const end = script.indexOf('\n    }', start);
  assert(end > -1, 'cannot find end of ' + signature);
  return script.slice(start, end + 6);
}
const bundle = [
  '    var NEW_ARRIVAL_WINDOW_DAYS = 7;',
  extract('parseReceivedAt(value)'),
  extract('getNewArrivalAgeDays(item)'),
  extract('isNewArrival(item)'),
  extract('getNewArrivalLabel(ageDays)')
].join('\n');
const api = new Function(bundle +
  '\nreturn { parseReceivedAt: parseReceivedAt, getNewArrivalAgeDays: getNewArrivalAgeDays,' +
  ' isNewArrival: isNewArrival, getNewArrivalLabel: getNewArrivalLabel };')();

const DAY = 24 * 60 * 60 * 1000;
function daysAgo(n) {
  const d = new Date(Date.now() - n * DAY);
  const p = function(x) { return String(x).padStart(2, '0'); };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
    ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}

// ── อยู่ในกรอบ 7 วัน = ของใหม่ ────────────────────────────────────────────────
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: daysAgo(0) }), 0, 'รับเข้าวันนี้ต้องเป็นของใหม่');
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: daysAgo(1) }), 1);
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: daysAgo(6) }), 6, 'วันที่ 6 ยังเป็นของใหม่');

// ── ครบ 7 วันแล้วต้องกลับเป็นการ์ดปกติเอง ────────────────────────────────────
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: daysAgo(7) }), null, 'ครบ 7 วันต้องเลิกเป็นของใหม่');
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: daysAgo(30) }), null);
assert.strictEqual(api.isNewArrival({ last_received_at: daysAgo(7) }), false);
assert.strictEqual(api.isNewArrival({ last_received_at: daysAgo(3) }), true);

// ── ไม่มีข้อมูล = ไม่ใช่ของใหม่ (ห้ามไฮไลต์ทั้งหน้า) ─────────────────────────
[undefined, null, '', '   ', 'ไม่ใช่วันที่'].forEach(function(v) {
  assert.strictEqual(api.isNewArrival({ last_received_at: v }), false,
    'ค่าว่าง/พังต้องไม่ถือเป็นของใหม่: ' + JSON.stringify(v));
});
assert.strictEqual(api.isNewArrival({}), false, 'รายการที่ยังไม่เคยรับเข้าต้องไม่ขึ้นป้าย');

// ── รับได้ทั้งสตริงเวลาไทย, ISO (Sheets แปลงเซลล์เป็นวันที่) และ Date object ──
assert(api.parseReceivedAt('2026-08-18 09:30:00') instanceof Date, 'ต้อง parse สตริงเวลาไทยได้');
assert(api.parseReceivedAt(new Date().toISOString()) instanceof Date, 'ต้อง parse ISO ได้');
assert(api.parseReceivedAt(new Date()) instanceof Date, 'ต้องรับ Date object ได้');
assert.strictEqual(api.isNewArrival({ last_received_at: new Date() }), true, 'Date object ต้องถือเป็นของใหม่');

// เวลาในอนาคตนิดหน่อย (นาฬิกาเครื่อง/ชีทคลาดกัน) ยังต้องถือเป็นของใหม่
assert.strictEqual(api.getNewArrivalAgeDays({ last_received_at: new Date(Date.now() + 2 * 60 * 60 * 1000) }), 0);
// แต่วันที่ในอนาคตไกลๆ (พิมพ์ผิด) ต้องไม่ขึ้นป้ายค้างตลอดไป
assert.strictEqual(api.isNewArrival({ last_received_at: new Date(Date.now() + 40 * DAY) }), false);

// ── ป้ายบอกอายุถูกต้อง ───────────────────────────────────────────────────────
assert.strictEqual(api.getNewArrivalLabel(0), 'เข้าวันนี้');
assert.strictEqual(api.getNewArrivalLabel(1), 'เข้าเมื่อวาน');
assert.strictEqual(api.getNewArrivalLabel(5), 'เข้า 5 วันก่อน');

// ── ต้องเดินสายข้อมูลครบตั้งแต่ชีทถึงการ์ด ───────────────────────────────────
const normalizeSrc = script.match(/function normalizeRecord\(item, index\)[\s\S]*?\n    \}/)[0];
assert(normalizeSrc.indexOf('last_received_at: firstDefined(') > -1,
  'normalizeRecord ต้อง map last_received_at ต่อ ไม่งั้นป้ายจะไม่ขึ้นเลย');

const visualSrc = extract('getCardVisualState(item)');
assert(visualSrc.indexOf('isNewArrival(item)') > -1, 'getCardVisualState ต้องเช็คของใหม่');
assert(visualSrc.indexOf('new-arrival-card') > -1, 'ของใหม่ต้องได้คลาสกรอบเรืองแสง');
assert(visualSrc.indexOf("status !== 'OUT' && isNewArrival(item)") > -1,
  'ของที่หมดสต็อกต้องไม่เรืองแสง');

assert(script.indexOf('renderNewArrivalBadge(item, false)') > -1, 'การ์ด desktop ต้องขึ้นป้ายแบบเต็ม (มีอายุของ)');
assert(script.indexOf('renderNewArrivalBadge(item, true)') > -1, 'การ์ด mobile ต้องขึ้นป้ายแบบย่อ');

// CSS ต้องมีจริง ไม่งั้นคลาสที่ใส่ไว้ไม่มีผล
['.new-arrival-card', '.new-arrival-badge', '@keyframes newArrivalGlow', '@keyframes newArrivalShine'].forEach(function(sel) {
  assert(html.indexOf(sel) > -1, 'ต้องมี CSS: ' + sel);
});
assert(/prefers-reduced-motion[\s\S]{0,240}new-arrival-card \{ animation: none/.test(html),
  'ต้องปิดอนิเมชันให้เครื่องที่ตั้งค่าลดการเคลื่อนไหว');

// ── backend: ประทับเวลาเฉพาะการรับเข้าจริง และส่งค่ากลับมาในโหมด lite ────────
assert(backend.indexOf('var NEW_ARRIVAL_WINDOW_DAYS = 7;') > -1, 'backend ต้องประกาศกรอบ 7 วันให้ตรงกับ frontend');
assert(/if \(signedQty > 0 && !payload\.skipPurchaseHistory\) \{\s*\r?\n\s*stampLastReceivedAt\(/.test(backend),
  'ต้องประทับเวลาเฉพาะรับเข้าจริง — การคืนของที่เบิกไปไม่ใช่ของใหม่');
assert(backend.indexOf("last_received_at: pickRowValue(row, map, NEW_ARRIVAL_ALIASES, '')") > -1,
  'backend ต้องส่ง last_received_at กลับมา');
assert(backend.indexOf('last_received_at: isLiteRead ?') === -1, 'ห้ามตัด last_received_at ทิ้งในโหมด lite');

// ประทับเวลาพลาดต้องไม่ทำให้การรับเข้าล้มทั้งรายการ (สต็อกสำคัญกว่าไฮไลต์)
const stampSrc = backend.match(/function stampLastReceivedAt\([\s\S]*?\n\}/)[0];
assert(stampSrc.indexOf('try {') > -1 && stampSrc.indexOf('catch (err)') > -1,
  'stampLastReceivedAt ต้องกัน error ไม่ให้ล้มการรับเข้า');

// alias ต้องเป็นรูปที่ normalize แล้วเท่านั้น (buildHeaderIndexMap ตัด _ และช่องว่างทิ้ง)
const aliasLine = backend.match(/var NEW_ARRIVAL_ALIASES = \[([^\]]*)\]/)[1];
aliasLine.split(',').forEach(function(raw) {
  const alias = raw.trim().replace(/^'|'$/g, '');
  if (!alias) return;
  assert(/^[a-z0-9]+$/.test(alias), 'alias ต้องเป็นตัวพิมพ์เล็ก/ตัวเลขล้วน ไม่งั้นจับคอลัมน์ไม่เจอ: ' + alias);
});

console.log('New arrival highlight checks passed');
