const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1];

// ── 1) ห้ามประกาศ ROLE_PRESETS / applyRolePreset ซ้ำที่ top-level ──────────────
// เคยมี 2 ชุด (คอมมิต 40715da) ชุดหลังทับชุดแรกเงียบๆ ทำให้สิทธิ์ไลน์หายไป
const topLevelDecls = {};
script.split('\n').forEach(function(line, i) {
  const m = line.match(/^ {4}(?:var|let|const) ([A-Za-z0-9_$]+)\s*=/) ||
            line.match(/^ {4}function ([A-Za-z0-9_$]+)\s*\(/);
  if (!m) return;
  (topLevelDecls[m[1]] = topLevelDecls[m[1]] || []).push(i + 1);
});
const dups = Object.keys(topLevelDecls).filter(function(k) { return topLevelDecls[k].length > 1; });
assert.deepStrictEqual(dups, [], 'มีตัวแปร/ฟังก์ชัน top-level ประกาศซ้ำ (ตัวหลังจะทับตัวแรกเงียบๆ): ' + dups.join(', '));

// ── 2) ทุก preset ต้องมี lines ─────────────────────────────────────────────────
const presetsSrc = script.match(/var ROLE_PRESETS = \{[\s\S]*?\n {4}\};/)[0];
const ROLE_PRESETS = vm.runInNewContext('(' + presetsSrc.replace(/^var ROLE_PRESETS = /, '').replace(/;$/, '') + ')');
const expectedLines = {
  viewer: 'view',
  leader_pd: 'view',
  user: 'view',
  leader_basic: 'managed',
  leader: 'managed',
  admin: 'all'
};
Object.keys(expectedLines).forEach(function(key) {
  assert(ROLE_PRESETS[key], 'ต้องมี preset: ' + key);
  assert.strictEqual(ROLE_PRESETS[key].lines, expectedLines[key],
    'preset ' + key + ' ต้องมี lines = ' + expectedLines[key]);
});

// role ที่ dropdown uRole ส่งเข้ามาต้องมี preset รองรับครบ (setUserEditorValues/ปุ่ม reset ส่งค่านี้)
['user', 'leader', 'admin'].forEach(function(roleValue) {
  assert(ROLE_PRESETS[roleValue], 'ค่า role "' + roleValue + '" ต้อง map กับ preset ได้');
});

// ── 3) applyRolePreset ต้อง sync สิทธิ์ไลน์ด้วย ───────────────────────────────
const applySrc = script.match(/function applyRolePreset\(presetKey\)[\s\S]*?\n {4}\}/)[0];
assert(/applyLineAccessPreset\(preset\.lines\)/.test(applySrc),
  'applyRolePreset ต้องเรียก applyLineAccessPreset(preset.lines) ไม่งั้นค่าไลน์ของ user คนก่อนจะค้าง');
assert(applySrc.indexOf('applyLineAccessPreset') < applySrc.indexOf('syncPermissionsJsonFromUI'),
  'ต้อง set สิทธิ์ไลน์ก่อนแล้วค่อย sync permissionsJson');

// ── 4) applyLineAccessPreset ต้องคุม checkbox ครบทั้ง 3 โหมด ──────────────────
const lineFnSrc = script.match(/function applyLineAccessPreset\(linesMode\)[\s\S]*?\n {4}\}/)[0];
assert(/uAllLinesAccess\.checked = linesMode === 'all'/.test(lineFnSrc), "โหมด 'all' ต้องติ๊ก uAllLinesAccess");
assert(/data-line-view/.test(lineFnSrc) && /data-line-manage/.test(lineFnSrc), 'ต้องแตะ checkbox ทั้ง View และ Manage');
assert(/linesMode === 'view' \|\| linesMode === 'managed'/.test(lineFnSrc), "โหมด view/managed ต้องติ๊ก View");
assert(/manage\.checked = linesMode === 'managed'/.test(lineFnSrc), "Manage ติ๊กเฉพาะโหมด managed");

console.log('Role preset line-access checks passed');
