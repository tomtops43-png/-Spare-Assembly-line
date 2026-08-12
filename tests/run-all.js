// รันไฟล์เทสต์ทุกตัวใน tests/ แยก process ละไฟล์ แล้วสรุปผลรวมทีเดียว
// ใช้ node ล้วน ไม่มี dependency (โปรเจกต์นี้ไม่มี build step / node_modules)
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const testsDir = __dirname;
const repoRoot = path.join(testsDir, '..');
const files = fs.readdirSync(testsDir)
  .filter(function(f) { return f.endsWith('.test.js'); })
  .sort();

let failed = 0;
files.forEach(function(file) {
  const res = spawnSync(process.execPath, [path.join(testsDir, file)], {
    cwd: repoRoot,
    encoding: 'utf8'
  });
  if (res.status === 0) {
    console.log('  PASS  ' + file);
    return;
  }
  failed += 1;
  console.log('  FAIL  ' + file);
  const output = ((res.stdout || '') + (res.stderr || '')).trim();
  output.split('\n').slice(0, 25).forEach(function(line) { console.log('        ' + line); });
});

console.log('\n' + (files.length - failed) + '/' + files.length + ' test files passed');
if (failed > 0) {
  console.error(failed + ' test file(s) failed');
  process.exit(1);
}
