const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');
const backendLf = backend.replace(/\r\n/g, '\n');

function grab(re, label) {
  const m = backendLf.match(re);
  assert(m, 'ต้องดึง ' + label + ' ออกมาได้');
  return m[0];
}

// ── นโยบายเก็บไฟล์: รายวัน 30 วัน + ของวันที่ 1 ของเดือนไว้ 1 ปี ───────────────
// รันฟังก์ชันจริงโดยจำลอง DriveApp ทั้งหมด — ตรรกะการเก็บ/ทิ้งพลาดแล้วข้อมูลหายจริง
const pruneSrc = grab(/^function pruneOldBackups\(folder\) \{[\s\S]*?\n\}/m, 'pruneOldBackups');
const pruneOldBackups = new Function('Logger', pruneSrc + '\nreturn pruneOldBackups;')({ log: function() {} });

function fakeFolder(names) {
  const files = names.map(function(n) {
    return { name: n, trashed: false, getName: function() { return n; }, setTrashed: function(v) { this.trashed = v; } };
  });
  let i = 0;
  return {
    files: files,
    getFiles: function() { return { hasNext: function() { return i < files.length; }, next: function() { return files[i++]; } }; }
  };
}
function dayString(daysAgo) {
  const d = new Date(Date.now() - daysAgo * 86400000);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
// ของเมื่อวานต้องอยู่ · ของ 100 วันก่อนต้องถูกทิ้ง (ยกเว้นเป็นวันที่ 1 ของเดือน)
const recent = 'SpareParts-Backup-' + dayString(1);
// เลี่ยงวันที่ 1 ของเดือน เพราะนั่นคือไฟล์รายเดือนที่ตั้งใจเก็บยาว 1 ปี
function oldDailyString() {
  for (var d = 45; d < 60; d += 1) {
    var s = dayString(d);
    if (!/-01$/.test(s)) return s;
  }
  return dayString(45);
}
const old100 = 'SpareParts-Backup-' + oldDailyString();
const monthly = 'SpareParts-Backup-2026-03-01';
const veryOldMonthly = 'SpareParts-Backup-2024-01-01';
const notABackup = 'อ่านฉันก่อนนะ.txt';
const folder = fakeFolder([recent, old100, monthly, veryOldMonthly, notABackup]);
const removed = pruneOldBackups(folder);
function trashed(name) { return folder.files.filter(function(f) { return f.name === name; })[0].trashed; }
assert.strictEqual(trashed(recent), false, 'ไฟล์สำรองล่าสุดห้ามถูกทิ้ง');
assert.strictEqual(trashed(old100), true, 'ไฟล์รายวันที่เกิน 30 วันต้องถูกทิ้ง');
assert.strictEqual(trashed(monthly), false, 'ของวันที่ 1 ของเดือน (ยังไม่เกิน 1 ปี) ต้องเก็บไว้');
assert.strictEqual(trashed(veryOldMonthly), true, 'ของวันที่ 1 ที่เกิน 1 ปีแล้วก็ต้องทิ้ง');
assert.strictEqual(trashed(notABackup), false, 'ไฟล์ที่ไม่ใช่ไฟล์สำรองห้ามไปยุ่ง');
assert(removed.indexOf(old100) > -1 && removed.length === 2);
// ทิ้งลงถังขยะ ไม่ใช่ลบถาวร — ตัดสินใจผิดยังกู้กลับได้
assert(/file\.setTrashed\(true\);/.test(backend) && !/removeFile|\.delete\(\)/.test(pruneSrc));

// ── รันซ้ำวันเดียวกันต้องไม่ได้สำเนาซ้ำ ──────────────────────────────────────
assert(/if \(folder\.getFilesByName\(name\)\.hasNext\(\)\) \{[\s\S]{0,220}skipped: true/.test(backend),
  'กดปุ่มสำรองหลัง trigger ทำไปแล้ว ต้องไม่ได้ไฟล์ซ้ำ');

// ── trigger ───────────────────────────────────────────────────────────────
assert(/ScriptApp\.newTrigger\('runDailyBackup'\)\.timeBased\(\)\.everyDays\(1\)\.atHour\(1\)\.create\(\);/.test(backend),
  'ต้องตั้ง trigger สำรองข้อมูลตี 1');
assert(/if \(fn === 'runDailyAutoJobs' \|\| fn === 'runDailyBackup'\) ScriptApp\.deleteTrigger\(t\);/.test(backend),
  'setupAutomation ต้องลบ trigger เดิมก่อน ไม่งั้นรันซ้ำหลายรอบต่อวัน');
assert(/ScriptApp\.newTrigger\('runDailyAutoJobs'\)\.timeBased\(\)\.everyDays\(1\)\.atHour\(7\)\.create\(\);/.test(backend),
  'งานเช้าเดิมต้องยังอยู่');

// ── backup ที่พังเงียบๆ อันตรายกว่าไม่มี backup — ต้องเตือนทุกเช้า ────────────
assert(/if \(backupState\.age_days < 0\) lines\.push\('⚠️ ยังไม่มีไฟล์สำรองข้อมูลเลย'\);/.test(backend));
assert(/else if \(backupState\.age_days >= 2\) lines\.push\('⚠️ ไม่ได้สำรองข้อมูลมา '/.test(backend));
// เช็ค backup พังต้องไม่ทำให้สรุปประจำวันทั้งฉบับหายไป
assert(/try \{[\s\S]{0,400}readBackupState\(\)[\s\S]{0,400}\} catch \(errBackup\)/.test(backend));

// ── สิทธิ์ + routing ───────────────────────────────────────────────────────
assert(/function getBackupStatus\(payload\) \{[\s\S]{0,120}requireAdminUser/.test(backend));
assert(/function runBackupNow\(payload\) \{[\s\S]{0,120}requireAdminUser/.test(backend));
['getBackupStatus', 'runBackupNow'].forEach(function(a) {
  assert(backend.includes("if (action === '" + a + "') return respond(" + a + "(e.parameter), e);"), a + ' GET');
  assert(backend.includes("if (action === '" + a + "') return respond(" + a + "(body), e);"), a + ' POST');
});

// ── หน้า Admin ────────────────────────────────────────────────────────────
assert(htmlLf.includes('id="backupStatusBadge"') && htmlLf.includes('id="backupRunNowBtn"') && htmlLf.includes('id="backupFolderLink"'));
assert(htmlLf.includes('💾 สำรองข้อมูลอัตโนมัติ'));
assert(/loadLineNotifyStatus\(\);\n\s+loadBackupStatus\(\);/.test(htmlLf), 'เข้าหน้า Admin ต้องเช็คสถานะสำรองให้เลย');
// ต้องบอก "ล่าสุดเมื่อไหร่/ค้างมากี่วัน" ไม่ใช่แค่ว่าเปิดใช้งานอยู่
assert(htmlLf.includes("'⚠️ ค้างมา ' + age + ' วัน'") && htmlLf.includes('⛔ ยังไม่มีไฟล์สำรอง'));
assert(htmlLf.includes('รัน setupAutomation ใน Apps Script อีกครั้ง'), 'ค้างแล้วต้องบอกวิธีแก้ ไม่ใช่เตือนลอยๆ');

console.log('Daily backup checks passed');
