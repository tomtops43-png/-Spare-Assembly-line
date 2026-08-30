const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const htmlLf = html.replace(/\r\n/g, '\n');

function grab(from, to, label) {
  const a = htmlLf.indexOf(from);
  const b = htmlLf.indexOf(to);
  assert(a > -1 && b > a, 'ต้องหาบล็อก ' + label + ' เจอ');
  return htmlLf.slice(a, b);
}

// ── ปุ่มและช่องเลือกไฟล์ต้องมีจริง และผูก event ไว้ ─────────────────────────────
assert(html.includes('id="scImportCsvBtn"'), 'ต้องมีปุ่มนำเข้า CSV');
assert(html.includes('id="scImportCsvInput"') && html.includes('accept=".csv,text/csv"'));
assert(htmlLf.includes("importCsvBtn.addEventListener('click', function(){ importCsvInput.click(); });"));
// เลือกไฟล์เดิมซ้ำต้องยังยิง change — ไม่งั้นกดนำเข้าไฟล์เดิมรอบสองแล้วเงียบ
assert(/importCsvInput\.value = '';[\s\S]{0,120}if \(file\) scImportCountCsv\(file\);/.test(htmlLf),
  'ต้องเคลียร์ input.value ก่อนเรียก import');

// นำเข้าได้เฉพาะตอนมี Session เปิดอยู่ — CSV ไม่มีคอลัมน์ชีทต้นทาง ซึ่งจำเป็นตอนปรับยอด
const importBlock = grab('function scImportCountCsv(file)', 'function scSortedForPrint()', 'scImportCountCsv');
assert(importBlock.includes('if (!scSession)'), 'ไม่มี Session ต้องไม่ให้นำเข้า');
assert(importBlock.includes('scSaveDraft();'), 'นำเข้าแล้วต้องเซฟ draft ทันที');
assert(importBlock.includes('showConfirmDialog('), 'ต้องถามก่อนทับยอดที่กรอกไว้');

// ── รันตัว parser/matcher จริง ────────────────────────────────────────────────
// ดึงเฉพาะฟังก์ชันที่ไม่พึ่ง DOM ออกมารันใน sandbox
const parserSrc = grab('function scParseCsv(text)', 'function scImportCountCsv(file)', 'scParseCsv');
const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(parserSrc, sandbox);

// รูปแบบไฟล์ต้องตรงกับที่ scExportCountCsv เขียนออกมา
const exportBlock = grab('function scExportCountCsv()', 'function scExportPdf(', 'scExportCountCsv');
assert(exportBlock.includes("var head = ['#','ชื่ออะไหล่','รุ่น','แบรนด์','ตำแหน่ง','หน่วย','ยอดในระบบ','นับได้จริง','หมายเหตุ','ลิงก์รูป'];"),
  'หัวคอลัมน์ CSV เปลี่ยน — ตัวนำเข้าหาคอลัมน์ด้วยชื่อหัว ต้องอัปเดตให้ตรงกัน');
const applySrc = grab('function scApplyImportedCsv(rows)', 'function scSortedForPrint()', 'scApplyImportedCsv');
assert(/col\['ชื่ออะไหล่'\]/.test(applySrc) && /col\['นับได้จริง'\]/.test(applySrc),
  'ต้องหาคอลัมน์จากชื่อหัว ไม่ใช่ตำแหน่งตายตัว');

// BOM + CRLF + เครื่องหมายคำพูดซ้อน (รูปแบบเดียวกับที่ export ออกมาเป๊ะๆ)
const csv = '﻿' + [
  '"#","ชื่ออะไหล่","รุ่น","แบรนด์","ตำแหน่ง","หน่วย","ยอดในระบบ","นับได้จริง","หมายเหตุ","ลิงก์รูป"',
  '"1","CM2/CM3 ROUND BODY CYLINDER***","CM2B20-5Z","SMC","-","PCS","2","1","","" ',
  '"2","Body fixed tool PF250","MF-DRWG-CWM","YHM","A1","PCS","3","3","ครบ",""',
  '"3","ของที่มี ""เครื่องหมาย"" ในชื่อ","-","-","-","PCS","5","","",""'
].join('\r\n');

const rows = sandbox.scParseCsv(csv);
assert.strictEqual(rows.length, 4, 'ต้องได้ 4 แถว (หัว + ข้อมูล 3)');
assert.strictEqual(rows[0][1], 'ชื่ออะไหล่', 'ต้องตัด BOM ออกจากหัวคอลัมน์');
assert.strictEqual(rows[1][1], 'CM2/CM3 ROUND BODY CYLINDER***');
assert.strictEqual(rows[1][7], '1', 'อ่านยอดที่นับได้');
assert.strictEqual(rows[2][8], 'ครบ', 'อ่านหมายเหตุ');
assert.strictEqual(rows[3][1], 'ของที่มี "เครื่องหมาย" ในชื่อ', 'ต้อง unescape "" เป็น "');
assert.strictEqual(rows[3][7], '', 'ยังไม่ได้นับต้องเป็นค่าว่าง ไม่ใช่ 0');

// จับคู่ด้วยชื่อ+รุ่น และ '-' ต้องถือว่าไม่มีรุ่น (ตรงกับที่ Backend ใช้)
assert.strictEqual(sandbox.scMatchKey('ABC', '-'), sandbox.scMatchKey('abc', ''),
  "'-' ต้องถือว่าไม่มีรุ่น และไม่สนตัวพิมพ์");
assert.notStrictEqual(sandbox.scMatchKey('ABC', 'M1'), sandbox.scMatchKey('ABC', 'M2'),
  'คนละรุ่นต้องไม่จับคู่กัน');
assert.strictEqual(sandbox.scMatchKey(' ABC ', ' M1 '), sandbox.scMatchKey('ABC', 'M1'),
  'ต้อง trim ช่องว่างหัวท้าย');

console.log('stock-count-csv-import: OK');
