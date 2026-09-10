// ตัวช่วยดึงโค้ดของ Export Center ออกมาจากไฟล์เดียว (index.html / Backend.gs) เพื่อรันทดสอบ
// โปรเจกต์นี้ไม่มี build step และ index.html เป็น SPA ไฟล์เดียว จึงทดสอบด้วยวิธีตัดฟังก์ชัน
// ออกมาประกอบใหม่แบบเดียวกับเทสต์เดิมในโฟลเดอร์นี้ (ดู misc-expense-ledger.test.js)
//
// ไฟล์นี้ชื่อไม่ลงท้าย .test.js จึงไม่ถูก run-all.js หยิบไปรันเป็นเทสต์เอง
const fs = require('fs');
const assert = require('assert');

const html = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8').replace(/\r\n/g, '\n');

// ดึงฟังก์ชันระดับบนสุดของสคริปต์ (ย่อหน้า 4 ช่อง ปิดด้วย '\n    }')
function grabFn(name, source, indent) {
  const src = source || html;
  const pad = indent === undefined ? 4 : indent;
  const sp = ' '.repeat(pad);
  const re = new RegExp('^' + sp + 'function ' + name + '\\([^)]*\\) \\{[\\s\\S]*?\\n' + sp + '\\}', 'm');
  const m = src.match(re);
  assert(m, 'ต้องดึงฟังก์ชัน ' + name + ' ออกมาได้');
  return m[0];
}

// ดึง `var NAME = [...]` หรือ `var NAME = {...}` โดยนับวงเล็บแบบข้ามเนื้อในสตริงและคอมเมนต์
// (ถ้านับดิบๆ เครื่องหมายในสตริงจะทำให้ตัดผิดตำแหน่ง)
function grabVar(name, source) {
  const src = source || html;
  const decl = 'var ' + name + ' = ';
  const at = src.indexOf(decl);
  assert(at > -1, 'ต้องหาตัวแปร ' + name + ' เจอ');
  let i = at + decl.length;
  const open = src[i];
  assert(open === '[' || open === '{', name + ' ต้องเป็น array หรือ object literal');
  const close = open === '[' ? ']' : '}';
  let depth = 0;
  let inStr = null;
  let inLineComment = false;
  for (; i < src.length; i += 1) {
    const ch = src[i];
    const prev = src[i - 1];
    if (inLineComment) {
      if (ch === '\n') inLineComment = false;
      continue;
    }
    if (inStr) {
      if (ch === inStr && prev !== '\\') inStr = null;
      continue;
    }
    if (ch === '/' && src[i + 1] === '/') { inLineComment = true; continue; }
    if (ch === '"' || ch === "'") { inStr = ch; continue; }
    if (ch === '[' || ch === '{') depth += 1;
    else if (ch === ']' || ch === '}') {
      depth -= 1;
      if (depth === 0) {
        assert.strictEqual(ch, close, name + ' วงเล็บปิดไม่ตรงชนิด');
        return src.slice(at, i + 1) + ';';
      }
    }
  }
  throw new Error('หาวงเล็บปิดของ ' + name + ' ไม่เจอ');
}

// ประกอบชิ้นส่วนเป็นโมดูลเดียวแล้วคืนค่าที่ขอ
function buildModule(pieces, returnExpr) {
  const src = pieces.join('\n') + '\nreturn (' + returnExpr + ');';
  return new Function(src)();
}

// ชุดฟังก์ชันพื้นฐานที่รายงาน/ตัวกรองแทบทุกตัวต้องใช้
// รวมชั้นภาษา (XP_LANG_EN + xpT) ด้วย เพราะทุกตารางที่ผ่าน xpTable ต้องแปลหัวคอลัมน์
function baseHelpers(extra) {
  const names = ['xpN', 'xpS', 'xpFmtInt', 'xpFmtMoney', 'xpDateStr', 'xpMonthKey', 'xpDayKey', 'xpWeekKey', 'xpBucketKey', 'xpInRange', 'xpLineMatches', 'xpIsSensitiveHeader', 'xpT', 'xpIsEn', 'xpLocaleStamp', 'xpDictKeyOf', 'xpTable', 'xpTableFromDump', 'xpPartKey', 'xpPriceMap', 'xpLookupPrice', 'xpIsOutput', 'xpIsInput', 'xpEscHtml'];
  const pieces = [grabVar('XP_SENSITIVE_PATTERNS'), grabVar('XP_LANG_EN'), 'var xpLangReverse = null;'];
  names.concat(extra || []).forEach(function(n) { pieces.push(grabFn(n)); });
  return pieces;
}

module.exports = { html, backend, grabFn, grabVar, buildModule, baseHelpers, assert };
