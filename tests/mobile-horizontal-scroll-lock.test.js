const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// อาการจริง: บนมือถือหน้าเลื่อนซ้าย-ขวาได้ ทำให้หัวข้อ/ช่องค้นหาโดนดันออกนอกจอ
// ต้นเหตุ 2 ชั้น — (1) แถวค้นหา+ปุ่มไอคอนกว้างเกินจอจริงๆ (2) ตัวกันของเดิมอยู่ที่ <body> อย่างเดียว
// ซึ่ง iOS Safari ไม่สนใจ ต้องใส่ที่ <html> ด้วย

// ---- 1. ตัวกันระดับเอกสาร ----
assert(/html \{ overflow-x: clip; \}/.test(html), '<html> ล็อกแกนนอนไว้ (iOS ไม่สนใจถ้ามีแค่ที่ body)');
assert(/body \{ overflow-x: clip; max-width: 100%; \}/.test(html), '<body> ล็อกแกนนอน + ไม่ให้กว้างเกินจอ');
// clip ไม่สร้าง scroll container ใหม่ → sticky header ยังเกาะ viewport ได้ปกติ
// hidden ทำได้เหมือนกันแต่ต้องอยู่ที่ <html> เท่านั้น (ถ้าไปอยู่ที่ body จะทำ sticky พัง)
const fallback = html.slice(html.indexOf('@supports not (overflow-x: clip)'));
assert(fallback.indexOf('html { overflow-x: hidden; }') > -1, 'fallback ของเบราว์เซอร์เก่าอยู่ที่ <html>');
assert(fallback.slice(0, fallback.indexOf('}\n', fallback.indexOf('html { overflow-x: hidden'))).indexOf('body { overflow-x: hidden') === -1,
  'fallback ต้องไม่ใส่ overflow-x:hidden ที่ body (จะทำให้ sticky header ไม่ติดขอบบน)');
// ต้องไม่กลับไปพึ่ง class ของ Tailwind บน body — specificity สูงกว่ากฎด้านบนจนทับกันเอง
assert(!/<body class="[^"]*overflow-x-hidden/.test(html), 'body ไม่ใช้ class overflow-x-hidden แล้ว');

// ---- 2. แถวค้นหาบนมือถือต้องย่อได้จริง ----
// flex item มี min-width:auto โดยปริยาย → <input> ไม่ยอมแคบกว่าความกว้างเริ่มต้น ~20 ตัวอักษร
// แถวเลยกว้างเกินจอและดันทั้งหน้าไปทางขวา ต้องมี min-w-0 ครบทุกชั้น
assert(/<div class="flex min-w-0 flex-1 items-center gap-1\.5 md:hidden">/.test(html),
  'แถว search+ปุ่มไอคอน (มือถือ) ย่อได้');
assert(/<div class="relative flex min-w-0 flex-1 items-center rounded-xl border border-slate-200 bg-white shadow-sm overflow-hidden">/.test(html),
  'กรอบช่องค้นหามือถือย่อได้');
['searchInputMobile', 'searchInput'].forEach(function(id) {
  const m = new RegExp('<input id="' + id + '"[^>]*class="([^"]*)"').exec(html);
  assert(m, 'เจอ input#' + id);
  assert(/\bmin-w-0\b/.test(m[1]), 'input#' + id + ' มี min-w-0');
  assert(/\bw-full\b/.test(m[1]), 'input#' + id + ' ไม่ยึดความกว้างเริ่มต้นของ input');
});

// ---- 3. เมนูย่อยบนมือถือต้องไม่ยื่นพ้นขอบขวา ----
// panel เดิมอ้างอิง .nav-group (min-width:200px) กลุ่มที่อยู่ค่อนไปทางขวาจึงล้นออกนอกจอ
const navMobile = html.slice(html.indexOf('#topNavTabs { position: relative; }'));
assert(navMobile.indexOf('.nav-group { position: static; }') > -1, 'มือถือ: เมนูย่อยอ้างอิงแถบเมนูแทนกลุ่ม');
assert(/\.nav-dropdown-panel\.nav-dropdown-right \{ left: 0; right: 0; width: auto; min-width: 0; \}/.test(navMobile),
  'มือถือ: เมนูย่อยกางเต็มความกว้างแถบเมนู ไม่ยึด min-width 200px');

console.log('PASS mobile-horizontal-scroll-lock');
