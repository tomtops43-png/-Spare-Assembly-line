const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ---- รูปใน Annex C ของใบ PR ที่ปริ้นจาก Inbox ต้องขึ้นจริง ----
// เคสจริง: PR-202608-STAR-189 ปริ้นออกมาแล้วช่อง Picture ขึ้น "No image available" ทั้ง 19 แถว

// 1) server ส่ง image_url มากับทุกบรรทัด PR อยู่แล้ว
assert(backend.includes("'remark', 'image_url'"), 'PR lines schema keeps image_url');
assert(/image_url: lIdx\.image_url !== undefined \? prStr\(r\[lIdx\.image_url\]\) : ''/.test(backend), 'getPRForApproval returns image_url per line');

// 2) ตัวปริ้นต้องส่ง image_url ต่อเข้า report — ห้ามพึ่ง master (partsData) อย่างเดียว
//    เพราะตอนปริ้นจาก Inbox ไลน์ของ PR มักไม่ใช่ไลน์ที่โหลดอยู่ → หา master ไม่เจอ → รูปหาย
const printMap = html.slice(html.indexOf('function openApprovedPrPrintWindow'), html.indexOf('function waitForPrintImages'));
assert(printMap.includes('mainImage: ln.image_url'), 'print items carry mainImage from PR line snapshot');
assert(printMap.includes('image: ln.image_url'), 'print items carry image from PR line snapshot');

// 3) report อ่าน mainImage/image จาก candidate ก่อน แล้วค่อย fallback ไป master
assert(/var selectedImg = firstDefined\(\[c\.mainImage, c\.image_main_url/.test(html), 'report prefers the selected row image');
assert(html.includes('firstDefined([selectedImg, masterImg]'), 'report falls back to master image');

// 4) รูปในเอกสารปริ้นห้าม lazy — เราปริ้นผ่าน iframe 0x0 ที่ซ่อนอยู่ รูป lazy จะไม่โหลดเลย
const annexImg = html.slice(html.indexOf('var imgHtml = c.image ?'), html.indexOf('counter.n += 1;'));
assert(!annexImg.includes('loading="lazy"'), 'annex picture must not be lazy-loaded');
assert(!annexImg.includes('decoding="async"'), 'annex picture must not decode async');

// 5) onerror ต้องไม่ทำ attribute พัง (เดิมใช้ \" ซึ่งปิด attribute กลางคัน)
assert(!/onerror="[^"]*\\"/.test(annexImg), 'annex onerror must not embed a raw double quote');
assert(annexImg.includes('&quot;No image available&quot;'), 'annex onerror uses entity-escaped fallback text');

// 6) ต้องรอรูปโหลดครบก่อนเปิด dialog ปริ้น ไม่ใช่หน่วงตายตัว 500ms
assert(html.includes('function waitForPrintImages(win, timeoutMs)'), 'image wait helper exists');
assert(html.includes('waitForPrintImages(frame.contentWindow, 12000).then('), 'print waits for images before print()');
assert(/img\.addEventListener\('load', one, \{ once: true \}\);/.test(html), 'waits on image load');
assert(/img\.addEventListener\('error', one, \{ once: true \}\);/.test(html), 'error images must not block printing');
assert(html.includes('setTimeout(finish, timeoutMs || 12000);'), 'wait has a timeout guard so print never hangs');

console.log('pr-print-annex-images: all assertions passed');
