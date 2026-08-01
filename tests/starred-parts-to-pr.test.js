const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');

// ⭐ อะไหล่ติดดาว → เปิด PR ได้โดยตรง
// ก่อนหน้านี้รายการติดดาวส่งต่อได้แค่ "ขอซื้อ" ทีละตัว ต้องยกทั้งรายการไปทำ PR ได้ด้วย

// ---- UI: preset card ที่ 5 ในหน้า PR Report ----
assert(html.includes('data-pr-preset="from_starred"'), 'PR Report has a from_starred preset card');
assert(html.includes('PR อะไหล่ติดดาว'), 'preset card labelled in Thai');
assert(html.includes('<option value="from_starred">'), '#prMode has from_starred option');
assert(html.includes('id="prFromStarredHint"'), 'from_starred mode has its own hint block');
// การ์ด preset มี 5 ใบแล้ว grid ต้องขยายจาก 4 คอลัมน์
assert(/grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-5/.test(html), 'preset grid widened to 5 columns');
assert((html.match(/class="pr-preset-btn/g) || []).length === 5, 'exactly 5 preset buttons');

// ---- UI: ปุ่มในป๊อปอัพอะไหล่ติดดาว ----
assert(html.includes('id="starredToPrBtn"'), 'starred modal has an open-PR button');
assert(html.includes('เปิด PR ทั้งรายการ'), 'open-PR button labelled');

// ---- candidate builder ----
assert(html.includes('function loadStarredPrCandidates()'), 'starred PR candidate builder exists');
assert(/loadStarredPrCandidates[\s\S]{0,600}getStarredEntriesSorted\(\)/.test(html), 'builder reads the starred store');
// ต้องพยายามใช้สต็อก/ราคาสดจาก master ก่อน ไม่ใช้ snapshot ตอนติดดาวเป็นหลัก
assert(/loadStarredPrCandidates[\s\S]{0,900}loadPrSourceItems\('assembly_monthly', 'all'\)/.test(html), 'builder pulls live master items');
assert(/live \? Number\(live\.stock \|\| 0\) : safeNum\(e\.stockAtStar\)/.test(html), 'live stock preferred, snapshot only as fallback');
// จับคู่ด้วย identity key ก่อน แล้ว fallback ชื่อ+รุ่น (identity key มี location ถ้าย้ายที่จะจับไม่เจอ)
assert(/byIdentity\[row\.key\][\s\S]{0,200}byNameModel\[/.test(html), 'identity-key match with name+model fallback');
// โน๊ตที่พิมพ์ตอนติดดาวคือเหตุผลจริง ต้องยกไปเป็น Purpose ของ PR
assert(/purpose: e\.note \? \('⭐ ' \+ e\.note\)/.test(html), 'starred note carried into PR purpose');
assert(/var reasons = \['STAR'\]/.test(html), 'candidates tagged with STAR reason code');

// ---- เดินสายเข้า runPrCandidateSelection ----
assert(html.includes("if (mode === 'from_request' || mode === 'from_starred')"), 'from_starred routed in runPrCandidateSelection');
assert(/mode === 'from_starred' \? loadStarredPrCandidates : loadOrderRequestPrCandidates/.test(html), 'picks the starred loader for from_starred');
// โหมดติดดาวต้องโหลด purchase history ไว้เทียบ "มีคำสั่งซื้อค้างอยู่" เหมือนโหมดปกติ
assert(/mode === 'from_starred'[\s\S]{0,160}loadPurchaseHistory\(true\)/.test(html), 'from_starred loads purchase history for PENDING check');
assert(html.includes('ยังไม่มีอะไหล่ติดดาว — กดปุ่ม'), 'empty starred list shows a helpful message instead of a blank list');

// ---- render / labels ----
assert(html.includes("currentPrMode === 'from_request' || currentPrMode === 'from_starred'"), 'from_starred renders flat (not grouped by line)');
assert(/STAR: \['⭐ ติดดาวไว้'/.test(html), 'STAR reason has a badge label');
assert(html.includes('<option value="STAR">'), 'STAR is filterable in the reason dropdown');
assert(/from_starred: \{ border: '#f59e0b'/.test(html), 'from_starred preset has its own highlight colour');
assert(/mode === 'from_starred' \? 'STAR'/.test(html), 'PR number uses a STAR mode code');

// ---- navigation ----
assert(html.includes('function openPrFromStarred()'), 'starred → PR entry point exists');
// ดูเฉพาะตัวฟังก์ชัน ไม่ใช่ระยะห่างในไฟล์ทั้งก้อน — จะได้ไม่พังเวลามีคนแก้บรรทัดรอบๆ
const openPrBody = (function() {
  const start = html.indexOf('function openPrFromStarred()');
  const end = html.indexOf('\n    }', start);
  assert(start > -1 && end > start, 'can slice openPrFromStarred body');
  return html.slice(start, end);
})();
assert(openPrBody.includes('switchToPrReportPage()'), 'navigates to the PR Report page');
assert(openPrBody.includes("hasPermission('view_logs')"), 'permission-gated like the PR tab itself');
assert(openPrBody.includes("__applyPrPreset('from_starred')"), 'switches the page into from_starred mode');
// ต้องกดปุ่ม Generate จริง ไม่ใช่เรียก runPrCandidateSelection() ตรงๆ ไม่งั้นขั้นตอนที่ 3/4 ยังซ่อนอยู่
assert(/getElementById\('prRunBtn'\)[\s\S]{0,80}runBtn\.click\(\)/.test(openPrBody), 'openPrFromStarred clicks the real Generate button');
assert(html.includes("window.__applyPrPreset = applyPreset"), 'preset switcher exposed for the starred modal');
assert(/starredToPrBtnEl\.classList\.toggle\('hidden', !canViewLogs\)/.test(html), 'open-PR button hidden when user cannot view the PR page');
assert(html.includes("starredToPrBtn.addEventListener('click', openPrFromStarred)"), 'open-PR button wired up');

console.log('starred-parts-to-pr: all assertions passed');
