const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// แท็บ "รับเข้า" มีไอคอน SVG + <span> ข้างใน ไม่ใช่ text ตรงๆ ใน <button> แล้ว
assert(/<button id="tabTxn"[\s\S]{0,600}?>รับเข้า<\/span>[\s\S]{0,80}?<\/button>/.test(html),
  'ต้องมีแท็บ id="tabTxn" ที่มีป้ายว่า "รับเข้า"');
assert(html.includes('<input id="txnType" type="hidden" value="Input"'));
assert(!html.includes('id="txnTypeOutputOption"'));
assert(html.includes('📥 ประเภทรายการ: รับเข้า'));
assert(html.includes('ผู้รับ <span class="text-emerald-700">(Auto จากผู้ Login)</span>'));
assert(html.includes('id="txnBy"') && html.includes('placeholder="ผู้รับ (Auto)" required readonly'));
assert(html.includes('Line ที่รับเข้า <span class="text-emerald-700">(Auto จากรายการอะไหล่)</span>'));
assert(html.includes('id="txnProcess"') && html.includes('placeholder="Line (Auto)" required readonly'));
assert(html.includes("type: 'Input'"));
assert(html.includes('syncReceiveFormAutoFields(selected);'));
assert(html.includes("qaIssueBtn.addEventListener('click', function()") && html.includes('openIssueCart();'));
assert(backend.includes("payload.by = getSessionUser({ authToken: payload.authToken }).user.username"));
assert(backend.includes("pickRowValue(targetRow, map, ['line', 'linearea', 'area', 'process', 'mainline']"));
assert(backend.includes("payload.type = 'Input'"));
console.log('Receive-only transaction page checks passed');
