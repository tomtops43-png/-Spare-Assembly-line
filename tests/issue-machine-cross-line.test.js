// อะไหล่ของไลน์หนึ่งเอาไปใส่เครื่องของอีกไลน์ได้ (เช่น Bearing ของ H9 → เครื่อง Coil Winding)
// เดิมรายชื่อเครื่องล็อกไว้ที่ไลน์ของอะไหล่ เลยเลือกเครื่อง Coil ไม่ได้
const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const script = html.match(/<script>([\s\S]*)<\/script>/)[1].replace(/\r\n/g, '\n');

function extract(sig) {
  const start = script.indexOf('    function ' + sig);
  assert(start > -1, 'cannot find ' + sig);
  const end = script.indexOf('\n    }', start);
  return script.slice(start, end + 6);
}

const LINES = ['H9', 'Arc Chute', 'Coil Winding', 'Lug&Screw'];
const api = new Function('realLineKeys', 'escHtml', extract('machineLineOptionsHtml(selectedLine)') + '\nreturn machineLineOptionsHtml;')(
  function() { return LINES.slice(); },
  function(v) { return String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }
);

// ทุกไลน์ต้องเลือกได้ และไลน์ของอะไหล่ถูกเลือกไว้เป็นค่าเริ่มต้น
const h9 = api('H9');
LINES.forEach(function(l) { assert(h9.indexOf('>' + l.replace('&', '&amp;') + '</option>') > -1, 'ต้องมีไลน์ ' + l); });
assert(/value="H9" selected/.test(h9), 'ค่าเริ่มต้นต้องเป็นไลน์ของอะไหล่');
assert(!/value="Coil Winding" selected/.test(h9));
// ไลน์แปลกที่ไม่อยู่ในรายการต้องไม่หาย
assert(/value="Assembly" selected/.test(api('Assembly')), 'ไลน์ที่ไม่อยู่ในรายการต้องยังถูกเลือกไว้');

// ── เบิกด่วน ─────────────────────────────────────────────────────────────────
assert(html.includes('id="quickIssueMachineLine"'), 'เบิกด่วนต้องมีช่องเลือกไลน์ของเครื่อง');
assert(script.includes("quickIssueMachineLineSel.innerHTML = machineLineOptionsHtml(quickIssueProcess.value);"));
const quickChange = script.slice(script.indexOf("quickIssueMachineLineEl.addEventListener('change'"), script.indexOf("var quickIssueMachineSelect = document.getElementById('quickIssueMachine');"));
assert(quickChange.includes('quickIssueProcess.value = line;'), 'เปลี่ยนไลน์ของเครื่องแล้ว ไลน์ที่บันทึกต้องเปลี่ยนตาม');
assert(quickChange.includes('populateQuickIssueMachineOptions(line, item)'), 'ต้องโหลดเครื่องของไลน์ที่เลือกใหม่');
assert(quickChange.includes('resetQuickIssueMachineField();'), 'เปลี่ยนไลน์แล้วต้องล้างเครื่องที่เลือกไว้เดิม');

// ── Issue Cart ───────────────────────────────────────────────────────────────
assert(script.includes("line: 'issueCartMachineLine_' + idx"), 'แต่ละกลุ่มในตะกร้าต้องมีช่องเลือกไลน์ของเครื่อง');
assert(script.includes('data-cart-machine-line-idx'), 'ช่องไลน์ต้องผูกอีเวนต์ผ่าน delegate');
const reload = extract('reloadIssueCartMachineGroup(idx)');
assert(reload.includes('resetMachineSelect(ids.sel, ids.other);'));
assert(reload.includes('getIssueCartMachineLineAt(i) || partLine'), 'ต้องโหลดเครื่องของไลน์ที่เลือก');
// สต็อกยังตัดจากชีทของอะไหล่เหมือนเดิม
assert(extract('createOutputTransactionPayload(').includes('sheetName: item.__sourceSheet || item.subLine || getCurrentSheetName()'));

console.log('Issue machine cross-line checks passed');
