const fs = require('fs');
const assert = require('assert');
const htmlLf = fs.readFileSync('index.html', 'utf8').replace(/\r\n/g, '\n');

// ── แถบแท็บแบบ browser ด้านบนสุด ─────────────────────────────────────────────
const shellAt = htmlLf.indexOf('<div id="appShell"');
const barAt = htmlLf.indexOf('<div id="browserTabBar"');
const headerAt = htmlLf.indexOf('<header', shellAt);
assert(shellAt > -1 && barAt > shellAt, 'แถบแท็บต้องอยู่ใน #appShell (โผล่เฉพาะตอน login แล้ว)');
assert(barAt < headerAt, 'แถบแท็บต้องอยู่บนสุด ก่อน header');
assert(htmlLf.includes('id="browserTabList"'), 'ต้องมีที่ใส่แท็บ');
assert(htmlLf.includes('id="browserTabNewBtn"'), 'ต้องมีปุ่ม + เปิดแท็บใหม่');

// ── ทุกหน้าที่แท็บจำได้ ต้องชี้ไปปุ่มเมนูที่มีจริง และเป็น key ที่ setTabStyles ใช้จริง ──
const pagesSrc = htmlLf.match(/^ {4}var BROWSER_TAB_PAGES = \{[\s\S]*?\n {4}\};/m);
assert(pagesSrc, 'ต้องดึง BROWSER_TAB_PAGES ออกมาได้');
const PAGES = new Function(pagesSrc[0] + '\nreturn BROWSER_TAB_PAGES;')();
const setTabStylesSrc = htmlLf.match(/function setTabStyles\(active\) \{[\s\S]*?\n    \}\n/)[0];
Object.keys(PAGES).forEach(function(key) {
  const p = PAGES[key];
  assert(p.btn && p.label && p.icon, key + ': ต้องมี btn/label/icon');
  assert(htmlLf.includes('id="' + p.btn + '"'), key + ': ไม่พบปุ่มเมนู id="' + p.btn + '"');
  assert(setTabStylesSrc.includes("active === '" + key + "'"), key + ': ไม่ใช่ key ที่ setTabStyles ส่งมา — แท็บจะไม่อัปเดตตามหน้า');
});
assert(PAGES.stock, 'ต้องมีหน้า stock เป็นหน้าตั้งต้น/สำรอง');

// ── hook: เปลี่ยนหน้า → แท็บ active อัปเดต, login/session เสร็จ → คืนแท็บเดิม ──
assert(setTabStylesSrc.includes('window.browserTabsOnPageChange(active)'), 'setTabStyles ต้องแจ้งแถบแท็บทุกครั้งที่เปลี่ยนหน้า');
const restoreCalls = htmlLf.match(/window\.browserTabsRestore\(\);/g) || [];
assert(restoreCalls.length >= 2, 'ต้องเรียก browserTabsRestore ทั้งหลัง login และหลังเช็ค session');

// ── behavior ของ state (ไม่ต้องมี DOM): เปิด/ปิด/สลับแท็บ ──────────────────
const blockStart = htmlLf.indexOf('    var BROWSER_TABS_STORAGE_KEY');
const blockEnd = htmlLf.indexOf('    var browserTabListEl = ');
assert(blockStart > -1 && blockEnd > blockStart, 'ต้องดึงโค้ดแท็บออกมาได้');
const clicked = [];
const store = {};
const fakeBtn = function(id) { return { id: id, classList: { contains: function() { return false; } }, click: function() { clicked.push(id); } }; };
const env = new Function('document', 'localStorage', 'window', 'escHtml', 'showQuickToast',
  pagesSrc[0] + '\n' + htmlLf.slice(blockStart, blockEnd) +
  '\nreturn { state: browserTabsState, open: openNewBrowserTab, close: closeBrowserTab, activate: activateBrowserTab };');
const win = {};
const api = env(
  { getElementById: function(id) { return /^browserTab/.test(id) ? null : fakeBtn(id); } },
  { getItem: function(k) { return store[k] || null; }, setItem: function(k, v) { store[k] = v; } },
  win, function(s) { return String(s); }, function() {}
);
// ก่อน restore: หน้าที่ระบบเปิดตอนบูตต้องไม่ถูกจด
win.browserTabsOnPageChange('dashboard');
win.browserTabsRestore();
assert.strictEqual(api.state.tabs.length, 1);
assert.strictEqual(api.state.tabs[0].page, 'dashboard', 'แท็บแรกใช้หน้าที่เปิดอยู่');
api.open('stock');
assert.strictEqual(api.state.tabs.length, 2);
assert.deepStrictEqual(clicked, ['tabStock'], 'เปิดแท็บใหม่ต้องพาไปหน้าของแท็บนั้น');
win.browserTabsOnPageChange('log'); // ผู้ใช้กดเมนู Log ในแท็บใหม่
assert.strictEqual(api.state.tabs[1].page, 'log', 'แท็บ active ต้องเปลี่ยนตามหน้า');
assert.strictEqual(api.state.tabs[0].page, 'dashboard', 'แท็บอื่นต้องไม่เปลี่ยน');
api.activate(api.state.tabs[0].id);
assert.strictEqual(clicked[clicked.length - 1], 'tabDashboard', 'สลับแท็บต้องกลับไปหน้าเดิมของแท็บ');
win.browserTabsOnPageChange('dashboard');
api.close(api.state.tabs[0].id);
assert.strictEqual(api.state.tabs.length, 1);
assert.strictEqual(clicked[clicked.length - 1], 'tabLog', 'ปิดแท็บ active ต้องย้ายไปแท็บข้างๆ');
api.close(api.state.tabs[0].id);
assert.strictEqual(api.state.tabs.length, 1, 'ห้ามปิดแท็บสุดท้าย');
// จำข้ามการรีเฟรช
const saved = JSON.parse(store.spare_browser_tabs);
assert.strictEqual(saved.tabs[0].page, 'log');

console.log('browser-tab-bar: ok');
