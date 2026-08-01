const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// เมื่อกดส่งอนุมัติ PR ต้องแนบกราฟงบ Spare part (แผง #prBudgetPanel ที่คนส่งเห็นอยู่ตรงๆ)
// ไปให้หัวหน้าเห็นตอนตรวจใน Inbox ด้วย ไม่ใช่แค่ตัวเลขรายบรรทัด

// ---- Backend: schema เก็บสแนปช็อตไว้ที่ PR header (additive column) ----
assert(backend.includes("'assigned_to', 'budget_snapshot_html']"), 'PR_HEADER_HEADERS gains budget_snapshot_html');

// ---- Backend: createPR เขียนสแนปช็อตลง header ----
const createPRBody = (function() {
  const start = backend.indexOf('function createPRUnlocked(payload)');
  const end = backend.indexOf('\nfunction getPRForApproval', start);
  assert(start > -1 && end > start, 'can slice createPRUnlocked body');
  return backend.slice(start, end);
})();
assert(createPRBody.includes('hIdx.budget_snapshot_html !== undefined'), 'createPR guards on additive column presence');
assert(createPRBody.includes("String(payload.budget_snapshot_html || '').slice(0, 20000)"), 'createPR stores the snapshot with a defensive length cap');

// ---- Backend: getPRForApproval returns it, but listPrCardsForUser (Inbox list/badge) does not ----
const getPRForApprovalBody = (function() {
  const start = backend.indexOf('function getPRForApproval(payload)');
  const end = backend.indexOf('\nfunction getPrApprovers', start);
  assert(start > -1 && end > start, 'can slice getPRForApproval body');
  return backend.slice(start, end);
})();
assert(getPRForApprovalBody.includes('header.budget_snapshot_html ='), 'getPRForApproval attaches the snapshot to the single-PR detail response');
const prHeaderRowToCardBody = (function() {
  const start = backend.indexOf('function prHeaderRowToCard(hIdx, row)');
  const end = backend.indexOf('\n}', start);
  assert(start > -1 && end > start, 'can slice prHeaderRowToCard body');
  return backend.slice(start, end);
})();
assert(!prHeaderRowToCardBody.includes('budget_snapshot_html'), 'prHeaderRowToCard (used by the polled Inbox list) stays lean and omits the snapshot');

// ---- Frontend: capture the live budget panel HTML at submit time ----
const prSubmitCreatePrBody = (function() {
  const start = html.indexOf('function prSubmitCreatePr()');
  const end = html.indexOf('\n    if (prPrintBtn)', start);
  assert(start > -1 && end > start, 'can slice prSubmitCreatePr body');
  return html.slice(start, end);
})();
assert(prSubmitCreatePrBody.includes("getElementById('prBudgetPanel')"), 'prSubmitCreatePr reads the rendered #prBudgetPanel');
assert(prSubmitCreatePrBody.includes('payload.budget_snapshot_html = prBudgetPanelEl.innerHTML'), 'prSubmitCreatePr attaches the panel HTML to the createPR payload');

// ---- Frontend: approval modal has a container and renders the snapshot when opened ----
assert(html.includes('id="prApprovalBudgetSnapshot"'), 'approval modal has a snapshot container');
const openPrApprovalBody = (function() {
  const start = html.indexOf('function openPrApproval(prId)');
  const end = html.indexOf('\n    function closePrApproval', start);
  assert(start > -1 && end > start, 'can slice openPrApproval body');
  return html.slice(start, end);
})();
assert(openPrApprovalBody.includes('res.header && res.header.budget_snapshot_html'), 'openPrApproval reads the snapshot from the PR header response');
assert(/budgetSnapshotEl\.innerHTML = snapshotHtml; budgetSnapshotEl\.classList\.remove\('hidden'\)/.test(openPrApprovalBody), 'openPrApproval shows the container when a snapshot exists');
assert(/budgetSnapshotEl\.innerHTML = ''; budgetSnapshotEl\.classList\.add\('hidden'\)/.test(openPrApprovalBody), 'openPrApproval hides/clears the container when there is no snapshot (older PRs)');

console.log('pr-budget-snapshot-attachment: all assertions passed');
