const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ---- Backend: schema (qty_requested locked, qty_approved separate + audit) ----
assert(backend.includes("var PR_HEADER_HEADERS = ['pr_id', 'status', 'created_by', 'created_at', 'line', 'dept', 'item_count', 'total_amount', 'approved_by', 'approved_at', 'reject_reason', 'updated_at', 'assigned_to', 'budget_snapshot_html']"), 'PR header schema');
assert(backend.includes("'qty_requested', 'qty_approved'"), 'PR lines keep qty_requested and qty_approved separately');
assert(backend.includes("var PR_AUDIT_HEADERS = ['timestamp', 'pr_id', 'action', 'actor', 'line', 'old_qty', 'new_qty', 'detail']"), 'PR audit schema');
assert(backend.includes("var PR_STATUSES = ['DRAFT', 'PENDING', 'APPROVED', 'REJECTED']"), 'PR statuses');

// ---- Backend: approvePR is atomic (LockService) and re-guards status ----
assert(/function approvePR\(payload\) \{\s*var lock = LockService\.getScriptLock\(\);\s*lock\.waitLock\(30000\);/.test(backend), 'approvePR uses LockService like borrow/transfer');
assert(backend.includes("if (curStatus !== 'PENDING') {"), 'approvePR rejects non-PENDING (กันกดอนุมัติซ้ำ)');
assert(backend.includes('อาจถูกดำเนินการไปแล้ว'), 'approvePR returns duplicate-approve error');
// writes qty_approved only, never qty_requested
assert(backend.includes('linesSheet.getRange(i + 1, lIdx.qty_approved + 1).setValue(newApproved);'), 'approvePR writes qty_approved per line');
assert(!/lIdx\.qty_requested \+ 1\).setValue/.test(backend), 'approvePR must never overwrite qty_requested');
// audit on edit + approve
assert(backend.includes("appendPrAudit(prId, 'EDIT_QTY'"), 'approvePR audits qty edits');
assert(backend.includes("appendPrAudit(prId, 'APPROVE'"), 'approvePR audits approval');
// header set to APPROVED with approver + time
assert(backend.includes("headerSheet.getRange(hRow + 1, hIdx.status + 1).setValue('APPROVED');"), 'approvePR sets APPROVED');
assert(backend.includes('hIdx.approved_by + 1'), 'approvePR records approver');

// ---- Backend: rejectPR requires reason, guards PENDING, audits ----
assert(backend.includes("throw new Error('กรุณากรอกเหตุผลการตีกลับ');"), 'rejectPR requires reason');
assert(backend.includes("headerSheet.getRange(hRow + 1, hIdx.status + 1).setValue('REJECTED');"), 'rejectPR sets REJECTED');
assert(backend.includes("appendPrAudit(prId, 'REJECT'"), 'rejectPR audits');

// ---- Backend: createPR persists PENDING, seeds qty_approved = qty_requested ----
assert(backend.includes("headerRowArr[hIdx.status] = 'PENDING';"), 'createPR creates PENDING header');
assert(backend.includes('rowArr[lIdx.qty_requested] = qtyReq;'), 'createPR stores locked qty_requested');
assert(backend.includes('rowArr[lIdx.qty_approved] = qtyReq;'), 'createPR seeds qty_approved from requested');

// ---- Backend: role/line filtering + permissions ----
assert(backend.includes('function prUserCanAccessLine(user, prLine)'), 'line-scope helper');
assert(backend.includes('pr_create: true, pr_view_own: true, pr_view_all: true, pr_approve: true'), 'admin/manager gets pr_view_all');
assert(backend.includes('pr_create: true, pr_view_own: true, pr_view_all: false, pr_approve: true'), 'leader/supervisor sees own line + approves');
assert(backend.includes('pr_create: true, pr_view_own: false, pr_view_all: false, pr_approve: false'), 'user can only create');

// ---- Backend: getPRStatus (live status recheck) + routing ----
assert(backend.includes('function getPRStatus(payload)'), 'getPRStatus exists');
['getInbox', 'createPR', 'getPRForApproval', 'approvePR', 'rejectPR', 'getPRStatus'].forEach(function(action) {
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(e.parameter), e);"), 'doGet routes ' + action);
  assert(backend.includes("if (action === '" + action + "') return respond(" + action + "(body), e);"), 'doPost routes ' + action);
});

// ---- Backend: Users gains a line column, surfaced to client ----
assert(backend.includes('function ensureUsersLineColumn(sheet)'), 'Users line column migration');
assert(backend.includes('line: String(user.line || \'\')'), 'sanitizeUserForClient exposes line');

// ---- Frontend: builder button now creates PENDING PR (no instant print/record) ----
assert(html.includes('📤 ส่งอนุมัติ (สร้าง PR)'), 'builder button relabelled to create PR');
assert(html.includes("action: 'createPR'"), 'builder posts createPR');
assert(html.includes('🔒 ต้องให้หัวหน้าอนุมัติใน Inbox ก่อน จึงจะปริ้นได้'), 'builder hint explains approval gate');

// ---- Frontend: Inbox UI + PR pending tab ----
assert(html.includes('id="inboxOverlay"'), 'inbox overlay present');
assert(html.includes('id="tabInbox"'), 'inbox nav tab present');
assert(html.includes("key: 'pr_pending', label: 'PR รออนุมัติ'"), 'PR pending tab implemented');
assert(html.includes('function openPrApproval(prId)'), 'approval editor');
assert(html.includes('data-appr-line'), 'per-line qty_approved inputs');
assert(html.includes('function submitApprove()') && html.includes('function submitReject()'), 'approve/reject actions');

// ---- Frontend: print gate — live status recheck, APPROVED only, uses qty_approved ----
assert(html.includes("action: 'getPRStatus'"), 'print rechecks live status');
assert(html.includes("st.pr_status !== 'APPROVED'"), 'print only when APPROVED');
assert(html.includes('ln.qty_approved') && html.includes('function openApprovedPrPrintWindow'), 'printed doc uses qty_approved');

// ---- Frontend: user admin gains line field ----
assert(html.includes('id="uLine"'), 'user form has line input');
assert(html.includes("line: (document.getElementById('uLine') || {}).value || ''"), 'upsertUser sends line');

console.log('PR approval + Inbox contract checks passed');
