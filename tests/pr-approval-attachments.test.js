const fs = require('fs');
const assert = require('assert');
const html = fs.readFileSync('index.html', 'utf8');
const backend = fs.readFileSync('scr/Backend.gs', 'utf8');

// ก่อนส่ง PR ให้อนุมัติ ต้องแนบเอกสาร (ใบเสนอราคา/รูป/สเปก) ไปให้ผู้อนุมัติเห็นใน Inbox ได้

// ---- Backend: คอลัมน์ต่อท้าย (additive) + อัปโหลดลง Drive แยกโฟลเดอร์ ----
assert(backend.includes("'closed_at', 'close_reason', 'attachments_json'];"), 'PR_HEADER_HEADERS ต่อท้ายด้วย attachments_json');
assert(backend.includes("if (action === 'uploadPrAttachment') return respond(uploadPrAttachment(body), e);"));
const uploadBody = backend.slice(backend.indexOf('function uploadPrAttachment(payload)'), backend.indexOf('// สร้าง PR ใหม่ -> สถานะ PENDING'));
assert(uploadBody.includes("requirePermission({ authToken: payload.authToken }, 'pr_create')"), 'ต้องมีสิทธิ์สร้าง PR ถึงจะอัปโหลดได้');
assert(uploadBody.includes("getOrCreateChildFolder(getOrCreateChildFolder(root, 'pr-attachments'), month)"));
assert(uploadBody.includes("'application/pdf': 'pdf'") && uploadBody.includes("'xlsx'") && uploadBody.includes("'docx'"));

// รับเฉพาะลิงก์ Drive + จำกัดจำนวน
const parseBody = backend.slice(backend.indexOf('function parsePrAttachmentsInput(raw)'), backend.indexOf('function uploadPrAttachment'));
eval(parseBody.replace('function parsePrAttachmentsInput', 'var PR_ATTACH_MAX_FILES = 10; global.parsePrAttachmentsInput = function'));
assert.deepStrictEqual(parsePrAttachmentsInput(''), []);
assert.deepStrictEqual(parsePrAttachmentsInput('not json'), []);
assert.deepStrictEqual(parsePrAttachmentsInput(JSON.stringify([
  { url: 'https://drive.google.com/file/d/abc/view', name: 'quote.pdf' },
  { url: 'javascript:alert(1)', name: 'x' }
])), [{ url: 'https://drive.google.com/file/d/abc/view', name: 'quote.pdf' }]);
assert.strictEqual(parsePrAttachmentsInput(Array.from({ length: 15 }, () => ({ url: 'https://drive.google.com/file/d/a/view' }))).length, 10);

const createBody = backend.slice(backend.indexOf('function createPRUnlocked(payload)'), backend.indexOf('\nfunction getPRForApproval'));
assert(createBody.includes('parsePrAttachmentsInput(payload.attachments_json)'), 'createPR บันทึกเอกสารแนบลง header');
const getBody = backend.slice(backend.indexOf('function getPRForApproval(payload)'), backend.indexOf('\nfunction getPrApprovers'));
assert(getBody.includes('header.attachments ='), 'getPRForApproval ส่งรายการเอกสารแนบกลับ');

// ---- Frontend: ปุ่มแนบเอกสารอยู่ในกล่องส่งอนุมัติ ----
assert(/<input id="prAttachFile" type="file" multiple/.test(html));
assert(html.includes('id="prAttachList"') && html.includes('📎 แนบเอกสาร'));
assert(html.includes("action: 'uploadPrAttachment'"));
const submitBody = html.slice(html.indexOf('function prSubmitCreatePr()'), html.indexOf('\n    if (prPrintBtn)'));
assert(submitBody.includes("if (prAttachUploading) { showQuickToast('รออัปโหลดเอกสารแนบให้เสร็จก่อน'"), 'ห้ามส่งระหว่างไฟล์ยังอัปโหลดไม่เสร็จ');
assert(submitBody.includes('payload.attachments_json = JSON.stringify(prAttachments)'));
assert(submitBody.includes('prAttachments = []; prRenderAttachments();'), 'ส่งสำเร็จแล้วล้างไฟล์แนบ');

// ---- Frontend: modal ผู้อนุมัติแสดงเอกสารแนบ ----
assert(html.includes('id="prApprovalAttachments"'));
const openBody = html.slice(html.indexOf('function openPrApproval(prId)'), html.indexOf('\n    function closePrApproval'));
assert(openBody.includes('res.header && res.header.attachments'));
assert(openBody.includes("escHtml(att.url)"), 'ลิงก์ต้อง escape');

console.log('pr-approval-attachments: all assertions passed');
