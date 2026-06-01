// =============================
// CONFIG
// =============================
var SPARE_APP_CONFIG = this.SPARE_APP_CONFIG || {};
SPARE_APP_CONFIG.readSheetName = SPARE_APP_CONFIG.readSheetName || 'Main List Stock';
SPARE_APP_CONFIG.writeSheetName = SPARE_APP_CONFIG.writeSheetName || 'Log';
SPARE_APP_CONFIG.usersSheetName = SPARE_APP_CONFIG.usersSheetName || 'Users';
SPARE_APP_CONFIG.requestSheetName = SPARE_APP_CONFIG.requestSheetName || 'OrderRequests';
var LOG_HEADERS = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After', 'Reason', 'Reason Remark'];
var USER_HEADERS = ['username', 'password', 'role', 'is_active', 'permissions_json', 'session_token', 'session_expiry'];
var ORDER_REQUEST_HEADERS = ['request_id', 'requested_date', 'requested_by', 'requester_role', 'item_id', 'item_name', 'model', 'brand', 'category', 'line', 'current_stock', 'min', 'max', 'request_qty', 'priority', 'reason', 'expected_use_date', 'remark', 'attachment_url', 'status', 'admin_comment', 'approved_by', 'approved_date', 'converted_pr_id', 'updated_at'];
var ORDER_REQUEST_STATUSES = ['Pending', 'Approved', 'Rejected', 'On Hold', 'Converted to PR', 'Purchased', 'Received', 'Closed'];
var STOCK_LOCATION_SHEETS = ['Main List Stock', 'Stock for MC', 'Standard Spare part', 'Arc chut', 'Common Gv.2', 'Gv.2 (6 plate)', 'Gv.2 (9 plate)', 'Coil Winding', 'Lug&Screw'];
var DRIVE_ROOT_FOLDER_ID = '1XWO5rGpku35gSTMAh4HDOCHa6GJIkoS3';
var DRAWING_STATUS_OPTIONS = ['Available', 'Not Available', 'Not Required', 'Pending Update'];
var PART_ATTACHMENT_COLUMNS = [
  { label: 'Photo URL', aliases: ['photourl', 'photo', 'sparepartsphotos'] },
  { label: 'Drawing URL', aliases: ['drawingurl', 'drawing_url'] },
  { label: 'Drawing File Name', aliases: ['drawingfilename', 'drawing_file_name'] },
  { label: 'Drawing Revision', aliases: ['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev'] },
  { label: 'Drawing Status', aliases: ['drawingstatus', 'drawing_status'] },
  { label: 'Datasheet URL', aliases: ['datasheeturl', 'datasheet_url'] },
  { label: 'Quotation URL', aliases: ['quotationurl', 'quotation_url'] }
];

// =============================
// HELPERS
// =============================

function buildErrorResponse(err) {
  var msg = err && err.message ? String(err.message) : String(err || 'Unknown error');
  var lower = msg.toLowerCase();
  var isDriveAuth = lower.indexOf('ไม่ได้รับอนุญาต') > -1 ||
    lower.indexOf('authorization') > -1 ||
    lower.indexOf('googleapis.com/auth/drive') > -1;
  var isDriveServiceError = lower.indexOf('ข้อผิดพลาดของบริการ: ไดรฟ์') > -1 ||
    lower.indexOf('service error: drive') > -1 ||
    lower.indexOf('drive_service_error') > -1 ||
    lower.indexOf('internal error encountered') > -1;

  if (isDriveAuth) {
    return {
      status: 'error',
      errorCode: 'DRIVE_AUTH_REQUIRED',
      message: 'ยังไม่ได้อนุญาตสิทธิ์ Google Drive ให้ Apps Script (DRIVE_AUTH_REQUIRED). กรุณาเปิด Apps Script แล้ว Run ฟังก์ชันที่ใช้ DriveApp 1 ครั้งเพื่ออนุญาตสิทธิ์ จากนั้น Deploy เว็บแอปใหม่และลองอีกครั้ง',
      detail: msg
    };
  }
  if (isDriveServiceError) {
    return {
      status: 'error',
      errorCode: 'DRIVE_SERVICE_ERROR',
      message: 'ระบบ Google Drive ขัดข้องชั่วคราว (DRIVE_SERVICE_ERROR) กรุณาลองอัปโหลดใหม่อีกครั้ง',
      detail: msg
    };
  }

  return { status: 'error', message: msg };
}

function normalizeHeaderName(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function buildHeaderIndexMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i += 1) {
    map[normalizeHeaderName(headers[i])] = i;
  }
  return map;
}

function getLocationOverrideKey(sheetName, noValue) {
  return 'location_override::' + String(sheetName || '').trim() + '::' + String(noValue || '').trim();
}

function setLocationOverride(sheetName, noValue, locationValue) {
  var props = PropertiesService.getScriptProperties();
  var key = getLocationOverrideKey(sheetName, noValue);
  var normalized = String(locationValue || '').trim();
  if (!normalized || normalized === '-') {
    props.deleteProperty(key);
    return;
  }
  props.setProperty(key, normalized);
}

function getLocationOverride(sheetName, noValue) {
  var props = PropertiesService.getScriptProperties();
  var key = getLocationOverrideKey(sheetName, noValue);
  return String(props.getProperty(key) || '').trim();
}

function ensureLocationColumnForSheet(sheetName) {
  var targetSheet = String(sheetName || '').trim();
  if (!targetSheet) return;
  var ctx = getMainSheetContext(targetSheet);
  ensureColumnInContext(ctx, 'Location', ['location', 'jrlocation']);
}

function ensureLocationColumnsForAllKnownSheets() {
  var candidates = [SPARE_APP_CONFIG.readSheetName].concat(STOCK_LOCATION_SHEETS);
  var uniqueSheets = Array.from(new Set(candidates.filter(function(name) { return !!String(name || '').trim(); })));
  uniqueSheets.forEach(function(sheetName) {
    try {
      ensureLocationColumnForSheet(sheetName);
    } catch (err) {
      Logger.log('ensureLocationColumnsForAllKnownSheets warning [' + sheetName + ']: ' + (err && err.message ? err.message : err));
    }
  });
}


function ensurePriceColumnsForSheet(sheetName) {
  var targetSheet = String(sheetName || '').trim();
  if (!targetSheet) return;
  var ctx = getMainSheetContext(targetSheet);
  ensureColumnInContext(ctx, 'Unit Price', ['unitprice', 'unit_price']);
  ensureColumnInContext(ctx, 'Currency', ['currency']);
  ensureColumnInContext(ctx, 'Supplier', ['supplier']);
  ensureColumnInContext(ctx, 'Price Updated At', ['priceupdatedat', 'price_updated_at']);
  ensureColumnInContext(ctx, 'Price Remark', ['priceremark', 'price_remark']);
}

function ensurePriceColumnsForAllKnownSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var discoveredSheets = spreadsheet.getSheets().map(function(s) { return s.getName(); });
  var candidates = [SPARE_APP_CONFIG.readSheetName].concat(STOCK_LOCATION_SHEETS).concat(discoveredSheets);
  var uniqueSheets = Array.from(new Set(candidates.filter(function(name) { return !!String(name || '').trim(); })));
  uniqueSheets.forEach(function(sheetName) {
    try {
      ensurePriceColumnsForSheet(sheetName);
    } catch (err) {
      Logger.log('ensurePriceColumnsForAllKnownSheets warning [' + sheetName + ']: ' + (err && err.message ? err.message : err));
    }
  });
}

function ensureAttachmentColumnsForSheet(sheetName) {
  var targetSheet = String(sheetName || '').trim();
  if (!targetSheet) return;
  var ctx = getMainSheetContext(targetSheet);
  PART_ATTACHMENT_COLUMNS.forEach(function(col) {
    ensureColumnInContext(ctx, col.label, col.aliases);
  });
}

function ensureAttachmentColumnsForAllKnownSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var discoveredSheets = spreadsheet.getSheets().map(function(s) { return s.getName(); });
  var candidates = [SPARE_APP_CONFIG.readSheetName].concat(STOCK_LOCATION_SHEETS).concat(discoveredSheets);
  var uniqueSheets = Array.from(new Set(candidates.filter(function(name) { return !!String(name || '').trim(); })));
  uniqueSheets.forEach(function(sheetName) {
    try {
      ensureAttachmentColumnsForSheet(sheetName);
    } catch (err) {
      Logger.log('ensureAttachmentColumnsForAllKnownSheets warning [' + sheetName + ']: ' + (err && err.message ? err.message : err));
    }
  });
}

function pickRowValue(row, map, keys, fallbackValue) {
  for (var i = 0; i < keys.length; i += 1) {
    var idx = map[keys[i]];
    if (idx !== undefined && row[idx] !== '' && row[idx] !== null && row[idx] !== undefined) {
      return row[idx];
    }
  }
  return fallbackValue;
}


function buildDriveViewUrlFromFileId(fileId) {
  var id = String(fileId || '').trim();
  return id ? ('https://drive.google.com/uc?export=view&id=' + id) : '';
}

function findHeaderRowIndex(data) {
  var requiredHints = ['no', 'name', 'description', 'category', 'brand', 'stock', 'qoh', 'model'];
  var maxScan = Math.min(data.length, 8);

  for (var r = 0; r < maxScan; r += 1) {
    var normalizedRow = data[r].map(function (cell) {
      return normalizeHeaderName(cell);
    });

    var hit = 0;
    for (var i = 0; i < requiredHints.length; i += 1) {
      var hint = requiredHints[i];
      var matched = normalizedRow.some(function (col) {
        return col.indexOf(hint) > -1;
      });
      if (matched) hit += 1;
    }

    if (hit >= 2) return r;
  }

  return 0;
}

function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

function getTemplateHeaders(spreadsheet) {
  var preferred = getSheetByFlexibleName(spreadsheet, SPARE_APP_CONFIG.readSheetName);
  var sheets = preferred ? [preferred].concat(spreadsheet.getSheets().filter(function(s){ return s.getName() !== preferred.getName(); })) : spreadsheet.getSheets();

  for (var i = 0; i < sheets.length; i += 1) {
    var data = sheets[i].getDataRange().getValues();
    if (!data || data.length === 0) continue;
    var headerRowIndex = findHeaderRowIndex(data);
    var headers = data[headerRowIndex] || [];
    if (headers.length >= 8) return headers;
  }

  return ['NO', 'Name / Description', 'Model', 'Line', 'Category', 'Brand', 'Location', 'Unit', 'Stock', 'Min', 'Max', 'Need to PO', 'image_main_url', 'image_main_file_id', 'image_install_url', 'image_install_file_id'];
}

function ensureSheetWithTemplate(spreadsheet, sheetName) {
  var sheet = getSheetByFlexibleName(spreadsheet, sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    var headers = getTemplateHeaders(spreadsheet);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function ensureUsersSheetHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(USER_HEADERS);
    return;
  }
  var firstRow = sheet.getRange(1, 1, 1, USER_HEADERS.length).getValues()[0];
  var same = true;
  for (var i = 0; i < USER_HEADERS.length; i += 1) {
    if (String(firstRow[i] || '') !== USER_HEADERS[i]) {
      same = false;
      break;
    }
  }
  if (!same) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  }
}

function getUsersSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.usersSheetName);
  ensureUsersSheetHeaders(usersSheet);
  return usersSheet;
}

function getOrderRequestSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.requestSheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ORDER_REQUEST_HEADERS);
    }
    return sheet;
  } catch (err) {
    Logger.log('getOrderRequestSheet error: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function toRequestObject(headers, row) {
  var out = {};
  headers.forEach(function(h, i) { out[h] = row[i]; });
  return out;
}

function ensureOrderRequestsSheetReady(payload) {
  try {
    requirePermission({ authToken: payload.authToken }, 'view');
    var sheet = getOrderRequestSheet();
    return { status: 'success', sheet: sheet.getName(), ready: true };
  } catch (err) {
    Logger.log('ensureOrderRequestsSheetReady error: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function createOrderRequest(payload) {
  try {
    var session = getSessionUser({ authToken: payload.authToken });
    var user = findUserByUsername(session.user.username);
    requirePermission({ authToken: payload.authToken }, 'request_order_create');
    var sheet = getOrderRequestSheet();
    var now = new Date();
    var requestId = 'REQ-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var attachmentUrl = String(payload.attachment_url || '');
    if (/^data:image\//i.test(attachmentUrl)) {
      attachmentUrl = uploadOrderRequestAttachmentToDrive({
        dataUrl: attachmentUrl,
        requestId: requestId,
        line: payload.line || '',
        requestedBy: user.username || ''
      });
    }
    var row = [
      requestId, payload.requested_date || Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), user.username, user.role,
      payload.item_id || '', payload.item_name || '', payload.model || '', payload.brand || '', payload.category || '',
      payload.line || '', Number(payload.current_stock || 0), Number(payload.min || 0), Number(payload.max || 0), Number(payload.request_qty || 0),
      payload.priority || 'Normal', payload.reason || '', payload.expected_use_date || '', payload.remark || '', attachmentUrl,
      'Pending', '', '', '', '', Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss')
    ];
    sheet.appendRow(row);
    return { status: 'success', request_id: requestId };
  } catch (err) {
    Logger.log('createOrderRequest error: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function uploadOrderRequestAttachmentToDrive(payload) {
  var dataUrl = String(payload.dataUrl || '');
  if (!dataUrl) return '';
  var mimeType = getDataUrlMimeType(dataUrl);
  if (!mimeType) throw new Error('รูปแบบไฟล์แนบไม่ถูกต้อง');
  var allowed = { 'image/jpeg': true, 'image/png': true, 'image/webp': true };
  if (!allowed[mimeType]) throw new Error('ไฟล์แนบรองรับเฉพาะ jpg, png, webp');

  var root = DriveApp.getFolderById('1XWO5rGpku35gSTMAh4HDOCHa6GJIkoS3');
  var reqRoot = getOrCreateChildFolder(root, 'order-requests');
  var lineFolder = getOrCreateChildFolder(reqRoot, String(payload.line || 'UnknownLine'));
  var requesterFolder = getOrCreateChildFolder(lineFolder, String(payload.requestedBy || 'unknown-user'));

  var ext = mimeType === 'image/png' ? 'png' : (mimeType === 'image/webp' ? 'webp' : 'jpg');
  var fileName = String(payload.requestId || ('REQ-' + Date.now())) + '-' + Date.now() + '.' + ext;
  var base64Content = dataUrl.split(',')[1] || '';
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, fileName);
  var file = requesterFolder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('uploadOrderRequestAttachmentToDrive setSharing warning: ' + (err && err.message ? err.message : err));
  }
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function uploadRequestAttachment(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  var user = findUserByUsername(session.user.username);
  requirePermission({ authToken: payload.authToken }, 'request_order_create');
  var dataUrl = String(payload.dataUrl || payload.fileBase64 || '');
  if (!dataUrl) throw new Error('ไม่พบข้อมูลรูปภาพ');
  var requestId = 'REQUPLOAD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var attachmentUrl = uploadOrderRequestAttachmentToDrive({
    dataUrl: dataUrl,
    requestId: requestId,
    line: payload.line || '',
    requestedBy: user.username || ''
  });
  var fileIdMatch = attachmentUrl.match(/id=([^&]+)/);
  return {
    status: 'success',
    attachment_url: attachmentUrl,
    file_id: fileIdMatch ? fileIdMatch[1] : '',
    view_url: fileIdMatch ? ('https://drive.google.com/file/d/' + fileIdMatch[1] + '/view') : ''
  };
}

function getOrderRequests(payload) {
  try {
    var session = getSessionUser({ authToken: payload.authToken });
    var user = findUserByUsername(session.user.username);
    var canViewAll = hasPermissionForUser(user, 'request_order_view_all');
    var canViewOwn = hasPermissionForUser(user, 'request_order_view_own');
    if (!canViewAll && !canViewOwn) throw new Error('ไม่มีสิทธิ์ดูคำขอซื้อ');
    var sheet = getOrderRequestSheet();
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    var headers = values[0];
    return values.slice(1).map(function(r) { return toRequestObject(headers, r); }).filter(function(item) {
      if (canViewAll) return true;
      return String(item.requested_by || '') === String(user.username || '');
    });
  } catch (err) {
    Logger.log('getOrderRequests error: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function updateOrderRequestStatus(payload, nextStatus) {
  if (ORDER_REQUEST_STATUSES.indexOf(nextStatus) === -1) throw new Error('สถานะไม่ถูกต้อง: ' + nextStatus);
  var session = getSessionUser({ authToken: payload.authToken });
  var user = findUserByUsername(session.user.username);
  var sheet = getOrderRequestSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][idx.request_id]) === String(payload.request_id)) {
      sheet.getRange(i + 1, idx.status + 1).setValue(nextStatus);
      sheet.getRange(i + 1, idx.admin_comment + 1).setValue(payload.admin_comment || '');
      sheet.getRange(i + 1, idx.updated_at + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'));
      if (nextStatus === 'Approved') {
        sheet.getRange(i + 1, idx.approved_by + 1).setValue(user.username || '');
        sheet.getRange(i + 1, idx.approved_date + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'));
      }
      return { status: 'success', request_id: payload.request_id, updated_status: nextStatus };
    }
  }
  throw new Error('ไม่พบ request_id');
}

function approveOrderRequest(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_approve'); return updateOrderRequestStatus(payload, 'Approved'); }
function rejectOrderRequest(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_reject'); return updateOrderRequestStatus(payload, 'Rejected'); }
function holdOrderRequest(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_approve'); return updateOrderRequestStatus(payload, 'On Hold'); }
function closeOrderRequest(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_close'); return updateOrderRequestStatus(payload, 'Closed'); }
function markOrderRequestPurchased(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_approve'); return updateOrderRequestStatus(payload, 'Purchased'); }
function markOrderRequestReceived(payload) { requirePermission({ authToken: payload.authToken }, 'request_order_approve'); return updateOrderRequestStatus(payload, 'Received'); }
function convertOrderRequestsToPR(payload) {
  requirePermission({ authToken: payload.authToken }, 'request_order_convert_pr');
  var ids = Array.isArray(payload.request_ids) ? payload.request_ids : [];
  if (!ids.length) throw new Error('ต้องมี request_ids อย่างน้อย 1 รายการ');
  var convertedPrId = payload.converted_pr_id || ('PR-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss'));
  var sheet = getOrderRequestSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0] || [];
  var idxRid = headers.indexOf('request_id');
  var idxPr = headers.indexOf('converted_pr_id');
  if (idxRid === -1 || idxPr === -1) throw new Error('โครงสร้างชีท OrderRequests ไม่ถูกต้อง');

  ids.forEach(function(id) {
    updateOrderRequestStatus({ authToken: payload.authToken, request_id: id, admin_comment: payload.admin_comment || '' }, 'Converted to PR');
    for (var i = 1; i < values.length; i += 1) {
      if (String(values[i][idxRid]) === String(id)) {
        sheet.getRange(i + 1, idxPr + 1).setValue(convertedPrId);
        break;
      }
    }
  });
  return { status: 'success', converted_pr_id: convertedPrId, count: ids.length };
}

function normalizeRole(role) {
  var val = String(role || 'user').toLowerCase().trim();
  if (val === 'admin' || val === 'leader' || val === 'user') return val;
  return 'user';
}

function getRoleDefaultPermissions(role) {
  var normalized = normalizeRole(role);
  if (normalized === 'admin') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: true, add_user: true, delete_user: true, manage_auth: true,
    request_order_create: true, request_order_view_own: true, request_order_view_all: true,
    request_order_approve: true, request_order_reject: true, request_order_convert_pr: true, request_order_close: true
  };
  if (normalized === 'leader') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false
  };
  return {
    view: true, transact: true, manage_items: false, delete_items: false,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false
  };
}

function parsePermissions(raw) {
  if (!raw) return { allow: [], deny: [] };
  try {
    var parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return {
      allow: Array.isArray(parsed.allow) ? parsed.allow : [],
      deny: Array.isArray(parsed.deny) ? parsed.deny : []
    };
  } catch (err) {
    return { allow: [], deny: [] };
  }
}

function mergePermissions(base, custom) {
  var out = {};
  for (var key in base) out[key] = !!base[key];
  (custom.allow || []).forEach(function(p) { out[p] = true; });
  (custom.deny || []).forEach(function(p) { out[p] = false; });
  return out;
}

function toBoolean(val, defaultValue) {
  if (val === undefined || val === null || val === '') return !!defaultValue;
  var s = String(val).toLowerCase().trim();
  return !(s === 'false' || s === '0' || s === 'no');
}

function ensureDefaultAdminUser() {
  var usersSheet = getUsersSheet();
  if (usersSheet.getLastRow() > 1) return;
  usersSheet.appendRow([
    'admin',
    'admin123',
    'admin',
    'true',
    JSON.stringify({ allow: [], deny: [] }),
    '',
    ''
  ]);
}

function getAllUsers() {
  ensureDefaultAdminUser();
  var usersSheet = getUsersSheet();
  var data = usersSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(row, idx) {
    var role = normalizeRole(row[2]);
    var custom = parsePermissions(row[4]);
    return {
      rowIndex: idx + 2,
      username: String(row[0] || ''),
      password: String(row[1] || ''),
      role: role,
      isActive: toBoolean(row[3], true),
      permissionsJson: JSON.stringify(custom),
      token: String(row[5] || ''),
      tokenExpiry: String(row[6] || ''),
      permissions: mergePermissions(getRoleDefaultPermissions(role), custom)
    };
  }).filter(function(u) { return !!u.username; });
}

function findUserByUsername(username) {
  var target = String(username || '').trim();
  if (!target) return null;
  var users = getAllUsers();
  for (var i = 0; i < users.length; i += 1) {
    if (users[i].username === target) return users[i];
  }
  return null;
}

function findUserByToken(token) {
  var target = String(token || '').trim();
  if (!target) return null;
  var users = getAllUsers();
  var now = Date.now();
  for (var i = 0; i < users.length; i += 1) {
    var u = users[i];
    var exp = Number(u.tokenExpiry || 0);
    if (u.token === target && exp > now && u.isActive) return u;
  }
  return null;
}

function sanitizeUserForClient(user) {
  return {
    username: user.username,
    role: user.role,
    isActive: user.isActive,
    permissions: user.permissions,
    permissionsJson: user.permissionsJson
  };
}

function loginUser(payload) {
  var username = String(payload.username || '').trim();
  var password = String(payload.password || '').trim();
  if (!username || !password) throw new Error('กรุณาระบุ username และ password');
  var user = findUserByUsername(username);
  if (!user || !user.isActive) throw new Error('ไม่พบผู้ใช้หรือผู้ใช้ถูกปิดใช้งาน');
  if (String(user.password) !== password) throw new Error('รหัสผ่านไม่ถูกต้อง');

  var token = Utilities.getUuid() + '-' + Date.now();
  var expiry = Date.now() + (8 * 60 * 60 * 1000);
  var usersSheet = getUsersSheet();
  usersSheet.getRange(user.rowIndex, 6).setValue(token);
  usersSheet.getRange(user.rowIndex, 7).setValue(String(expiry));

  user.token = token;
  user.tokenExpiry = String(expiry);
  return {
    status: 'success',
    token: token,
    expiry: expiry,
    user: sanitizeUserForClient(user)
  };
}

function logoutUser(payload) {
  var token = String(payload.authToken || payload.token || '').trim();
  if (!token) return { status: 'success' };
  var user = findUserByToken(token);
  if (!user) return { status: 'success' };
  var usersSheet = getUsersSheet();
  usersSheet.getRange(user.rowIndex, 6).setValue('');
  usersSheet.getRange(user.rowIndex, 7).setValue('');
  return { status: 'success' };
}

function getSessionUser(payload) {
  var token = String(payload.authToken || payload.token || '').trim();
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  var user = findUserByToken(token);
  if (!user) throw new Error('session หมดอายุหรือไม่ถูกต้อง');
  return { status: 'success', user: sanitizeUserForClient(user) };
}

function requirePermission(payload, permissionName) {
  var session = getSessionUser(payload);
  var user = findUserByUsername(session.user.username);
  if (!user) throw new Error('ไม่พบผู้ใช้');
  if (!user.permissions[permissionName]) {
    throw new Error('ไม่มีสิทธิ์ใช้งานฟังก์ชันนี้ (' + permissionName + ')');
  }
  return user;
}

function hasPermissionForUser(user, permissionName) {
  return !!(user && user.permissions && user.permissions[permissionName]);
}

function requireAdminUser(payload) {
  var user = requirePermission(payload, 'manage_users');
  if (normalizeRole(user.role) !== 'admin') {
    throw new Error('เฉพาะ Admin เท่านั้นที่เข้าถึงหน้านี้ได้');
  }
  return user;
}

function listUsers(payload) {
  requireAdminUser(payload);
  return {
    status: 'success',
    users: getAllUsers().map(function(u) { return sanitizeUserForClient(u); })
  };
}

function upsertUser(payload) {
  var actor = requireAdminUser(payload);
  var username = String(payload.username || '').trim();
  if (!username) throw new Error('ต้องระบุ username');
  var role = normalizeRole(payload.role || 'user');
  var isActive = toBoolean(payload.isActive, true);
  var password = String(payload.password || '').trim();
  var permissionsObj = parsePermissions(payload.permissionsJson || payload.permissions || '');
  var permissionsJson = JSON.stringify(permissionsObj);

  var usersSheet = getUsersSheet();
  var existing = findUserByUsername(username);
  if (existing) {
    if (String(payload.password || '') !== '') {
      usersSheet.getRange(existing.rowIndex, 2).setValue(password);
    }
    usersSheet.getRange(existing.rowIndex, 3).setValue(role);
    usersSheet.getRange(existing.rowIndex, 4).setValue(String(isActive));
    usersSheet.getRange(existing.rowIndex, 5).setValue(permissionsJson);
    return { status: 'success', mode: 'update', username: username };
  }

  if (!actor.permissions.add_user) throw new Error('ไม่มีสิทธิ์เพิ่มผู้ใช้');
  if (!password) throw new Error('ต้องระบุ password สำหรับผู้ใช้ใหม่');
  usersSheet.appendRow([username, password, role, String(isActive), permissionsJson, '', '']);
  return { status: 'success', mode: 'create', username: username };
}

function deleteUser(payload) {
  var actor = requireAdminUser(payload);
  var username = String(payload.username || '').trim();
  if (!username) throw new Error('ต้องระบุ username');
  if (username === actor.username) throw new Error('ไม่สามารถลบ user ตัวเองได้');
  var existing = findUserByUsername(username);
  if (!existing) throw new Error('ไม่พบผู้ใช้');
  var usersSheet = getUsersSheet();
  usersSheet.deleteRow(existing.rowIndex);
  return { status: 'success', username: username };
}

function getSheetByFlexibleName(spreadsheet, requestedName) {
  var exact = spreadsheet.getSheetByName(requestedName);
  if (exact) return exact;

  var target = normalizeHeaderName(requestedName);
  if (!target) return null;

  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i += 1) {
    var normalized = normalizeHeaderName(sheets[i].getName());
    if (normalized === target) return sheets[i];
  }

  for (var x = 0; x < sheets.length; x += 1) {
    var n = normalizeHeaderName(sheets[x].getName());
    if (n.indexOf(target) > -1 || target.indexOf(n) > -1) return sheets[x];
  }
  return null;
}

function ensureLogSheetHeaders(historySheet) {
  var lastRow = historySheet.getLastRow();
  if (lastRow === 0) {
    historySheet.appendRow(LOG_HEADERS);
    return;
  }

  var oldHeaders = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After'];
  var currentWidth = Math.max(historySheet.getLastColumn(), LOG_HEADERS.length);
  var firstRow = historySheet.getRange(1, 1, 1, currentWidth).getValues()[0];
  var oldMatch = true;
  for (var oh = 0; oh < oldHeaders.length; oh += 1) {
    if (String(firstRow[oh] || '') !== oldHeaders[oh]) { oldMatch = false; break; }
  }
  if (oldMatch) {
    historySheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    return;
  }
  var isSame = true;
  for (var i = 0; i < LOG_HEADERS.length; i += 1) {
    if (String(firstRow[i] || '') !== LOG_HEADERS[i]) {
      isSame = false;
      break;
    }
  }

  if (!isSame) {
    historySheet.insertRowBefore(1);
    historySheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  }
}

function parseTransactionPayloadFromGet(e) {
  return {
    partNo: e.parameter.partNo,
    type: e.parameter.type,
    process: e.parameter.process,
    category: e.parameter.category,
    partName: e.parameter.partName,
    model: e.parameter.model,
    brand: e.parameter.brand,
    qty: e.parameter.qty,
    unit: e.parameter.unit,
    by: e.parameter.by,
    reason: e.parameter.reason,
    reasonRemark: e.parameter.reasonRemark,
    sheetName: e.parameter.sheet
  };
}

function getLogRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.writeSheetName);
  ensureLogSheetHeaders(historySheet);

  var data = historySheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headerMap = buildHeaderIndexMap(data[0] || []);

  function pick(row, keys, fallback) {
    for (var i = 0; i < keys.length; i += 1) {
      var idx = headerMap[keys[i]];
      if (idx !== undefined) return row[idx];
    }
    return fallback;
  }

  return data.slice(1).map(function (row, idx) {
    return {
      no: idx + 1,
      timestamp: pick(row, ['timestamp'], ''),
      type: pick(row, ['type'], ''),
      process: pick(row, ['process'], ''),
      category: pick(row, ['category'], ''),
      partName: pick(row, ['partname'], ''),
      model: pick(row, ['model'], ''),
      brand: pick(row, ['brand'], ''),
      qty: pick(row, ['qty'], 0),
      unit: pick(row, ['unit'], ''),
      by: pick(row, ['by'], ''),
      reason: pick(row, ['reason'], ''),
      reasonRemark: pick(row, ['reasonremark'], ''),
      partNo: pick(row, ['partno'], ''),
      stockBefore: pick(row, ['stockbefore'], 0),
      stockAfter: pick(row, ['stockafter'], 0)
    };
  }).reverse();
}

function processTransaction(payload) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.writeSheetName);
  var resolvedSheetName = resolveReadSheetName({ sheet: payload.sheetName });
  var mainSheet = ensureSheetWithTemplate(spreadsheet, resolvedSheetName);

  ensureLogSheetHeaders(historySheet);
  if (!payload.partName || !payload.qty) throw new Error('ต้องมี partName และ qty');

  var qty = Number(payload.qty);
  if (!qty || qty <= 0) throw new Error('qty ต้องมากกว่า 0');

  var signedQty = qty;
  if (payload.type && String(payload.type).indexOf('Output') > -1) signedQty = -Math.abs(qty);
  else signedQty = Math.abs(qty);

  var mainData = mainSheet.getDataRange().getValues();
  if (!mainData.length || mainData.length <= 1) throw new Error('ยังไม่มีข้อมูลอะไหล่ในชีท ' + resolvedSheetName);

  var headerRowIndex = findHeaderRowIndex(mainData);
  var headers = mainData[headerRowIndex];
  var map = buildHeaderIndexMap(headers);
  var rows = mainData.slice(headerRowIndex + 1);

  var stockCol = map.stockqty !== undefined ? map.stockqty : (map.qtystock !== undefined ? map.qtystock : (map.qoh !== undefined ? map.qoh : map.stock));
  var minCol = map.min;
  var needPoCol = map.needtopo !== undefined ? map.needtopo : map.needpo;

  if (stockCol === undefined) throw new Error('ไม่พบคอลัมน์ stock/stock qty');

  var targetIndex = -1;
  for (var i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    var rowNo = pickRowValue(row, map, ['no'], '');
    var rowName = pickRowValue(row, map, ['namedescriptions', 'name', 'description', 'partname', 'jrpartname', 'jrpartnameolderp'], '');
    var rowModel = pickRowValue(row, map, ['model', 'codeno', 'jrcodeno'], '');
    var noMatch = payload.partNo !== undefined && String(rowNo) === String(payload.partNo);
    var nameMatch = String(rowName) === String(payload.partName);
    var modelMatch = !payload.model || String(rowModel) === String(payload.model);
    var strictNoMatch = noMatch && modelMatch && (!payload.partName || nameMatch);
    if (strictNoMatch || (nameMatch && modelMatch)) {
      targetIndex = i;
      break;
    }
  }

  if (targetIndex === -1) throw new Error('ไม่พบอะไหล่ที่ต้องการเบิก/คืนในชีทหลัก');

  var targetRow = rows[targetIndex];
  var stockBefore = Number(targetRow[stockCol]) || 0;
  var stockAfter = stockBefore + signedQty;
  if (stockAfter < 0) throw new Error('สต็อกไม่พอสำหรับการเบิกออก');

  var sheetRowNumber = headerRowIndex + 2 + targetIndex;
  mainSheet.getRange(sheetRowNumber, stockCol + 1).setValue(stockAfter);

  if (needPoCol !== undefined) {
    var minValue = minCol !== undefined ? Number(targetRow[minCol]) || 0 : 0;
    var needPoValue = Math.max(minValue - stockAfter, 0);
    mainSheet.getRange(sheetRowNumber, needPoCol + 1).setValue(needPoValue);
  }

  historySheet.appendRow([
    Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss"),
    payload.type || 'Input',
    payload.process || '-',
    payload.category || 'General',
    payload.partName,
    payload.model || '-',
    payload.brand || '-',
    signedQty,
    payload.unit || 'PCS',
    payload.by || 'Unknown',
    payload.partNo || '',
    stockBefore,
    stockAfter,
    payload.reason || '',
    payload.reasonRemark || ''
  ]);

  return {
    status: 'success',
    stockBefore: stockBefore,
    stockAfter: stockAfter,
    qty: signedQty
  };
}


function resolveReadSheetName(source) {
  var candidate = source && source.sheet ? String(source.sheet).trim() : '';
  return candidate || SPARE_APP_CONFIG.readSheetName;
}

function getMainSheetContext(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureSheetWithTemplate(spreadsheet, sheetName);

  var data = sheet.getDataRange().getValues();
  if (!data.length) {
    var headers = getTemplateHeaders(spreadsheet);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    data = sheet.getDataRange().getValues();
  }

  var headerRowIndex = findHeaderRowIndex(data);
  var headers = data[headerRowIndex];
  var rows = data.slice(headerRowIndex + 1);
  var map = buildHeaderIndexMap(headers);

  return {
    sheet: sheet,
    data: data,
    headerRowIndex: headerRowIndex,
    headers: headers,
    rows: rows,
    map: map
  };
}

function ensureColumnInContext(ctx, headerLabel, aliases) {
  var aliasList = aliases || [normalizeHeaderName(headerLabel)];
  for (var i = 0; i < aliasList.length; i += 1) {
    if (ctx.map[aliasList[i]] !== undefined) return ctx.map[aliasList[i]];
  }

  var newColIndex = ctx.headers.length;
  ctx.sheet.getRange(ctx.headerRowIndex + 1, newColIndex + 1).setValue(headerLabel);
  ctx.headers.push(headerLabel);
  ctx.map = buildHeaderIndexMap(ctx.headers);
  ctx.rows = ctx.rows.map(function(row) {
    row.push('');
    return row;
  });
  return newColIndex;
}

function getOrCreateChildFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getUploadTargetFolder(line, itemId, imageType) {
  var root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var safeLine = String(line || '').trim() || 'UnknownLine';
  var safeItemId = String(itemId || '').trim() || 'UNKNOWN';
  var typeName = imageType === 'install' ? 'install' : 'main';

  var lineFolder = getOrCreateChildFolder(root, safeLine);
  var itemFolder = getOrCreateChildFolder(lineFolder, 'item-' + safeItemId);
  var typeFolder = getOrCreateChildFolder(itemFolder, typeName);

  return {
    folder: typeFolder,
    drivePath: safeLine + '/item-' + safeItemId + '/' + typeName + '/'
  };
}

function getDataUrlMimeType(dataUrl) {
  var m = String(dataUrl || '').match(/^data:([^;]+);base64,/i);
  return m ? m[1].toLowerCase() : '';
}

function uploadImageToDrive(payload) {
  payload = payload || {};
  if (!payload.itemId && !payload.no && !payload.dataUrl && !payload.fileBase64) {
    throw new Error('uploadImageToDrive ต้องรับ payload เช่น { itemId, line, imageType/kind, dataUrl }');
  }

  var itemId = String(payload.itemId || payload.no || '').trim();
  var line = String(payload.line || payload.mainLine || '').trim();
  var kind = String(payload.kind || payload.imageType || 'main').toLowerCase();
  var dataUrl = String(payload.dataUrl || payload.fileBase64 || '');
  if (!itemId) throw new Error('ต้องมี itemId');
  if (!dataUrl) throw new Error('ไม่พบข้อมูลไฟล์');
  if (kind !== 'main' && kind !== 'install') throw new Error('kind ต้องเป็น main หรือ install');

  var mimeType = getDataUrlMimeType(dataUrl);
  if (!mimeType) throw new Error('รูปแบบไฟล์ไม่ถูกต้อง');
  var allowed = { 'image/jpeg': true, 'image/png': true, 'image/webp': true };
  if (!allowed[mimeType]) throw new Error('รองรับเฉพาะ jpg, png, webp');

  var base64Content = dataUrl.split(',')[1] || '';
  var bytes = Utilities.base64Decode(base64Content);
  var ext = mimeType === 'image/png' ? 'png' : (mimeType === 'image/webp' ? 'webp' : 'jpg');
  var fileName = (kind === 'main' ? 'main-' : 'install-') + Date.now() + '.' + ext;
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  var target = getUploadTargetFolder(line, itemId, kind);
  var folder = target.folder;
  var existing = folder.getFiles();
  while (existing.hasNext()) {
    var oldFile = existing.next();
    if (!oldFile.isTrashed()) oldFile.setTrashed(true);
  }

  var file = null;
  var createErr = null;
  var maxCreateAttempts = 3;
  for (var attempt = 0; attempt < maxCreateAttempts; attempt += 1) {
    try {
      file = folder.createFile(blob);
      createErr = null;
      break;
    } catch (err) {
      createErr = err;
      if (attempt < maxCreateAttempts - 1) Utilities.sleep(250 * (attempt + 1));
    }
  }
  if (!file) {
    throw new Error('DRIVE_SERVICE_ERROR: ไม่สามารถสร้างไฟล์ใน Google Drive ได้ (' + (createErr && createErr.message ? createErr.message : 'unknown') + ')');
  }
  var sharingWarning = '';
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    sharingWarning = shareErr && shareErr.message ? String(shareErr.message) : String(shareErr);
    Logger.log('setSharing warning: ' + sharingWarning);
  }

  return {
    ok: true,
    status: 'success',
    itemId: itemId,
    kind: kind,
    fileId: file.getId(),
    imageUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    drivePath: target.drivePath,
    warning: sharingWarning
  };
}


function getSparePartsAttachmentFolder(kind) {
  var normalized = String(kind || '').toLowerCase();
  var folderNameMap = {
    photo: 'Photos',
    photos: 'Photos',
    drawing: 'Drawings',
    drawings: 'Drawings',
    datasheet: 'Datasheets',
    datasheets: 'Datasheets',
    quotation: 'Quotations',
    quotations: 'Quotations'
  };
  var folderName = folderNameMap[normalized];
  if (!folderName) throw new Error('ชนิดไฟล์แนบไม่ถูกต้อง');
  var root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var spareRoot = getOrCreateChildFolder(root, 'SpareParts');
  return {
    folder: getOrCreateChildFolder(spareRoot, folderName),
    folderName: folderName,
    drivePath: 'SpareParts/' + folderName + '/'
  };
}

function getFileExtensionFromName(fileName) {
  var match = String(fileName || '').toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : '';
}

function validateDrawingAttachmentFile(fileName, mimeType) {
  var ext = getFileExtensionFromName(fileName);
  var allowedExt = { pdf: true, dwg: true, dxf: true, jpg: true, jpeg: true, png: true, step: true, stp: true };
  var allowedMime = {
    'application/pdf': true,
    'image/jpeg': true,
    'image/png': true,
    'application/acad': true,
    'application/autocad_dwg': true,
    'application/dwg': true,
    'application/dxf': true,
    'application/octet-stream': true,
    'application/step': true,
    'model/step': true,
    'model/stp': true,
    'text/plain': true,
    'image/vnd.dwg': true,
    'image/vnd.dxf': true
  };
  if (!allowedExt[ext]) throw new Error('รองรับเฉพาะ PDF, DWG, DXF, JPG, PNG, STEP, STP');
  if (mimeType && !allowedMime[String(mimeType || '').toLowerCase()]) {
    Logger.log('validateDrawingAttachmentFile warning unknown mime: ' + mimeType + ' for ' + fileName);
  }
}

function uploadPartAttachmentToDrive(payload) {
  payload = payload || {};
  var kind = String(payload.kind || payload.attachmentType || 'drawing').toLowerCase();
  var itemId = String(payload.itemId || payload.no || '').trim();
  var itemName = String(payload.itemName || payload.name || '').trim();
  var dataUrl = String(payload.dataUrl || payload.fileBase64 || '');
  var originalName = String(payload.fileName || '').trim() || (kind + '-' + Date.now());
  if (!itemId) throw new Error('ต้องมี itemId');
  if (!dataUrl) throw new Error('ไม่พบข้อมูลไฟล์');
  var mimeType = getDataUrlMimeType(dataUrl) || String(payload.mimeType || '').toLowerCase() || 'application/octet-stream';
  if (kind === 'drawing') validateDrawingAttachmentFile(originalName, mimeType);
  var base64Content = dataUrl.split(',')[1] || '';
  var bytes = Utilities.base64Decode(base64Content);
  var safeOriginalName = originalName.replace(/[\\/:*?"<>|#%{}~&]/g, '-');
  var fileName = itemId + '-' + Date.now() + '-' + safeOriginalName;
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var target = getSparePartsAttachmentFolder(kind);
  var file = target.folder.createFile(blob);
  var sharingWarning = '';
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    sharingWarning = shareErr && shareErr.message ? String(shareErr.message) : String(shareErr);
    Logger.log('uploadPartAttachmentToDrive setSharing warning: ' + sharingWarning);
  }
  return {
    ok: true,
    status: 'success',
    itemId: itemId,
    kind: kind,
    fileId: file.getId(),
    fileName: safeOriginalName,
    storedFileName: fileName,
    url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    drivePath: target.drivePath,
    warning: sharingWarning
  };
}

function extractNumericNo(value) {
  var raw = String(value || '').trim();
  if (!raw) return NaN;
  if (/^\d+$/.test(raw)) return Number(raw);
  var match = raw.match(/(\d+)(?!.*\d)/);
  return match ? Number(match[1]) : NaN;
}

function getNextNoBySheet(sheetName) {
  var targetSheet = resolveReadSheetName({ sheet: sheetName });
  var ctx = getMainSheetContext(targetSheet);
  var noCol = ctx.map.no;
  var maxNo = 0;

  if (noCol === undefined) {
    return {
      status: 'success',
      sheet: targetSheet,
      nextNo: '1',
      maxNo: 0,
      scannedRows: 0
    };
  }

  for (var i = 0; i < ctx.rows.length; i += 1) {
    var candidate = extractNumericNo(ctx.rows[i][noCol]);
    if (Number.isFinite(candidate) && candidate > maxNo) maxNo = candidate;
  }

  return {
    status: 'success',
    sheet: targetSheet,
    nextNo: String(maxNo + 1),
    maxNo: maxNo,
    scannedRows: ctx.rows.length
  };
}

function upsertMainItem(payload) {
  var sheetName = resolveReadSheetName({ sheet: payload.sheetName });
  var ctx = getMainSheetContext(sheetName);
  var map = ctx.map;
  var noValue = String(payload.no || '').trim();
  if (!noValue) throw new Error('ต้องมีรหัส NO');

  function findCol(aliases) {
    for (var i = 0; i < aliases.length; i += 1) {
      if (map[aliases[i]] !== undefined) return map[aliases[i]];
    }
    return undefined;
  }

  var fieldCols = {
    no: findCol(['no']),
    name: findCol(['namedescriptions', 'name', 'description', 'partname', 'jrpartname', 'jrpartnameolderp']),
    model: findCol(['model', 'codeno', 'jrcodeno']),
    line: findCol(['mainline', 'line', 'process', 'ไลน์']),
    location: findCol(['location', 'jrlocation']),
    category: findCol(['category']),
    brand: findCol(['brand']),
    photo: findCol(['sparepartsphotos', 'photo', 'photourl', 'photo_url', 'image', 'imageurl', 'picture', 'pic']),
    drawing_url: findCol(['drawingurl', 'drawing_url']),
    drawing_file_name: findCol(['drawingfilename', 'drawing_file_name']),
    drawing_revision: findCol(['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev']),
    drawing_status: findCol(['drawingstatus', 'drawing_status']),
    datasheet_url: findCol(['datasheeturl', 'datasheet_url']),
    quotation_url: findCol(['quotationurl', 'quotation_url']),
    image_main_url: findCol(['image_main_url', 'imagemainurl', 'image_main', 'imagemain']),
    image_main_file_id: findCol(['image_main_file_id', 'imagemainfileid']),
    image_install_url: findCol(['image_install_url', 'imageinstallurl', 'image_install', 'imageinstall']),
    image_install_file_id: findCol(['image_install_file_id', 'imageinstallfileid']),
    image_main: findCol(['image_main', 'imagemain', 'mainimage', 'main_image']),
    image_install: findCol(['image_install', 'imageinstall', 'installimage', 'install_image']),
    max: findCol(['max', 'qtymax']),
    min: findCol(['min', 'qtymin']),
    unit: findCol(['unit']),
    stock: findCol(['stockqty', 'qtystock', 'qoh', 'stock', 'initialstock']),
    unit_price: findCol(['unitprice', 'unit_price']),
    currency: findCol(['currency']),
    supplier: findCol(['supplier']),
    price_updated_at: findCol(['priceupdatedat', 'price_updated_at']),
    price_remark: findCol(['priceremark', 'price_remark']),
    coil_size: findCol(['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size'])
  };

  if (fieldCols.brand === undefined) {
    fieldCols.brand = ensureColumnInContext(ctx, 'Brand', ['brand']);
  }
  if (fieldCols.line === undefined) {
    fieldCols.line = ensureColumnInContext(ctx, 'Line', ['line', 'mainline', 'process', 'ไลน์']);
  }
  if (fieldCols.location === undefined) {
    fieldCols.location = ensureColumnInContext(ctx, 'Location', ['location', 'jrlocation']);
  }
  if (fieldCols.category === undefined) {
    fieldCols.category = ensureColumnInContext(ctx, 'Category', ['category']);
  }
  if (fieldCols.photo === undefined) {
    fieldCols.photo = ensureColumnInContext(ctx, 'Photo URL', ['photourl', 'photo', 'sparepartsphotos']);
  }
  if (fieldCols.drawing_url === undefined) {
    fieldCols.drawing_url = ensureColumnInContext(ctx, 'Drawing URL', ['drawingurl', 'drawing_url']);
  }
  if (fieldCols.drawing_file_name === undefined) {
    fieldCols.drawing_file_name = ensureColumnInContext(ctx, 'Drawing File Name', ['drawingfilename', 'drawing_file_name']);
  }
  if (fieldCols.drawing_revision === undefined) {
    fieldCols.drawing_revision = ensureColumnInContext(ctx, 'Drawing Revision', ['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev']);
  }
  if (fieldCols.drawing_status === undefined) {
    fieldCols.drawing_status = ensureColumnInContext(ctx, 'Drawing Status', ['drawingstatus', 'drawing_status']);
  }
  if (fieldCols.datasheet_url === undefined) {
    fieldCols.datasheet_url = ensureColumnInContext(ctx, 'Datasheet URL', ['datasheeturl', 'datasheet_url']);
  }
  if (fieldCols.quotation_url === undefined) {
    fieldCols.quotation_url = ensureColumnInContext(ctx, 'Quotation URL', ['quotationurl', 'quotation_url']);
  }
  if (fieldCols.image_main === undefined) {
    fieldCols.image_main = ensureColumnInContext(ctx, 'image_main', ['image_main', 'imagemain']);
  }
  if (fieldCols.image_install === undefined) {
    fieldCols.image_install = ensureColumnInContext(ctx, 'image_install', ['image_install', 'imageinstall']);
  }
  if (fieldCols.image_main_url === undefined) {
    fieldCols.image_main_url = ensureColumnInContext(ctx, 'image_main_url', ['image_main_url', 'imagemainurl']);
  }
  if (fieldCols.image_main_file_id === undefined) {
    fieldCols.image_main_file_id = ensureColumnInContext(ctx, 'image_main_file_id', ['image_main_file_id', 'imagemainfileid']);
  }
  if (fieldCols.image_install_url === undefined) {
    fieldCols.image_install_url = ensureColumnInContext(ctx, 'image_install_url', ['image_install_url', 'imageinstallurl']);
  }
  if (fieldCols.unit_price === undefined) {
    fieldCols.unit_price = ensureColumnInContext(ctx, 'Unit Price', ['unitprice', 'unit_price']);
  }
  if (fieldCols.currency === undefined) {
    fieldCols.currency = ensureColumnInContext(ctx, 'Currency', ['currency']);
  }
  if (fieldCols.supplier === undefined) {
    fieldCols.supplier = ensureColumnInContext(ctx, 'Supplier', ['supplier']);
  }
  if (fieldCols.price_updated_at === undefined) {
    fieldCols.price_updated_at = ensureColumnInContext(ctx, 'Price Updated At', ['priceupdatedat', 'price_updated_at']);
  }
  if (fieldCols.price_remark === undefined) {
    fieldCols.price_remark = ensureColumnInContext(ctx, 'Price Remark', ['priceremark', 'price_remark']);
  }
  if (fieldCols.coil_size === undefined) {
    fieldCols.coil_size = ensureColumnInContext(ctx, 'Coil Size', ['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size']);
  }

  if (fieldCols.image_install_file_id === undefined) {
    fieldCols.image_install_file_id = ensureColumnInContext(ctx, 'image_install_file_id', ['image_install_file_id', 'imageinstallfileid']);
  }

  if (fieldCols.no === undefined) throw new Error('ไม่พบคอลัมน์ NO');

  var targetIndex = -1;
  for (var i = 0; i < ctx.rows.length; i += 1) {
    if (String(ctx.rows[i][fieldCols.no]) === noValue) {
      targetIndex = i;
      break;
    }
  }

  var values = {
    no: noValue,
    name: payload.name || '',
    model: payload.model || '',
    line: payload.line || '',
    location: payload.location || '',
    category: payload.category || '',
    brand: payload.brand || '',
    photo: payload.photo || '',
    drawing_url: payload.drawing_url || '',
    drawing_file_name: payload.drawing_file_name || '',
    drawing_revision: payload.drawing_revision || '',
    drawing_status: DRAWING_STATUS_OPTIONS.indexOf(payload.drawing_status) > -1 ? payload.drawing_status : (payload.drawing_status || ''),
    datasheet_url: payload.datasheet_url || '',
    quotation_url: payload.quotation_url || '',
    image_main: payload.image_main || payload.image_main_url || payload.photo || '',
    image_install: payload.image_install || payload.image_install_url || '',
    image_main_url: payload.image_main_url || payload.image_main || payload.photo || '',
    image_main_file_id: payload.image_main_file_id || '',
    image_install_url: payload.image_install_url || payload.image_install || '',
    image_install_file_id: payload.image_install_file_id || '',
    max: payload.max || '',
    min: payload.min || '',
    unit: payload.unit || '',
    stock: payload.stock || '',
    unit_price: payload.unit_price === undefined ? '' : payload.unit_price,
    currency: payload.currency || 'THB',
    supplier: payload.supplier || '',
    price_updated_at: payload.price_updated_at || '',
    price_remark: payload.price_remark || '',
    coil_size: payload.coil_size !== undefined ? payload.coil_size : (payload.machine_model !== undefined ? payload.machine_model : '')
  };
  Logger.log('[upsertMainItem] sheet=%s no=%s location=%s', sheetName, noValue, values.location);
  setLocationOverride(sheetName, noValue, values.location);

  if (targetIndex > -1) {
    var sheetRow = ctx.headerRowIndex + 2 + targetIndex;
    var imageValueKeys = {
      photo: true,
      image_main: true,
      image_install: true,
      image_main_url: true,
      image_main_file_id: true,
      image_install_url: true,
      image_install_file_id: true
    };
    for (var key in fieldCols) {
      if (fieldCols[key] !== undefined) {
        var nextValue = values[key];
        if (key === 'coil_size' && String(nextValue || '').trim() === '') continue;
        if (imageValueKeys[key] && String(nextValue || '').trim() === '') continue;
        ctx.sheet.getRange(sheetRow, fieldCols[key] + 1).setValue(nextValue);
      }
    }
    return { status: 'success', mode: 'update', no: noValue, sheet: sheetName, location: values.location };
  }

  var newRow = new Array(ctx.headers.length);
  for (var x = 0; x < newRow.length; x += 1) newRow[x] = '';
  for (var k in fieldCols) {
    if (fieldCols[k] !== undefined) newRow[fieldCols[k]] = values[k];
  }
  ctx.sheet.appendRow(newRow);
  return { status: 'success', mode: 'create', no: noValue, sheet: sheetName, location: values.location };
}

function deleteMainItem(payload) {
  var sheetName = resolveReadSheetName({ sheet: payload.sheetName });
  var ctx = getMainSheetContext(sheetName);
  var noCol = ctx.map.no;
  var noValue = String(payload.no || '').trim();
  if (!noValue) throw new Error('ต้องระบุ NO เพื่อการลบ');
  if (noCol === undefined) throw new Error('ไม่พบคอลัมน์ NO');

  for (var i = 0; i < ctx.rows.length; i += 1) {
    if (String(ctx.rows[i][noCol]) === noValue) {
      var rowNumber = ctx.headerRowIndex + 2 + i;
      ctx.sheet.deleteRow(rowNumber);
      setLocationOverride(sheetName, noValue, '');
      return { status: 'success', mode: 'delete', no: noValue };
    }
  }

  throw new Error('ไม่พบรายการ NO: ' + noValue);
}

// =============================
// GET (stock + logs + JSONP transaction)
// =============================


function authorizeGoogleDriveAccess() {
  // รันฟังก์ชันนี้จาก Apps Script Editor 1 ครั้งเพื่อให้ Google แสดงหน้าขอสิทธิ์
  var root = DriveApp.getRootFolder();
  return {
    ok: true,
    status: 'success',
    authorized: true,
    message: 'อนุญาตสิทธิ์ Google Drive สำเร็จ',
    rootFolderName: root.getName()
  };
}

function getDriveAuthStatus() {
  try {
    var root = DriveApp.getRootFolder();
    return {
      ok: true,
      status: 'success',
      authorized: true,
      rootFolderName: root.getName()
    };
  } catch (err) {
    return {
      ok: false,
      status: 'error',
      authorized: false,
      message: err && err.message ? err.message : String(err)
    };
  }
}

function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : '';
    var authToken = e && e.parameter ? (e.parameter.authToken || e.parameter.token || '') : '';
    var authPayload = { authToken: authToken };
    if (action === 'login') return respond(loginUser({ username: e.parameter.username, password: e.parameter.password }), e);
    if (action === 'logout') return respond(logoutUser(authPayload), e);
    if (action === 'session') return respond(getSessionUser(authPayload), e);
    if (action === 'listUsers') return respond(listUsers(authPayload), e);
    if (action === 'upsertUser') return respond(upsertUser({
      authToken: authToken,
      username: e.parameter.username,
      password: e.parameter.password,
      role: e.parameter.role,
      isActive: e.parameter.isActive,
      permissionsJson: e.parameter.permissionsJson
    }), e);
    if (action === 'deleteUser') return respond(deleteUser({ authToken: authToken, username: e.parameter.username }), e);
    if (action === 'createOrderRequest') return respond(createOrderRequest(e.parameter), e);
    if (action === 'uploadRequestAttachment') return respond(uploadRequestAttachment(e.parameter), e);
    if (action === 'getOrderRequests') return respond(getOrderRequests(e.parameter), e);
    if (action === 'ensureOrderRequestsSheet') return respond(ensureOrderRequestsSheetReady(e.parameter), e);
    if (action === 'approveOrderRequest') return respond(approveOrderRequest(e.parameter), e);
    if (action === 'rejectOrderRequest') return respond(rejectOrderRequest(e.parameter), e);
    if (action === 'holdOrderRequest') return respond(holdOrderRequest(e.parameter), e);
    if (action === 'closeOrderRequest') return respond(closeOrderRequest(e.parameter), e);
    if (action === 'markOrderRequestPurchased') return respond(markOrderRequestPurchased(e.parameter), e);
    if (action === 'markOrderRequestReceived') return respond(markOrderRequestReceived(e.parameter), e);
    if (action === 'updateOrderRequestStatus') return respond(updateOrderRequestStatus(e.parameter, e.parameter.status), e);
    requirePermission(authPayload, 'view');
    if (action === 'transact') {
      requirePermission(authPayload, 'transact');
      var txnPayload = parseTransactionPayloadFromGet(e);
      var txnType = String(txnPayload.type || '');
      if (txnType.indexOf('Input') > -1) requirePermission(authPayload, 'receive_part');
      if (txnType.indexOf('Output') > -1) requirePermission(authPayload, 'issue_part');
      return respond(processTransaction(txnPayload), e);
    }
    if (action === 'logs') return respond(getLogRows(), e);
    if (action === 'nextNo') {
      requirePermission(authPayload, 'manage_items');
      return respond(getNextNoBySheet(e.parameter.sheet), e);
    }
    if (action === 'authStatus') return respond(getDriveAuthStatus(), e);
    if (action === 'uploadDrawing' || action === 'uploadAttachment') {
      requirePermission(authPayload, 'manage_items');
      e.parameter.kind = e.parameter.kind || 'drawing';
      return respond(uploadPartAttachmentToDrive(e.parameter), e);
    }
    if (action === 'authorizeDrive') return respond(authorizeGoogleDriveAccess(), e);
    if (action === 'upsertItem') {
      requirePermission(authPayload, 'manage_items');
      ensureLocationColumnsForAllKnownSheets();
      ensurePriceColumnsForAllKnownSheets();
      ensureAttachmentColumnsForAllKnownSheets();
      return respond(upsertMainItem({
      sheetName: e.parameter.sheet,
      no: e.parameter.no,
      name: e.parameter.name,
      model: e.parameter.model,
      line: e.parameter.line,
      location: e.parameter.location,
      category: e.parameter.category,
      brand: e.parameter.brand,
      photo: e.parameter.photo,
      image_main: e.parameter.image_main,
      image_install: e.parameter.image_install,
      image_main_url: e.parameter.image_main_url,
      image_main_file_id: e.parameter.image_main_file_id,
      image_install_url: e.parameter.image_install_url,
      image_install_file_id: e.parameter.image_install_file_id,
      max: e.parameter.max,
      min: e.parameter.min,
      unit: e.parameter.unit,
      stock: e.parameter.stock,
      unit_price: e.parameter.unit_price,
      currency: e.parameter.currency,
      supplier: e.parameter.supplier,
      price_updated_at: e.parameter.price_updated_at,
      price_remark: e.parameter.price_remark,
      drawing_url: e.parameter.drawing_url,
      drawing_file_name: e.parameter.drawing_file_name,
      drawing_revision: e.parameter.drawing_revision,
      drawing_status: e.parameter.drawing_status,
      datasheet_url: e.parameter.datasheet_url,
      quotation_url: e.parameter.quotation_url
    }), e);
    }
    if (action === 'deleteItem') {
      requirePermission(authPayload, 'delete_items');
      return respond(deleteMainItem({ sheetName: e.parameter.sheet, no: e.parameter.no }), e);
    }

    var sheetName = resolveReadSheetName({ sheet: e.parameter.sheet });
    ensureLocationColumnsForAllKnownSheets();
    ensurePriceColumnsForAllKnownSheets();
    ensureAttachmentColumnsForAllKnownSheets();
    var ctx = getMainSheetContext(sheetName);
    ensureColumnInContext(ctx, 'Location', ['location', 'jrlocation']);
    var map = ctx.map;
    var rows = ctx.rows;
    if (!rows.length) return respond([], e);

    var result = rows.map(function (row, index) {
      var stockValue = Number(pickRowValue(row, map, ['stockqty', 'qtystock', 'qoh', 'stock'], 0)) || 0;
      var minValue = Number(pickRowValue(row, map, ['min', 'qtymin'], 0)) || 0;
      var needToPOValue = Math.max(minValue - stockValue, 0);

      var noText = String(pickRowValue(row, map, ['no'], index + 1));
      var rawLocation = pickRowValue(row, map, ['location', 'jrlocation'], '-');
      var locationOverride = getLocationOverride(sheetName, noText);
      var photoValue = pickRowValue(row, map, ['sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], '');
      var mainFileIdValue = pickRowValue(row, map, ['image_main_file_id', 'imagemainfileid'], '');
      var installFileIdValue = pickRowValue(row, map, ['image_install_file_id', 'imageinstallfileid'], '');
      var mainImageValue = pickRowValue(row, map, ['image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image', 'sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], '') || buildDriveViewUrlFromFileId(mainFileIdValue);
      var installImageValue = pickRowValue(row, map, ['image_install_url', 'imageinstallurl', 'image_install', 'imageinstall', 'installimage', 'install_image'], '') || buildDriveViewUrlFromFileId(installFileIdValue);
      if (!photoValue && mainImageValue) photoValue = mainImageValue;
      return {
        no: noText,
        name: pickRowValue(row, map, ['namedescriptions', 'name', 'description', 'partname', 'jrpartname', 'jrpartnameolderp'], '-'),
        model: pickRowValue(row, map, ['model', 'codeno', 'jrcodeno'], '-'),
        line: pickRowValue(row, map, ['mainline', 'line', 'process', 'ไลน์'], '-'),
        location: locationOverride || rawLocation,
        category: pickRowValue(row, map, ['category'], 'General'),
        brand: pickRowValue(row, map, ['brand'], '-'),
        stock: stockValue,
        max: pickRowValue(row, map, ['max', 'qtymax'], 0),
        min: minValue,
        needToPO: needToPOValue,
        unit: pickRowValue(row, map, ['unit'], 'PCS'),
        remark: pickRowValue(row, map, ['remark'], ''),
        photo: photoValue,
        image_main: mainImageValue,
        image_install: installImageValue,
        image_main_url: mainImageValue,
        image_main_file_id: mainFileIdValue,
        image_install_url: installImageValue,
        image_install_file_id: installFileIdValue,
        drawing_url: pickRowValue(row, map, ['drawingurl', 'drawing_url'], ''),
        drawing_file_name: pickRowValue(row, map, ['drawingfilename', 'drawing_file_name'], ''),
        drawing_revision: pickRowValue(row, map, ['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev'], ''),
        drawing_status: pickRowValue(row, map, ['drawingstatus', 'drawing_status'], ''),
        datasheet_url: pickRowValue(row, map, ['datasheeturl', 'datasheet_url'], ''),
        quotation_url: pickRowValue(row, map, ['quotationurl', 'quotation_url'], ''),
        unit_price: pickRowValue(row, map, ['unitprice', 'unit_price'], ''),
        currency: pickRowValue(row, map, ['currency'], 'THB'),
        supplier: pickRowValue(row, map, ['supplier'], ''),
        price_updated_at: pickRowValue(row, map, ['priceupdatedat', 'price_updated_at'], ''),
        price_remark: pickRowValue(row, map, ['priceremark', 'price_remark'], ''),
        coil_size: pickRowValue(row, map, ['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size'], '-')
      };
    }).filter(function (item) {
      return item.name && item.name !== '-';
    });

    return respond(result, e);
  } catch (err) {
    return respond(buildErrorResponse(err), e);
  }
}

// =============================
// POST (transaction)
// =============================
function doPost(e) {
  try {
    function parseMultipartFields(rawText, contentType) {
      var out = {};
      var m = String(contentType || '').match(/boundary=([^;]+)/i);
      if (!m || !m[1]) return out;
      var boundary = '--' + m[1];
      var parts = String(rawText || '').split(boundary);
      for (var i = 0; i < parts.length; i += 1) {
        var part = parts[i];
        if (!part || part === '--' || part === '--\r\n') continue;
        var nameMatch = part.match(/name=\"([^\"]+)\"/i);
        if (!nameMatch || !nameMatch[1]) continue;
        var key = nameMatch[1];
        var splitIndex = part.indexOf('\r\n\r\n');
        if (splitIndex < 0) continue;
        var value = part.substring(splitIndex + 4).replace(/\r\n--$/, '').replace(/\r\n$/, '');
        if (!/filename=\"/i.test(part)) out[key] = value;
      }
      return out;
    }

    var body = {};
    var raw = e && e.postData ? String(e.postData.contents || '') : '';
    try {
      body = raw ? JSON.parse(raw) : {};
    } catch (jsonErr) {
      body = e && e.parameter ? e.parameter : {};
      if (!body || !Object.keys(body).length) {
        body = parseMultipartFields(raw, e && e.postData ? e.postData.type : '');
      }
      body.dataUrl = body.dataUrl || body.file || body.fileBase64 || '';
    }
    var action = body && body.action ? String(body.action) : '';
    var authPayload = { authToken: body.authToken || body.token || '' };
    if (!action && (body.itemId || body.imageType || body.kind || body.dataUrl)) action = 'uploadImage';
    if (action === 'login') {
      return respond(loginUser({ username: body.username, password: body.password }), e);
    }
    if (action === 'logout') {
      return respond(logoutUser(authPayload), e);
    }
    if (action === 'session') {
      return respond(getSessionUser(authPayload), e);
    }
    if (action === 'listUsers') {
      return respond(listUsers(authPayload), e);
    }
    if (action === 'upsertUser') {
      return respond(upsertUser({
        authToken: authPayload.authToken,
        username: body.username,
        password: body.password,
        role: body.role,
        isActive: body.isActive,
        permissionsJson: body.permissionsJson
      }), e);
    }
    if (action === 'deleteUser') {
      return respond(deleteUser({ authToken: authPayload.authToken, username: body.username }), e);
    }
    if (action === 'createOrderRequest') return respond(createOrderRequest(body), e);
    if (action === 'uploadRequestAttachment') return respond(uploadRequestAttachment(body), e);
    if (action === 'getOrderRequests') return respond(getOrderRequests(body), e);
    if (action === 'ensureOrderRequestsSheet') return respond(ensureOrderRequestsSheetReady(body), e);
    if (action === 'approveOrderRequest') return respond(approveOrderRequest(body), e);
    if (action === 'rejectOrderRequest') return respond(rejectOrderRequest(body), e);
    if (action === 'holdOrderRequest') return respond(holdOrderRequest(body), e);
    if (action === 'closeOrderRequest') return respond(closeOrderRequest(body), e);
    if (action === 'markOrderRequestPurchased') return respond(markOrderRequestPurchased(body), e);
    if (action === 'markOrderRequestReceived') return respond(markOrderRequestReceived(body), e);
    if (action === 'convertOrderRequestsToPR') return respond(convertOrderRequestsToPR(body), e);
    if (action === 'updateOrderRequestStatus') return respond(updateOrderRequestStatus(body, body.status), e);
    requirePermission(authPayload, 'view');
    if (action === 'upsertItem') {
      requirePermission(authPayload, 'manage_items');
      ensureLocationColumnsForAllKnownSheets();
      ensurePriceColumnsForAllKnownSheets();
      ensureAttachmentColumnsForAllKnownSheets();
      return respond(upsertMainItem({
        sheetName: body.sheet || body.sheetName,
        no: body.no,
        name: body.name,
        model: body.model,
        line: body.line,
        location: body.location,
        category: body.category,
        brand: body.brand,
        photo: body.photo,
        image_main: body.image_main,
        image_install: body.image_install,
        image_main_url: body.image_main_url,
        image_main_file_id: body.image_main_file_id,
        image_install_url: body.image_install_url,
        image_install_file_id: body.image_install_file_id,
        max: body.max,
        min: body.min,
        unit: body.unit,
        stock: body.stock,
        unit_price: body.unit_price,
        currency: body.currency,
        supplier: body.supplier,
        price_updated_at: body.price_updated_at,
        price_remark: body.price_remark,
        coil_size: body.coil_size || body.machine_model,
        drawing_url: body.drawing_url,
        drawing_file_name: body.drawing_file_name,
        drawing_revision: body.drawing_revision,
        drawing_status: body.drawing_status,
        datasheet_url: body.datasheet_url,
        quotation_url: body.quotation_url
      }), e);
    }
    if (action === 'uploadDrawing' || action === 'uploadAttachment') {
      requirePermission(authPayload, 'manage_items');
      body.kind = body.kind || 'drawing';
      return respond(uploadPartAttachmentToDrive(body), e);
    }
    if (action === 'uploadImage' || action === 'upload') {
      requirePermission(authPayload, 'manage_items');
      return respond(uploadImageToDrive(body), e);
    }
    if (action === 'authStatus') {
      return respond(getDriveAuthStatus(), e);
    }
    if (action === 'authorizeDrive') {
      return respond(authorizeGoogleDriveAccess(), e);
    }
    if (action === 'deleteItem') {
      requirePermission(authPayload, 'delete_items');
      return respond(deleteMainItem({
        sheetName: body.sheet || body.sheetName,
        no: body.no
      }), e);
    }
    requirePermission(authPayload, 'transact');
    var postTxnType = String(body.type || '');
    if (postTxnType.indexOf('Input') > -1) requirePermission(authPayload, 'receive_part');
    if (postTxnType.indexOf('Output') > -1) requirePermission(authPayload, 'issue_part');
    return respond(processTransaction(body), e);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return respond(buildErrorResponse(err), e);
  }
}

// =============================
// RESPONSE HELPERS
// =============================
function respond(data, e) {
  var callback = e && e.parameter ? e.parameter.callback : null;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
