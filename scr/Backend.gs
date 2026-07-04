// =============================
// CONFIG
// =============================
var SPARE_APP_CONFIG = this.SPARE_APP_CONFIG || {};
SPARE_APP_CONFIG.readSheetName = SPARE_APP_CONFIG.readSheetName || 'Main List Stock';
SPARE_APP_CONFIG.writeSheetName = SPARE_APP_CONFIG.writeSheetName || 'Log';
SPARE_APP_CONFIG.usersSheetName = SPARE_APP_CONFIG.usersSheetName || 'Users';
SPARE_APP_CONFIG.requestSheetName = SPARE_APP_CONFIG.requestSheetName || 'OrderRequests';
SPARE_APP_CONFIG.purchaseHistorySheetName = SPARE_APP_CONFIG.purchaseHistorySheetName || 'PurchaseHistory';
SPARE_APP_CONFIG.purchaseHistoryAuditSheetName = SPARE_APP_CONFIG.purchaseHistoryAuditSheetName || 'PurchaseHistoryLog';
SPARE_APP_CONFIG.purchaseHistoryImportLogSheetName = SPARE_APP_CONFIG.purchaseHistoryImportLogSheetName || 'PurchaseHistoryImportLog';
SPARE_APP_CONFIG.productionVolumeSheetName = SPARE_APP_CONFIG.productionVolumeSheetName || 'ProductionVolume';
SPARE_APP_CONFIG.productionCostConfigSheetName = SPARE_APP_CONFIG.productionCostConfigSheetName || 'ProductionCostConfig';
// ชีต "ProductionLog" ของแต่ละไลน์ที่มีระบบติดตามยอดผลิตแยกต่างหาก (Google Sheet คนละไฟล์กับ
// สเปรดชีตอะไหล่) — ถ้าตั้งค่าไว้ ระบบจะดึงยอดผลิตจริงมาคำนวณอัตโนมัติแทนการกรอกมือ
// ต้องให้บัญชีที่รัน Apps Script นี้มีสิทธิ์ View/Edit ไฟล์ปลายทางด้วย ไม่งั้นจะ fallback ไปใช้
// ค่าที่กรอกมือแทนเงียบๆ
SPARE_APP_CONFIG.productionLogSources = SPARE_APP_CONFIG.productionLogSources || {
  'Lug&Screw': '1Xx2XEGtT-KbnvVP_9gzkW9kuyFUBpj1H-oIsT3zUx1U'
};
SPARE_APP_CONFIG.sessionDurationMs = SPARE_APP_CONFIG.sessionDurationMs || (7 * 24 * 60 * 60 * 1000);
SPARE_APP_CONFIG.sessionRefreshThresholdMs = SPARE_APP_CONFIG.sessionRefreshThresholdMs || (24 * 60 * 60 * 1000);
var SESSION_PROPERTY_PREFIX = 'spare_session::';
var LOG_HEADERS = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After', 'Reason', 'Reason Remark'];
var USER_HEADERS = ['username', 'password', 'role', 'is_active', 'permissions_json', 'session_token', 'session_expiry'];
var ORDER_REQUEST_HEADERS = ['request_id', 'requested_date', 'requested_by', 'requester_role', 'item_id', 'item_name', 'model', 'brand', 'category', 'line', 'current_stock', 'min', 'max', 'request_qty', 'priority', 'reason', 'expected_use_date', 'remark', 'attachment_url', 'status', 'admin_comment', 'approved_by', 'approved_date', 'converted_pr_id', 'updated_at', 'unit', 'unit_price', 'currency'];
var ORDER_REQUEST_STATUSES = ['Pending', 'Approved', 'Rejected', 'On Hold', 'Converted to PR', 'Purchased', 'Received', 'Closed'];
var PURCHASE_HISTORY_HEADERS = ['History ID', 'Request ID', 'Source', 'Requested Date', 'Month', 'Line', 'Part ID', 'Part Name', 'Brand', 'Model / Part No.', 'Qty Ordered', 'Unit', 'Unit Price', 'Currency', 'Total Amount', 'Requested By', 'Status', 'Ordered Date', 'Received Date', 'Received Qty', 'Updated By', 'Remark', 'Deleted', 'Created At', 'Updated At', 'Import Batch ID', 'Source File Name', 'Source File Hash', 'Price Status', 'Created By', 'Request Period'];
var PURCHASE_HISTORY_AUDIT_HEADERS = ['Date Time', 'User', 'History ID', 'Action Type', 'Old Value', 'New Value', 'Reason'];
var PURCHASE_HISTORY_IMPORT_LOG_HEADERS = ['Import Batch ID', 'File Name', 'Imported By', 'Imported At', 'Total Rows Detected', 'Imported Rows', 'Skipped Duplicate Rows', 'Review Required Rows'];
var PURCHASE_HISTORY_STATUSES = ['Requested', 'PR Created', 'Ordered', 'Partial Received', 'Received', 'Cancelled'];
var PURCHASE_HISTORY_SOURCES = ['Purchase Request', 'PR Report', 'Manual', 'Auto PR'];
// ยอดผลิตรายเดือนต่อไลน์ (รวมทุกเครื่อง) — ใช้คู่กับ ProductionCostConfig เพื่อคำนวณ
// "รายจ่ายอะไหล่ / มูลค่าผลิต" เป็น % สำหรับควบคุมต้นทุนการสั่งซื้อรายเดือน
var PRODUCTION_VOLUME_HEADERS = ['Month', 'Line', 'Actual Qty', 'Updated By', 'Updated At'];
// ราคาต่อหน่วย (บาท/ชิ้น) และเป้าหมาย % ที่ต้องการควบคุมรายจ่ายอะไหล่ให้อยู่ในกรอบ ต่อไลน์
var PRODUCTION_COST_CONFIG_HEADERS = ['Line', 'Unit Price', 'Target Pct', 'Updated By', 'Updated At'];
var STOCK_LOCATION_SHEETS = ['Main List Stock', 'Stock for MC', 'Standard Spare part', 'Arc chut', 'Common Gv.2', 'Gv.2 (6 plate)', 'Gv.2 (9 plate)', 'Coil Winding', 'Lug&Screw'];
var DRIVE_ROOT_FOLDER_ID = '1XWO5rGpku35gSTMAh4HDOCHa6GJIkoS3';
var DRAWING_STATUS_OPTIONS = ['Available', 'Missing', 'Not Required', 'Access Required'];
var PART_ATTACHMENT_COLUMNS = [
  // Photo is already stored in existing image_main/image_main_url columns.
  // Do not create another Photo URL column because it duplicates current image columns.
  { label: 'Drawing URL', aliases: ['drawingurl', 'drawing_url'] },
  { label: 'Drawing File Name', aliases: ['drawingfilename', 'drawing_file_name'] },
  { label: 'Drawing Revision', aliases: ['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev'] },
  { label: 'Drawing Status', aliases: ['drawingstatus', 'drawing_status'] },
  { label: 'Datasheet URL', aliases: ['datasheeturl', 'datasheet_url'] },
  { label: 'Quotation URL', aliases: ['quotationurl', 'quotation_url'] }
];

function normalizeDrawingStatusValue(value) {
  var raw = String(value || '').trim();
  var normalized = raw.toLowerCase();
  if (normalized === 'available') return 'Available';
  if (normalized === 'missing' || normalized === 'not available') return 'Missing';
  if (normalized === 'not required' || normalized === 'n/a' || normalized === 'na') return 'Not Required';
  if (normalized === 'access required' || normalized === 'pending update' || normalized === 'pending') return 'Access Required';
  return raw;
}

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

  return {
    status: 'error',
    errorCode: err && err.code ? String(err.code) : '',
    message: msg
  };
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

function getFirstColumnIndexByAliases(map, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    if (map[aliases[i]] !== undefined) return map[aliases[i]];
  }
  return undefined;
}

function cleanupRedundantPhotoUrlColumn(ctx) {
  if (!ctx || !ctx.headers || !ctx.headers.length) return;
  var targetCol = getFirstColumnIndexByAliases(ctx.map, [
    'image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image', 'sparepartsphotos', 'photo', 'imageurl', 'image'
  ]);
  if (targetCol === undefined) return;

  var redundantCols = [];
  for (var i = 0; i < ctx.headers.length; i += 1) {
    var normalized = normalizeHeaderName(ctx.headers[i]);
    if (normalized === 'photourl' && i !== targetCol) redundantCols.push(i);
  }
  if (!redundantCols.length) return;

  for (var r = 0; r < ctx.rows.length; r += 1) {
    var row = ctx.rows[r];
    for (var c = 0; c < redundantCols.length; c += 1) {
      var photoValue = row[redundantCols[c]];
      if (photoValue !== '' && photoValue !== null && photoValue !== undefined && (row[targetCol] === '' || row[targetCol] === null || row[targetCol] === undefined)) {
        ctx.sheet.getRange(ctx.headerRowIndex + 2 + r, targetCol + 1).setValue(photoValue);
        row[targetCol] = photoValue;
      }
    }
  }

  redundantCols.sort(function(a, b) { return b - a; }).forEach(function(colIndex) {
    ctx.sheet.deleteColumn(colIndex + 1);
  });
}

function ensureAttachmentColumnsForSheet(sheetName) {
  var targetSheet = String(sheetName || '').trim();
  if (!targetSheet) return;
  var ctx = getMainSheetContext(targetSheet);
  PART_ATTACHMENT_COLUMNS.forEach(function(col) {
    ensureColumnInContext(ctx, col.label, col.aliases);
  });
  cleanupRedundantPhotoUrlColumn(ctx);
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

function getPurchaseHistoryHeaderMap(headers) {
  var map = {};
  (headers || []).forEach(function(header, index) { map[normalizeHeaderName(header)] = index; });
  return map;
}

function getPurchaseHistoryCell(row, map, aliases, fallback) {
  for (var i = 0; i < aliases.length; i += 1) {
    var index = map[normalizeHeaderName(aliases[i])];
    if (index !== undefined) return row[index];
  }
  return fallback;
}

function formatPurchaseHistoryDate(value, includeTime) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value || '');
  return Utilities.formatDate(date, 'Asia/Bangkok', includeTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
}

// Google Sheets บางครั้ง auto-convert ค่า "yyyy-MM" ในคอลัมน์ Month ให้กลายเป็น Date object
// จริง (ตีความเป็นวันที่ 1 ของเดือนนั้น) ทำให้ String(row[4]) คืนค่าเป็น toString() แบบเต็ม
// ("Wed Jul 01 2026 00:00:00 GMT+0700 ...") แทนที่จะเป็น "2026-07" ต้อง format ให้ถูก timezone เสมอ
function formatPurchaseHistoryMonth(value) {
  if (!value) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM');
  }
  return String(value).trim();
}

function calculatePurchaseHistoryTotal(qty, unitPrice) {
  var priceText = String(unitPrice === undefined || unitPrice === null ? '' : unitPrice).trim();
  if (!priceText) return '';
  var price = Number(priceText);
  if (!isFinite(price) || price < 0) return '';
  return Number(qty || 0) * price;
}

function migratePurchaseHistoryRows(rows, oldHeaders) {
  var oldMap = getPurchaseHistoryHeaderMap(oldHeaders);
  return rows.filter(function(row) {
    return String(getPurchaseHistoryCell(row, oldMap, ['History ID', 'history_id'], '') || '').trim();
  }).map(function(row) {
    var date = getPurchaseHistoryCell(row, oldMap, ['Date', 'Requested Date'], '');
    var requestId = getPurchaseHistoryCell(row, oldMap, ['Request ID', 'request_id'], '');
    var historyId = getPurchaseHistoryCell(row, oldMap, ['History ID', 'history_id'], '') || buildPurchaseHistoryId('PH');
    if (!requestId && String(historyId).indexOf('PH-REQ-') === 0) requestId = String(historyId).substring(3);
    var qty = Number(getPurchaseHistoryCell(row, oldMap, ['Qty Ordered', 'qty_ordered'], 0) || 0);
    var unitPrice = getPurchaseHistoryCell(row, oldMap, ['Unit Price', 'unit_price'], '');
    var source = String(getPurchaseHistoryCell(row, oldMap, ['Source', 'Request Type'], 'Manual') || 'Manual');
    if (source === 'PR' || source === 'PO') source = 'PR Report';
    if (PURCHASE_HISTORY_SOURCES.indexOf(source) === -1) source = 'Manual';
    return [
      historyId, requestId, source,
      formatPurchaseHistoryDate(date, true),
      getPurchaseHistoryCell(row, oldMap, ['Month'], date ? formatPurchaseHistoryDate(date, false).slice(0, 7) : ''),
      getPurchaseHistoryCell(row, oldMap, ['Line'], ''), getPurchaseHistoryCell(row, oldMap, ['Part ID', 'part_id'], ''),
      getPurchaseHistoryCell(row, oldMap, ['Part Name', 'part_name'], ''), getPurchaseHistoryCell(row, oldMap, ['Brand'], ''),
      getPurchaseHistoryCell(row, oldMap, ['Model / Part No.', 'Model', 'Part No'], ''), qty,
      getPurchaseHistoryCell(row, oldMap, ['Unit'], ''), unitPrice,
      getPurchaseHistoryCell(row, oldMap, ['Currency'], unitPrice === '' ? '' : 'THB'), calculatePurchaseHistoryTotal(qty, unitPrice),
      getPurchaseHistoryCell(row, oldMap, ['Requested By', 'requested_by'], ''), getPurchaseHistoryCell(row, oldMap, ['Status'], 'Requested') || 'Requested',
      formatPurchaseHistoryDate(getPurchaseHistoryCell(row, oldMap, ['Ordered Date'], ''), false),
      formatPurchaseHistoryDate(getPurchaseHistoryCell(row, oldMap, ['Received Date'], ''), false),
      Number(getPurchaseHistoryCell(row, oldMap, ['Received Qty'], 0) || 0), getPurchaseHistoryCell(row, oldMap, ['Updated By'], ''),
      getPurchaseHistoryCell(row, oldMap, ['Remark'], ''), toBoolean(getPurchaseHistoryCell(row, oldMap, ['Deleted'], false), false),
      formatPurchaseHistoryDate(getPurchaseHistoryCell(row, oldMap, ['Created At'], ''), true),
      formatPurchaseHistoryDate(getPurchaseHistoryCell(row, oldMap, ['Updated At'], ''), true),
      getPurchaseHistoryCell(row, oldMap, ['Import Batch ID'], ''), getPurchaseHistoryCell(row, oldMap, ['Source File Name'], ''),
      getPurchaseHistoryCell(row, oldMap, ['Source File Hash'], ''), getPurchaseHistoryCell(row, oldMap, ['Price Status'], unitPrice === '' ? 'TBC' : 'Confirmed'),
      getPurchaseHistoryCell(row, oldMap, ['Created By'], ''), getPurchaseHistoryCell(row, oldMap, ['Request Period'], '')
    ];
  });
}

function getPurchaseHistorySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.purchaseHistorySheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(PURCHASE_HISTORY_HEADERS);
    return sheet;
  }
  var width = Math.max(sheet.getLastColumn(), PURCHASE_HISTORY_HEADERS.length);
  var values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
  var currentHeaders = values[0] || [];
  var headersMatch = PURCHASE_HISTORY_HEADERS.every(function(header, index) { return String(currentHeaders[index] || '') === header; });
  if (headersMatch) return sheet;
  var migratedRows = migratePurchaseHistoryRows(values.slice(1), currentHeaders);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, PURCHASE_HISTORY_HEADERS.length).setValues([PURCHASE_HISTORY_HEADERS]);
  if (migratedRows.length) sheet.getRange(2, 1, migratedRows.length, PURCHASE_HISTORY_HEADERS.length).setValues(migratedRows);
  return sheet;
}

function getPurchaseHistoryAuditSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.purchaseHistoryAuditSheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(PURCHASE_HISTORY_AUDIT_HEADERS);
  return sheet;
}

function appendPurchaseHistoryAudit(user, historyId, actionType, oldValue, newValue, reason) {
  getPurchaseHistoryAuditSheet().appendRow([
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), user || '', historyId || '', actionType || '',
    JSON.stringify(oldValue || {}), JSON.stringify(newValue || {}), reason || ''
  ]);
}

function buildPurchaseHistoryId(prefix) { return String(prefix || 'PH') + '-' + Utilities.getUuid(); }
function normalizePurchaseHistoryKeyPart(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g, ' '); }
// เทียบรุ่น/Part Number โดยตัดอักขระที่ไม่ใช่ตัวอักษร-ตัวเลขออก (กัน "-", ช่องว่าง, จุด ที่เพี้ยนจากการ import PDF)
function normalizePurchaseHistoryModel(value) { return String(value || '').toLowerCase().replace(/[^a-z0-9฀-๿]/g, ''); }
function isMeaningfulPurchaseHistoryModel(value) { var m = normalizePurchaseHistoryModel(value); return m.length >= 3 && m !== 'na'; }
// เทียบชื่อแบบไม่สนช่องว่าง (กันชื่อไทยที่ถูกแยกด้วยช่องว่างตอน import PDF)
function normalizePurchaseHistoryName(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g, ''); }
// ยี่ห้อที่ "มีความหมาย" (ไม่ใช่ค่าว่าง/Unknown) ใช้เป็นตัวกันชื่อซ้ำต่างยี่ห้อ
function isMeaningfulPurchaseHistoryBrand(value) { var b = normalizePurchaseHistoryName(value); return !!b && b !== '-' && b !== 'na' && b !== 'unknown' && b !== 'unknownbrand' && b !== 'nobrand'; }

function purchaseHistoryRowsMatch(row, payload) {
  var requestId = normalizePurchaseHistoryKeyPart(payload.request_id || payload.requestId);
  var partId = normalizePurchaseHistoryKeyPart(payload.part_id || payload.partId);
  var line = normalizePurchaseHistoryKeyPart(payload.line);
  var model = normalizePurchaseHistoryKeyPart(payload.model || payload.part_no || payload.partNo);
  if (!requestId || normalizePurchaseHistoryKeyPart(row[1]) !== requestId) return false;
  if (partId) return normalizePurchaseHistoryKeyPart(row[6]) === partId;
  if (line && model) return normalizePurchaseHistoryKeyPart(row[5]) === line && normalizePurchaseHistoryKeyPart(row[9]) === model;
  return true;
}

function findOpenPurchaseHistoryRow(values, payload) {
  var partId = normalizePurchaseHistoryKeyPart(payload.part_id || payload.partId);
  var line = normalizePurchaseHistoryKeyPart(payload.line);
  var model = normalizePurchaseHistoryKeyPart(payload.model || payload.part_no || payload.partNo);
  var openStatuses = { 'Requested': true, 'PR Created': true, 'Ordered': true, 'Partial Received': true };
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    if (toBoolean(row[22], false) || !openStatuses[String(row[16] || '')]) continue;
    if (partId && normalizePurchaseHistoryKeyPart(row[6]) === partId) return i;
    if (line && model && normalizePurchaseHistoryKeyPart(row[5]) === line && normalizePurchaseHistoryKeyPart(row[9]) === model) return i;
  }
  return -1;
}

function upsertPurchaseHistoryRecord(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return upsertPurchaseHistoryRecordUnlocked(payload); } finally { lock.releaseLock(); }
}

function upsertPurchaseHistoryRecordUnlocked(payload) {
  payload = payload || {};
  var status = String(payload.status || 'Requested').trim();
  if (PURCHASE_HISTORY_STATUSES.indexOf(status) === -1) throw new Error('Purchase History status ไม่ถูกต้อง: ' + status);
  var source = String(payload.source || '').trim();
  if (source && PURCHASE_HISTORY_SOURCES.indexOf(source) === -1) throw new Error('Purchase History source ไม่ถูกต้อง: ' + source);
  var qtyProvided = payload.qty_ordered !== undefined || payload.qtyOrdered !== undefined;
  var qty = Number(payload.qty_ordered !== undefined ? payload.qty_ordered : payload.qtyOrdered);
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < values.length; i += 1) if (purchaseHistoryRowsMatch(values[i], payload)) { rowIndex = i; break; }
  if (rowIndex === -1 && payload.match_open_item) rowIndex = findOpenPurchaseHistoryRow(values, payload);
  if (rowIndex === -1 && (!qtyProvided || !isFinite(qty) || qty <= 0)) return { skipped: true, reason: 'QTY_NOT_POSITIVE' };

  var now = new Date();
  var date = payload.date ? new Date(payload.date) : now;
  if (isNaN(date.getTime())) date = now;
  var timestamp = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var requestedAt = Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var orderedDate = payload.ordered_date || (status === 'Ordered' ? Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd') : '');
  var receivedDate = payload.received_date || ((status === 'Received' || status === 'Partial Received') ? Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd') : '');

  if (rowIndex !== -1) {
    var existing = values[rowIndex];
    var statusRank = { 'Requested': 1, 'PR Created': 2, 'Ordered': 3, 'Partial Received': 4, 'Received': 5, 'Cancelled': 6 };
    var effectiveStatus = status;
    if (!payload.force_status && status !== 'Cancelled' && Number(statusRank[String(existing[16] || 'Requested')] || 0) > Number(statusRank[status] || 0)) effectiveStatus = existing[16];
    var effectiveQty = qtyProvided && isFinite(qty) && qty > 0 ? qty : Number(existing[10] || 0);
    var effectivePrice = payload.unit_price !== undefined ? payload.unit_price : existing[12];
    var merged = [
      existing[0] || payload.history_id || buildPurchaseHistoryId('PH'),
      (payload.match_open_item && existing[1]) ? existing[1] : (payload.request_id || payload.requestId || existing[1] || ''),
      payload.source || existing[2] || 'Manual', existing[3] || requestedAt, existing[4] || requestedAt.slice(0, 7),
      payload.line !== undefined ? payload.line : existing[5], payload.part_id !== undefined ? payload.part_id : (payload.partId !== undefined ? payload.partId : existing[6]),
      payload.part_name !== undefined ? payload.part_name : (payload.partName !== undefined ? payload.partName : existing[7]),
      payload.brand !== undefined ? payload.brand : existing[8], payload.model !== undefined ? payload.model : existing[9], effectiveQty,
      payload.unit !== undefined ? payload.unit : existing[11], effectivePrice,
      payload.currency !== undefined ? payload.currency : existing[13], calculatePurchaseHistoryTotal(effectiveQty, effectivePrice),
      payload.requested_by !== undefined ? payload.requested_by : (payload.requestedBy !== undefined ? payload.requestedBy : existing[15]),
      effectiveStatus, orderedDate || existing[17] || '', receivedDate || existing[18] || '',
      payload.received_qty !== undefined ? Number(payload.received_qty || 0) : Number(existing[19] || 0),
      payload.updated_by || payload.updatedBy || existing[20] || '', payload.remark !== undefined ? payload.remark : existing[21],
      payload.deleted !== undefined ? toBoolean(payload.deleted, false) : toBoolean(existing[22], false), existing[23] || timestamp, timestamp,
      payload.import_batch_id !== undefined ? payload.import_batch_id : existing[25], payload.source_file_name !== undefined ? payload.source_file_name : existing[26],
      payload.source_file_hash !== undefined ? payload.source_file_hash : existing[27], payload.price_status !== undefined ? payload.price_status : existing[28],
      existing[29] || payload.created_by || payload.requested_by || payload.requestedBy || '', payload.request_period !== undefined ? payload.request_period : existing[30]
    ];
    sheet.getRange(rowIndex + 1, 1, 1, PURCHASE_HISTORY_HEADERS.length).setValues([merged]);
    return { history_id: merged[0], mode: 'update', row: merged };
  }

  var unitPrice = payload.unit_price === undefined ? '' : payload.unit_price;
  var row = [
    String(payload.history_id || payload.historyId || buildPurchaseHistoryId('PH')).trim(), payload.request_id || payload.requestId || '',
    payload.source || 'Manual', requestedAt, requestedAt.slice(0, 7), payload.line || '', payload.part_id || payload.partId || '',
    payload.part_name || payload.partName || '', payload.brand || '', payload.model || payload.part_no || payload.partNo || '', qty,
    payload.unit || '', unitPrice, payload.currency || (String(unitPrice).trim() ? 'THB' : ''), calculatePurchaseHistoryTotal(qty, unitPrice),
    payload.requested_by || payload.requestedBy || '', status, orderedDate, receivedDate,
    Number(payload.received_qty || payload.receivedQty || 0), payload.updated_by || payload.updatedBy || '', payload.remark || '', false, timestamp, timestamp,
    payload.import_batch_id || '', payload.source_file_name || '', payload.source_file_hash || '',
    payload.price_status || (String(unitPrice).trim() ? 'Confirmed' : 'TBC'), payload.created_by || payload.requested_by || payload.requestedBy || '', payload.request_period || ''
  ];
  sheet.appendRow(row);
  return { history_id: row[0], mode: 'insert', row: row };
}

function purchaseHistoryRowToObject(row) {
  return {
    history_id: String(row[0] || ''), request_id: String(row[1] || ''), source: String(row[2] || ''), date: formatPurchaseHistoryDate(row[3], true), month: formatPurchaseHistoryMonth(row[4]),
    line: String(row[5] || ''), part_id: String(row[6] || ''), part_name: String(row[7] || ''), brand: String(row[8] || ''), model: String(row[9] || ''),
    qty_ordered: Number(row[10] || 0), unit: String(row[11] || ''), unit_price: String(row[12] === undefined ? '' : row[12]), currency: String(row[13] || ''),
    total_amount: String(row[14] === undefined ? '' : row[14]), requested_by: String(row[15] || ''), status: String(row[16] || ''),
    ordered_date: formatPurchaseHistoryDate(row[17], false), received_date: formatPurchaseHistoryDate(row[18], false), received_qty: Number(row[19] || 0),
    updated_by: String(row[20] || ''), remark: String(row[21] || ''), deleted: toBoolean(row[22], false),
    created_at: formatPurchaseHistoryDate(row[23], true), updated_at: formatPurchaseHistoryDate(row[24], true),
    import_batch_id: String(row[25] || ''), source_file_name: String(row[26] || ''), source_file_hash: String(row[27] || ''),
    price_status: String(row[28] || (String(row[12] || '').trim() ? 'Confirmed' : 'TBC')), created_by: String(row[29] || ''), request_period: String(row[30] || '')
  };
}

function getPurchaseHistory(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var values = getPurchaseHistorySheet().getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(purchaseHistoryRowToObject).filter(function(item) { return item.history_id && !item.deleted && item.qty_ordered > 0; })
    .sort(function(a, b) { return String(b.updated_at).localeCompare(String(a.updated_at)); });
}

function requirePurchaseHistoryEditor(payload) {
  var session = getSessionUser(payload);
  var user = findUserByUsername(session.user.username);
  var role = normalizeRole(user && user.role);
  if (role !== 'admin' && role !== 'leader') throw new Error('เฉพาะ Admin / Engineer เท่านั้นที่แก้ไข Purchase History ได้');
  return user;
}

function editPurchaseHistory(payload) {
  var user = requirePurchaseHistoryEditor({ authToken: payload.authToken });
  var historyId = String(payload.history_id || '').trim();
  var reason = String(payload.reason || '').trim();
  if (!historyId) throw new Error('ไม่พบ History ID');
  if (!reason) throw new Error('กรุณาระบุเหตุผลการแก้ไข');
  var qty = Number(payload.qty_ordered);
  if (!isFinite(qty) || qty <= 0) throw new Error('Qty Ordered ต้องมากกว่า 0');
  var status = String(payload.status || '').trim();
  if (PURCHASE_HISTORY_STATUSES.indexOf(status) === -1) throw new Error('Status ไม่ถูกต้อง');
  var priceText = String(payload.unit_price === undefined ? '' : payload.unit_price).trim();
  if (priceText && (!isFinite(Number(priceText)) || Number(priceText) < 0)) throw new Error('Unit Price ไม่ถูกต้อง');
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') !== historyId || toBoolean(values[i][22], false)) continue;
    var oldObject = purchaseHistoryRowToObject(values[i]);
    var updated = upsertPurchaseHistoryRecord({
      request_id: values[i][1], part_id: values[i][6], qty_ordered: qty, unit_price: priceText,
      currency: priceText ? (payload.currency || values[i][13] || 'THB') : '', status: status, force_status: true,
      remark: payload.remark || '', updated_by: user.username
    });
    var newObject = purchaseHistoryRowToObject(updated.row);
    appendPurchaseHistoryAudit(user.username, historyId, 'EDIT', oldObject, newObject, reason);
    return { status: 'success', history: newObject };
  }
  throw new Error('ไม่พบ Purchase History');
}

function deletePurchaseHistory(payload) {
  var user = requirePurchaseHistoryEditor({ authToken: payload.authToken });
  var historyId = String(payload.history_id || '').trim();
  var reason = String(payload.reason || '').trim();
  if (!historyId) throw new Error('ไม่พบ History ID');
  if (!reason) throw new Error('กรุณาระบุเหตุผลการลบ');
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') !== historyId || toBoolean(values[i][22], false)) continue;
    var oldObject = purchaseHistoryRowToObject(values[i]);
    values[i][22] = true; values[i][20] = user.username; values[i][24] = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(i + 1, 1, 1, PURCHASE_HISTORY_HEADERS.length).setValues([values[i]]);
    var newObject = purchaseHistoryRowToObject(values[i]);
    appendPurchaseHistoryAudit(user.username, historyId, 'DELETE', oldObject, newObject, reason);
    return { status: 'success', history_id: historyId };
  }
  throw new Error('ไม่พบ Purchase History');
}

// ── Production Volume & Cost Ratio ─────────────────────────────
// ยอดผลิตรายเดือนต่อไลน์ (นำเข้า/กรอกมือจากชีต ProductionLog ภายนอก) x ราคา/หน่วย
// = มูลค่าผลิต ใช้เทียบกับรายจ่ายอะไหล่ (Purchase History) เป็น % สำหรับควบคุมต้นทุน
function getProductionVolumeSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.productionVolumeSheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(PRODUCTION_VOLUME_HEADERS);
  return sheet;
}

function getProductionCostConfigSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.productionCostConfigSheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(PRODUCTION_COST_CONFIG_HEADERS);
  return sheet;
}

function requireProductionCostEditor(payload) {
  var session = getSessionUser(payload);
  var user = findUserByUsername(session.user.username);
  var role = normalizeRole(user && user.role);
  if (role !== 'admin' && role !== 'leader') throw new Error('เฉพาะ Admin / Engineer เท่านั้นที่แก้ไขข้อมูลยอดผลิต/ราคาต่อหน่วยได้');
  return user;
}

// อ่านชีต "ProductionLog" ของสเปรดชีตภายนอก (คนละไฟล์กับสเปรดชีตอะไหล่) แล้วรวมยอด ActualQty
// เป็นรายเดือน (รวมทุกเครื่อง) — cache ไว้ 30 นาทีเพราะชีตต้นทางมีเป็นพันแถว อ่านทุกครั้งจะช้า
// คืนค่า null ถ้าไม่มีสิทธิ์เข้าถึง/ไม่พบชีต เพื่อให้ผู้เรียกใช้ fallback ไปใช้ค่าที่กรอกมือแทน
function getExternalProductionVolumeForLine(line) {
  var sourceId = (SPARE_APP_CONFIG.productionLogSources || {})[line];
  if (!sourceId) return null;
  var cacheKey = 'prod_volume_external::' + line;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (cacheReadErr) {
    Logger.log('getExternalProductionVolumeForLine cache read warning: ' + (cacheReadErr && cacheReadErr.message ? cacheReadErr.message : cacheReadErr));
  }
  try {
    var sourceSheet = SpreadsheetApp.openById(sourceId).getSheetByName('ProductionLog');
    if (!sourceSheet) return null;
    var values = sourceSheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    var headers = values[0];
    var dateCol = -1, qtyCol = -1;
    for (var i = 0; i < headers.length; i += 1) {
      var h = String(headers[i] || '').trim().toLowerCase();
      if (h === 'date') dateCol = i;
      if (h === 'actualqty') qtyCol = i;
    }
    if (dateCol === -1 || qtyCol === -1) return null;
    var byMonth = {};
    for (var r = 1; r < values.length; r += 1) {
      var dateVal = values[r][dateCol];
      var qty = Number(values[r][qtyCol] || 0);
      if (!dateVal || !isFinite(qty)) continue;
      var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
      if (isNaN(d.getTime())) continue;
      var monthKey = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
      byMonth[monthKey] = (byMonth[monthKey] || 0) + qty;
    }
    var result = Object.keys(byMonth).map(function(m) { return { month: m, actual_qty: byMonth[m] }; });
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 1800);
    } catch (cacheWriteErr) {
      Logger.log('getExternalProductionVolumeForLine cache write warning: ' + (cacheWriteErr && cacheWriteErr.message ? cacheWriteErr.message : cacheWriteErr));
    }
    return result;
  } catch (err) {
    Logger.log('getExternalProductionVolumeForLine error [' + line + ']: ' + (err && err.message ? err.message : err));
    return null;
  }
}

function getProductionVolume(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var values = getProductionVolumeSheet().getDataRange().getValues();
  var manual = values.length <= 1 ? [] : values.slice(1).filter(function(r) { return String(r[0] || '').trim(); }).map(function(r) {
    return { month: String(r[0] || ''), line: String(r[1] || ''), actual_qty: Number(r[2] || 0), updated_by: String(r[3] || ''), updated_at: String(r[4] || ''), source: 'manual' };
  });
  // ไลน์ที่ตั้งค่า productionLogSources ไว้ ให้ยอดจาก ProductionLog จริงทับค่าที่กรอกมือเสมอ
  // (ค่ากรอกมือยังอยู่เป็น fallback เผื่อดึงข้อมูลจริงไม่สำเร็จ)
  Object.keys(SPARE_APP_CONFIG.productionLogSources || {}).forEach(function(line) {
    var external = getExternalProductionVolumeForLine(line);
    if (!external) return;
    external.forEach(function(row) {
      var idx = -1;
      for (var i = 0; i < manual.length; i += 1) {
        if (manual[i].line === line && manual[i].month === row.month) { idx = i; break; }
      }
      var entry = { month: row.month, line: line, actual_qty: row.actual_qty, updated_by: 'ProductionLog (auto)', updated_at: '', source: 'auto' };
      if (idx === -1) manual.push(entry); else manual[idx] = entry;
    });
  });
  return manual;
}

function upsertProductionVolume(payload) {
  var user = requireProductionCostEditor({ authToken: payload.authToken });
  var month = String(payload.month || '').trim();
  var line = String(payload.line || '').trim();
  var qty = Number(payload.actual_qty);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('รูปแบบเดือนไม่ถูกต้อง (yyyy-MM)');
  if (!line) throw new Error('กรุณาระบุ Line');
  if (!isFinite(qty) || qty < 0) throw new Error('ยอดผลิตต้องเป็นตัวเลขไม่ติดลบ');
  var sheet = getProductionVolumeSheet();
  var values = sheet.getDataRange().getValues();
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') === month && String(values[i][1] || '') === line) {
      sheet.getRange(i + 1, 1, 1, PRODUCTION_VOLUME_HEADERS.length).setValues([[month, line, qty, user.username, now]]);
      return { status: 'success', mode: 'update' };
    }
  }
  sheet.appendRow([month, line, qty, user.username, now]);
  return { status: 'success', mode: 'insert' };
}

function getProductionCostConfig(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var values = getProductionCostConfigSheet().getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(function(r) { return String(r[0] || '').trim(); }).map(function(r) {
    return { line: String(r[0] || ''), unit_price: Number(r[1] || 0), target_pct: Number(r[2] || 0), updated_by: String(r[3] || ''), updated_at: String(r[4] || '') };
  });
}

function upsertProductionCostConfig(payload) {
  var user = requireProductionCostEditor({ authToken: payload.authToken });
  var line = String(payload.line || '').trim();
  var unitPrice = Number(payload.unit_price);
  var targetPct = Number(payload.target_pct);
  if (!line) throw new Error('กรุณาระบุ Line');
  if (!isFinite(unitPrice) || unitPrice < 0) throw new Error('ราคา/หน่วยต้องเป็นตัวเลขไม่ติดลบ');
  if (!isFinite(targetPct) || targetPct < 0) throw new Error('เป้าหมาย % ต้องเป็นตัวเลขไม่ติดลบ');
  var sheet = getProductionCostConfigSheet();
  var values = sheet.getDataRange().getValues();
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') === line) {
      sheet.getRange(i + 1, 1, 1, PRODUCTION_COST_CONFIG_HEADERS.length).setValues([[line, unitPrice, targetPct, user.username, now]]);
      return { status: 'success', mode: 'update' };
    }
  }
  sheet.appendRow([line, unitPrice, targetPct, user.username, now]);
  return { status: 'success', mode: 'insert' };
}

function getPurchaseHistoryImportLogSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.purchaseHistoryImportLogSheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(PURCHASE_HISTORY_IMPORT_LOG_HEADERS);
  return sheet;
}

function normalizePurchaseImportValue(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function buildPurchaseImportDuplicateKeys(item) {
  var line = normalizePurchaseImportValue(item.line);
  var model = normalizePurchaseImportValue(item.model);
  var qty = Number(item.qty_ordered || 0);
  var hash = normalizePurchaseImportValue(item.source_file_hash);
  var fileName = normalizePurchaseImportValue(item.source_file_name);
  var period = normalizePurchaseImportValue(item.request_period);
  return {
    hashKey: hash && line && model && qty > 0 ? [hash, line, model, qty].join('|') : '',
    fallbackKey: fileName && period && line && model && qty > 0 ? [fileName, period, line, model, qty].join('|') : ''
  };
}

function buildPurchaseImportRequestId(item) {
  var hashOrFile = normalizePurchaseImportValue(item.source_file_hash).slice(0, 12) || normalizePurchaseImportValue(item.source_file_name);
  var parts = [hashOrFile, item.line, item.model, Number(item.qty_ordered || 0)].map(function(value) {
    return normalizePurchaseImportValue(value).replace(/[^a-z0-9ก-๙._-]+/g, '-').replace(/^-+|-+$/g, '');
  }).filter(Boolean);
  return ('PDF-' + (parts.join('-') || Utilities.getUuid())).slice(0, 180);
}

function findPurchaseImportDuplicate(values, item) {
  var target = buildPurchaseImportDuplicateKeys(item);
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    if (toBoolean(row[22], false)) continue;
    var current = buildPurchaseImportDuplicateKeys({
      line: row[5], model: row[9], qty_ordered: row[10], source_file_name: row[26], source_file_hash: row[27], request_period: row[30]
    });
    if (target.hashKey && current.hashKey === target.hashKey) return { rowIndex: i, historyId: String(row[0] || ''), keyType: 'hash' };
    if (!target.hashKey && target.fallbackKey && current.fallbackKey === target.fallbackKey) return { rowIndex: i, historyId: String(row[0] || ''), keyType: 'fallback' };
  }
  return null;
}

function checkPurchaseHistoryImportDuplicates(payload) {
  requirePurchaseHistoryEditor({ authToken: payload.authToken });
  var items = Array.isArray(payload.items) ? payload.items : [];
  var values = getPurchaseHistorySheet().getDataRange().getValues();
  var seen = {};
  return {
    status: 'success',
    duplicates: items.map(function(item, index) {
      var duplicate = findPurchaseImportDuplicate(values, item);
      if (duplicate) return { index: index, history_id: duplicate.historyId, key_type: duplicate.keyType };
      var keys = buildPurchaseImportDuplicateKeys(item);
      var key = keys.hashKey || keys.fallbackKey;
      if (key && seen[key] !== undefined) return { index: index, history_id: '', key_type: 'batch', duplicate_of_index: seen[key] };
      if (key) seen[key] = index;
      return null;
    }).filter(Boolean)
  };
}

function importPurchaseHistoryPdfBatch(payload) {
  var user = requirePurchaseHistoryEditor({ authToken: payload.authToken });
  var items = Array.isArray(payload.items) ? payload.items : [];
  var files = Array.isArray(payload.files) ? payload.files : [];
  var batchId = String(payload.import_batch_id || ('PHIMP-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8)));
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  var imported = 0;
  var skipped = 0;
  var reviewRequired = 0;
  var fileStats = {};

  files.forEach(function(file) {
    fileStats[file.file_name] = { total: Number(file.total_rows || 0), imported: 0, skipped: 0, review: Number(file.review_required_rows || 0), reviewItemsSeen: 0 };
  });

  items.forEach(function(item) {
    var fileName = String(item.source_file_name || '');
    if (!fileStats[fileName]) fileStats[fileName] = { total: 0, imported: 0, skipped: 0, review: 0, reviewItemsSeen: 0 };
    if (item.review_required || !String(item.part_name || '').trim() || !String(item.model || '').trim() || Number(item.qty_ordered || 0) <= 0) {
      reviewRequired += 1;
      fileStats[fileName].reviewItemsSeen += 1;
      if (fileStats[fileName].reviewItemsSeen > fileStats[fileName].review) fileStats[fileName].review += 1;
      return;
    }
    var duplicate = findPurchaseImportDuplicate(values, item);
    var duplicateAction = String(item.duplicate_action || payload.duplicate_action || 'skip').toLowerCase();
    if (duplicate && duplicateAction !== 'replace') {
      skipped += 1;
      fileStats[fileName].skipped += 1;
      return;
    }
    if (duplicate && duplicateAction === 'replace') {
      item.request_id = values[duplicate.rowIndex][1];
      item.part_id = values[duplicate.rowIndex][6] || item.part_id;
    }
    var priceText = String(item.unit_price === undefined || item.unit_price === null ? '' : item.unit_price).trim();
    var result = upsertPurchaseHistoryRecord({
      request_id: item.request_id || buildPurchaseImportRequestId(item),
      history_id: duplicate ? duplicate.historyId : undefined,
      source: 'PR Report', date: item.requested_date || new Date(), line: item.line || '', part_id: item.part_id || '',
      part_name: item.part_name || '', brand: item.brand || '', model: item.model || '', qty_ordered: Number(item.qty_ordered || 0), unit: item.unit || '',
      unit_price: priceText, currency: priceText ? (item.currency || 'THB') : '', status: item.status === 'Requested' ? 'Requested' : 'PR Created',
      requested_by: item.requested_by || '', updated_by: user.username, remark: item.remark || '', import_batch_id: batchId,
      source_file_name: fileName, source_file_hash: item.source_file_hash || '', price_status: priceText ? 'Confirmed' : 'TBC',
      created_by: user.username, request_period: item.request_period || '', force_status: duplicateAction === 'replace'
    });
    if (!result.skipped) {
      imported += 1;
      fileStats[fileName].imported += 1;
      // append new row to local copy so duplicate detection in subsequent iterations stays accurate
      // without making an extra Sheets API call per item
      if (result.row) values.push(result.row);
    }
  });

  var importedAt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var logSheet = getPurchaseHistoryImportLogSheet();
  Object.keys(fileStats).forEach(function(fileName) {
    var stat = fileStats[fileName];
    logSheet.appendRow([batchId, fileName, user.username, importedAt, stat.total, stat.imported, stat.skipped, stat.review]);
  });
  reviewRequired = Object.keys(fileStats).reduce(function(total, fileName) { return total + Number(fileStats[fileName].review || 0); }, 0);
  return { status: 'success', import_batch_id: batchId, imported_rows: imported, skipped_duplicate_rows: skipped, review_required_rows: reviewRequired };
}


function createPurchaseHistoryBatch(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  requirePermission({ authToken: payload.authToken }, 'view_logs');
  var items = Array.isArray(payload.items) ? payload.items : [];
  if (!items.length) throw new Error('ไม่พบรายการสำหรับบันทึก Purchase History');
  var month = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
  var results = [];
  items.forEach(function(item, index) {
    var qty = Number(item.qty_ordered || item.qtyOrdered || 0);
    if (!isFinite(qty) || qty <= 0) return;
    var partKey = String(item.part_id || item.partId || item.model || index + 1).replace(/[^a-zA-Z0-9_-]/g, '_');
    var requestId = item.request_id || item.requestId || ('AUTOPR-' + month + '-' + partKey);
    var result = upsertPurchaseHistoryRecord({
      request_id: requestId, history_id: 'PH-' + requestId, match_open_item: true, source: item.request_id ? 'PR Report' : 'Auto PR',
      date: payload.date || new Date(), line: item.line || '', part_id: item.part_id || item.partId || '', part_name: item.part_name || item.partName || '',
      brand: item.brand || '', model: item.model || '', qty_ordered: qty, unit: item.unit || '', unit_price: item.unit_price,
      currency: item.currency || '', status: 'PR Created', requested_by: item.requested_by || session.user.username,
      updated_by: session.user.username, remark: item.remark || ''
    });
    if (!result.skipped) results.push(result);
  });
  return { status: 'success', count: results.length, history_ids: results.map(function(result) { return result.history_id; }) };
}

function addManualPurchaseHistory(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  requirePermission({ authToken: payload.authToken }, 'view_logs');
  var qty = Number(payload.qty_ordered || 0);
  if (!isFinite(qty) || qty <= 0) throw new Error('Qty Ordered ต้องมากกว่า 0');
  if (!String(payload.part_name || '').trim()) throw new Error('กรุณาระบุ Part Name');
  var status = String(payload.status || 'Requested').trim();
  if (PURCHASE_HISTORY_STATUSES.indexOf(status) === -1) throw new Error('Status ไม่ถูกต้อง');
  var result = upsertPurchaseHistoryRecord({
    source: 'Manual',
    date: payload.date || new Date(),
    line: payload.line || '',
    part_name: String(payload.part_name || '').trim(),
    brand: payload.brand || '',
    model: payload.model || '',
    qty_ordered: qty,
    unit: payload.unit || '',
    unit_price: payload.unit_price !== undefined && payload.unit_price !== '' ? Number(payload.unit_price) : undefined,
    currency: payload.currency || (payload.unit_price ? 'THB' : ''),
    status: status,
    requested_by: String(payload.requested_by || session.user.username || '').trim(),
    updated_by: session.user.username,
    remark: payload.remark || '',
    force_status: true
  });
  return { status: 'success', history: result };
}

function syncPurchaseHistoryForRequest(requestRow, status, updatedBy, remark, preserveExistingQty) {
  var statusMap = { Pending: 'Requested', Approved: 'Requested', 'On Hold': 'Requested', 'Converted to PR': 'PR Created', Purchased: 'Ordered', Received: 'Received', Rejected: 'Cancelled', Closed: 'Cancelled' };
  var purchaseStatus = statusMap[status] || 'Requested';
  return upsertPurchaseHistoryRecord({
    request_id: requestRow.request_id, history_id: 'PH-' + requestRow.request_id, source: 'Purchase Request', date: requestRow.requested_date,
    line: requestRow.line, part_id: requestRow.item_id, part_name: requestRow.item_name, brand: requestRow.brand, model: requestRow.model,
    qty_ordered: preserveExistingQty ? undefined : requestRow.request_qty, unit: requestRow.unit || '', unit_price: requestRow.unit_price,
    currency: requestRow.currency || '', requested_by: requestRow.requested_by, status: purchaseStatus,
    ordered_date: purchaseStatus === 'Ordered' ? new Date() : '', received_date: purchaseStatus === 'Received' ? new Date() : '',
    received_qty: purchaseStatus === 'Received' ? Number(requestRow.request_qty || 0) : undefined,
    updated_by: updatedBy || '', remark: remark !== undefined ? remark : requestRow.remark
  });
}

function syncPurchaseHistoryOnReceive(payload, updatedBy) {
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  var remaining = Number(payload.qty || 0);
  var partId = normalizePurchaseHistoryKeyPart(payload.partNo);
  var line = normalizePurchaseHistoryKeyPart(payload.process);
  var modelNorm = normalizePurchaseHistoryModel(payload.model);
  var modelOk = isMeaningfulPurchaseHistoryModel(payload.model);
  var nameNorm = normalizePurchaseHistoryName(payload.partName);
  var brandOk = isMeaningfulPurchaseHistoryBrand(payload.brand);
  var brandNorm = normalizePurchaseHistoryName(payload.brand);
  var openStatuses = { 'Requested': true, 'PR Created': true, 'Ordered': true, 'Partial Received': true };
  var matched = 0;
  for (var i = 1; i < values.length && remaining > 0; i += 1) {
    var row = values[i];
    if (toBoolean(row[22], false) || !openStatuses[String(row[16] || '')]) continue;
    // จับคู่โดยดึงทุกฟิลด์มาเทียบตามลำดับ: NO+Line → รุ่น → ชื่อ (+ยี่ห้อกันชื่อซ้ำต่างยี่ห้อ)
    var matches = false;
    var rowLineMatch = !line || normalizePurchaseHistoryKeyPart(row[5]) === line;
    // 1) part_id ตรงกัน = match ทันที — แต่เลข NO เป็นแค่ลำดับในแต่ละ Sheet/Line เท่านั้น
    //    ไม่ unique ทั้งระบบ จึงต้องอยู่ Line เดียวกันด้วยเสมอ ไม่งั้นข้ามไปเทียบรุ่น/ชื่อแทน
    if (partId && line && rowLineMatch && normalizePurchaseHistoryKeyPart(row[6]) === partId) {
      matches = true;
    } else if (rowLineMatch) {
      // ต้องอยู่ Line เดียวกัน เพื่อกันหักยอดผิดไลน์
      var rowModelOk = isMeaningfulPurchaseHistoryModel(row[9]);
      if (modelOk && rowModelOk) {
        // 2) ทั้งคู่มีรุ่น -> ตัดสินด้วยรุ่นเป็นหลัก (รุ่นต่างกัน = คนละตัว แม้ชื่อจะซ้ำ)
        matches = normalizePurchaseHistoryModel(row[9]) === modelNorm;
      } else if (nameNorm && normalizePurchaseHistoryName(row[7]) === nameNorm) {
        // 3) ฝั่งใดไม่มีรุ่น (อะไหล่ที่ไม่มี Model/Brand) -> เทียบชื่อ
        //    ถ้าทั้งคู่มียี่ห้อชัดเจนแต่คนละยี่ห้อ -> ถือว่าคนละตัว
        var rowBrandOk = isMeaningfulPurchaseHistoryBrand(row[8]);
        matches = !(brandOk && rowBrandOk && normalizePurchaseHistoryName(row[8]) !== brandNorm);
      }
    }
    if (!matches) continue;
    var orderedQty = Number(row[10] || 0), receivedBefore = Number(row[19] || 0), outstanding = Math.max(orderedQty - receivedBefore, 0);
    if (outstanding <= 0) continue;
    var applied = Math.min(outstanding, remaining), receivedTotal = receivedBefore + applied;
    upsertPurchaseHistoryRecord({
      request_id: row[1], part_id: row[6], status: receivedTotal >= orderedQty ? 'Received' : 'Partial Received',
      received_date: new Date(), received_qty: receivedTotal, updated_by: updatedBy || payload.by || '', remark: row[21]
    });
    remaining -= applied; matched += 1;
  }
  return { matched: matched, unmatched_qty: remaining };
}

// ตรวจสอบ/ Reconcile: หักยอด PO ที่ยังค้างอยู่ ย้อนหลังกับประวัติการรับเข้าทั้งหมด
// (ครอบคลุมเคสที่คำขอซื้อถูกสร้างขึ้น "ทีหลัง" การรับของจริง ซึ่งระบบหักยอดแบบ
// real-time ตอนรับของ (syncPurchaseHistoryOnReceive) ไม่มีโอกาสจับคู่ให้ เพราะ
// มันทำงานเฉพาะตอนมีการรับของใหม่เกิดขึ้นเท่านั้น ไม่ย้อนสแกนของเก่า)
function purchaseHistoryReconcileKey(line, partId, model, name) {
  var l = normalizePurchaseHistoryKeyPart(line);
  var pid = normalizePurchaseHistoryKeyPart(partId);
  if (pid) return 'P|' + l + '|' + pid;
  if (isMeaningfulPurchaseHistoryModel(model)) return 'M|' + l + '|' + normalizePurchaseHistoryModel(model);
  return 'N|' + l + '|' + normalizePurchaseHistoryName(name);
}
function reconcilePurchaseHistory(payload) {
  var user = requirePurchaseHistoryEditor({ authToken: payload.authToken });
  var sheet = getPurchaseHistorySheet();
  var values = sheet.getDataRange().getValues();
  var openStatuses = { 'Requested': true, 'PR Created': true, 'Ordered': true, 'Partial Received': true };

  var openByKey = {};
  for (var i = 1; i < values.length; i += 1) {
    var row = values[i];
    if (toBoolean(row[22], false) || !openStatuses[String(row[16] || '')]) continue;
    var key = purchaseHistoryReconcileKey(row[5], row[6], row[9], row[7]);
    (openByKey[key] = openByKey[key] || []).push({ rowIndex: i, orderedQty: Number(row[10] || 0), receivedBefore: Number(row[19] || 0), date: row[3] });
  }
  var openKeys = Object.keys(openByKey);
  if (!openKeys.length) return { status: 'success', updated: 0, qty_applied: 0 };
  openKeys.forEach(function(k) { openByKey[k].sort(function(a, b) { return new Date(a.date) - new Date(b.date); }); });

  var logRows = getLogRows().filter(function(l) { return String(l.type || '').toLowerCase().indexOf('input') > -1; });
  var inputByKey = {};
  logRows.forEach(function(l) {
    var key = purchaseHistoryReconcileKey(l.process, l.partNo, l.model, l.partName);
    if (openByKey[key] === undefined) return; // ข้ามถ้าไม่มี PO เปิดค้างสำหรับชิ้นนี้ ลดงานคำนวณ
    inputByKey[key] = (inputByKey[key] || 0) + Math.abs(Number(l.qty || 0));
  });

  var updatedCount = 0, totalApplied = 0;
  openKeys.forEach(function(key) {
    var avail = Number(inputByKey[key] || 0);
    if (avail <= 0) return;
    var rows = openByKey[key];
    for (var j = 0; j < rows.length && avail > 0; j += 1) {
      var r = rows[j];
      var outstanding = Math.max(r.orderedQty - r.receivedBefore, 0);
      if (outstanding <= 0) continue;
      var applied = Math.min(outstanding, avail);
      var receivedTotal = r.receivedBefore + applied;
      var existingRow = values[r.rowIndex];
      upsertPurchaseHistoryRecord({
        request_id: existingRow[1], part_id: existingRow[6], status: receivedTotal >= r.orderedQty ? 'Received' : 'Partial Received',
        received_date: new Date(), received_qty: receivedTotal, updated_by: user.username, remark: existingRow[21]
      });
      avail -= applied; updatedCount += 1; totalApplied += applied;
    }
  });
  return { status: 'success', updated: updatedCount, qty_applied: totalApplied };
}

function bulkUpdateOrderRequestStatus(payload) {
  var ids = Array.isArray(payload.request_ids) ? payload.request_ids : [];
  var targetStatus = String(payload.status || '');
  var allowed = { Purchased: 'request_order_approve', Received: 'request_order_approve', Rejected: 'request_order_reject' };
  if (!ids.length) throw new Error('กรุณาเลือกรายการอย่างน้อย 1 รายการ');
  if (!allowed[targetStatus]) throw new Error('สถานะ Bulk Update ไม่ถูกต้อง');
  var user = requirePermission({ authToken: payload.authToken }, allowed[targetStatus]);
  if (ORDER_REQUEST_STATUSES.indexOf(targetStatus) === -1) throw new Error('สถานะไม่ถูกต้อง: ' + targetStatus);

  var sheet = getOrderRequestSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0] || [];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var idSet = {};
  ids.forEach(function(id) { idSet[String(id)] = true; });
  var updatedCount = 0;

  for (var i = 1; i < values.length; i += 1) {
    if (!idSet[String(values[i][idx.request_id])]) continue;
    var updatedRow = values[i].slice();
    updatedRow[idx.status] = targetStatus;
    updatedRow[idx.admin_comment] = payload.admin_comment || '';
    updatedRow[idx.updated_at] = now;
    if (targetStatus === 'Approved') {
      updatedRow[idx.approved_by] = user.username || '';
      updatedRow[idx.approved_date] = now;
    }
    sheet.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
    try {
      syncPurchaseHistoryForRequest(toRequestObject(headers, updatedRow), targetStatus, user.username, payload.admin_comment || values[i][idx.remark] || '', true);
    } catch (historyErr) {
      Logger.log('bulkUpdateOrderRequestStatus PurchaseHistory warning [' + values[i][idx.request_id] + ']: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
    }
    updatedCount += 1;
  }

  return { status: 'success', updated_status: targetStatus, count: updatedCount };
}


function getOrderRequestSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.requestSheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ORDER_REQUEST_HEADERS);
    } else {
      var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), ORDER_REQUEST_HEADERS.length)).getValues()[0];
      ORDER_REQUEST_HEADERS.forEach(function(header) {
        if (currentHeaders.indexOf(header) !== -1) return;
        var nextColumn = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextColumn).setValue(header);
        currentHeaders.push(header);
      });
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
    var requestId = 'REQ-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8);
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
      'Pending', '', '', '', '', Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'), payload.unit || '', payload.unit_price === undefined ? '' : payload.unit_price, payload.currency || ''
    ];
    sheet.appendRow(row);
    var purchaseHistoryRecorded = true;
    try {
      syncPurchaseHistoryForRequest(toRequestObject(ORDER_REQUEST_HEADERS, row), 'Pending', user.username, payload.remark || '');
    } catch (historyErr) {
      purchaseHistoryRecorded = false;
      Logger.log('createOrderRequest PurchaseHistory warning: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
    }
    return { status: 'success', request_id: requestId, purchase_history_recorded: purchaseHistoryRecorded };
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
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][idx.request_id]) === String(payload.request_id)) {
      var updatedRow = values[i].slice();
      updatedRow[idx.status] = nextStatus;
      updatedRow[idx.admin_comment] = payload.admin_comment || '';
      updatedRow[idx.updated_at] = now;
      if (nextStatus === 'Approved') {
        updatedRow[idx.approved_by] = user.username || '';
        updatedRow[idx.approved_date] = now;
      }
      sheet.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
      try {
        syncPurchaseHistoryForRequest(toRequestObject(headers, updatedRow), nextStatus, user.username, payload.admin_comment || values[i][idx.remark] || '', true);
      } catch (historyErr) {
        Logger.log('updateOrderRequestStatus PurchaseHistory warning: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
      }
      return { status: 'success', request_id: payload.request_id, updated_status: nextStatus };
    }
  }
  throw new Error('ไม่พบ request_id');
}

var ORDER_REQUEST_EDITABLE_FIELDS = ['item_name', 'model', 'brand', 'category', 'line', 'request_qty', 'unit', 'reason', 'remark'];
var ORDER_REQUEST_EDITABLE_STATUSES = { Pending: true, 'On Hold': true };

// แก้ไขรายละเอียดคำขอซื้อ (ชื่อ/รุ่น/แบรนด์/Line/จำนวน/หน่วย/เหตุผล/หมายเหตุ) —
// ให้ Admin แก้ไขข้อมูลที่ผู้ขอกรอกผิด/ไม่ครบ ก่อนอนุมัติ แก้ได้เฉพาะรายการที่ยังไม่อนุมัติ (Pending/On Hold)
function editOrderRequest(payload) {
  requirePermission({ authToken: payload.authToken }, 'request_order_edit');
  var session = getSessionUser({ authToken: payload.authToken });
  var user = findUserByUsername(session.user.username);
  var requestId = String(payload.request_id || '').trim();
  if (!requestId) throw new Error('ไม่พบ request_id');
  var sheet = getOrderRequestSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][idx.request_id]) !== requestId) continue;
    var row = values[i].slice();
    var currentStatus = String(row[idx.status] || 'Pending');
    if (!ORDER_REQUEST_EDITABLE_STATUSES[currentStatus]) {
      throw new Error('แก้ไขได้เฉพาะรายการที่ยังไม่อนุมัติ (Pending / On Hold) เท่านั้น — สถานะปัจจุบัน: ' + currentStatus);
    }
    ORDER_REQUEST_EDITABLE_FIELDS.forEach(function(field) {
      if (payload[field] === undefined) return;
      var col = idx[field];
      if (col === undefined) return;
      row[col] = (field === 'request_qty') ? Number(payload[field] || 0) : String(payload[field]).trim();
    });
    if (!String(row[idx.item_name] || '').trim()) throw new Error('กรุณาระบุชื่ออะไหล่');
    if (!(Number(row[idx.request_qty]) > 0)) throw new Error('จำนวนต้องมากกว่า 0');
    row[idx.updated_at] = now;
    var editNote = '✏️ แก้ไขโดย ' + (user.username || '') + ' เมื่อ ' + now;
    var existingComment = String(row[idx.admin_comment] || '').trim();
    row[idx.admin_comment] = existingComment ? (existingComment + ' | ' + editNote) : editNote;
    sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    try {
      syncPurchaseHistoryForRequest(toRequestObject(headers, row), currentStatus, user.username, row[idx.remark] || '', true);
    } catch (historyErr) {
      Logger.log('editOrderRequest PurchaseHistory warning: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
    }
    return { status: 'success', request_id: requestId };
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
        try {
          updatePurchaseHistoryForRequest(id, { request_type: 'PR' });
        } catch (historyErr) {
          Logger.log('convertOrderRequestsToPR PurchaseHistory warning: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
        }
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
    request_order_approve: true, request_order_reject: true, request_order_convert_pr: true, request_order_close: true,
    request_order_edit: true, delete_logs: true
  };
  if (normalized === 'leader') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false,
    request_order_edit: false, delete_logs: false
  };
  return {
    view: true, transact: true, manage_items: false, delete_items: false,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false,
    request_order_edit: false, delete_logs: false
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

function hashPassword(password) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function isPasswordHashed(stored) {
  return /^[0-9a-f]{64}$/.test(String(stored || ''));
}

function ensureDefaultAdminUser() {
  var usersSheet = getUsersSheet();
  if (usersSheet.getLastRow() > 1) return;
  usersSheet.appendRow([
    'admin',
    hashPassword('admin123'),
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

function getSessionPropertyKey(token) {
  return SESSION_PROPERTY_PREFIX + String(token || '').trim();
}

function readSessionRecord(token) {
  var target = String(token || '').trim();
  if (!target) return null;
  var raw = PropertiesService.getScriptProperties().getProperty(getSessionPropertyKey(target));
  if (!raw) return null;
  try {
    var record = JSON.parse(raw);
    if (!record || !record.username || !Number(record.expiry)) return null;
    return record;
  } catch (err) {
    return null;
  }
}

function writeSessionRecord(token, username, expiry) {
  PropertiesService.getScriptProperties().setProperty(getSessionPropertyKey(token), JSON.stringify({
    username: String(username || '').trim(),
    expiry: Number(expiry || 0)
  }));
}

function deleteSessionRecord(token) {
  var target = String(token || '').trim();
  if (!target) return;
  PropertiesService.getScriptProperties().deleteProperty(getSessionPropertyKey(target));
}

function revokeUserSessions(username) {
  var target = String(username || '').trim();
  if (!target) return;
  var props = PropertiesService.getScriptProperties();
  var allProperties = props.getProperties();
  Object.keys(allProperties).forEach(function(key) {
    if (key.indexOf(SESSION_PROPERTY_PREFIX) !== 0) return;
    try {
      var record = JSON.parse(allProperties[key]);
      if (record && record.username === target) props.deleteProperty(key);
    } catch (err) {
      props.deleteProperty(key);
    }
  });
}

function cleanupExpiredSessionRecords() {
  var props = PropertiesService.getScriptProperties();
  var allProperties = props.getProperties();
  var now = Date.now();
  Object.keys(allProperties).forEach(function(key) {
    if (key.indexOf(SESSION_PROPERTY_PREFIX) !== 0) return;
    try {
      var record = JSON.parse(allProperties[key]);
      if (!record || Number(record.expiry || 0) <= now) props.deleteProperty(key);
    } catch (err) {
      props.deleteProperty(key);
    }
  });
}

function findUserByToken(token) {
  var target = String(token || '').trim();
  if (!target) return null;
  var now = Date.now();
  var record = readSessionRecord(target);
  if (record) {
    if (Number(record.expiry) <= now) {
      deleteSessionRecord(target);
      return null;
    }
    var sessionUser = findUserByUsername(record.username);
    if (!sessionUser || !sessionUser.isActive) {
      deleteSessionRecord(target);
      return null;
    }
    if (Number(record.expiry) - now <= Number(SPARE_APP_CONFIG.sessionRefreshThresholdMs)) {
      writeSessionRecord(target, sessionUser.username, now + Number(SPARE_APP_CONFIG.sessionDurationMs));
    }
    return sessionUser;
  }

  // Migrate a still-valid session created by an older deployment. Keeping this
  // fallback avoids forcing every signed-in user to log in during deployment.
  var users = getAllUsers();
  for (var i = 0; i < users.length; i += 1) {
    var user = users[i];
    var legacyExpiry = Number(user.tokenExpiry || 0);
    if (user.token === target && legacyExpiry > now && user.isActive) {
      writeSessionRecord(target, user.username, now + Number(SPARE_APP_CONFIG.sessionDurationMs));
      return user;
    }
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
  var storedPw = String(user.password || '');
  var alreadyHashed = isPasswordHashed(storedPw);
  var passwordMatch = alreadyHashed
    ? (storedPw === hashPassword(password))
    : (storedPw === password);
  if (!passwordMatch) throw new Error('รหัสผ่านไม่ถูกต้อง');
  if (!alreadyHashed) {
    // auto-migrate plain-text password to hash
    getUsersSheet().getRange(user.rowIndex, 2).setValue(hashPassword(password));
  }

  cleanupExpiredSessionRecords();
  var token = Utilities.getUuid() + '-' + Date.now();
  var expiry = Date.now() + Number(SPARE_APP_CONFIG.sessionDurationMs);
  var usersSheet = getUsersSheet();
  usersSheet.getRange(user.rowIndex, 6).setValue(token);
  usersSheet.getRange(user.rowIndex, 7).setValue(String(expiry));
  writeSessionRecord(token, user.username, expiry);

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
  deleteSessionRecord(token);

  // Clear the legacy sheet columns only when they contain this exact token.
  var users = getAllUsers();
  for (var i = 0; i < users.length; i += 1) {
    if (users[i].token !== token) continue;
    var usersSheet = getUsersSheet();
    usersSheet.getRange(users[i].rowIndex, 6).setValue('');
    usersSheet.getRange(users[i].rowIndex, 7).setValue('');
    break;
  }
  return { status: 'success' };
}

function createAuthSessionError(message) {
  var err = new Error(message);
  err.code = 'AUTH_SESSION_INVALID';
  return err;
}

function getSessionUser(payload) {
  var token = String(payload.authToken || payload.token || '').trim();
  if (!token) throw createAuthSessionError('กรุณาเข้าสู่ระบบ');
  var user = findUserByToken(token);
  if (!user) throw createAuthSessionError('session หมดอายุหรือไม่ถูกต้อง');
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
      if (password.length < 4) throw new Error('password ต้องมีอย่างน้อย 4 ตัวอักษร');
      usersSheet.getRange(existing.rowIndex, 2).setValue(hashPassword(password));
    }
    usersSheet.getRange(existing.rowIndex, 3).setValue(role);
    usersSheet.getRange(existing.rowIndex, 4).setValue(String(isActive));
    usersSheet.getRange(existing.rowIndex, 5).setValue(permissionsJson);
    if (String(payload.password || '') !== '' || !isActive) revokeUserSessions(username);
    return { status: 'success', mode: 'update', username: username };
  }

  if (!actor.permissions.add_user) throw new Error('ไม่มีสิทธิ์เพิ่มผู้ใช้');
  if (!password) throw new Error('ต้องระบุ password สำหรับผู้ใช้ใหม่');
  if (password.length < 4) throw new Error('password ต้องมีอย่างน้อย 4 ตัวอักษร');
  usersSheet.appendRow([username, hashPassword(password), role, String(isActive), permissionsJson, '', '']);
  return { status: 'success', mode: 'create', username: username };
}

function deleteUser(payload) {
  var actor = requireAdminUser(payload);
  var username = String(payload.username || '').trim();
  if (!username) throw new Error('ต้องระบุ username');
  if (username === actor.username) throw new Error('ไม่สามารถลบ user ตัวเองได้');
  var existing = findUserByUsername(username);
  if (!existing) throw new Error('ไม่พบผู้ใช้');
  revokeUserSessions(username);
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
    sheetName: e.parameter.sheet,
    authToken: e.parameter.authToken || e.parameter.token || ''
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

// ลบรายการ Log (ประวัติเบิก/รับเข้า) — ใช้สำหรับล้างรายการที่บันทึกผิด/ซ้ำ
// เป็นการลบ "ประวัติ" เท่านั้น ไม่กระทบ Stock ปัจจุบัน (Stock ถูกตัด/เติมไปแล้วตอน
// ทำรายการจริง การลบ log ย้อนหลังจึงไม่ควรไปแก้ stock ซ้ำ)
function deleteLogEntry(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return deleteLogEntryUnlocked(payload);
  } finally {
    lock.releaseLock();
  }
}

function deleteLogEntryUnlocked(payload) {
  requirePermission({ authToken: payload.authToken }, 'delete_logs');
  var no = Number(payload.no);
  if (!no || no <= 0) throw new Error('ไม่พบรายการที่จะลบ');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.writeSheetName);
  ensureLogSheetHeaders(historySheet);
  var rowIndex = no + 1; // แถว 1 คือ header, ข้อมูลแถวแรก (no=1) จึงอยู่แถวชีตที่ 2
  if (rowIndex > historySheet.getLastRow()) throw new Error('ไม่พบรายการที่จะลบ (แถวอาจถูกลบไปแล้ว)');
  var data = historySheet.getDataRange().getValues();
  var headerMap = buildHeaderIndexMap(data[0] || []);
  var row = data[rowIndex - 1] || [];
  // เช็คซ้ำว่า timestamp/partName ที่ frontend ส่งมาตรงกับแถวจริงก่อนลบ กัน race
  // condition กรณีมีรายการอื่นถูกเพิ่ม/ลบสลับแถวไปแล้วตั้งแต่ตอนโหลดหน้าเว็บ
  var tsIdx = headerMap['timestamp'];
  var nameIdx = headerMap['partname'];
  var expectedTs = String(payload.timestamp || '').trim();
  var expectedName = String(payload.partName || payload.partname || '').trim();
  if (expectedTs && tsIdx !== undefined && String(row[tsIdx] || '').trim() !== expectedTs) {
    throw new Error('ข้อมูล Log มีการเปลี่ยนแปลงตั้งแต่โหลดหน้านี้ กรุณารีเฟรชแล้วลองใหม่');
  }
  if (expectedName && nameIdx !== undefined && String(row[nameIdx] || '').trim() !== expectedName) {
    throw new Error('ข้อมูล Log มีการเปลี่ยนแปลงตั้งแต่โหลดหน้านี้ กรุณารีเฟรชแล้วลองใหม่');
  }
  historySheet.deleteRow(rowIndex);
  return { status: 'success', deleted_no: no };
}

function processTransaction(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return processTransactionUnlocked(payload);
  } finally {
    lock.releaseLock();
  }
}

function processTransactionUnlocked(payload) {
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
  var isReceiveTransaction = !(payload.type && String(payload.type).indexOf('Output') > -1);
  if (isReceiveTransaction) {
    if (payload.authToken) payload.by = getSessionUser({ authToken: payload.authToken }).user.username;
    var masterLine = pickRowValue(targetRow, map, ['line', 'linearea', 'area', 'process', 'mainline'], '');
    if (String(masterLine || '').trim()) payload.process = String(masterLine).trim();
    payload.type = 'Input';
  }
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

  var purchaseHistorySync = null;
  if (signedQty > 0) {
    try {
      var transactionUser = payload.authToken ? getSessionUser({ authToken: payload.authToken }).user.username : (payload.by || '');
      purchaseHistorySync = syncPurchaseHistoryOnReceive(payload, transactionUser);
    } catch (historyErr) {
      Logger.log('processTransaction PurchaseHistory warning: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
    }
  }

  return {
    status: 'success',
    stockBefore: stockBefore,
    stockAfter: stockAfter,
    qty: signedQty,
    purchaseHistorySync: purchaseHistorySync
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

// โครงสร้างโฟลเดอร์ใน Drive: ไลน์ > Model > ชื่อไฟล์รูป (ไม่มีโฟลเดอร์ item-N / main/install
// ซ้อนอีกแล้วเหมือนเดิม) — ถ้าไม่มี model (รายการเก่าที่ยังไม่กรอก) fallback เป็น "item-{itemId}"
// กันชื่อโฟลเดอร์ว่าง/ชนกัน
function getUploadTargetFolder(line, itemId, imageType, model) {
  var root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var safeLine = String(line || '').trim() || 'UnknownLine';
  var safeItemId = String(itemId || '').trim() || 'UNKNOWN';
  var safeModel = sanitizeDrivePathSegment(model, 'item-' + safeItemId);

  var lineFolder = getOrCreateChildFolder(root, safeLine);
  var modelFolder = getOrCreateChildFolder(lineFolder, safeModel);

  return {
    folder: modelFolder,
    drivePath: safeLine + '/' + safeModel + '/'
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
  var model = String(payload.model || payload.partNo || payload.part_no || '').trim();
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
  // ใส่ itemId ไว้ในชื่อไฟล์เสมอ เพราะตอนนี้หลายรายการที่ Model เดียวกันจะแชร์โฟลเดอร์เดียวกัน
  // (โครงสร้างใหม่ ไลน์ > Model > ไฟล์ ไม่มีโฟลเดอร์แยกราย item แล้ว) ต้องกันชื่อไฟล์ชนกัน/
  // archive ผิดรายการ
  var namePrefix = kind + '-' + sanitizeDrivePathSegment(itemId, 'UNKNOWN') + '-';
  var fileName = namePrefix + Date.now() + '.' + ext;
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  var target = getUploadTargetFolder(line, itemId, kind, model);
  var folder = target.folder;

  // ย้ายรูปเก่าไป _archive แทนที่จะ Trash (ป้องกันรูปหายถาวร) — กรองด้วย namePrefix เท่านั้น
  // เพราะโฟลเดอร์นี้อาจมีรูปของรายการอื่น (Model เดียวกัน) หรือรูปอีก kind ปนอยู่ด้วย
  var archiveFolderName = '_archive';
  var existing = folder.getFiles();
  var hasOldFiles = false;
  var oldFiles = [];
  while (existing.hasNext()) {
    var f = existing.next();
    if (!f.isTrashed() && f.getName().indexOf(namePrefix) === 0) { oldFiles.push(f); hasOldFiles = true; }
  }
  if (hasOldFiles) {
    try {
      var archiveFolder = getOrCreateChildFolder(folder, archiveFolderName);
      for (var oi = 0; oi < oldFiles.length; oi++) {
        oldFiles[oi].moveTo(archiveFolder);
      }
    } catch (archiveErr) {
      Logger.log('archive warning: ' + (archiveErr && archiveErr.message ? archiveErr.message : archiveErr));
      // ถ้า archive ไม่ได้ ไม่ลบ ปล่อยทับไปก็ได้
    }
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


function sanitizeDrivePathSegment(value, fallbackValue) {
  var fallback = fallbackValue === undefined ? 'Unknown' : String(fallbackValue || '');
  var text = String(value || '').trim();
  if (!text) text = fallback;
  text = text
    .replace(/&/g, '_')
    .replace(/[\\/:*?"<>|#%{}~]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return text || fallback;
}

function normalizeAttachmentLineFolder(line) {
  var raw = String(line || '').trim();
  var normalized = raw.toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, '');
  if (normalized === 'h9') return 'H9';
  if (normalized === 'arcchute' || normalized === 'arcchut') return 'Arc_Chute';
  if (normalized === 'coilwinding') return 'Coil_Winding';
  if (normalized === 'lugscrew' || normalized === 'lugandscrew') return 'Lug_Screw';
  return sanitizeDrivePathSegment(raw, 'Unknown_Line');
}

function getAttachmentKindFolderName(kind) {
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
  return folderName;
}

function buildAttachmentPartFolderName(payload, itemId, itemName) {
  var modelOrPartNo = String(payload.model || payload.partNo || payload.part_no || itemId || '').trim();
  if (!modelOrPartNo) throw new Error('ต้องมี Part ID หรือ Model/Part No. สำหรับจัดเก็บ Drawing');
  var base = sanitizeDrivePathSegment(modelOrPartNo, 'UNKNOWN_PART');
  var suffix = sanitizeDrivePathSegment(itemName, '');
  return suffix ? (base + '_' + suffix) : base;
}

function buildAttachmentStoredFileName(kind, payload, itemId, originalName) {
  var ext = getFileExtensionFromName(originalName);
  var modelOrPartNo = String(payload.model || payload.partNo || payload.part_no || itemId || '').trim();
  var base = sanitizeDrivePathSegment(modelOrPartNo, 'UNKNOWN_PART');
  var revision = sanitizeDrivePathSegment(payload.revision || payload.drawingRevision || payload.drawing_revision || payload.rev || '', '');
  var labelMap = { drawing: 'Drawing', drawings: 'Drawing', datasheet: 'Datasheet', datasheets: 'Datasheet', quotation: 'Quotation', quotations: 'Quotation', photo: 'Photo', photos: 'Photo' };
  var label = labelMap[String(kind || '').toLowerCase()] || 'Attachment';
  var fileName = base + '_' + label + (revision ? '_Rev' + revision.replace(/^rev_?/i, '') : '');
  if (ext) fileName += '.' + ext;
  return fileName;
}

function getSparePartsAttachmentFolder(kind, payload) {
  payload = payload || {};
  var folderName = getAttachmentKindFolderName(kind);
  var root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var spareRoot = getOrCreateChildFolder(root, 'SpareParts');
  var kindRoot = getOrCreateChildFolder(spareRoot, folderName);
  var targetFolder = kindRoot;
  var drivePath = 'SpareParts/' + folderName + '/';

  if (folderName === 'Drawings') {
    var lineFolderName = normalizeAttachmentLineFolder(payload.line || payload.mainLine || payload.sourceLine || payload.sheet || payload.sourceSheet || '');
    var partFolderName = buildAttachmentPartFolderName(payload, String(payload.itemId || payload.no || '').trim(), String(payload.itemName || payload.name || '').trim());
    targetFolder = getOrCreateChildFolder(getOrCreateChildFolder(kindRoot, lineFolderName), partFolderName);
    drivePath += lineFolderName + '/' + partFolderName + '/';
  }

  return {
    folder: targetFolder,
    folderName: folderName,
    drivePath: drivePath
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
  var safeOriginalName = sanitizeDrivePathSegment(originalName, kind + '-' + Date.now());
  var fileName = buildAttachmentStoredFileName(kind, payload, itemId, originalName);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var target = getSparePartsAttachmentFolder(kind, payload);
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
    photo: findCol(['sparepartsphotos', 'photo', 'photo_url', 'image', 'imageurl', 'picture', 'pic']),
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
  // Reuse the existing image_main/image_main_url columns for photos.
  // Do not create a separate Photo URL column; it duplicates existing image data.
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
    drawing_status: normalizeDrawingStatusValue(payload.drawing_status),
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
    if (action === 'getPurchaseHistory') return respond(getPurchaseHistory(e.parameter), e);
    if (action === 'editPurchaseHistory') return respond(editPurchaseHistory(e.parameter), e);
    if (action === 'deletePurchaseHistory') return respond(deletePurchaseHistory(e.parameter), e);
    if (action === 'getProductionVolume') return respond(getProductionVolume(e.parameter), e);
    if (action === 'upsertProductionVolume') return respond(upsertProductionVolume(e.parameter), e);
    if (action === 'getProductionCostConfig') return respond(getProductionCostConfig(e.parameter), e);
    if (action === 'upsertProductionCostConfig') return respond(upsertProductionCostConfig(e.parameter), e);
    if (action === 'checkPurchaseHistoryImportDuplicates') return respond(checkPurchaseHistoryImportDuplicates(e.parameter), e);
    if (action === 'addManualPurchaseHistory') return respond(addManualPurchaseHistory(e.parameter), e);
    if (action === 'bulkUpdateOrderRequestStatus') return respond(bulkUpdateOrderRequestStatus(e.parameter), e);
    if (action === 'ensureOrderRequestsSheet') return respond(ensureOrderRequestsSheetReady(e.parameter), e);
    if (action === 'approveOrderRequest') return respond(approveOrderRequest(e.parameter), e);
    if (action === 'rejectOrderRequest') return respond(rejectOrderRequest(e.parameter), e);
    if (action === 'holdOrderRequest') return respond(holdOrderRequest(e.parameter), e);
    if (action === 'closeOrderRequest') return respond(closeOrderRequest(e.parameter), e);
    if (action === 'markOrderRequestPurchased') return respond(markOrderRequestPurchased(e.parameter), e);
    if (action === 'markOrderRequestReceived') return respond(markOrderRequestReceived(e.parameter), e);
    if (action === 'editOrderRequest') return respond(editOrderRequest(e.parameter), e);
    if (action === 'updateOrderRequestStatus') return respond(updateOrderRequestStatus(e.parameter, e.parameter.status), e);
    if (action === 'saveStockCountResult') return respond(saveStockCountResult(e.parameter), e);
    if (action === 'getStockCountHistory') return respond(getStockCountHistory(e.parameter), e);
    if (action === 'adjustStockFromCount') return respond(adjustStockFromCount(e.parameter), e);
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
    if (action === 'deleteLogEntry') return respond(deleteLogEntry(e.parameter), e);
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
    if (action === 'getSparePartsLite') e.parameter.lite = '1';
    if (action === 'getSparePartDetail') e.parameter.lite = '';

    var sheetName = resolveReadSheetName({ sheet: e.parameter.sheet });
    var isDetailRead = action === 'getSparePartDetail';
    var detailLookup = String(e.parameter.partId || e.parameter.no || e.parameter.id || e.parameter.model || '').trim();
    var isLiteRead = String(e.parameter.lite || e.parameter.mode || '').toLowerCase() === '1' || String(e.parameter.lite || e.parameter.mode || '').toLowerCase() === 'lite';
    var liteCache = null;
    var liteCacheKey = '';
    if (isLiteRead && String(e.parameter.refresh || '') !== '1') {
      try {
        liteCache = CacheService.getScriptCache();
        liteCacheKey = 'spare_parts_lite::' + sheetName;
        var cachedLite = liteCache.get(liteCacheKey);
        if (cachedLite) return respond(JSON.parse(cachedLite), e);
      } catch (cacheReadErr) {
        Logger.log('lite cache read warning [' + sheetName + ']: ' + (cacheReadErr && cacheReadErr.message ? cacheReadErr.message : cacheReadErr));
      }
    }
    var ctx = getMainSheetContext(sheetName);
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
        drawing_url: isLiteRead ? '' : pickRowValue(row, map, ['drawingurl', 'drawing_url'], ''),
        drawing_file_name: isLiteRead ? '' : pickRowValue(row, map, ['drawingfilename', 'drawing_file_name'], ''),
        drawing_revision: isLiteRead ? '' : pickRowValue(row, map, ['drawingrevision', 'drawingrev', 'drawing_revision', 'drawing_rev'], ''),
        drawing_status: normalizeDrawingStatusValue(pickRowValue(row, map, ['drawingstatus', 'drawing_status'], '')),
        datasheet_url: isLiteRead ? '' : pickRowValue(row, map, ['datasheeturl', 'datasheet_url'], ''),
        quotation_url: isLiteRead ? '' : pickRowValue(row, map, ['quotationurl', 'quotation_url'], ''),
        unit_price: isLiteRead ? '' : pickRowValue(row, map, ['unitprice', 'unit_price'], ''),
        currency: pickRowValue(row, map, ['currency'], 'THB'),
        supplier: isLiteRead ? '' : pickRowValue(row, map, ['supplier'], ''),
        price_updated_at: isLiteRead ? '' : pickRowValue(row, map, ['priceupdatedat', 'price_updated_at'], ''),
        price_remark: isLiteRead ? '' : pickRowValue(row, map, ['priceremark', 'price_remark'], ''),
        coil_size: pickRowValue(row, map, ['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size'], '-')
      };
    }).filter(function (item) {
      return item.name && item.name !== '-';
    });

    if (isDetailRead && detailLookup) {
      result = result.filter(function(item) {
        return String(item.no || '') === detailLookup || String(item.id || '') === detailLookup || String(item.partId || '') === detailLookup || String(item.model || '') === detailLookup;
      });
    }

    if (isLiteRead && liteCache && liteCacheKey) {
      try {
        liteCache.put(liteCacheKey, JSON.stringify(result), 300);
      } catch (cacheWriteErr) {
        Logger.log('lite cache write warning [' + sheetName + ']: ' + (cacheWriteErr && cacheWriteErr.message ? cacheWriteErr.message : cacheWriteErr));
      }
    }

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
    if (action === 'saveStockCountResult') return respond(saveStockCountResult(body), e);
    if (action === 'adjustStockFromCount') return respond(adjustStockFromCount(body), e);
    if (action === 'createOrderRequest') return respond(createOrderRequest(body), e);
    if (action === 'uploadRequestAttachment') return respond(uploadRequestAttachment(body), e);
    if (action === 'getOrderRequests') return respond(getOrderRequests(body), e);
    if (action === 'getPurchaseHistory') return respond(getPurchaseHistory(body), e);
    if (action === 'editPurchaseHistory') return respond(editPurchaseHistory(body), e);
    if (action === 'deletePurchaseHistory') return respond(deletePurchaseHistory(body), e);
    if (action === 'getProductionVolume') return respond(getProductionVolume(body), e);
    if (action === 'upsertProductionVolume') return respond(upsertProductionVolume(body), e);
    if (action === 'getProductionCostConfig') return respond(getProductionCostConfig(body), e);
    if (action === 'upsertProductionCostConfig') return respond(upsertProductionCostConfig(body), e);
    if (action === 'checkPurchaseHistoryImportDuplicates') return respond(checkPurchaseHistoryImportDuplicates(body), e);
    if (action === 'addManualPurchaseHistory') return respond(addManualPurchaseHistory(body), e);
    if (action === 'importPurchaseHistoryPdfBatch') return respond(importPurchaseHistoryPdfBatch(body), e);
    if (action === 'createPurchaseHistoryBatch') return respond(createPurchaseHistoryBatch(body), e);
    if (action === 'reconcilePurchaseHistory') return respond(reconcilePurchaseHistory(body), e);
    if (action === 'bulkUpdateOrderRequestStatus') return respond(bulkUpdateOrderRequestStatus(body), e);
    if (action === 'ensureOrderRequestsSheet') return respond(ensureOrderRequestsSheetReady(body), e);
    if (action === 'approveOrderRequest') return respond(approveOrderRequest(body), e);
    if (action === 'rejectOrderRequest') return respond(rejectOrderRequest(body), e);
    if (action === 'holdOrderRequest') return respond(holdOrderRequest(body), e);
    if (action === 'closeOrderRequest') return respond(closeOrderRequest(body), e);
    if (action === 'markOrderRequestPurchased') return respond(markOrderRequestPurchased(body), e);
    if (action === 'markOrderRequestReceived') return respond(markOrderRequestReceived(body), e);
    if (action === 'editOrderRequest') return respond(editOrderRequest(body), e);
    if (action === 'deleteLogEntry') return respond(deleteLogEntry(body), e);
    if (action === 'convertOrderRequestsToPR') return respond(convertOrderRequestsToPR(body), e);
    if (action === 'updateOrderRequestStatus') return respond(updateOrderRequestStatus(body, body.status), e);
    requirePermission(authPayload, 'view');
    if (action === 'upsertItem') {
      requirePermission(authPayload, 'manage_items');
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
// STOCK COUNT
// =============================
var STOCK_COUNT_SHEET_NAME = 'StockCount';
var STOCK_COUNT_HEADERS = ['session_id','month','line','category','sheets','created_by','created_at','submitted_at','status','total_items','matched','diff_count','items_json'];

function getOrCreateStockCountSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(STOCK_COUNT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STOCK_COUNT_SHEET_NAME);
    sheet.appendRow(STOCK_COUNT_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,STOCK_COUNT_HEADERS.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
  }
  return sheet;
}

function saveStockCountResult(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  requirePermission({ authToken: payload.authToken }, 'view_logs');
  var sessionId = 'SC-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + (String(payload.line||'ALL').replace(/[^A-Za-z0-9]/g,'')).toUpperCase().substring(0,4);
  var sheet = getOrCreateStockCountSheet();
  var items = payload.items;
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
  sheet.appendRow([
    sessionId,
    String(payload.month || ''),
    String(payload.line || 'all'),
    String(payload.category || 'all'),
    String(payload.sheets || ''),
    String(payload.created_by || session.user.username),
    String(payload.created_at || ''),
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
    'submitted',
    Number(payload.total_items || 0),
    Number(payload.matched || 0),
    Number(payload.diff_count || 0),
    JSON.stringify(items || [])
  ]);
  return { status: 'success', session_id: sessionId, message: 'บันทึกผลเช็คสต็อกแล้ว' };
}

function getStockCountHistory(payload) {
  getSessionUser({ authToken: payload.authToken });
  requirePermission({ authToken: payload.authToken }, 'view_logs');
  var sheet = getOrCreateStockCountSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[String(h)] = row[i]; });
    obj.items_json = undefined; // ไม่ส่ง items ทั้งหมด (ใหญ่เกิน)
    return obj;
  }).reverse();
}

function adjustStockFromCount(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  requirePermission({ authToken: payload.authToken }, 'view_logs');
  var diffItems = payload.diff_items;
  if (typeof diffItems === 'string') { try { diffItems = JSON.parse(diffItems); } catch(e) { diffItems = []; } }
  if (!diffItems || !diffItems.length) return { status: 'success', adjusted: 0, results: [] };
  var results = [];
  diffItems.forEach(function(item) {
    try {
      var variance = Number(item.counted) - Number(item.system_qty);
      if (variance === 0) return;
      var txnPayload = {
        authToken: payload.authToken,
        partName: String(item.name || ''),
        model: String(item.model || '-'),
        brand: String(item.brand || '-'),
        type: variance > 0 ? 'Input' : 'Output',
        qty: Math.abs(variance),
        unit: String(item.unit || 'PCS'),
        by: session.user.username,
        process: String(payload.line || item.line || ''),
        reason: 'Stock Adjustment',
        reasonRemark: 'Stock Count: ' + String(payload.session_id || '') + ' | ' + String(item.reason || 'ปรับจากการนับจริง'),
        partNo: String(item.id || ''),
        category: String(item.category || 'General')
      };
      processTransaction(txnPayload);
      results.push({ name: item.name, variance: variance, status: 'adjusted' });
    } catch(err) {
      results.push({ name: item.name, status: 'error', error: err.message || String(err) });
    }
  });
  return { status: 'success', adjusted: results.filter(function(r){ return r.status==='adjusted'; }).length, results: results };
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
