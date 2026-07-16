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
SPARE_APP_CONFIG.machinesSheetName = SPARE_APP_CONFIG.machinesSheetName || 'Machines';
// Purchase Request (PR) — ระบบอนุมัติก่อนปริ้น (Header/Lines/Audit แยกชีท)
SPARE_APP_CONFIG.prHeaderSheetName = SPARE_APP_CONFIG.prHeaderSheetName || 'PRHeaders';
SPARE_APP_CONFIG.prLinesSheetName = SPARE_APP_CONFIG.prLinesSheetName || 'PRLines';
SPARE_APP_CONFIG.prAuditSheetName = SPARE_APP_CONFIG.prAuditSheetName || 'PRAudit';
// ชีตติดตามยอดผลิตของแต่ละไลน์ (Google Sheet คนละไฟล์กับสเปรดชีตอะไหล่) — ถ้าตั้งค่าไว้ ระบบจะดึง
// ยอดผลิตจริงมาคำนวณอัตโนมัติแทนการกรอกมือ ต้องให้บัญชีที่รัน Apps Script นี้มีสิทธิ์ View ไฟล์
// ปลายทางด้วย ไม่งั้นจะ fallback ไปใช้ค่าที่กรอกมือแทนเงียบๆ
// รูปแบบค่า: string = spreadsheet id (ใช้ค่า default sheet/คอลัมน์แบบ Lug&Screw) หรือ object
// { id, sheet, dateCol, qtyCol, [modelCol] } เพื่อระบุชื่อชีต/คอลัมน์เอง (แต่ละไลน์โครงสร้างต่างกัน)
// modelCol (ถ้ามี) = คอลัมน์รุ่นสินค้า สำหรับไลน์ที่ราคาต่างกันตาม Model เช่น H9 — เก็บ breakdown
// รายรุ่นไว้คำนวณมูลค่าผลิตแบบต่อรุ่นภายหลัง
SPARE_APP_CONFIG.productionLogSources = SPARE_APP_CONFIG.productionLogSources || {
  'Lug&Screw': { id: '1Xx2XEGtT-KbnvVP_9gzkW9kuyFUBpj1H-oIsT3zUx1U', sheet: 'ProductionLog', dateCol: 'Date', qtyCol: 'ActualQty', modelCol: 'ProductCode' },
  'H9': { id: '1PYcAatoJ4QX28uQ_LF8dDC6oTiMWbfPs5TZDfGJVa4U', sheet: 'Plan', dateCol: 'Actual complete date', qtyCol: 'Actual', modelCol: 'Order model' },
  'Coil Winding': { id: '11NGAEXnTZIXMseO_0vfA-yRWxBXEiWpNkCIdIQq2ftQ', sheet: 'Production_Data', dateCol: 'Date', qtyCol: 'FG', modelCol: 'Product' }
};
// ชีตราคาต่อชิ้นล่าสุด (ทุกไลน์รวมกัน) — ใช้จับคู่กับ Model ของแต่ละไลน์เพื่อคำนวณมูลค่าผลิต
// อัตโนมัติ ไม่ต้องกรอกราคาเอง ชื่อรหัสรุ่นในชีตนี้อาจไม่ตรงกับชีตผลิตเป๊ะ (เช่น H9 ในชีตผลิต
// ไม่มี "-S" ต่อท้าย, Coil มี " (10A)" ต่อท้าย) — matchModelPrice() จัดการความต่างเหล่านี้
SPARE_APP_CONFIG.unitPriceSource = SPARE_APP_CONFIG.unitPriceSource || {
  id: '1yFmVV3qJmCyKGzlRUZ99vlfhps-WGZ9ehSM87dMDFlE',
  sheet: 'Unit Price Latest',
  specCol: 'Specification (รหัสรุ่น)',
  priceCol: 'Unit Price (Latest)'
};
SPARE_APP_CONFIG.sessionDurationMs = SPARE_APP_CONFIG.sessionDurationMs || (7 * 24 * 60 * 60 * 1000);
SPARE_APP_CONFIG.sessionRefreshThresholdMs = SPARE_APP_CONFIG.sessionRefreshThresholdMs || (24 * 60 * 60 * 1000);
var SESSION_PROPERTY_PREFIX = 'spare_session::';
var LOG_HEADERS = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After', 'Reason', 'Reason Remark', 'Machine'];
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
// เครื่องจักรของแต่ละไลน์ (master list ให้ Admin กรอกเอง) — อะไหล่แต่ละชิ้นผูกได้หลายเครื่องจักร
// โดยเก็บชื่อเครื่องจักรแบบ comma-separated ไว้ในคอลัมน์ "Machines" ของชีตอะไหล่แต่ละไลน์
var MACHINE_HEADERS = ['Machine ID', 'Line', 'Machine Name', 'Active', 'Created By', 'Created At', 'Updated By', 'Updated At'];
// Purchase Request (PR) schema — เก็บ qty_requested (ล็อก) แยกจาก qty_approved (หัวหน้าแก้ได้)
var PR_HEADER_HEADERS = ['pr_id', 'status', 'created_by', 'created_at', 'line', 'dept', 'item_count', 'total_amount', 'approved_by', 'approved_at', 'reject_reason', 'updated_at', 'assigned_to'];
var PR_LINE_HEADERS = ['pr_id', 'line_no', 'part_no', 'part_name', 'model', 'brand', 'category', 'unit', 'qty_requested', 'qty_approved', 'unit_price', 'remark', 'image_url'];
var PR_AUDIT_HEADERS = ['timestamp', 'pr_id', 'action', 'actor', 'line', 'old_qty', 'new_qty', 'detail'];
var PR_STATUSES = ['DRAFT', 'PENDING', 'APPROVED', 'REJECTED'];
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

// รายการ getSparePartsLite cache ไว้ 300 วิ (ดู doGet) — ต้องล้าง cache ของชีตนั้นทุกครั้งที่มีการ
// แก้ไข/ลบ/ทำรายการเบิก-รับ ไม่งั้นค่าที่เพิ่งบันทึกจะไม่ขึ้นในหน้ารายการ/Detail จนกว่า cache จะหมดอายุเอง
function invalidateSparePartsLiteCache(sheetName) {
  var name = String(sheetName || '').trim();
  if (!name) return;
  try {
    CacheService.getScriptCache().remove('spare_parts_lite::' + name);
  } catch (err) {
    Logger.log('invalidateSparePartsLiteCache warning [' + name + ']: ' + (err && err.message ? err.message : err));
  }
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
  ensureUsersLineColumn(usersSheet);
  return usersSheet;
}

// เพิ่มคอลัมน์ 'line' ต่อท้าย Users แบบไม่ทำลายข้อมูลเดิม (ใช้ระบุไลน์ของหัวหน้า)
// คืนค่า index แบบ 0-based ของคอลัมน์ line
function ensureUsersLineColumn(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf('line');
  if (idx !== -1) return idx;
  sheet.getRange(1, lastCol + 1).setValue('line');
  return lastCol;
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
// แปลง config เป็น object มาตรฐานเสมอ (รองรับทั้งแบบเก่าที่เป็น string id และแบบใหม่ที่เป็น object)
function normalizeProductionSourceConfig(raw) {
  if (!raw) return null;
  if (typeof raw === 'string') return { id: raw, sheet: 'ProductionLog', dateCol: 'Date', qtyCol: 'ActualQty', modelCol: '' };
  return {
    id: raw.id || '',
    sheet: raw.sheet || 'ProductionLog',
    dateCol: raw.dateCol || 'Date',
    qtyCol: raw.qtyCol || 'ActualQty',
    modelCol: raw.modelCol || ''
  };
}

function getExternalProductionVolumeForLine(line) {
  var cfg = normalizeProductionSourceConfig((SPARE_APP_CONFIG.productionLogSources || {})[line]);
  if (!cfg || !cfg.id) return null;
  var cacheKey = 'prod_volume_external::' + line;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (cacheReadErr) {
    Logger.log('getExternalProductionVolumeForLine cache read warning: ' + (cacheReadErr && cacheReadErr.message ? cacheReadErr.message : cacheReadErr));
  }
  try {
    var sourceSheet = SpreadsheetApp.openById(cfg.id).getSheetByName(cfg.sheet);
    if (!sourceSheet) return null;
    var values = sourceSheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    var headers = values[0];
    // จับคู่คอลัมน์แบบตรงตัว (lowercased+trim) ไม่ใช่ substring — กัน "Actual" ไปชน "Actual Scan"/
    // "Actual complete date" และ "Actual complete date" ไปชน "Actual complete date 0"
    var wantDate = String(cfg.dateCol).trim().toLowerCase();
    var wantQty = String(cfg.qtyCol).trim().toLowerCase();
    var wantModel = cfg.modelCol ? String(cfg.modelCol).trim().toLowerCase() : '';
    var dateCol = -1, qtyCol = -1, modelCol = -1;
    for (var i = 0; i < headers.length; i += 1) {
      var h = String(headers[i] || '').trim().toLowerCase();
      if (h === wantDate && dateCol === -1) dateCol = i;
      if (h === wantQty && qtyCol === -1) qtyCol = i;
      if (wantModel && h === wantModel && modelCol === -1) modelCol = i;
    }
    if (dateCol === -1 || qtyCol === -1) return null;
    var byMonth = {}, byMonthModel = {};
    for (var r = 1; r < values.length; r += 1) {
      var dateVal = values[r][dateCol];
      var qty = Number(values[r][qtyCol] || 0);
      if (!dateVal || !isFinite(qty) || qty <= 0) continue;
      var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
      if (isNaN(d.getTime())) continue;
      // ข้ามงานที่ยังไม่ได้บันทึกวันเสร็จ (Actual complete date ว่าง/เป็น 0 → Excel epoch ปี 1899)
      if (d.getFullYear() < 2000) continue;
      var monthKey = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
      byMonth[monthKey] = (byMonth[monthKey] || 0) + qty;
      if (modelCol !== -1) {
        var model = String(values[r][modelCol] || '').trim();
        if (model) {
          if (!byMonthModel[monthKey]) byMonthModel[monthKey] = {};
          byMonthModel[monthKey][model] = (byMonthModel[monthKey][model] || 0) + qty;
        }
      }
    }
    var result = Object.keys(byMonth).map(function(m) {
      var row = { month: m, actual_qty: byMonth[m] };
      // แนบ breakdown รายรุ่น (สำหรับไลน์ที่ราคาต่างกันตาม Model เช่น H9) ไว้คำนวณมูลค่าผลิตต่อรุ่น
      if (byMonthModel[m]) row.by_model = byMonthModel[m];
      return row;
    });
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

// อ่านชีตราคาต่อชิ้น → { specUpper: { spec, price } } (cache 30 นาที) — ชื่อ spec เก็บ uppercase
// ไว้ทำ case-insensitive lookup แต่คง spec เดิมไว้แสดงผล
function getUnitPriceMap() {
  var cfg = SPARE_APP_CONFIG.unitPriceSource;
  if (!cfg || !cfg.id) return null;
  var cacheKey = 'unit_price_map::' + cfg.id;
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { Logger.log('getUnitPriceMap cache read warning: ' + (e && e.message ? e.message : e)); }
  try {
    var sheet = SpreadsheetApp.openById(cfg.id).getSheetByName(cfg.sheet);
    if (!sheet) return null;
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return {};
    var headers = values[0];
    var wantSpec = String(cfg.specCol).trim().toLowerCase();
    var wantPrice = String(cfg.priceCol).trim().toLowerCase();
    var specCol = -1, priceCol = -1;
    for (var i = 0; i < headers.length; i += 1) {
      var h = String(headers[i] || '').trim().toLowerCase();
      if (h === wantSpec && specCol === -1) specCol = i;
      if (h === wantPrice && priceCol === -1) priceCol = i;
    }
    if (specCol === -1 || priceCol === -1) return null;
    var map = {};
    for (var r = 1; r < values.length; r += 1) {
      var spec = String(values[r][specCol] || '').trim();
      var price = Number(values[r][priceCol] || 0);
      if (!spec) continue;
      map[spec.toUpperCase()] = { spec: spec, price: price };
    }
    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(map), 1800); }
    catch (e2) { Logger.log('getUnitPriceMap cache write warning: ' + (e2 && e2.message ? e2.message : e2)); }
    return map;
  } catch (err) {
    Logger.log('getUnitPriceMap error: ' + (err && err.message ? err.message : err));
    return null;
  }
}

// จับคู่ Model จากชีตผลิตกับ spec ในชีตราคา — ทดสอบ transform แบบ deterministic ตามลำดับ
// (verify แล้วกับข้อมูลจริง: Lug 2/2, Coil 4/4, H9 130/135 โดยไม่มีการเดามั่ว):
//   exact → ตัด " (ขนาด)" ท้าย (Coil) → เติม "-S"/"-T-S"/"T-S"/"-T" (H9) → ตัด "-S" → สลับ 0↔O → ตัด "*"
// คืน { price, matched } หรือ null ถ้าไม่เจอ (ไม่เดา เพื่อไม่ให้ราคาผิด)
function matchModelPrice(model, priceUpper) {
  var s0 = String(model || '').trim();
  if (!s0 || !priceUpper) return null;
  var stripStar = s0.replace(/\*+$/, '').trim();
  var bases = [s0];
  if (stripStar !== s0) bases.push(stripStar);
  bases.slice().forEach(function (b) {
    if (/0/.test(b)) bases.push(b.replace(/0/g, 'O'));
    if (/O/.test(b)) bases.push(b.replace(/O/g, '0'));
  });
  var cand = [];
  bases.forEach(function (b) {
    var noSize = b.replace(/\s+\([^)]*\)\s*$/, '').trim();
    [b, noSize].forEach(function (x) {
      cand.push(x); cand.push(x + '-S'); cand.push(x + '-T-S'); cand.push(x + 'T-S'); cand.push(x + '-T');
      if (/-S$/i.test(x)) cand.push(x.replace(/-S$/i, ''));
    });
  });
  var seen = {};
  for (var i = 0; i < cand.length; i += 1) {
    var key = cand[i]; if (!key || seen[key]) continue; seen[key] = 1;
    var hit = priceUpper[key.toUpperCase()];
    if (hit) return { price: hit.price, matched: hit.spec };
  }
  return null;
}

function getProductionVolume(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var values = getProductionVolumeSheet().getDataRange().getValues();
  var manual = values.length <= 1 ? [] : values.slice(1).filter(function(r) { return String(r[0] || '').trim(); }).map(function(r) {
    return { month: String(r[0] || ''), line: String(r[1] || ''), actual_qty: Number(r[2] || 0), updated_by: String(r[3] || ''), updated_at: String(r[4] || ''), source: 'manual' };
  });
  var priceUpper = getUnitPriceMap(); // อาจเป็น null ถ้าเข้าถึงชีตราคาไม่ได้
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
      // คำนวณมูลค่าผลิต = ผลรวม (ยอดแต่ละรุ่น × ราคาที่จับคู่ได้) จาก breakdown รายรุ่น
      if (row.by_model && priceUpper) {
        var value = 0, unpricedQty = 0, unpriced = {};
        Object.keys(row.by_model).forEach(function(model) {
          var q = Number(row.by_model[model] || 0);
          var hit = matchModelPrice(model, priceUpper);
          if (hit) value += q * hit.price;
          else { unpricedQty += q; unpriced[model] = q; }
        });
        entry.production_value = value;
        entry.unpriced_qty = unpricedQty;
        if (Object.keys(unpriced).length) entry.unpriced_models = unpriced;
      }
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

// =============================
// MACHINES (master list ต่อไลน์ + ผูกกับอะไหล่แบบหลายเครื่อง)
// =============================
function getMachinesSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.machinesSheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(MACHINE_HEADERS);
  return sheet;
}

function rowToMachine(r) {
  return {
    machine_id: String(r[0] || ''),
    line: String(r[1] || ''),
    machine_name: String(r[2] || ''),
    active: String(r[3]).toUpperCase() !== 'FALSE',
    created_by: String(r[4] || ''),
    created_at: String(r[5] || ''),
    updated_by: String(r[6] || ''),
    updated_at: String(r[7] || '')
  };
}

function getMachines(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var values = getMachinesSheet().getDataRange().getValues();
  var list = values.length <= 1 ? [] : values.slice(1)
    .filter(function(r) { return String(r[0] || '').trim(); })
    .map(rowToMachine);
  var line = String(payload.line || '').trim();
  if (line) list = list.filter(function(m) { return m.line === line; });
  if (String(payload.includeInactive || '') !== '1') list = list.filter(function(m) { return m.active; });
  return list;
}

function addMachine(payload) {
  var user = requireAdminUser({ authToken: payload.authToken });
  var line = String(payload.line || '').trim();
  var name = String(payload.machine_name || '').trim();
  if (!line) throw new Error('กรุณาระบุ Line');
  if (!name) throw new Error('กรุณาระบุชื่อเครื่องจักร');
  var sheet = getMachinesSheet();
  var values = sheet.getDataRange().getValues();
  var dup = values.slice(1).some(function(r) {
    return String(r[1] || '').trim() === line && String(r[2] || '').trim().toLowerCase() === name.toLowerCase();
  });
  if (dup) throw new Error('มีชื่อเครื่องจักรนี้อยู่แล้วในไลน์นี้');
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var machineId = 'MC-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([machineId, line, name, true, user.username, now, user.username, now]);
  return { status: 'success', machine_id: machineId };
}

function updateMachine(payload) {
  var user = requireAdminUser({ authToken: payload.authToken });
  var machineId = String(payload.machine_id || '').trim();
  if (!machineId) throw new Error('ไม่พบรหัสเครื่องจักร');
  var sheet = getMachinesSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') === machineId) {
      var name = payload.machine_name !== undefined ? String(payload.machine_name).trim() : String(values[i][2] || '');
      if (!name) throw new Error('กรุณาระบุชื่อเครื่องจักร');
      var active = payload.active !== undefined
        ? (String(payload.active) === '1' || String(payload.active).toLowerCase() === 'true')
        : (String(values[i][3]).toUpperCase() !== 'FALSE');
      var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(i + 1, 3).setValue(name);
      sheet.getRange(i + 1, 4).setValue(active);
      sheet.getRange(i + 1, 7, 1, 2).setValues([[user.username, now]]);
      return { status: 'success' };
    }
  }
  throw new Error('ไม่พบเครื่องจักรนี้');
}

function deleteMachine(payload) {
  requireAdminUser({ authToken: payload.authToken });
  var machineId = String(payload.machine_id || '').trim();
  if (!machineId) throw new Error('ไม่พบรหัสเครื่องจักร');
  var sheet = getMachinesSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '') === machineId) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  throw new Error('ไม่พบเครื่องจักรนี้');
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

// =============================
// PURCHASE REQUEST (PR) — ประตูอนุมัติก่อนปริ้น + audit
// เก็บ 3 ชีทแยก: PRHeaders / PRLines / PRAudit
// กติกา: qty_requested ล็อก ห้ามเขียนทับ, qty_approved เท่านั้นที่หัวหน้าแก้ได้
// สถานะ: DRAFT -> PENDING -> APPROVED / REJECTED (ปริ้นได้เฉพาะ APPROVED)
// =============================
function prStr(v) { return v === undefined || v === null ? '' : String(v); }

function getPrSheetWithHeaders(sheetName, headers) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(spreadsheet, sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    // เพิ่มคอลัมน์ที่ขาดแบบ additive (ไม่ทำลายของเดิม) เหมือน getOrderRequestSheet
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach(function(header) {
      if (currentHeaders.indexOf(header) !== -1) return;
      var nextColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextColumn).setValue(header);
      currentHeaders.push(header);
    });
  }
  return sheet;
}
function getPrHeaderSheet() { return getPrSheetWithHeaders(SPARE_APP_CONFIG.prHeaderSheetName, PR_HEADER_HEADERS); }
function getPrLinesSheet() { return getPrSheetWithHeaders(SPARE_APP_CONFIG.prLinesSheetName, PR_LINE_HEADERS); }
function getPrAuditSheet() { return getPrSheetWithHeaders(SPARE_APP_CONFIG.prAuditSheetName, PR_AUDIT_HEADERS); }

function prIndexMap(headerRow) {
  var idx = {};
  (headerRow || []).forEach(function(h, i) { idx[String(h)] = i; });
  return idx;
}

function appendPrAudit(prId, action, actor, line, oldQty, newQty, detail) {
  try {
    getPrAuditSheet().appendRow([
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
      prStr(prId), prStr(action), prStr(actor), prStr(line),
      (oldQty === undefined || oldQty === null ? '' : oldQty),
      (newQty === undefined || newQty === null ? '' : newQty),
      prStr(detail)
    ]);
  } catch (auditErr) {
    Logger.log('appendPrAudit warning [' + prId + '/' + action + ']: ' + (auditErr && auditErr.message ? auditErr.message : auditErr));
  }
}

function parsePrLinesInput(raw) {
  if (!raw) return [];
  var parsed = raw;
  if (typeof raw === 'string') {
    try { parsed = JSON.parse(raw); } catch (err) { throw new Error('รูปแบบ lines_json ไม่ถูกต้อง'); }
  }
  return Array.isArray(parsed) ? parsed : [];
}

// หัวหน้า (leader/supervisor) เห็น+อนุมัติได้เฉพาะไลน์ตัวเอง; admin/manager (pr_view_all) ข้ามไลน์ได้
// หัวหน้าที่ยังไม่ถูกกำหนดไลน์ (user.line ว่าง) จะไม่ถูกจำกัด เพื่อไม่ให้ระบบล็อกตัวเอง
function prUserCanAccessLine(user, prLine) {
  if (user.permissions && user.permissions.pr_view_all) return true;
  var myLine = String(user.line || '').trim();
  if (!myLine) return true;
  return String(prLine || '').trim() === myLine;
}

function prHeaderRowToCard(hIdx, row) {
  return {
    pr_id: prStr(row[hIdx.pr_id]),
    status: prStr(row[hIdx.status]),
    created_by: prStr(row[hIdx.created_by]),
    created_at: prStr(row[hIdx.created_at]),
    line: prStr(row[hIdx.line]),
    dept: prStr(row[hIdx.dept]),
    item_count: Number(row[hIdx.item_count] || 0),
    total_amount: Number(row[hIdx.total_amount] || 0),
    approved_by: prStr(row[hIdx.approved_by]),
    approved_at: prStr(row[hIdx.approved_at]),
    reject_reason: prStr(row[hIdx.reject_reason]),
    updated_at: prStr(row[hIdx.updated_at]),
    assigned_to: hIdx.assigned_to !== undefined ? prStr(row[hIdx.assigned_to]) : ''
  };
}

// อ่านการ์ด PR ตามสถานะที่ต้องการ (filter ตาม role/line ของผู้ใช้)
function listPrCardsForUser(user, statuses) {
  var wanted = {};
  (statuses || []).forEach(function(s) { wanted[s] = true; });
  var sheet = getPrHeaderSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var hIdx = prIndexMap(data[0]);
  var cards = [];
  for (var i = 1; i < data.length; i += 1) {
    var row = data[i];
    var status = prStr(row[hIdx.status]);
    if (!wanted[status]) continue;
    if (!prUserCanAccessLine(user, prStr(row[hIdx.line]))) continue;
    // PR ที่ระบุผู้อนุมัติไว้ — โชว์เฉพาะผู้ถูกมอบหมาย / คนเห็นทุกไลน์ / คนสร้างเอง (ติดตามใบตัวเอง)
    var assignedTo = hIdx.assigned_to !== undefined ? prStr(row[hIdx.assigned_to]).trim() : '';
    if (assignedTo && assignedTo !== user.username && prStr(row[hIdx.created_by]) !== user.username && !(user.permissions && user.permissions.pr_view_all)) continue;
    cards.push(prHeaderRowToCard(hIdx, row));
  }
  cards.sort(function(a, b) {
    return String(b.created_at || '').localeCompare(String(a.created_at || ''));
  });
  return cards;
}

function findPrHeaderRow(data, hIdx, prId) {
  for (var i = 1; i < data.length; i += 1) {
    if (String(data[i][hIdx.pr_id]) === String(prId)) return i;
  }
  return -1;
}

// สร้าง PR ใหม่ -> สถานะ PENDING (ปริ้นไม่ได้จนกว่าจะอนุมัติ) แบบ atomic
function createPR(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return createPRUnlocked(payload); } finally { lock.releaseLock(); }
}
function createPRUnlocked(payload) {
  var user = requirePermission({ authToken: payload.authToken }, 'pr_create');
  var lines = parsePrLinesInput(payload.lines_json || payload.lines);
  if (!lines.length) throw new Error('ต้องมีรายการอย่างน้อย 1 บรรทัดใน PR');

  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var prId = String(payload.pr_id || '').trim() ||
    ('PR-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100));
  var headerLine = String(payload.line || user.line || '').trim();
  var dept = String(payload.dept || '').trim();

  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  if (findPrHeaderRow(hData, hIdx, prId) !== -1) throw new Error('pr_id ซ้ำในระบบ: ' + prId);

  var linesSheet = getPrLinesSheet();
  var lHeaderRow = linesSheet.getRange(1, 1, 1, linesSheet.getLastColumn()).getValues()[0];
  var lIdx = prIndexMap(lHeaderRow);
  var totalAmount = 0;
  var lineRows = lines.map(function(ln, i) {
    var qtyReq = Number(ln.qty_requested !== undefined ? ln.qty_requested : ln.qty) || 0;
    if (qtyReq <= 0) throw new Error('qty_requested ต้องมากกว่า 0 (บรรทัด ' + (i + 1) + ')');
    var price = Number(ln.unit_price || 0) || 0;
    totalAmount += qtyReq * price;
    var rowArr = new Array(lHeaderRow.length).fill('');
    rowArr[lIdx.pr_id] = prId;
    rowArr[lIdx.line_no] = i + 1;
    rowArr[lIdx.part_no] = prStr(ln.part_no || ln.no);
    rowArr[lIdx.part_name] = prStr(ln.part_name || ln.name);
    rowArr[lIdx.model] = prStr(ln.model);
    rowArr[lIdx.brand] = prStr(ln.brand);
    rowArr[lIdx.category] = prStr(ln.category);
    rowArr[lIdx.unit] = prStr(ln.unit || 'PCS');
    rowArr[lIdx.qty_requested] = qtyReq;       // ล็อก — ห้ามแก้หลังจากนี้
    rowArr[lIdx.qty_approved] = qtyReq;          // เริ่มต้นเท่ากับที่ขอ ให้หัวหน้าปรับลดได้
    rowArr[lIdx.unit_price] = price;
    rowArr[lIdx.remark] = prStr(ln.remark);
    // สแนปช็อตรูปไว้ตอนสร้าง PR — ผู้อนุมัติจะได้เห็นรูปประกอบตอนตรวจ/แก้ยอด โดยไม่ต้องพึ่ง
    // cache ฝั่ง client ของไลน์นั้น (ซึ่งอาจไม่มีถ้าไม่เคยเปิดไลน์นี้มาก่อน) หรือ master ที่อาจถูกแก้ทีหลัง
    if (lIdx.image_url !== undefined) rowArr[lIdx.image_url] = prStr(ln.image_url || ln.image || '');
    return rowArr;
  });
  linesSheet.getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, lHeaderRow.length).setValues(lineRows);

  // ผู้อนุมัติที่ถูกระบุ (optional) — ถ้าระบุมา ต้องเป็น user ที่มีสิทธิ์อนุมัติไลน์นี้จริง
  // PR ที่ระบุคนไว้จะโชว์ใน Inbox เฉพาะคนนั้น (+ คนเห็นทุกไลน์ + คนสร้างเอง) — เป็นการ routing
  // ไม่ใช่ lock: คนที่มีสิทธิ์คนอื่นยังอนุมัติแทนได้ถ้าเปิด PR ตรงๆ (กันใบค้างตอนคนถูกระบุไม่อยู่)
  var assignedTo = String(payload.assign_to || '').trim();
  if (assignedTo) {
    var assignee = findUserByUsername(assignedTo);
    if (!assignee || !assignee.isActive || !assignee.permissions || !assignee.permissions.pr_approve || !prUserCanAccessLine(assignee, headerLine)) {
      throw new Error('ผู้อนุมัติที่ระบุ (' + assignedTo + ') ไม่มีสิทธิ์อนุมัติ PR ของไลน์นี้');
    }
  }

  var headerRowArr = new Array(hData[0].length).fill('');
  headerRowArr[hIdx.pr_id] = prId;
  headerRowArr[hIdx.status] = 'PENDING';
  headerRowArr[hIdx.created_by] = user.username;
  headerRowArr[hIdx.created_at] = now;
  headerRowArr[hIdx.line] = headerLine;
  headerRowArr[hIdx.dept] = dept;
  headerRowArr[hIdx.item_count] = lines.length;
  headerRowArr[hIdx.total_amount] = totalAmount;
  headerRowArr[hIdx.updated_at] = now;
  if (hIdx.assigned_to !== undefined) headerRowArr[hIdx.assigned_to] = assignedTo;
  headerSheet.appendRow(headerRowArr);

  appendPrAudit(prId, 'CREATE', user.username, headerLine, '', '', lines.length + ' รายการ' + (assignedTo ? ' · มอบหมายให้ ' + assignedTo : ''));
  // บันทึกร่องรอยการ Bypass งบ Spare part (ถ้าผู้ใช้ยืนยันเปิด PR ทั้งที่เกินงบ) — เก็บว่าใคร/
  // เกินไลน์ไหนเท่าไหร่/เหตุผล เพื่อให้หัวหน้าเห็นตอนอนุมัติและตรวจสอบย้อนหลังได้
  if (toBoolean(payload.budget_bypass, false)) {
    var bypassReason = String(payload.bypass_reason || '').trim() || '(ไม่ระบุเหตุผล)';
    var bypassDetail = '⚠️ เปิด PR ทั้งที่เกินงบ Spare part — เหตุผล: ' + bypassReason;
    var overInfo = String(payload.budget_over_detail || '').trim();
    if (overInfo) bypassDetail += ' | ' + overInfo;
    appendPrAudit(prId, 'BUDGET_BYPASS', user.username, headerLine, '', '', bypassDetail);
  }
  return { status: 'success', pr_id: prId, pr_status: 'PENDING', item_count: lines.length, total_amount: totalAmount };
}

// ดึง PR ใบเดียว (header + lines) สำหรับหน้าตรวจ+แก้ยอด
function getPRForApproval(payload) {
  var user = requirePermission({ authToken: payload.authToken }, 'view');
  var prId = String(payload.pr_id || '').trim();
  if (!prId) throw new Error('ต้องระบุ pr_id');

  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  var hRow = findPrHeaderRow(hData, hIdx, prId);
  if (hRow === -1) throw new Error('ไม่พบ PR: ' + prId);
  var header = prHeaderRowToCard(hIdx, hData[hRow]);
  if (!prUserCanAccessLine(user, header.line)) throw new Error('ไม่มีสิทธิ์ดู PR ของไลน์นี้');

  var linesSheet = getPrLinesSheet();
  var lData = linesSheet.getDataRange().getValues();
  var lIdx = prIndexMap(lData[0]);
  var lines = [];
  for (var i = 1; i < lData.length; i += 1) {
    if (String(lData[i][lIdx.pr_id]) !== prId) continue;
    var r = lData[i];
    lines.push({
      line_no: Number(r[lIdx.line_no] || 0),
      part_no: prStr(r[lIdx.part_no]),
      part_name: prStr(r[lIdx.part_name]),
      model: prStr(r[lIdx.model]),
      brand: prStr(r[lIdx.brand]),
      category: prStr(r[lIdx.category]),
      unit: prStr(r[lIdx.unit]),
      qty_requested: Number(r[lIdx.qty_requested] || 0),
      qty_approved: Number(r[lIdx.qty_approved] || 0),
      unit_price: Number(r[lIdx.unit_price] || 0),
      remark: prStr(r[lIdx.remark]),
      image_url: lIdx.image_url !== undefined ? prStr(r[lIdx.image_url]) : ''
    });
  }
  lines.sort(function(a, b) { return a.line_no - b.line_no; });
  return { status: 'success', header: header, lines: lines };
}

// รายชื่อผู้มีสิทธิ์อนุมัติ PR ของไลน์ที่ระบุ — โชว์ใต้ปุ่มส่งอนุมัติในหน้า PR builder
// ให้คนส่งรู้ว่าใบนี้จะเข้า Inbox ของใครบ้าง จะได้ตามถูกคน (ระบบไม่มี notification ช่องทางอื่น)
function getPrApprovers(payload) {
  requirePermission({ authToken: payload.authToken }, 'pr_create');
  var line = String(payload.line || '').trim();
  var approvers = getAllUsers().filter(function(u) {
    return u.isActive && u.permissions && u.permissions.pr_approve && prUserCanAccessLine(u, line);
  }).map(function(u) {
    return {
      username: u.username,
      role: u.role,
      // ไลน์ว่าง หรือมี pr_view_all = เห็นทุกไลน์
      line: (u.permissions.pr_view_all || !u.line) ? '' : u.line
    };
  });
  return { status: 'success', line: line, approvers: approvers };
}

// อนุมัติ PR — เขียน status + qty_approved ทุกบรรทัด + audit เป็น operation เดียว (atomic)
function approvePR(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return approvePRUnlocked(payload); } finally { lock.releaseLock(); }
}
function approvePRUnlocked(payload) {
  var user = requirePermission({ authToken: payload.authToken }, 'pr_approve');
  var prId = String(payload.pr_id || '').trim();
  if (!prId) throw new Error('ต้องระบุ pr_id');

  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  var hRow = findPrHeaderRow(hData, hIdx, prId);
  if (hRow === -1) throw new Error('ไม่พบ PR: ' + prId);
  var prLine = prStr(hData[hRow][hIdx.line]);
  if (!prUserCanAccessLine(user, prLine)) throw new Error('ไม่มีสิทธิ์อนุมัติ PR ของไลน์นี้');

  // กันกดอนุมัติซ้ำ / กันสถานะไม่ใช่ PENDING
  var curStatus = prStr(hData[hRow][hIdx.status]);
  if (curStatus !== 'PENDING') {
    throw new Error('PR นี้ไม่อยู่สถานะรออนุมัติ (สถานะปัจจุบัน: ' + curStatus + ') — อาจถูกดำเนินการไปแล้ว');
  }

  // map ยอดที่หัวหน้าปรับ: { line_no: qty_approved }
  var edits = {};
  parsePrLinesInput(payload.lines_json || payload.lines).forEach(function(ln) {
    if (ln && ln.line_no !== undefined) edits[String(ln.line_no)] = ln.qty_approved;
  });

  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  var linesSheet = getPrLinesSheet();
  var lData = linesSheet.getDataRange().getValues();
  var lIdx = prIndexMap(lData[0]);
  var totalAmount = 0;
  var itemCount = 0;
  for (var i = 1; i < lData.length; i += 1) {
    if (String(lData[i][lIdx.pr_id]) !== prId) continue;
    itemCount += 1;
    var lineNo = String(lData[i][lIdx.line_no]);
    var oldApproved = Number(lData[i][lIdx.qty_approved] || 0);
    var newApproved = oldApproved;
    if (edits.hasOwnProperty(lineNo)) {
      newApproved = Number(edits[lineNo]);
      if (!isFinite(newApproved) || newApproved < 0) throw new Error('qty_approved ต้องเป็นตัวเลข >= 0 (บรรทัด ' + lineNo + ')');
    }
    // เขียนเฉพาะ qty_approved — ไม่แตะ qty_requested เด็ดขาด
    linesSheet.getRange(i + 1, lIdx.qty_approved + 1).setValue(newApproved);
    if (newApproved !== oldApproved) {
      appendPrAudit(prId, 'EDIT_QTY', user.username, prLine, oldApproved, newApproved, 'line ' + lineNo);
    }
    var price = Number(lData[i][lIdx.unit_price] || 0) || 0;
    totalAmount += newApproved * price;

    // บันทึก Purchase History ด้วยยอด qty_approved (ย้ายจากตอนปริ้นมาที่ตอนอนุมัติ)
    if (newApproved > 0) {
      try {
        upsertPurchaseHistoryRecord({
          request_id: prId + '-L' + lineNo, history_id: 'PH-' + prId + '-L' + lineNo, source: 'PR Report',
          date: new Date(), line: prLine, part_id: prStr(lData[i][lIdx.part_no]), part_name: prStr(lData[i][lIdx.part_name]),
          brand: prStr(lData[i][lIdx.brand]), model: prStr(lData[i][lIdx.model]), qty_ordered: newApproved,
          unit: prStr(lData[i][lIdx.unit]), unit_price: price, currency: 'THB', status: 'PR Created',
          requested_by: prStr(hData[hRow][hIdx.created_by]), updated_by: user.username, remark: prStr(lData[i][lIdx.remark])
        });
      } catch (historyErr) {
        Logger.log('approvePR PurchaseHistory warning [' + prId + '/L' + lineNo + ']: ' + (historyErr && historyErr.message ? historyErr.message : historyErr));
      }
    }
  }

  headerSheet.getRange(hRow + 1, hIdx.status + 1).setValue('APPROVED');
  headerSheet.getRange(hRow + 1, hIdx.approved_by + 1).setValue(user.username);
  headerSheet.getRange(hRow + 1, hIdx.approved_at + 1).setValue(now);
  headerSheet.getRange(hRow + 1, hIdx.item_count + 1).setValue(itemCount);
  headerSheet.getRange(hRow + 1, hIdx.total_amount + 1).setValue(totalAmount);
  headerSheet.getRange(hRow + 1, hIdx.updated_at + 1).setValue(now);
  appendPrAudit(prId, 'APPROVE', user.username, prLine, '', '', itemCount + ' รายการ');

  return { status: 'success', pr_id: prId, pr_status: 'APPROVED', item_count: itemCount, total_amount: totalAmount };
}

// ตีกลับ PR พร้อมเหตุผล -> REJECTED (คนสั่งไปแก้เอง)
function rejectPR(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return rejectPRUnlocked(payload); } finally { lock.releaseLock(); }
}
function rejectPRUnlocked(payload) {
  var user = requirePermission({ authToken: payload.authToken }, 'pr_approve');
  var prId = String(payload.pr_id || '').trim();
  if (!prId) throw new Error('ต้องระบุ pr_id');
  var reason = String(payload.reject_reason || payload.reason || '').trim();
  if (!reason) throw new Error('กรุณากรอกเหตุผลการตีกลับ');

  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  var hRow = findPrHeaderRow(hData, hIdx, prId);
  if (hRow === -1) throw new Error('ไม่พบ PR: ' + prId);
  var prLine = prStr(hData[hRow][hIdx.line]);
  if (!prUserCanAccessLine(user, prLine)) throw new Error('ไม่มีสิทธิ์ตีกลับ PR ของไลน์นี้');
  var curStatus = prStr(hData[hRow][hIdx.status]);
  if (curStatus !== 'PENDING') {
    throw new Error('PR นี้ไม่อยู่สถานะรออนุมัติ (สถานะปัจจุบัน: ' + curStatus + ')');
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  headerSheet.getRange(hRow + 1, hIdx.status + 1).setValue('REJECTED');
  headerSheet.getRange(hRow + 1, hIdx.reject_reason + 1).setValue(reason);
  headerSheet.getRange(hRow + 1, hIdx.updated_at + 1).setValue(now);
  appendPrAudit(prId, 'REJECT', user.username, prLine, '', '', reason);
  return { status: 'success', pr_id: prId, pr_status: 'REJECTED' };
}

// endpoint เบา ให้ปุ่มปริ้นเช็คสถานะสดก่อนเปิดหน้าปริ้น (กันเคสสถานะถูกแก้กลับ)
function getPRStatus(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var prId = String(payload.pr_id || '').trim();
  if (!prId) throw new Error('ต้องระบุ pr_id');
  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  var hRow = findPrHeaderRow(hData, hIdx, prId);
  if (hRow === -1) return { status: 'success', found: false, pr_id: prId, pr_status: '' };
  return { status: 'success', found: true, pr_id: prId, pr_status: prStr(hData[hRow][hIdx.status]) };
}

// Inbox rollup — คืน count ต่อ tab + การ์ด PR (filter ตาม role/line)
// tab อื่นเป็น placeholder (count 0) ค่อยเติมทีหลังตามขอบเขตงาน
function getInbox(payload) {
  var user = requirePermission({ authToken: payload.authToken }, 'view');
  var canViewPr = !!(user.permissions && (user.permissions.pr_view_all || user.permissions.pr_view_own || user.permissions.pr_approve));
  var prPending = canViewPr ? listPrCardsForUser(user, ['PENDING']) : [];
  var prApproved = canViewPr ? listPrCardsForUser(user, ['APPROVED']) : [];
  var prRejected = canViewPr ? listPrCardsForUser(user, ['REJECTED']) : [];
  var anomalies = [];
  try {
    anomalies = getIssueAnomalies({ authToken: payload.authToken }).anomalies || [];
  } catch (anomalyErr) {
    Logger.log('getInbox anomaly warning: ' + (anomalyErr && anomalyErr.message ? anomalyErr.message : anomalyErr));
  }
  // แจ้งเตือนการ Bypass งบ Spare part — เห็นเฉพาะ role ระดับสูง (ดู PR ได้ทุกไลน์)
  // เป็นข้อมูล informational ให้ผู้บริหารรู้ว่ามี PR ที่เปิดทั้งที่เกินงบ
  var canSeeBypass = !!(user.permissions && (user.permissions.pr_view_all || user.permissions.manage_users));
  var budgetBypasses = [];
  if (canSeeBypass) {
    try { budgetBypasses = listBudgetBypasses(30); } catch (bpErr) {
      Logger.log('getInbox bypass warning: ' + (bpErr && bpErr.message ? bpErr.message : bpErr));
    }
  }
  // สต็อกต่ำกว่า Min — scan ชุดเดียวกับ Auto-PR พร้อมธงว่ารายการอยู่ใน PR ค้างอนุมัติแล้วหรือยัง
  // เรียงตามความหนัก (สัดส่วน stock/min น้อยสุดก่อน — ของหมดขึ้นบนสุด)
  var belowMin = [];
  try {
    var pendingPartKeys = listPendingPrPartKeys();
    belowMin = readAllStockItemsLean().filter(function(it) {
      return it.min > 0 && it.stock < it.min;
    }).map(function(it) {
      return {
        sheet: it.sheet, no: it.no, name: it.name, model: it.model, unit: it.unit,
        stock: it.stock, min: it.min,
        in_pending_pr: !!pendingPartKeys[partKeyOf(it.name, it.model)]
      };
    }).sort(function(a, b) { return (a.stock / a.min) - (b.stock / b.min); }).slice(0, 200);
  } catch (bmErr) {
    Logger.log('getInbox belowMin warning: ' + (bmErr && bmErr.message ? bmErr.message : bmErr));
  }

  // ของเข้ารอตรวจรับ — Purchase History ที่สั่งแล้วแต่ยังรับไม่ครบ (Ordered / Partial Received)
  var qcIncoming = [];
  try {
    var phValues = getPurchaseHistorySheet().getDataRange().getValues();
    for (var qi = 1; qi < phValues.length; qi += 1) {
      var ph = purchaseHistoryRowToObject(phValues[qi]);
      if (!ph.history_id || ph.deleted || ph.qty_ordered <= 0) continue;
      if (['Ordered', 'Partial Received'].indexOf(ph.status) === -1) continue;
      qcIncoming.push({
        history_id: ph.history_id, request_id: ph.request_id, line: ph.line,
        part_name: ph.part_name, brand: ph.brand, model: ph.model,
        qty_ordered: ph.qty_ordered, received_qty: ph.received_qty, unit: ph.unit,
        status: ph.status, ordered_date: ph.ordered_date, requested_date: ph.date
      });
    }
    qcIncoming.sort(function(a, b) { return String(b.ordered_date || b.requested_date || '').localeCompare(String(a.ordered_date || a.requested_date || '')); });
    qcIncoming = qcIncoming.slice(0, 200);
  } catch (qcErr) {
    Logger.log('getInbox qcIncoming warning: ' + (qcErr && qcErr.message ? qcErr.message : qcErr));
  }

  var tabs = {
    pr_pending: prPending.length,
    issue_anomaly: anomalies.length,
    budget_bypass: budgetBypasses.length,
    stock_below_min: belowMin.length,
    qc_incoming: qcIncoming.length
  };
  // badge นับเฉพาะงานที่ต้องจัดการ (PR รอ/เบิกผิดปกติ/Bypass) — สองแท็บหลังเป็นสถานะคงค้าง
  // ถ้านับรวมตัวเลขจะค้างสูงตลอดจนคนเลิกสนใจ badge
  var total = tabs.pr_pending + tabs.issue_anomaly + tabs.budget_bypass;
  return {
    status: 'success',
    role: user.role,
    line: String(user.line || ''),
    can_approve: !!(user.permissions && user.permissions.pr_approve),
    total: total,
    tabs: tabs,
    can_see_bypass: canSeeBypass,
    pr_pending: prPending,
    pr_approved: prApproved,
    pr_rejected: prRejected,
    anomalies: anomalies,
    budget_bypasses: budgetBypasses,
    below_min: belowMin,
    qc_incoming: qcIncoming
  };
}

// อ่านรายการ Bypass งบจาก PRAudit (action = BUDGET_BYPASS) ย้อนหลัง N วัน ใหม่สุดก่อน
function listBudgetBypasses(days) {
  var sheet = getPrAuditSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var idx = prIndexMap(data[0]);
  var cutoff = Date.now() - (days || 30) * 86400000;
  var out = [];
  for (var i = 1; i < data.length; i += 1) {
    var row = data[i];
    if (String(row[idx.action]) !== 'BUDGET_BYPASS') continue;
    var tsRaw = String(row[idx.timestamp] || '');
    var ts = new Date(tsRaw.replace(' ', 'T') + '+07:00').getTime();
    if (!isNaN(ts) && ts < cutoff) continue;
    out.push({
      timestamp: tsRaw,
      pr_id: prStr(row[idx.pr_id]),
      actor: prStr(row[idx.actor]),
      line: prStr(row[idx.line]),
      detail: prStr(row[idx.detail])
    });
  }
  out.sort(function(a, b) { return String(b.timestamp).localeCompare(String(a.timestamp)); });
  return out;
}

// ยอด PR ที่ยัง "รออนุมัติ" (PENDING) ต่อไลน์ต่อเดือน (เดือนจาก created_at) —
// ใช้บวกเข้า "ใช้ไปแล้ว" ในการคุมงบ กันเปิด PR เล็กๆ หลายใบที่รวมกันเกินงบ
function getPendingPrUsage(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var sheet = getPrHeaderSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'success', pending: [] };
  var hIdx = prIndexMap(data[0]);
  var out = [];
  for (var i = 1; i < data.length; i += 1) {
    var row = data[i];
    if (String(row[hIdx.status]) !== 'PENDING') continue;
    var createdAt = String(row[hIdx.created_at] || '');
    var month = createdAt.slice(0, 7); // yyyy-MM
    out.push({
      pr_id: prStr(row[hIdx.pr_id]),
      line: prStr(row[hIdx.line]),
      month: month,
      amount: Number(row[hIdx.total_amount] || 0)
    });
  }
  return { status: 'success', pending: out };
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
    request_order_edit: true, delete_logs: true,
    pr_create: true, pr_view_own: true, pr_view_all: true, pr_approve: true
  };
  if (normalized === 'leader') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false,
    request_order_edit: false, delete_logs: false,
    pr_create: true, pr_view_own: true, pr_view_all: false, pr_approve: true
  };
  return {
    view: true, transact: true, manage_items: false, delete_items: false,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false,
    request_order_create: true, request_order_view_own: true, request_order_view_all: false,
    request_order_approve: false, request_order_reject: false, request_order_convert_pr: false, request_order_close: false,
    request_order_edit: false, delete_logs: false,
    pr_create: true, pr_view_own: false, pr_view_all: false, pr_approve: false
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
  var lineCol = (data[0] || []).indexOf('line');
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
      line: lineCol !== -1 ? String(row[lineCol] || '').trim() : '',
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
    line: String(user.line || ''),
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
  var line = String(payload.line || '').trim();
  var lineProvided = payload.line !== undefined && payload.line !== null;

  var usersSheet = getUsersSheet();
  var lineCol = ensureUsersLineColumn(usersSheet); // 0-based
  var existing = findUserByUsername(username);
  if (existing) {
    if (String(payload.password || '') !== '') {
      if (password.length < 4) throw new Error('password ต้องมีอย่างน้อย 4 ตัวอักษร');
      usersSheet.getRange(existing.rowIndex, 2).setValue(hashPassword(password));
    }
    usersSheet.getRange(existing.rowIndex, 3).setValue(role);
    usersSheet.getRange(existing.rowIndex, 4).setValue(String(isActive));
    usersSheet.getRange(existing.rowIndex, 5).setValue(permissionsJson);
    if (lineProvided && lineCol !== -1) usersSheet.getRange(existing.rowIndex, lineCol + 1).setValue(line);
    if (String(payload.password || '') !== '' || !isActive) revokeUserSessions(username);
    return { status: 'success', mode: 'update', username: username };
  }

  if (!actor.permissions.add_user) throw new Error('ไม่มีสิทธิ์เพิ่มผู้ใช้');
  if (!password) throw new Error('ต้องระบุ password สำหรับผู้ใช้ใหม่');
  if (password.length < 4) throw new Error('password ต้องมีอย่างน้อย 4 ตัวอักษร');
  usersSheet.appendRow([username, hashPassword(password), role, String(isActive), permissionsJson, '', '']);
  if (lineCol !== -1) usersSheet.getRange(usersSheet.getLastRow(), lineCol + 1).setValue(line);
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
    machine: e.parameter.machine,
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
      machine: pick(row, ['machine'], ''),
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
    payload.reasonRemark || '',
    payload.machine || ''
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

  invalidateSparePartsLiteCache(resolvedSheetName);

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

// โครงสร้างโฟลเดอร์ใน Drive: ไลน์ > "ชื่อรายการ (Model)" > ชื่อไฟล์รูป
// ตั้งชื่อโฟลเดอร์จากชื่อรายการเป็นหลักเพื่อให้ค้นหาใน Drive ได้ตรงๆ (ห้ามใช้ item-1/2/3
// ที่ไล่หาไม่ได้ว่าเป็นของชิ้นไหน) — ถ้าไม่มีชื่อ fallback เป็น Model แล้วค่อยเป็น item-{itemId}
function getUploadTargetFolder(line, itemId, imageType, model, itemName) {
  var root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var safeLine = String(line || '').trim() || 'UnknownLine';
  var safeItemId = String(itemId || '').trim() || 'UNKNOWN';
  var name = String(itemName || '').trim();
  var modelTxt = String(model || '').trim();
  var segmentRaw = name ? (modelTxt ? name + ' (' + modelTxt + ')' : name) : modelTxt;
  var safeSegment = sanitizeDrivePathSegment(segmentRaw, 'item-' + safeItemId);

  var lineFolder = getOrCreateChildFolder(root, safeLine);
  var itemFolder = getOrCreateChildFolder(lineFolder, safeSegment);

  return {
    folder: itemFolder,
    drivePath: safeLine + '/' + safeSegment + '/'
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
  var itemName = String(payload.itemName || payload.name || '').trim();
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

  var target = getUploadTargetFolder(line, itemId, kind, model, itemName);
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
    // thumbnail endpoint เป็น URL หลัก — uc?export=view จะ redirect ไป lh3.googleusercontent
    // ที่ต้องล็อกอิน Google ทำให้รูปไม่ขึ้นบนเว็บสาธารณะ/ค้าง pending
    imageUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200',
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
    coil_size: findCol(['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size']),
    machines: findCol(['machines', 'machinelist', 'machine_list', 'machinenames'])
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
  if (fieldCols.machines === undefined) {
    fieldCols.machines = ensureColumnInContext(ctx, 'Machines', ['machines', 'machinelist', 'machine_list', 'machinenames']);
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
    coil_size: payload.coil_size !== undefined ? payload.coil_size : (payload.machine_model !== undefined ? payload.machine_model : ''),
    machines: payload.machines !== undefined ? payload.machines : ''
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
    // ฟิลด์ราคา/เอกสารที่โหมด lite ตัดทิ้ง — ฟอร์มแก้ไขที่เปิดจากรายการ lite จะส่งค่าว่างกลับมา
    // ถ้าเขียนทับตามนั้นราคา/Supplier/Drawing ที่ลงไว้จะถูกลบเงียบๆ ทุกครั้งที่มีคนกดแก้ไข
    // จึงถือว่า "ค่าว่าง = ไม่เปลี่ยนแปลง" เหมือนรูปภาพ (จะล้างค่าจริงให้แก้ในชีทโดยตรง)
    var detailValueKeys = {
      unit_price: true,
      supplier: true,
      price_updated_at: true,
      price_remark: true,
      drawing_url: true,
      drawing_file_name: true,
      drawing_revision: true,
      datasheet_url: true,
      quotation_url: true
    };
    for (var key in fieldCols) {
      if (fieldCols[key] !== undefined) {
        var nextValue = values[key];
        if (key === 'coil_size' && String(nextValue || '').trim() === '') continue;
        if (imageValueKeys[key] && String(nextValue || '').trim() === '') continue;
        if (detailValueKeys[key] && String(nextValue || '').trim() === '') continue;
        ctx.sheet.getRange(sheetRow, fieldCols[key] + 1).setValue(nextValue);
      }
    }
    invalidateSparePartsLiteCache(sheetName);
    return { status: 'success', mode: 'update', no: noValue, sheet: sheetName, location: values.location };
  }

  var newRow = new Array(ctx.headers.length);
  for (var x = 0; x < newRow.length; x += 1) newRow[x] = '';
  for (var k in fieldCols) {
    if (fieldCols[k] !== undefined) newRow[fieldCols[k]] = values[k];
  }
  ctx.sheet.appendRow(newRow);
  invalidateSparePartsLiteCache(sheetName);
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
      invalidateSparePartsLiteCache(sheetName);
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

// ซ่อมสิทธิ์เผยแพร่ไฟล์รูป/แนบทั้งหมดในสเปรดชีตนี้ (ทุกชีต ทุกแถว) — เดิม setSharing() ที่รันตอน
// อัปโหลดแต่ละครั้งอาจ fail แบบเงียบๆ (ถูก catch ทิ้งแค่ log warning ไม่เคยแจ้งผู้ใช้) ทำให้ไฟล์
// ถูกสร้างและบันทึก URL ไว้ในชีตแล้ว แต่เว็บสาธารณะเปิดรูปไม่ได้ (ต่างจาก Sheets ที่ล็อกอินด้วย
// บัญชีเจ้าของไฟล์อยู่แล้วเลยเห็น preview ได้ปกติ) ฟังก์ชันนี้ไล่ setSharing ซ้ำให้ทุกไฟล์ที่เจอ
function repairImageSharingPermissions(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  var user = findUserByUsername(session.user.username);
  var role = normalizeRole(user && user.role);
  if (role !== 'admin') throw new Error('เฉพาะ Admin เท่านั้นที่รันการซ่อมสิทธิ์รูปภาพได้');

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var seenFileIds = {};
  var checked = 0, fixed = 0, broken = 0;
  var brokenDetails = [];

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headerRowIndex = findHeaderRowIndex(data);
    var headers = data[headerRowIndex];
    var map = buildHeaderIndexMap(headers);
    var mainCol = map[normalizeHeaderName('image_main_file_id')];
    var installCol = map[normalizeHeaderName('image_install_file_id')];
    if (mainCol === undefined && installCol === undefined) return; // ไม่ใช่ชีตอะไหล่หลัก ข้าม

    data.slice(headerRowIndex + 1).forEach(function(row) {
      [mainCol, installCol].forEach(function(col) {
        if (col === undefined) return;
        var fileId = String(row[col] || '').trim();
        if (!fileId || seenFileIds[fileId]) return;
        seenFileIds[fileId] = true;
        checked += 1;
        try {
          DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fixed += 1;
        } catch (err) {
          broken += 1;
          brokenDetails.push({ sheet: sheetName, fileId: fileId, error: err && err.message ? err.message : String(err) });
        }
      });
    });
  });

  return {
    status: 'success',
    checked: checked,
    fixed: fixed,
    broken: broken,
    brokenDetails: brokenDetails.slice(0, 50)
  };
}

function extractDriveFileIdBackend(input) {
  var value = String(input || '').trim();
  if (!value) return '';
  var directMatch = value.match(/[?&]id=([A-Za-z0-9_-]+)/i);
  if (directMatch && directMatch[1]) return directMatch[1];
  var viewMatch = value.match(/\/file\/d\/([A-Za-z0-9_-]+)/i);
  if (viewMatch && viewMatch[1]) return viewMatch[1];
  if (/^[A-Za-z0-9_-]{20,}$/.test(value)) return value;
  return '';
}

// ตรวจสุขภาพรูปทุกรายการในทุกชีตสต็อก — ตอบว่ารายการไหน "ไม่มีรูปในชีต" (NO_URL),
// "มี URL แต่ไฟล์หายจาก Drive" (FILE_MISSING) หรือ "ไฟล์อยู่ในถังขยะ" (FILE_TRASHED)
// ใช้ไล่เก็บกวาดเคสรูปหาย/อัปโหลดไม่ติดให้ครบทุกรายการ — รันจากเว็บ (admin) หรือจาก
// Apps Script editor โดยตรงก็ได้ (เรียก auditItemImagesFromEditor)
function auditItemImages(payload) {
  var session = getSessionUser({ authToken: payload.authToken });
  var user = findUserByUsername(session.user.username);
  var role = normalizeRole(user && user.role);
  if (role !== 'admin') throw new Error('เฉพาะ Admin เท่านั้นที่รันตรวจสอบรูปภาพได้');
  return runItemImageAudit();
}

function auditItemImagesFromEditor() {
  var result = runItemImageAudit();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function runItemImageAudit() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = Array.from(new Set([SPARE_APP_CONFIG.readSheetName].concat(STOCK_LOCATION_SHEETS)));
  var summary = { totalItems: 0, ok: 0, noUrl: 0, fileMissing: 0, fileTrashed: 0 };
  var problems = [];
  var fileStateCache = {}; // fileId -> 'ok' | 'missing' | 'trashed' (รายการซ้ำใช้ไฟล์เดียวกันไม่ต้องเช็คซ้ำ)

  function checkFileState(fileId) {
    if (fileStateCache[fileId]) return fileStateCache[fileId];
    var state;
    try {
      state = DriveApp.getFileById(fileId).isTrashed() ? 'trashed' : 'ok';
    } catch (err) {
      state = 'missing';
    }
    fileStateCache[fileId] = state;
    return state;
  }

  sheetNames.forEach(function(sheetName) {
    var sheet = getSheetByFlexibleName(spreadsheet, sheetName);
    if (!sheet) return;
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headerRowIndex = findHeaderRowIndex(data);
    var map = buildHeaderIndexMap(data[headerRowIndex]);
    var noCol = getFirstColumnIndexByAliases(map, ['no', 'itemno', 'item_no', 'id']);
    var nameCol = getFirstColumnIndexByAliases(map, ['name', 'itemname', 'item_name', 'partname']);
    var modelCol = getFirstColumnIndexByAliases(map, ['model', 'partno', 'part_no']);
    var urlCol = getFirstColumnIndexByAliases(map, ['image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image']);
    var fileIdCol = getFirstColumnIndexByAliases(map, ['image_main_file_id', 'imagemainfileid']);
    if (nameCol === undefined && urlCol === undefined) return; // ไม่ใช่ชีตสต็อก

    data.slice(headerRowIndex + 1).forEach(function(row, i) {
      var name = nameCol !== undefined ? String(row[nameCol] || '').trim() : '';
      if (!name) return; // แถวว่าง
      summary.totalItems += 1;
      var entry = {
        sheet: sheetName,
        rowNumber: headerRowIndex + 2 + i,
        no: noCol !== undefined ? String(row[noCol] || '').trim() : '',
        name: name,
        model: modelCol !== undefined ? String(row[modelCol] || '').trim() : ''
      };
      var fileId = '';
      if (fileIdCol !== undefined) fileId = String(row[fileIdCol] || '').trim();
      if (!fileId && urlCol !== undefined) fileId = extractDriveFileIdBackend(row[urlCol]);
      if (!fileId) {
        summary.noUrl += 1;
        entry.problem = 'NO_URL';
        problems.push(entry);
        return;
      }
      var state = checkFileState(fileId);
      if (state === 'ok') { summary.ok += 1; return; }
      entry.fileId = fileId;
      if (state === 'trashed') {
        summary.fileTrashed += 1;
        entry.problem = 'FILE_TRASHED'; // กู้ได้จากถังขยะ Drive ภายใน 30 วัน
      } else {
        summary.fileMissing += 1;
        entry.problem = 'FILE_MISSING'; // ไฟล์ถูกลบถาวรแล้ว ต้องอัปโหลดใหม่
      }
      problems.push(entry);
    });
  });

  return { status: 'success', summary: summary, problems: problems };
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
    if (action === 'repairImageSharingPermissions') return respond(repairImageSharingPermissions(e.parameter), e);
    if (action === 'auditItemImages') return respond(auditItemImages(e.parameter), e);
    if (action === 'runAutoPrNow') return respond(runAutoPrNow(e.parameter), e);
    if (action === 'getIssueAnomalies') return respond(getIssueAnomalies(e.parameter), e);
    if (action === 'acknowledgeIssueAnomaly') return respond(acknowledgeIssueAnomaly(e.parameter), e);
    if (action === 'machineSpareReport') return respond(machineSpareReport(e.parameter), e);
    if (action === 'checkCrossLineStock') return respond(checkCrossLineStock(e.parameter), e);
    if (action === 'getPendingPrUsage') return respond(getPendingPrUsage(e.parameter), e);
    if (action === 'getMachines') return respond(getMachines(e.parameter), e);
    if (action === 'addMachine') return respond(addMachine(e.parameter), e);
    if (action === 'updateMachine') return respond(updateMachine(e.parameter), e);
    if (action === 'deleteMachine') return respond(deleteMachine(e.parameter), e);
    if (action === 'upsertUser') return respond(upsertUser({
      authToken: authToken,
      username: e.parameter.username,
      password: e.parameter.password,
      role: e.parameter.role,
      isActive: e.parameter.isActive,
      line: e.parameter.line,
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
    if (action === 'getInbox') return respond(getInbox(e.parameter), e);
    if (action === 'createPR') return respond(createPR(e.parameter), e);
    if (action === 'getPRForApproval') return respond(getPRForApproval(e.parameter), e);
    if (action === 'approvePR') return respond(approvePR(e.parameter), e);
    if (action === 'rejectPR') return respond(rejectPR(e.parameter), e);
    if (action === 'getPRStatus') return respond(getPRStatus(e.parameter), e);
    if (action === 'getPrApprovers') return respond(getPrApprovers(e.parameter), e);
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
      quotation_url: e.parameter.quotation_url,
      machines: e.parameter.machines
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
        // unit_price ต้องมาในโหมด lite ด้วย — การ์ดหน้ารายการใช้เช็คว่า "ยังไม่มีราคา" จริงหรือไม่
        // (เดิมตัดทิ้งทำให้การ์ดฟ้องไม่มีราคาทั้งที่ลงไว้แล้ว) เป็นตัวเลขสั้นๆ ไม่ทำ payload บวมแบบพวก URL
        unit_price: pickRowValue(row, map, ['unitprice', 'unit_price'], ''),
        currency: pickRowValue(row, map, ['currency'], 'THB'),
        supplier: isLiteRead ? '' : pickRowValue(row, map, ['supplier'], ''),
        price_updated_at: isLiteRead ? '' : pickRowValue(row, map, ['priceupdatedat', 'price_updated_at'], ''),
        price_remark: isLiteRead ? '' : pickRowValue(row, map, ['priceremark', 'price_remark'], ''),
        coil_size: pickRowValue(row, map, ['coilsize', 'machine_model', 'machinemodel', 'machinemodelcoilsize', 'model_size'], '-'),
        // รายชื่อเครื่องจักรที่ใช้อะไหล่นี้ได้ (comma-separated) — มาจาก master list ในชีต Machines
        machines: pickRowValue(row, map, ['machines', 'machinelist', 'machine_list', 'machinenames'], '')
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
    if (action === 'repairImageSharingPermissions') {
      return respond(repairImageSharingPermissions(body), e);
    }
    if (action === 'auditItemImages') {
      return respond(auditItemImages(body), e);
    }
    if (action === 'runAutoPrNow') {
      return respond(runAutoPrNow(body), e);
    }
    if (action === 'getIssueAnomalies') {
      return respond(getIssueAnomalies(body), e);
    }
    if (action === 'acknowledgeIssueAnomaly') {
      return respond(acknowledgeIssueAnomaly(body), e);
    }
    if (action === 'machineSpareReport') {
      return respond(machineSpareReport(body), e);
    }
    if (action === 'checkCrossLineStock') {
      return respond(checkCrossLineStock(body), e);
    }
    if (action === 'getPendingPrUsage') {
      return respond(getPendingPrUsage(body), e);
    }
    if (action === 'getMachines') return respond(getMachines(body), e);
    if (action === 'addMachine') return respond(addMachine(body), e);
    if (action === 'updateMachine') return respond(updateMachine(body), e);
    if (action === 'deleteMachine') return respond(deleteMachine(body), e);
    if (action === 'aiReadNameplate') {
      return respond(aiReadNameplate(body), e);
    }
    if (action === 'upsertUser') {
      return respond(upsertUser({
        authToken: authPayload.authToken,
        username: body.username,
        password: body.password,
        role: body.role,
        isActive: body.isActive,
        line: body.line,
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
    if (action === 'getInbox') return respond(getInbox(body), e);
    if (action === 'createPR') return respond(createPR(body), e);
    if (action === 'getPRForApproval') return respond(getPRForApproval(body), e);
    if (action === 'approvePR') return respond(approvePR(body), e);
    if (action === 'rejectPR') return respond(rejectPR(body), e);
    if (action === 'getPRStatus') return respond(getPRStatus(body), e);
    if (action === 'getPrApprovers') return respond(getPrApprovers(body), e);
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
        quotation_url: body.quotation_url,
        machines: body.machines
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
// SMART AUTOMATION + AI FEATURES
// =============================
// อ่านรายการอะไหล่แบบเบา (เฉพาะฟิลด์ที่งานอัตโนมัติใช้) จากทุกชีตสต็อก — dedupe ด้วยชื่อชีตจริง
// เพราะรายชื่อ candidate มีชื่อซ้ำ/ชื่อเก่าปนอยู่ (getSheetByFlexibleName จับคู่หลายแบบ)
function readAllStockItemsLean() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var candidates = Array.from(new Set([SPARE_APP_CONFIG.readSheetName].concat(STOCK_LOCATION_SHEETS)));
  var seenSheetIds = {};
  var items = [];
  candidates.forEach(function(sheetName) {
    var sheet = getSheetByFlexibleName(spreadsheet, sheetName);
    if (!sheet || seenSheetIds[sheet.getSheetId()]) return;
    seenSheetIds[sheet.getSheetId()] = true;
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headerRowIndex = findHeaderRowIndex(data);
    var map = buildHeaderIndexMap(data[headerRowIndex]);
    data.slice(headerRowIndex + 1).forEach(function(row) {
      var name = String(pickRowValue(row, map, ['namedescriptions', 'name', 'description', 'partname', 'jrpartname', 'jrpartnameolderp'], '')).trim();
      if (!name || name === '-') return;
      items.push({
        sheet: sheet.getName(),
        no: String(pickRowValue(row, map, ['no'], '')).trim(),
        name: name,
        model: String(pickRowValue(row, map, ['model', 'codeno', 'jrcodeno'], '')).trim(),
        brand: String(pickRowValue(row, map, ['brand'], '')).trim(),
        category: String(pickRowValue(row, map, ['category'], 'General')).trim(),
        line: String(pickRowValue(row, map, ['mainline', 'line', 'process', 'ไลน์'], '')).trim(),
        unit: String(pickRowValue(row, map, ['unit'], 'PCS')).trim(),
        stock: Number(pickRowValue(row, map, ['stockqty', 'qtystock', 'qoh', 'stock'], 0)) || 0,
        min: Number(pickRowValue(row, map, ['min', 'qtymin'], 0)) || 0,
        max: Number(pickRowValue(row, map, ['max', 'qtymax'], 0)) || 0,
        unit_price: Number(pickRowValue(row, map, ['unitprice', 'unit_price'], 0)) || 0,
        supplier: String(pickRowValue(row, map, ['supplier'], '')).trim(),
        // ให้ runAutoPrJob สแนปช็อตรูปแนบ PR ได้เหมือน PR ที่สร้างจากหน้าเว็บ (alias เดียวกับ lite/detail reader)
        image_main_url: String(pickRowValue(row, map, ['image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image', 'sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], '') || buildDriveViewUrlFromFileId(pickRowValue(row, map, ['image_main_file_id', 'imagemainfileid'], ''))).trim()
      });
    });
  });
  return items;
}

function normalizePartKeyText(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function partKeyOf(name, model) {
  return normalizePartKeyText(name) + '||' + normalizePartKeyText(model === '-' ? '' : model);
}

// ── F6: เช็คสต็อกข้ามไลน์ ─────────────────────────────────────────
// หา "ของตัวเดียวกัน" ในชีตอื่นที่ยังมีเหลือเกิน Min — จับคู่ด้วย model ก่อน (แม่นสุด)
// ถ้าไม่มี model ใช้ชื่อเทียบแบบ normalize
function findCrossLineStockMatches(allItems, name, model, excludeSheet) {
  var modelKey = normalizePartKeyText(model === '-' ? '' : model);
  var nameKey = normalizePartKeyText(name);
  return allItems.filter(function(it) {
    if (String(it.sheet) === String(excludeSheet)) return false;
    var surplus = it.stock - it.min;
    if (surplus <= 0) return false;
    var itModel = normalizePartKeyText(it.model === '-' ? '' : it.model);
    // ถ้าฝั่งที่ขาดมี model ให้จับคู่ด้วย model เท่านั้น — ชื่อเดียวกันแต่คนละรุ่นห้ามนับว่าโอนแทนกันได้
    if (modelKey) return itModel === modelKey;
    return normalizePartKeyText(it.name) === nameKey;
  }).map(function(it) {
    return { sheet: it.sheet, line: it.line || it.sheet, stock: it.stock, min: it.min, surplus: it.stock - it.min, no: it.no };
  });
}

function checkCrossLineStock(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var allItems = readAllStockItemsLean();
  var matches = findCrossLineStockMatches(allItems, payload.name, payload.model, String(payload.sheet || ''));
  return { status: 'success', matches: matches };
}

// ── อัตราการใช้จาก Log (ใช้ทั้ง Auto-PR และตรวจจับเบิกผิดปกติ) ────────
function computeIssueUsageStats(logRows, nowMs) {
  var now = nowMs || Date.now();
  var d30 = now - 30 * 86400000, d7 = now - 7 * 86400000, d63 = now - 63 * 86400000;
  var stats = {}; // key -> { qty30, qty7, qtyBase56, users7: {user: qty}, lastTs }
  (logRows || []).forEach(function(r) {
    if (String(r.type || '').toLowerCase() !== 'output') return;
    var ts = new Date(String(r.timestamp || '').replace(' ', 'T') + '+07:00').getTime();
    if (isNaN(ts) || ts < d63) return;
    var key = partKeyOf(r.partName, r.model);
    if (!stats[key]) stats[key] = { name: r.partName, model: r.model, qty30: 0, qty7: 0, qtyBase56: 0, users7: {}, lastTs: 0 };
    var s = stats[key];
    var qty = Math.abs(Number(r.qty) || 0);
    if (ts >= d30) s.qty30 += qty;
    if (ts >= d7) {
      s.qty7 += qty;
      var user = String(r.by || 'Unknown');
      s.users7[user] = (s.users7[user] || 0) + qty;
    } else {
      s.qtyBase56 += qty; // 8 สัปดาห์ก่อนหน้า (ไม่รวม 7 วันล่าสุด)
    }
    if (ts > s.lastTs) s.lastTs = ts;
  });
  return stats;
}

// ── F1: Auto-PR ทุกเช้า ───────────────────────────────────────────
// จำนวนแนะนำ: เติมให้พอใช้ 45 วันตามอัตราเบิกจริง 30 วันล่าสุด หรือเติมถึง Max — เอาค่ามากกว่า
// (ถ้าไม่มีข้อมูลการใช้เลย ใช้ min-stock+ให้ถึง max ตามเดิม)
function computeSuggestedOrderQty(item, usage30) {
  var dailyUse = (usage30 || 0) / 30;
  var coverTarget = Math.ceil(dailyUse * 45);
  var refillToMax = item.max > 0 ? item.max - item.stock : 0;
  var refillToMin = Math.max(item.min - item.stock, 0) + Math.max(Math.ceil(item.min * 0.5), 1);
  var suggested = Math.max(coverTarget - item.stock, refillToMax, refillToMin);
  return Math.max(suggested, 1);
}

function listPendingPrPartKeys() {
  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  if (hData.length <= 1) return {};
  var hIdx = prIndexMap(hData[0]);
  var pendingIds = {};
  for (var i = 1; i < hData.length; i += 1) {
    if (String(hData[i][hIdx.status]) === 'PENDING') pendingIds[String(hData[i][hIdx.pr_id])] = true;
  }
  var keys = {};
  if (!Object.keys(pendingIds).length) return keys;
  var linesSheet = getPrLinesSheet();
  var lData = linesSheet.getDataRange().getValues();
  if (lData.length <= 1) return keys;
  var lIdx = prIndexMap(lData[0]);
  for (var j = 1; j < lData.length; j += 1) {
    if (pendingIds[String(lData[j][lIdx.pr_id])]) {
      keys[partKeyOf(lData[j][lIdx.part_name], lData[j][lIdx.model])] = true;
    }
  }
  return keys;
}

function runAutoPrJob() {
  var allItems = readAllStockItemsLean();
  var usageStats = computeIssueUsageStats(getLogRows());
  var pendingKeys = listPendingPrPartKeys();
  var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');

  var candidates = allItems.filter(function(it) {
    return it.min > 0 && it.stock < it.min && !pendingKeys[partKeyOf(it.name, it.model)];
  });
  if (!candidates.length) return { status: 'success', created: [], itemCount: 0, message: 'ไม่มีรายการต่ำกว่า Min ที่ยังไม่อยู่ใน PR ค้างอนุมัติ' };

  // จัดกลุ่มตามชีต (= ไลน์) → PR ใบละไลน์ อ่านง่ายตอนอนุมัติ
  var bySheet = {};
  candidates.forEach(function(it) {
    if (!bySheet[it.sheet]) bySheet[it.sheet] = [];
    bySheet[it.sheet].push(it);
  });

  var headerSheet = getPrHeaderSheet();
  var hData = headerSheet.getDataRange().getValues();
  var hIdx = prIndexMap(hData[0]);
  var linesSheet = getPrLinesSheet();
  var lHeaderRow = linesSheet.getRange(1, 1, 1, linesSheet.getLastColumn()).getValues()[0];
  var lIdx = prIndexMap(lHeaderRow);

  var createdPrs = [];
  var totalItems = 0;
  Object.keys(bySheet).forEach(function(sheetName) {
    var items = bySheet[sheetName];
    var prId = 'PR-AUTO-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd') + '-' + sanitizeDrivePathSegment(sheetName, 'LINE').slice(0, 12);
    if (findPrHeaderRow(headerSheet.getDataRange().getValues(), hIdx, prId) !== -1) return; // วันนี้สร้างของไลน์นี้ไปแล้ว
    var totalAmount = 0;
    var lineRows = items.map(function(it, i) {
      var usage = usageStats[partKeyOf(it.name, it.model)];
      var qty = computeSuggestedOrderQty(it, usage ? usage.qty30 : 0);
      var price = it.unit_price || 0;
      totalAmount += qty * price;
      var remarkParts = ['🤖 Auto: คงเหลือ ' + it.stock + '/Min ' + it.min];
      if (usage && usage.qty30 > 0) remarkParts.push('ใช้ ' + usage.qty30 + ' ชิ้น/30วัน');
      if (it.supplier) remarkParts.push('เจ้าเดิม: ' + it.supplier);
      var crossMatches = findCrossLineStockMatches(allItems, it.name, it.model, it.sheet);
      if (crossMatches.length) {
        var cm = crossMatches[0];
        remarkParts.push('♻️ ' + cm.sheet + ' มีเหลือ ' + cm.surplus + ' ชิ้น — พิจารณาโอนก่อนสั่ง');
      }
      var rowArr = new Array(lHeaderRow.length).fill('');
      rowArr[lIdx.pr_id] = prId;
      rowArr[lIdx.line_no] = i + 1;
      rowArr[lIdx.part_no] = it.no;
      rowArr[lIdx.part_name] = it.name;
      rowArr[lIdx.model] = it.model;
      rowArr[lIdx.brand] = it.brand;
      rowArr[lIdx.category] = it.category;
      rowArr[lIdx.unit] = it.unit;
      rowArr[lIdx.qty_requested] = qty;
      rowArr[lIdx.qty_approved] = qty;
      rowArr[lIdx.unit_price] = price;
      rowArr[lIdx.remark] = remarkParts.join(' | ');
      if (lIdx.image_url !== undefined) rowArr[lIdx.image_url] = it.image_main_url || '';
      return rowArr;
    });
    linesSheet.getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, lHeaderRow.length).setValues(lineRows);

    var headerRowArr = new Array(hData[0].length).fill('');
    headerRowArr[hIdx.pr_id] = prId;
    headerRowArr[hIdx.status] = 'PENDING';
    headerRowArr[hIdx.created_by] = 'AUTO-BOT';
    headerRowArr[hIdx.created_at] = now;
    headerRowArr[hIdx.line] = String(items[0].line || sheetName);
    headerRowArr[hIdx.dept] = 'Auto Reorder';
    headerRowArr[hIdx.item_count] = items.length;
    headerRowArr[hIdx.total_amount] = totalAmount;
    headerRowArr[hIdx.updated_at] = now;
    headerSheet.appendRow(headerRowArr);
    appendPrAudit(prId, 'CREATE', 'AUTO-BOT', sheetName, '', '', items.length + ' รายการ (สร้างอัตโนมัติจากรายการต่ำกว่า Min)');
    createdPrs.push(prId);
    totalItems += items.length;
  });

  return { status: 'success', created: createdPrs, itemCount: totalItems };
}

function runAutoPrNow(payload) {
  requirePermission({ authToken: payload.authToken }, 'pr_approve');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return runAutoPrJob(); } finally { lock.releaseLock(); }
}

// ── F4: ตรวจจับการเบิกผิดปกติ ─────────────────────────────────────
var ANOMALY_SHEET_NAME = 'AnomalyAlerts';
var ANOMALY_HEADERS = ['detected_at', 'part_name', 'model', 'qty_7d', 'weekly_avg', 'ratio', 'top_user', 'top_user_qty', 'detail', 'status'];

function getAnomalySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(ANOMALY_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ANOMALY_SHEET_NAME);
    sheet.appendRow(ANOMALY_HEADERS);
  }
  return sheet;
}

function runIssueAnomalyScan() {
  var stats = computeIssueUsageStats(getLogRows());
  var sheet = getAnomalySheet();
  var data = sheet.getDataRange().getValues();
  var idx = prIndexMap(data[0] || ANOMALY_HEADERS);
  var recentFlagged = {}; // กันแจ้งซ้ำ: part เดียวกันที่ยังสถานะ NEW ใน 7 วันล่าสุด
  var now = Date.now();
  for (var i = 1; i < data.length; i += 1) {
    var ts = new Date(String(data[i][idx.detected_at] || '').replace(' ', 'T') + '+07:00').getTime();
    if (String(data[i][idx.status]) === 'NEW' && !isNaN(ts) && now - ts < 7 * 86400000) {
      recentFlagged[partKeyOf(data[i][idx.part_name], data[i][idx.model])] = true;
    }
  }

  var flagged = [];
  Object.keys(stats).forEach(function(key) {
    var s = stats[key];
    if (recentFlagged[key]) return;
    var weeklyAvg = s.qtyBase56 / 8;
    // เกณฑ์: 7 วันล่าสุดเบิก >= 5 ชิ้น และมากกว่า 3 เท่าของค่าเฉลี่ยรายสัปดาห์ 8 สัปดาห์ก่อนหน้า
    // (ถ้าไม่เคยเบิกเลยในอดีต ใช้เกณฑ์ >= 10 ชิ้นใน 7 วัน กันของใหม่เข้าระบบโดน flag มั่ว)
    var isSpike = s.qty7 >= 5 && weeklyAvg > 0 && s.qty7 > weeklyAvg * 3;
    var isNewHeavy = s.qty7 >= 10 && weeklyAvg === 0;
    if (!isSpike && !isNewHeavy) return;
    var topUser = '', topUserQty = 0;
    Object.keys(s.users7).forEach(function(u) {
      if (s.users7[u] > topUserQty) { topUser = u; topUserQty = s.users7[u]; }
    });
    var ratio = weeklyAvg > 0 ? (s.qty7 / weeklyAvg) : 0;
    var detail = 'เบิก ' + s.qty7 + ' ชิ้นใน 7 วัน' +
      (weeklyAvg > 0 ? ' (ปกติ ~' + weeklyAvg.toFixed(1) + ' ชิ้น/สัปดาห์ = ' + ratio.toFixed(1) + ' เท่า)' : ' (ไม่เคยมีการเบิกใน 8 สัปดาห์ก่อนหน้า)') +
      (topUser && topUserQty >= s.qty7 * 0.7 ? ' — ผู้เบิกหลัก: ' + topUser + ' (' + topUserQty + ' ชิ้น)' : '');
    flagged.push([
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss'),
      s.name, s.model || '', s.qty7, Number(weeklyAvg.toFixed(2)), Number(ratio.toFixed(2)),
      topUser, topUserQty, detail, 'NEW'
    ]);
  });

  if (flagged.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, flagged.length, ANOMALY_HEADERS.length).setValues(flagged);
  }
  return { status: 'success', flagged: flagged.length };
}

function getIssueAnomalies(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var sheet = getAnomalySheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'success', anomalies: [] };
  var idx = prIndexMap(data[0]);
  var now = Date.now();
  var out = [];
  for (var i = 1; i < data.length; i += 1) {
    var row = data[i];
    if (String(row[idx.status]) !== 'NEW') continue;
    var ts = new Date(String(row[idx.detected_at] || '').replace(' ', 'T') + '+07:00').getTime();
    if (isNaN(ts) || now - ts > 14 * 86400000) continue;
    out.push({
      rowNumber: i + 1,
      detected_at: String(row[idx.detected_at] || ''),
      part_name: String(row[idx.part_name] || ''),
      model: String(row[idx.model] || ''),
      qty_7d: Number(row[idx.qty_7d] || 0),
      weekly_avg: Number(row[idx.weekly_avg] || 0),
      ratio: Number(row[idx.ratio] || 0),
      top_user: String(row[idx.top_user] || ''),
      detail: String(row[idx.detail] || '')
    });
  }
  out.sort(function(a, b) { return b.ratio - a.ratio; });
  return { status: 'success', anomalies: out };
}

function acknowledgeIssueAnomaly(payload) {
  requirePermission({ authToken: payload.authToken }, 'pr_approve');
  var rowNumber = Number(payload.rowNumber || 0);
  if (!rowNumber || rowNumber < 2) throw new Error('ไม่พบรายการแจ้งเตือน');
  var sheet = getAnomalySheet();
  var idx = prIndexMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  sheet.getRange(rowNumber, idx.status + 1).setValue('ACKNOWLEDGED');
  return { status: 'success' };
}

// ── ตัวติดตั้ง trigger รายวัน — รันครั้งเดียวจาก Apps Script editor ──
function setupAutomation() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyAutoJobs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyAutoJobs').timeBased().everyDays(1).atHour(7).create();
  Logger.log('ติดตั้ง trigger รายวัน 07:00 สำเร็จ — ระบบจะสร้าง Auto-PR และสแกนการเบิกผิดปกติทุกเช้า');
  return 'OK';
}

function runDailyAutoJobs() {
  var results = {};
  try { results.autoPr = runAutoPrJob(); } catch (e1) { results.autoPr = { error: String(e1 && e1.message || e1) }; }
  try { results.anomaly = runIssueAnomalyScan(); } catch (e2) { results.anomaly = { error: String(e2 && e2.message || e2) }; }
  Logger.log(JSON.stringify(results));
  return results;
}

// ── F5: รายงานอะไหล่ต่อเครื่องจักร ─────────────────────────────────
function machineSpareReport(payload) {
  requirePermission({ authToken: payload.authToken }, 'view');
  var logRows = getLogRows();
  var allItems = readAllStockItemsLean();
  var priceByKey = {};
  allItems.forEach(function(it) {
    var key = partKeyOf(it.name, it.model);
    if (!priceByKey[key]) priceByKey[key] = it.unit_price || 0;
  });
  var byMachine = {};
  logRows.forEach(function(r) {
    if (String(r.type || '').toLowerCase() !== 'output') return;
    var machine = String(r.machine || '').trim();
    if (!machine) return;
    if (!byMachine[machine]) byMachine[machine] = { machine: machine, issueCount: 0, totalQty: 0, estCost: 0, byReason: {}, parts: {}, lastIssue: '' };
    var m = byMachine[machine];
    var qty = Math.abs(Number(r.qty) || 0);
    m.issueCount += 1;
    m.totalQty += qty;
    m.estCost += qty * (priceByKey[partKeyOf(r.partName, r.model)] || 0);
    var reason = String(r.reason || 'ไม่ระบุ');
    m.byReason[reason] = (m.byReason[reason] || 0) + qty;
    var pk = String(r.partName || '') + (r.model && r.model !== '-' ? ' (' + r.model + ')' : '');
    m.parts[pk] = (m.parts[pk] || 0) + qty;
    if (String(r.timestamp) > m.lastIssue) m.lastIssue = String(r.timestamp);
  });
  var report = Object.keys(byMachine).map(function(k) {
    var m = byMachine[k];
    var topParts = Object.keys(m.parts).map(function(p) { return { part: p, qty: m.parts[p] }; })
      .sort(function(a, b) { return b.qty - a.qty; }).slice(0, 5);
    return {
      machine: m.machine, issueCount: m.issueCount, totalQty: m.totalQty,
      estCost: Math.round(m.estCost), byReason: m.byReason, topParts: topParts, lastIssue: m.lastIssue
    };
  }).sort(function(a, b) { return b.estCost - a.estCost; });
  return { status: 'success', machines: report };
}

// ── AI helpers (F2 + F3) — Claude API ผ่าน UrlFetchApp ─────────────
// ตั้งค่า: Apps Script > Project Settings > Script Properties เพิ่ม ANTHROPIC_API_KEY
// (เลือกรุ่นได้ด้วย property AI_MODEL — default claude-sonnet-5, ประหยัดกว่าใช้ claude-haiku-4-5)
function getAnthropicApiKey() {
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('AI_KEY_MISSING: ยังไม่ได้ตั้งค่า ANTHROPIC_API_KEY ใน Script Properties (Apps Script > Project Settings)');
  return key;
}

function callClaudeApi(systemPrompt, userContent, maxTokens) {
  var model = PropertiesService.getScriptProperties().getProperty('AI_MODEL') || 'claude-sonnet-5';
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': getAnthropicApiKey(), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: model,
      max_tokens: maxTokens || 1024,
      system: systemPrompt,
      messages: [{ role: 'user', content: userContent }]
    }),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText() || '{}');
  if (code !== 200) {
    throw new Error('AI_API_ERROR: ' + (body && body.error && body.error.message ? body.error.message : ('HTTP ' + code)));
  }
  var parts = (body.content || []).filter(function(b) { return b.type === 'text'; });
  return parts.map(function(b) { return b.text; }).join('\n');
}

function parseAiJson(text) {
  var cleaned = String(text || '').replace(/```json/gi, '').replace(/```/g, '').trim();
  var start = cleaned.indexOf('{');
  var end = cleaned.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('AI ตอบกลับมาไม่ใช่รูปแบบ JSON');
  return JSON.parse(cleaned.slice(start, end + 1));
}

// ── F2: อ่านป้ายอะไหล่จากรูป ──────────────────────────────────────
function aiReadNameplate(payload) {
  requirePermission({ authToken: payload.authToken }, 'manage_items');
  var dataUrl = String(payload.dataUrl || '');
  var mimeType = getDataUrlMimeType(dataUrl);
  var allowed = { 'image/jpeg': true, 'image/png': true, 'image/webp': true };
  if (!allowed[mimeType]) throw new Error('รองรับเฉพาะรูป jpg, png, webp');
  var base64Content = dataUrl.split(',')[1] || '';
  if (!base64Content) throw new Error('ไม่พบข้อมูลรูป');

  var text = callClaudeApi(
    'คุณเป็นผู้เชี่ยวชาญอ่านป้ายอะไหล่/nameplate ของชิ้นส่วนอุตสาหกรรม ตอบเป็น JSON เท่านั้น',
    [
      { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Content } },
      { type: 'text', text: 'อ่านข้อมูลจากป้าย/ฉลาก/ตัวพิมพ์บนอะไหล่ในรูปนี้ แล้วตอบ JSON: {"name": "ชื่อประเภทอะไหล่ภาษาไทยหรืออังกฤษสั้นๆ เช่น Solenoid Valve, Bearing", "model": "รหัสรุ่นที่พิมพ์บนป้าย", "brand": "ยี่ห้อ", "category": "หมวดหมู่ เช่น Electrical, Mechanical, Pneumatic", "specs": "สเปคสำคัญที่อ่านได้ เช่น แรงดัน กระแส ขนาด"} — ช่องไหนอ่านไม่ได้ให้ใส่ค่าว่าง ห้ามเดา' }
    ],
    500
  );
  var parsed = parseAiJson(text);
  return {
    status: 'success',
    name: String(parsed.name || ''),
    model: String(parsed.model || ''),
    brand: String(parsed.brand || ''),
    category: String(parsed.category || ''),
    specs: String(parsed.specs || '')
  };
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
