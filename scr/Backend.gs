// =============================
// CONFIG
// =============================
var SPARE_APP_CONFIG = this.SPARE_APP_CONFIG || {};
SPARE_APP_CONFIG.readSheetName = SPARE_APP_CONFIG.readSheetName || 'Main List Stock';
SPARE_APP_CONFIG.writeSheetName = SPARE_APP_CONFIG.writeSheetName || 'Log';
SPARE_APP_CONFIG.usersSheetName = SPARE_APP_CONFIG.usersSheetName || 'Users';
var LOG_HEADERS = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After'];
var USER_HEADERS = ['username', 'password', 'role', 'is_active', 'permissions_json', 'session_token', 'session_expiry'];

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

function pickRowValue(row, map, keys, fallbackValue) {
  for (var i = 0; i < keys.length; i += 1) {
    var idx = map[keys[i]];
    if (idx !== undefined && row[idx] !== '' && row[idx] !== null && row[idx] !== undefined) {
      return row[idx];
    }
  }
  return fallbackValue;
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

function normalizeRole(role) {
  var val = String(role || 'user').toLowerCase().trim();
  if (val === 'admin' || val === 'leader' || val === 'user') return val;
  return 'user';
}

function getRoleDefaultPermissions(role) {
  var normalized = normalizeRole(role);
  if (normalized === 'admin') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: true, add_user: true, delete_user: true, manage_auth: true
  };
  if (normalized === 'leader') return {
    view: true, transact: true, manage_items: true, delete_items: true,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false
  };
  return {
    view: true, transact: true, manage_items: false, delete_items: false,
    manage_users: false, add_user: false, delete_user: false, manage_auth: false
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

  var firstRow = historySheet.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
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
    sheetName: e.parameter.sheet
  };
}

function getLogRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = getOrCreateSheet(spreadsheet, SPARE_APP_CONFIG.writeSheetName);
  ensureLogSheetHeaders(historySheet);

  var data = historySheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(function (row, idx) {
    return {
      no: idx + 1,
      timestamp: row[0],
      type: row[1],
      process: row[2],
      category: row[3],
      partName: row[4],
      model: row[5],
      brand: row[6],
      qty: row[7],
      unit: row[8],
      by: row[9],
      partNo: row[10],
      stockBefore: row[11],
      stockAfter: row[12]
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
    if (noMatch || (nameMatch && modelMatch)) {
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
    new Date(),
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
    stockAfter
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
  var ROOT_FOLDER_ID = '1XWO5rGpku35gSTMAh4HDOCHa6GJIkoS3';
  var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
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
    line: findCol(['mainline', 'line', 'location', 'jrlocation']),
    category: findCol(['category']),
    brand: findCol(['brand']),
    photo: findCol(['sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic']),
    image_main_url: findCol(['image_main_url', 'imagemainurl', 'image_main', 'imagemain']),
    image_main_file_id: findCol(['image_main_file_id', 'imagemainfileid']),
    image_install_url: findCol(['image_install_url', 'imageinstallurl', 'image_install', 'imageinstall']),
    image_install_file_id: findCol(['image_install_file_id', 'imageinstallfileid']),
    image_main: findCol(['image_main', 'imagemain', 'mainimage', 'main_image']),
    image_install: findCol(['image_install', 'imageinstall', 'installimage', 'install_image']),
    max: findCol(['max', 'qtymax']),
    min: findCol(['min', 'qtymin']),
    unit: findCol(['unit']),
    stock: findCol(['stockqty', 'qtystock', 'qoh', 'stock', 'initialstock'])
  };

  if (fieldCols.brand === undefined) {
    fieldCols.brand = ensureColumnInContext(ctx, 'Brand', ['brand']);
  }
  if (fieldCols.line === undefined) {
    fieldCols.line = ensureColumnInContext(ctx, 'Line', ['line', 'mainline']);
  }
  if (fieldCols.category === undefined) {
    fieldCols.category = ensureColumnInContext(ctx, 'Category', ['category']);
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
    category: payload.category || '',
    brand: payload.brand || '',
    photo: payload.photo || '',
    image_main: payload.image_main || payload.image_main_url || payload.photo || '',
    image_install: payload.image_install || payload.image_install_url || '',
    image_main_url: payload.image_main_url || payload.image_main || payload.photo || '',
    image_main_file_id: payload.image_main_file_id || '',
    image_install_url: payload.image_install_url || payload.image_install || '',
    image_install_file_id: payload.image_install_file_id || '',
    max: payload.max || '',
    min: payload.min || '',
    unit: payload.unit || '',
    stock: payload.stock || ''
  };

  if (targetIndex > -1) {
    var sheetRow = ctx.headerRowIndex + 2 + targetIndex;
    for (var key in fieldCols) {
      if (fieldCols[key] !== undefined) {
        ctx.sheet.getRange(sheetRow, fieldCols[key] + 1).setValue(values[key]);
      }
    }
    return { status: 'success', mode: 'update', no: noValue };
  }

  var newRow = new Array(ctx.headers.length);
  for (var x = 0; x < newRow.length; x += 1) newRow[x] = '';
  for (var k in fieldCols) {
    if (fieldCols[k] !== undefined) newRow[fieldCols[k]] = values[k];
  }
  ctx.sheet.appendRow(newRow);
  return { status: 'success', mode: 'create', no: noValue };
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
    requirePermission(authPayload, 'view');
    if (action === 'transact') {
      requirePermission(authPayload, 'transact');
      return respond(processTransaction(parseTransactionPayloadFromGet(e)), e);
    }
    if (action === 'logs') return respond(getLogRows(), e);
    if (action === 'authStatus') return respond(getDriveAuthStatus(), e);
    if (action === 'authorizeDrive') return respond(authorizeGoogleDriveAccess(), e);
    if (action === 'upsertItem') {
      requirePermission(authPayload, 'manage_items');
      return respond(upsertMainItem({
      sheetName: e.parameter.sheet,
      no: e.parameter.no,
      name: e.parameter.name,
      model: e.parameter.model,
      line: e.parameter.line,
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
      stock: e.parameter.stock
    }), e);
    }
    if (action === 'deleteItem') {
      requirePermission(authPayload, 'delete_items');
      return respond(deleteMainItem({ sheetName: e.parameter.sheet, no: e.parameter.no }), e);
    }

    var sheetName = resolveReadSheetName({ sheet: e.parameter.sheet });
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureSheetWithTemplate(spreadsheet, sheetName);

    var data = sheet.getDataRange().getValues();
    if (!data.length || data.length <= 1) return respond([], e);

    var headerRowIndex = findHeaderRowIndex(data);
    var headers = data[headerRowIndex];
    var rows = data.slice(headerRowIndex + 1);
    var map = buildHeaderIndexMap(headers);

    var result = rows.map(function (row, index) {
      var stockValue = Number(pickRowValue(row, map, ['stockqty', 'qtystock', 'qoh', 'stock'], 0)) || 0;
      var minValue = Number(pickRowValue(row, map, ['min', 'qtymin'], 0)) || 0;
      var needToPOValue = Math.max(minValue - stockValue, 0);

      return {
        no: pickRowValue(row, map, ['no'], index + 1),
        name: pickRowValue(row, map, ['namedescriptions', 'name', 'description', 'partname', 'jrpartname', 'jrpartnameolderp'], '-'),
        model: pickRowValue(row, map, ['model', 'codeno', 'jrcodeno'], '-'),
        line: pickRowValue(row, map, ['mainline', 'line', 'location', 'jrlocation'], '-'),
        location: pickRowValue(row, map, ['location', 'jrlocation'], '-'),
        category: pickRowValue(row, map, ['category'], 'General'),
        brand: pickRowValue(row, map, ['brand'], '-'),
        stock: stockValue,
        max: pickRowValue(row, map, ['max', 'qtymax'], 0),
        min: minValue,
        needToPO: needToPOValue,
        unit: pickRowValue(row, map, ['unit'], 'PCS'),
        remark: pickRowValue(row, map, ['remark'], ''),
        photo: pickRowValue(row, map, ['sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], ''),
        image_main: pickRowValue(row, map, ['image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image', 'sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], ''),
        image_install: pickRowValue(row, map, ['image_install_url', 'imageinstallurl', 'image_install', 'imageinstall', 'installimage', 'install_image'], ''),
        image_main_url: pickRowValue(row, map, ['image_main_url', 'imagemainurl', 'image_main', 'imagemain', 'mainimage', 'main_image', 'sparepartsphotos', 'photo', 'photourl', 'image', 'imageurl', 'picture', 'pic'], ''),
        image_main_file_id: pickRowValue(row, map, ['image_main_file_id', 'imagemainfileid'], ''),
        image_install_url: pickRowValue(row, map, ['image_install_url', 'imageinstallurl', 'image_install', 'imageinstall', 'installimage', 'install_image'], ''),
        image_install_file_id: pickRowValue(row, map, ['image_install_file_id', 'imageinstallfileid'], '')
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
    requirePermission(authPayload, 'view');
    if (action === 'upsertItem') {
      requirePermission(authPayload, 'manage_items');
      return respond(upsertMainItem({
        sheetName: body.sheet || body.sheetName,
        no: body.no,
        name: body.name,
        model: body.model,
        line: body.line,
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
        stock: body.stock
      }), e);
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
