// =============================
// CONFIG
// =============================
var SPARE_APP_CONFIG = this.SPARE_APP_CONFIG || {};
SPARE_APP_CONFIG.readSheetName = SPARE_APP_CONFIG.readSheetName || 'Main List Stock';
SPARE_APP_CONFIG.writeSheetName = SPARE_APP_CONFIG.writeSheetName || 'Log';
var LOG_HEADERS = ['Timestamp', 'Type', 'Process', 'Category', 'Part Name', 'Model', 'Brand', 'Qty', 'Unit', 'By', 'Part No', 'Stock Before', 'Stock After'];

// =============================
// HELPERS
// =============================

function buildErrorResponse(err) {
  var msg = err && err.message ? String(err.message) : String(err || 'Unknown error');
  var lower = msg.toLowerCase();
  var isDriveAuth = lower.indexOf('ไม่ได้รับอนุญาต') > -1 ||
    lower.indexOf('authorization') > -1 ||
    lower.indexOf('googleapis.com/auth/drive') > -1 ||
    lower.indexOf('driveapp') > -1;

  if (isDriveAuth) {
    return {
      status: 'error',
      errorCode: 'DRIVE_AUTH_REQUIRED',
      message: 'ยังไม่ได้อนุญาตสิทธิ์ Google Drive ให้ Apps Script (DRIVE_AUTH_REQUIRED). กรุณาเปิด Apps Script แล้ว Run ฟังก์ชันที่ใช้ DriveApp 1 ครั้งเพื่ออนุญาตสิทธิ์ จากนั้น Deploy เว็บแอปใหม่และลองอีกครั้ง',
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
  var mainSheet = getSheetByFlexibleName(spreadsheet, resolvedSheetName);

  ensureLogSheetHeaders(historySheet);
  if (!mainSheet) throw new Error('ไม่พบชีทชื่อ ' + resolvedSheetName);
  if (!payload.partName || !payload.qty) throw new Error('ต้องมี partName และ qty');

  var qty = Number(payload.qty);
  if (!qty || qty <= 0) throw new Error('qty ต้องมากกว่า 0');

  var signedQty = qty;
  if (payload.type && String(payload.type).indexOf('Output') > -1) signedQty = -Math.abs(qty);
  else signedQty = Math.abs(qty);

  var mainData = mainSheet.getDataRange().getValues();
  if (!mainData.length) throw new Error('ไม่พบข้อมูลในชีทหลัก');

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
  var sheet = getSheetByFlexibleName(spreadsheet, sheetName);
  if (!sheet) throw new Error('ไม่พบชีทชื่อ ' + sheetName);

  var data = sheet.getDataRange().getValues();
  if (!data.length) throw new Error('ไม่พบข้อมูลในชีท ' + sheetName);

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

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    ok: true,
    status: 'success',
    itemId: itemId,
    kind: kind,
    fileId: file.getId(),
    imageUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    viewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    drivePath: target.drivePath
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
    if (action === 'transact') return respond(processTransaction(parseTransactionPayloadFromGet(e)), e);
    if (action === 'logs') return respond(getLogRows(), e);
    if (action === 'authStatus') return respond(getDriveAuthStatus(), e);
    if (action === 'upsertItem') return respond(upsertMainItem({
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
    if (action === 'deleteItem') return respond(deleteMainItem({ sheetName: e.parameter.sheet, no: e.parameter.no }), e);

    var sheetName = resolveReadSheetName({ sheet: e.parameter.sheet });
    var sheet = getSheetByFlexibleName(SpreadsheetApp.getActiveSpreadsheet(), sheetName);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ ' + sheetName);

    var data = sheet.getDataRange().getValues();
    if (!data.length) return respond([], e);

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
    if (!action && (body.itemId || body.imageType || body.kind || body.dataUrl)) action = 'uploadImage';
    if (action === 'upsertItem') {
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
      return respond(uploadImageToDrive(body), e);
    }
    if (action === 'authStatus') {
      return respond(getDriveAuthStatus(), e);
    }
    if (action === 'deleteItem') {
      return respond(deleteMainItem({
        sheetName: body.sheet || body.sheetName,
        no: body.no
      }), e);
    }
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
