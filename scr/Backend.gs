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
  var requiredHints = ['no', 'name', 'category', 'brand', 'stock'];
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

    if (hit >= 3) return r;
  }

  return 0;
}

function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
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
  var mainSheet = spreadsheet.getSheetByName(resolvedSheetName);

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

  var stockCol = map.stockqty !== undefined ? map.stockqty : map.stock;
  var minCol = map.min;
  var needPoCol = map.needtopo !== undefined ? map.needtopo : map.needpo;

  if (stockCol === undefined) throw new Error('ไม่พบคอลัมน์ stock/stock qty');

  var targetIndex = -1;
  for (var i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    var rowNo = pickRowValue(row, map, ['no'], '');
    var rowName = pickRowValue(row, map, ['namedescriptions', 'name', 'description'], '');
    var rowModel = pickRowValue(row, map, ['model'], '');
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
  var sheet = spreadsheet.getSheetByName(sheetName);
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
    name: findCol(['namedescriptions', 'name', 'description']),
    model: findCol(['model']),
    line: findCol(['mainline', 'line']),
    category: findCol(['category']),
    brand: findCol(['brand']),
    photo: findCol(['sparepartsphotos', 'photo', 'image', 'imageurl', 'picture']),
    max: findCol(['max', 'qtymax']),
    min: findCol(['min', 'qtymin']),
    unit: findCol(['unit']),
    stock: findCol(['stockqty', 'stock', 'initialstock'])
  };

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
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : '';
    if (action === 'transact') return respond(processTransaction(parseTransactionPayloadFromGet(e)), e);
    if (action === 'logs') return respond(getLogRows(), e);
    if (action === 'upsertItem') return respond(upsertMainItem({
      sheetName: e.parameter.sheet,
      no: e.parameter.no,
      name: e.parameter.name,
      model: e.parameter.model,
      line: e.parameter.line,
      category: e.parameter.category,
      brand: e.parameter.brand,
      photo: e.parameter.photo,
      max: e.parameter.max,
      min: e.parameter.min,
      unit: e.parameter.unit,
      stock: e.parameter.stock
    }), e);
    if (action === 'deleteItem') return respond(deleteMainItem({ sheetName: e.parameter.sheet, no: e.parameter.no }), e);

    var sheetName = resolveReadSheetName({ sheet: e.parameter.sheet });
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ ' + sheetName);

    var data = sheet.getDataRange().getValues();
    if (!data.length) return respond([], e);

    var headerRowIndex = findHeaderRowIndex(data);
    var headers = data[headerRowIndex];
    var rows = data.slice(headerRowIndex + 1);
    var map = buildHeaderIndexMap(headers);

    var result = rows.map(function (row, index) {
      return {
        no: pickRowValue(row, map, ['no'], index + 1),
        name: pickRowValue(row, map, ['namedescriptions', 'name', 'description'], '-'),
        model: pickRowValue(row, map, ['model'], '-'),
        line: pickRowValue(row, map, ['mainline', 'line'], '-'),
        category: pickRowValue(row, map, ['category'], 'General'),
        brand: pickRowValue(row, map, ['brand'], '-'),
        stock: pickRowValue(row, map, ['stockqty', 'stock'], 0),
        max: pickRowValue(row, map, ['max', 'qtymax'], 0),
        min: pickRowValue(row, map, ['min', 'qtymin'], 0),
        needToPO: pickRowValue(row, map, ['needtopo', 'needpo'], 0),
        unit: pickRowValue(row, map, ['unit'], 'PCS'),
        remark: pickRowValue(row, map, ['remark'], ''),
        photo: pickRowValue(row, map, ['sparepartsphotos', 'photo', 'image', 'imageurl', 'picture'], '')
      };
    }).filter(function (item) {
      if (!item.name || item.name === '-') return false;
      var lineText = String(item.line || '').toLowerCase();
      return lineText.indexOf('h9') > -1 || lineText === '' || item.line === '-';
    });

    return respond(result, e);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
  }
}

// =============================
// POST (transaction)
// =============================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body && body.action ? String(body.action) : '';
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
        max: body.max,
        min: body.min,
        unit: body.unit,
        stock: body.stock
      }), e);
    }
    if (action === 'deleteItem') {
      return respond(deleteMainItem({
        sheetName: body.sheet || body.sheetName,
        no: body.no
      }), e);
    }
    return respond(processTransaction(body), e);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
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
