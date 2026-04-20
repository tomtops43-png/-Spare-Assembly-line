// =============================
// CONFIG
// =============================
var SPARE_APP_CONFIG = this.SPARE_APP_CONFIG || {};
SPARE_APP_CONFIG.readSheetName = SPARE_APP_CONFIG.readSheetName || 'Main List Stock';
SPARE_APP_CONFIG.writeSheetName = SPARE_APP_CONFIG.writeSheetName || 'Record';

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

// =============================
// GET (ดึงข้อมูลจาก Main List Stock)
// =============================
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPARE_APP_CONFIG.readSheetName);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ ' + SPARE_APP_CONFIG.readSheetName);

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
        photo: pickRowValue(row, map, ['sparepartsphotos', 'photo'], '')
      };
    }).filter(function (item) {
      return item.name && item.name !== '-';
    });

    return respond(result, e);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
  }
}

// =============================
// POST (เพิ่มข้อมูลลง Record)
// =============================
function doPost(e) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var historySheet = spreadsheet.getSheetByName(SPARE_APP_CONFIG.writeSheetName);
    var mainSheet = spreadsheet.getSheetByName(SPARE_APP_CONFIG.readSheetName);

    if (!historySheet) throw new Error('ไม่พบชีทชื่อ ' + SPARE_APP_CONFIG.writeSheetName);
    if (!mainSheet) throw new Error('ไม่พบชีทชื่อ ' + SPARE_APP_CONFIG.readSheetName);

    var body = JSON.parse(e.postData.contents);
    if (!body.partName || !body.qty) throw new Error('ต้องมี partName และ qty');

    var qty = Number(body.qty);
    if (!qty || qty <= 0) throw new Error('qty ต้องมากกว่า 0');

    var signedQty = qty;
    if (body.type && String(body.type).indexOf('Output') > -1) {
      signedQty = -Math.abs(qty);
    } else {
      signedQty = Math.abs(qty);
    }

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

      var noMatch = body.partNo !== undefined && String(rowNo) === String(body.partNo);
      var nameMatch = String(rowName) === String(body.partName);
      var modelMatch = !body.model || String(rowModel) === String(body.model);

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
      body.type || 'Input',
      body.process || '-',
      body.category || 'General',
      body.partName,
      body.model || '-',
      body.brand || '-',
      signedQty,
      body.unit || 'PCS',
      body.by || 'Unknown',
      body.partNo || '',
      stockBefore,
      stockAfter
    ]);

    return respond({
      status: 'success',
      stockBefore: stockBefore,
      stockAfter: stockAfter,
      qty: signedQty
    }, e);
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
