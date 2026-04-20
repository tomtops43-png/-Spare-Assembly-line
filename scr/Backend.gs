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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPARE_APP_CONFIG.writeSheetName);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ ' + SPARE_APP_CONFIG.writeSheetName);

    var body = JSON.parse(e.postData.contents);

    if (!body.partName || !body.qty) {
      throw new Error('ต้องมี partName และ qty');
    }

    var qty = Number(body.qty);
    if (body.type && String(body.type).indexOf('Output') > -1) {
      qty = -Math.abs(qty);
    } else {
      qty = Math.abs(qty);
    }

    sheet.appendRow([
      new Date(),
      body.type || 'Input',
      body.process || '-',
      body.category || 'General',
      body.partName,
      body.model || '-',
      body.brand || '-',
      qty,
      body.unit || 'PCS',
      body.by || 'Unknown'
    ]);

    return respond({ status: 'success' }, e);
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
