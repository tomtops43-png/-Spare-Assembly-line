// =============================
// CONFIG
// =============================
var SHEET_NAME = 'Record';

// =============================
// GET (ดึงข้อมูล)
// =============================
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ Record');

    var data = sheet.getDataRange().getValues();
    var rows = data.slice(1);

    var result = rows.map(function (row) {
      return {
        timestamp: row[0],
        type: row[1],
        process: row[2],
        category: row[3],
        partName: row[4],
        model: row[5],
        brand: row[6],
        qty: row[7],
        unit: row[8],
        by: row[9]
      };
    });

    return respond(result, e);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
  }
}

// =============================
// POST (เพิ่มข้อมูล)
// =============================
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ Record');

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
