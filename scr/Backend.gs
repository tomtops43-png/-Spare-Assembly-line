// =============================
// CONFIG
// =============================
const SHEET_NAME = 'Record';

// =============================
// GET (ดึงข้อมูล)
// =============================
function doGet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ Record');

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);

    const result = rows.map((row) => ({
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
    }));

    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

// =============================
// POST (เพิ่มข้อมูล)
// =============================
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีทชื่อ Record');

    const body = JSON.parse(e.postData.contents || '{}');

    if (!body.partName || body.qty === undefined || body.qty === null || body.qty === '') {
      throw new Error('ต้องมี partName และ qty');
    }

    let qty = Number(body.qty);
    if (body.type && body.type.includes('Output')) {
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

    return jsonOutput({ status: 'success' });
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

// =============================
// OPTIONS
// =============================
function doOptions() {
  return jsonOutput({ status: 'ok' });
}

// =============================
// JSON RESPONSE
// =============================
function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
