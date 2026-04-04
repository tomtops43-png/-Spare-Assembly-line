/**
 * Google Apps Script backend for H9 spare-parts portal.
 * Deploy as Web app: Execute as Me, access Anyone.
 */
function doGet() {
  const sheetName = 'Stock H9';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    return jsonResponse({
      ok: false,
      error: `Sheet not found: ${sheetName}`,
      data: []
    });
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) {
    return jsonResponse({ ok: true, data: [] });
  }

  const headers = values[0].map((h) => String(h || '').trim());
  const data = values.slice(1).map((row) => {
    const item = {};
    headers.forEach((header, i) => {
      item[header] = row[i] ?? '';
    });
    return item;
  });

  return jsonResponse({ ok: true, data });
}

function doPost(e) {
  return jsonResponse({ ok: true, message: 'POST ready', payload: e && e.postData ? e.postData.contents : null });
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
