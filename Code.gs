const SHEET_NAME = 'Sheet1';
const LOG_SHEET_NAME = 'LoginLog';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Date', 'Timestamp', 'Data']);
  }
  return sheet;
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'IP']);
  }
  return sheet;
}

function doGet() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    try {
      const entry = JSON.parse(row[2]);
      entry.date = row[0];
      entry.timestamp = row[1];
      entries.push(entry);
    } catch (e) {}
  }
  return ContentService.createTextOutput(JSON.stringify(entries))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'login') {
      const logSheet = getLogSheet();
      logSheet.appendRow([body.timestamp || new Date().toISOString(), '']);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'delete') {
      const sheet = getSheet();
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        const rowDate = String(data[i][0]);
        const targetDate = String(body.date);

        const d1 = new Date(rowDate);
        const d2 = new Date(targetDate);
        if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
          if (d1.getTime() === d2.getTime()) {
            sheet.deleteRow(i + 1);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
        if (rowDate === targetDate) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'clear') {
      const sheet = getSheet();
      const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      sheet.clearContents();
      if (header.some(c => c)) sheet.appendRow(header);
      else sheet.appendRow(['Date', 'Timestamp', 'Data']);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: save entry
    const sheet = getSheet();
    const { date, ...rest } = body;
    const dataStr = JSON.stringify(rest);
    const existing = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < existing.length; i++) {
      const rowDate = String(existing[i][0]);
      if (rowDate === date) {
        sheet.getRange(i + 1, 3).setValue(dataStr);
        updated = true;
        break;
      }
    }
    if (!updated) {
      sheet.appendRow([date, new Date().toISOString(), dataStr]);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
