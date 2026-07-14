const SHEET_NAME = 'Sheet1';
const LOG_SHEET_NAME = 'LoginLog';
const RATES_SHEET_NAME = 'Rates';

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

function getRatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RATES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RATES_SHEET_NAME);
    sheet.appendRow(['ConfigJSON']);
  }
  return sheet;
}

function readConfig() {
  const sheet = getRatesSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) return null;
  try {
    return JSON.parse(String(data[1][0]));
  } catch (e) {
    return null;
  }
}

function writeConfig(config) {
  const sheet = getRatesSheet();
  sheet.clearContents();
  sheet.appendRow(['ConfigJSON']);
  sheet.appendRow([JSON.stringify(config)]);
}

function readEntries() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    try {
      const entry = JSON.parse(row[2]);
      // Normalize date to yyyy-mm-dd string
      let d = row[0];
      if (d instanceof Date) {
        d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        d = String(d);
      }
      entry.date = d;
      entry.timestamp = row[1] instanceof Date
        ? row[1].toISOString()
        : String(row[1]);
      entries.push(entry);
    } catch (e) {}
  }
  return entries;
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    entries: readEntries(),
    config: readConfig()
  })).setMimeType(ContentService.MimeType.JSON);
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

    if (body.action === 'saveConfig') {
      if (!body.config) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'No config' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      writeConfig(body.config);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'delete') {
      const sheet = getSheet();
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        const rowDate = data[i][0];
        let rowDateStr = rowDate instanceof Date
          ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(rowDate);
        const targetDate = String(body.date);
        if (rowDateStr === targetDate) {
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
      let rowDate = existing[i][0];
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        rowDate = String(rowDate);
      }
      if (rowDate === date) {
        sheet.getRange(i + 1, 2).setValue(new Date().toISOString());
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
