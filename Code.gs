const SHEET_NAME = 'Sheet1';
const LOG_SHEET_NAME = 'LoginLog';
const RATES_SHEET_NAME = 'Rates';

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // Prefer first non-meta sheet that already has rows
    const skip = { [LOG_SHEET_NAME]: true, [RATES_SHEET_NAME]: true };
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (skip[s.getName()]) continue;
      if (s.getLastRow() > 1) {
        sheet = s;
        break;
      }
    }
    if (!sheet) {
      for (let i = 0; i < sheets.length; i++) {
        if (!skip[sheets[i].getName()]) {
          sheet = sheets[i];
          break;
        }
      }
    }
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Date', 'Timestamp', 'Data']);
    }
  }
  return sheet;
}

function getLogSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'IP']);
  }
  return sheet;
}

function getRatesSheet() {
  const ss = getSpreadsheet();
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

function normalizeDate(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(val || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function tryParseEntry(cell) {
  if (cell === null || cell === undefined || cell === '') return null;
  if (typeof cell === 'object' && !(cell instanceof Date)) {
    // Already an object somehow
    return cell;
  }
  const s = String(cell).trim();
  if (!s || s.charAt(0) !== '{') return null;
  try {
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

function readEntriesFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return { entries: [], parseErrors: 0, rows: 0 };

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const entries = [];
  let parseErrors = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Find JSON blob in any column
    let entry = null;
    let jsonCol = -1;
    for (let c = 0; c < row.length; c++) {
      entry = tryParseEntry(row[c]);
      if (!entry) entry = tryParseEntry(display[i][c]);
      if (entry) {
        jsonCol = c;
        break;
      }
    }
    if (!entry) {
      // Skip empty rows
      const nonempty = row.some(function (v) { return v !== '' && v !== null; });
      if (nonempty) parseErrors++;
      continue;
    }

    // Date: prefer column A, then entry.date
    let d = null;
    if (row[0] !== '' && row[0] !== null) d = normalizeDate(row[0]);
    else if (display[i][0]) d = normalizeDate(display[i][0]);
    if (!d || d === 'Invalid Date') d = normalizeDate(entry.date);
    entry.date = d;

    // Timestamp: column B if present
    if (row.length > 1 && row[1] !== '' && row[1] !== null) {
      entry.timestamp = row[1] instanceof Date
        ? row[1].toISOString()
        : String(row[1]);
    } else if (!entry.timestamp) {
      entry.timestamp = new Date().toISOString();
    }

    entries.push(entry);
  }

  return { entries: entries, parseErrors: parseErrors, rows: lastRow - 1 };
}

function readEntries() {
  const ss = getSpreadsheet();
  const skip = {};
  skip[LOG_SHEET_NAME] = true;
  skip[RATES_SHEET_NAME] = true;

  // Prefer Sheet1 if it has parseable entries
  let primary = ss.getSheetByName(SHEET_NAME);
  if (primary) {
    const r = readEntriesFromSheet(primary);
    if (r.entries.length) return r.entries;
  }

  // Scan all other sheets for entry data
  const sheets = ss.getSheets();
  let best = { entries: [], parseErrors: 0, rows: 0, name: null };
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    if (skip[s.getName()]) continue;
    const r = readEntriesFromSheet(s);
    if (r.entries.length > best.entries.length) {
      best = { entries: r.entries, parseErrors: r.parseErrors, rows: r.rows, name: s.getName() };
    }
  }
  return best.entries;
}

function sheetDiagnostics() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const out = [];
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const lastRow = s.getLastRow();
    const lastCol = s.getLastColumn();
    let sample = [];
    if (lastRow >= 1 && lastCol >= 1) {
      const vals = s.getRange(1, 1, Math.min(lastRow, 3), Math.min(lastCol, 4)).getDisplayValues();
      sample = vals;
    }
    const r = (s.getName() === LOG_SHEET_NAME || s.getName() === RATES_SHEET_NAME)
      ? { entries: [], parseErrors: 0, rows: lastRow }
      : readEntriesFromSheet(s);
    out.push({
      name: s.getName(),
      lastRow: lastRow,
      lastCol: lastCol,
      parsedEntries: r.entries ? r.entries.length : 0,
      parseErrors: r.parseErrors || 0,
      sample: sample
    });
  }
  return out;
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.debug === '1') {
    return ContentService.createTextOutput(JSON.stringify({
      spreadsheet: getSpreadsheet().getName(),
      spreadsheetId: getSpreadsheet().getId(),
      sheets: sheetDiagnostics(),
      entries: readEntries(),
      config: readConfig()
    })).setMimeType(ContentService.MimeType.JSON);
  }

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
      // Delete from any sheet that has the date
      const ss = getSpreadsheet();
      const skip = {};
      skip[LOG_SHEET_NAME] = true;
      skip[RATES_SHEET_NAME] = true;
      const targetDate = normalizeDate(body.date);
      const sheets = ss.getSheets();
      for (let s = 0; s < sheets.length; s++) {
        const sheet = sheets[s];
        if (skip[sheet.getName()]) continue;
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i >= 1; i--) {
          // Match by col A date OR JSON date
          let rowDateStr = normalizeDate(data[i][0]);
          if (!rowDateStr || rowDateStr === String(data[i][0])) {
            for (let c = 0; c < data[i].length; c++) {
              const ent = tryParseEntry(data[i][c]);
              if (ent && ent.date) {
                rowDateStr = normalizeDate(ent.date);
                break;
              }
            }
          }
          if (rowDateStr === targetDate) {
            sheet.deleteRow(i + 1);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'clear') {
      const sheet = getSheet();
      const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
      sheet.clearContents();
      if (header.some(function (c) { return c; })) sheet.appendRow(header);
      else sheet.appendRow(['Date', 'Timestamp', 'Data']);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: save entry onto primary data sheet
    const sheet = getSheet();
    // Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Timestamp', 'Data']);
    }
    const date = normalizeDate(body.date);
    const rest = {};
    for (const k in body) {
      if (k !== 'date') rest[k] = body[k];
    }
    const dataStr = JSON.stringify(rest);
    const existing = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < existing.length; i++) {
      const rowDate = normalizeDate(existing[i][0]);
      if (rowDate === date) {
        sheet.getRange(i + 1, 1).setValue(date);
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
