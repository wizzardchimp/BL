// Primary daily-entry tab (your data lives here)
const SHEET_NAME = 'Data';
// Also try these if Data is empty / missing
const SHEET_FALLBACKS = ['Sheet1', 'Entries'];
const LOG_SHEET_NAME = 'LoginLog';
const RATES_SHEET_NAME = 'Rates';

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    for (let i = 0; i < SHEET_FALLBACKS.length; i++) {
      sheet = ss.getSheetByName(SHEET_FALLBACKS[i]);
      if (sheet) break;
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Date', 'Timestamp', 'Data']);
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
  if (typeof cell === 'object' && !(cell instanceof Date)) return cell;
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
    let entry = null;
    // Prefer column C (index 2), then scan all columns
    if (row.length > 2) {
      entry = tryParseEntry(row[2]) || tryParseEntry(display[i][2]);
    }
    if (!entry) {
      for (let c = 0; c < row.length; c++) {
        entry = tryParseEntry(row[c]) || tryParseEntry(display[i][c]);
        if (entry) break;
      }
    }
    if (!entry) {
      const nonempty = row.some(function (v) { return v !== '' && v !== null; });
      if (nonempty) parseErrors++;
      continue;
    }

    let d = null;
    if (row[0] !== '' && row[0] !== null) d = normalizeDate(row[0]);
    else if (display[i][0]) d = normalizeDate(display[i][0]);
    if (!d) d = normalizeDate(entry.date);
    entry.date = d;

    if (row.length > 1 && row[1] !== '' && row[1] !== null) {
      entry.timestamp = row[1] instanceof Date
        ? row[1].toISOString()
        : String(row[1]);
    } else if (!entry.timestamp) {
      entry.timestamp = new Date().toISOString();
    }

    // Skip non-entry rows (e.g. login events wrongly saved to Data)
    if (entry.action) continue;
    if (!entry.locations || !entry.locations.length) {
      if (!entry.grandTotal && !entry.grandCL) continue;
    }
    // Ensure stable id for duplicate-date rows
    if (!entry.id) entry.id = entry.timestamp || entry.date;

    entries.push(entry);
  }
  return { entries: entries, parseErrors: parseErrors, rows: lastRow - 1 };
}

function readEntries() {
  const ss = getSpreadsheet();
  // Prefer Data, then fallbacks — use first sheet that has parseable entries
  const names = [SHEET_NAME].concat(SHEET_FALLBACKS);
  let best = [];
  for (let i = 0; i < names.length; i++) {
    const sheet = ss.getSheetByName(names[i]);
    if (!sheet) continue;
    const r = readEntriesFromSheet(sheet);
    if (r.entries.length > best.length) best = r.entries;
  }
  // Last resort: any other non-meta sheet
  if (!best.length) {
    const skip = {};
    skip[LOG_SHEET_NAME] = true;
    skip[RATES_SHEET_NAME] = true;
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (skip[s.getName()]) continue;
      const r = readEntriesFromSheet(s);
      if (r.entries.length > best.length) best = r.entries;
    }
  }
  return best;
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
      sample = s.getRange(1, 1, Math.min(lastRow, 3), Math.min(lastCol, 4)).getDisplayValues();
    }
    let parsed = 0;
    let parseErrors = 0;
    if (s.getName() !== LOG_SHEET_NAME && s.getName() !== RATES_SHEET_NAME) {
      const r = readEntriesFromSheet(s);
      parsed = r.entries.length;
      parseErrors = r.parseErrors;
    }
    out.push({
      name: s.getName(),
      lastRow: lastRow,
      lastCol: lastCol,
      parsedEntries: parsed,
      parseErrors: parseErrors,
      sample: sample
    });
  }
  return out;
}

/**
 * Run once from the Apps Script editor (select cleanupDataSheet → Run)
 * Removes login/control rows and empty £0 rows from Data (and Sheet1).
 */
function cleanupDataSheet() {
  const ss = getSpreadsheet();
  const names = [SHEET_NAME].concat(SHEET_FALLBACKS);
  let totalRemoved = 0;
  const seen = {};
  for (let n = 0; n < names.length; n++) {
    const name = names[n];
    if (seen[name]) continue;
    seen[name] = true;
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    totalRemoved += cleanupSheetRows_(sheet);
  }
  Logger.log('Removed ' + totalRemoved + ' junk row(s)');
  return { success: true, removed: totalRemoved };
}

function cleanupSheetRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return 0;

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  let removed = 0;

  // Delete from bottom so row indexes stay valid
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    let entry = null;
    if (row.length > 2) {
      entry = tryParseEntry(row[2]) || tryParseEntry(display[i][2]);
    }
    if (!entry) {
      for (let c = 0; c < row.length; c++) {
        entry = tryParseEntry(row[c]) || tryParseEntry(display[i][c]);
        if (entry) break;
      }
    }

    let junk = false;
    if (entry) {
      if (entry.action) junk = true;
      else if (!entry.locations || !entry.locations.length) {
        if (!entry.grandTotal && !entry.grandCL) junk = true;
      }
    } else {
      // Non-empty row with no parseable JSON
      const nonempty = row.some(function (v) { return v !== '' && v !== null; });
      if (nonempty) {
        // Keep header-like rows; remove pure timestamp-only leftovers if col C empty
        const hasDate = row[0] !== '' && row[0] !== null;
        const hasJson = false;
        if (hasDate && !hasJson && (row[2] === '' || row[2] === null)) {
          // leave normal empty template rows alone only if entirely empty after col A/B
          // Don't delete unknown non-JSON data blindly
        }
      }
    }

    if (junk) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  return removed;
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.debug === '1') {
    return ContentService.createTextOutput(JSON.stringify({
      spreadsheet: getSpreadsheet().getName(),
      spreadsheetId: getSpreadsheet().getId(),
      primarySheet: SHEET_NAME,
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

    if (body.action === 'cleanup') {
      const result = cleanupDataSheet();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'delete') {
      // Prefer unique id/timestamp so same-date entries can coexist
      const targetId = body.id ? String(body.id) : '';
      const targetTs = body.timestamp ? String(body.timestamp) : '';
      const targetDate = body.date ? normalizeDate(body.date) : '';
      const ss = getSpreadsheet();
      const skip = {};
      skip[LOG_SHEET_NAME] = true;
      skip[RATES_SHEET_NAME] = true;
      const sheets = ss.getSheets();
      for (let s = 0; s < sheets.length; s++) {
        const sheet = sheets[s];
        if (skip[sheet.getName()]) continue;
        const data = sheet.getDataRange().getValues();
        const display = sheet.getDataRange().getDisplayValues();
        for (let i = data.length - 1; i >= 1; i--) {
          let entry = null;
          if (data[i].length > 2) {
            entry = tryParseEntry(data[i][2]) || tryParseEntry(display[i][2]);
          }
          if (!entry) {
            for (let c = 0; c < data[i].length; c++) {
              entry = tryParseEntry(data[i][c]) || tryParseEntry(display[i][c]);
              if (entry) break;
            }
          }
          if (!entry) continue;

          const rowId = entry.id ? String(entry.id) : '';
          const rowTs = entry.timestamp ? String(entry.timestamp) : String(data[i][1] || '');
          const rowDate = normalizeDate(data[i][0] || entry.date);

          let match = false;
          if (targetId && rowId && rowId === targetId) match = true;
          else if (targetTs && rowTs && rowTs === targetTs) match = true;
          else if (!targetId && !targetTs && targetDate && rowDate === targetDate) match = true;

          if (match) {
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

    // Default: always APPEND entry (allow multiple entries on same date)
    const sheet = getSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Timestamp', 'Data']);
    }
    const date = normalizeDate(body.date);
    const ts = body.timestamp || new Date().toISOString();
    const id = body.id || ts;
    const rest = {};
    for (const k in body) {
      if (k !== 'date') rest[k] = body[k];
    }
    rest.timestamp = ts;
    rest.id = id;
    const dataStr = JSON.stringify(rest);
    sheet.appendRow([date, ts, dataStr]);
    return ContentService.createTextOutput(JSON.stringify({ success: true, id: id, timestamp: ts }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
