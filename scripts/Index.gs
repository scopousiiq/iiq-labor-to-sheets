/**
 * Index.gs - Persistent index sheets for fast row lookups
 */

const TICKET_INDEX_SHEET = 'TicketIndex';
const ACTIVITY_INDEX_SHEET = 'ActivityIndex';

function getTicketIndexSheet() {
  return getIndexSheet(TICKET_INDEX_SHEET, ['TicketId', 'RowNumber', 'UpdatedAt']);
}

function getActivityIndexSheet() {
  return getIndexSheet(ACTIVITY_INDEX_SHEET, ['ActivityId', 'RowNumber', 'UpdatedAt']);
}

function getIndexSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.hideSheet();
  }
  return sheet;
}

function loadIndexMap(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(row => {
    const id = row[0];
    const rowNumber = row[1];
    if (id && rowNumber) {
      map[id] = rowNumber;
    }
  });

  return map;
}

function rebuildTicketIndex(rawSheet) {
  const sheet = getTicketIndexSheet();
  const lastRow = rawSheet.getLastRow();
  clearSheetData(sheet);

  if (lastRow < 2) return;

  const ids = rawSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rows = [];
  ids.forEach((row, idx) => {
    const id = row[0];
    if (id) {
      rows.push([id, idx + 2, new Date().toISOString()]);
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function rebuildActivityIndex(activitySheet) {
  const sheet = getActivityIndexSheet();
  const lastRow = activitySheet.getLastRow();
  clearSheetData(sheet);

  if (lastRow < 2) return;

  const ids = activitySheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rows = [];
  ids.forEach((row, idx) => {
    const id = row[0];
    if (id) {
      rows.push([id, idx + 2, new Date().toISOString()]);
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function appendIndexRows(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function ensureDataSheetsProtected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('RawData');
  const activity = ss.getSheetByName('ActivityLog');
  if (raw) ensureSheetProtected(raw, 'Data Locked');
  if (activity) ensureSheetProtected(activity, 'Data Locked');

  const ticketIndex = ss.getSheetByName(TICKET_INDEX_SHEET);
  const activityIndex = ss.getSheetByName(ACTIVITY_INDEX_SHEET);
  if (ticketIndex) ensureSheetProtected(ticketIndex, 'Index Locked');
  if (activityIndex) ensureSheetProtected(activityIndex, 'Index Locked');
}

function ensureSheetProtected(sheet, description) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let protection = protections.find(p => p.getDescription() === description);
  if (!protection) {
    protection = sheet.protect();
    protection.setDescription(description);
  }

  protection.setWarningOnly(false);

  const editors = protection.getEditors();
  if (editors.length > 0) {
    protection.removeEditors(editors);
  }

  const user = Session.getEffectiveUser();
  if (user) {
    protection.addEditor(user);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
