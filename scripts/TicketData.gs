/**
 * TicketData.gs - Ticket loader and open refresh
 */

const MAX_RUNTIME_MS = 5.5 * 60 * 1000;

function loadTicketsPaginated() {
  ensureSchoolYearLocked();
  assertSchoolYearUnchanged();
  cacheConfigRowPositions();
  const startTime = Date.now();
  const pageSize = getPageSize();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RawData');
  if (!sheet) throw new Error('RawData sheet not found.');

  setLoadState(DATA_LOAD_TYPES.TICKETS, LOAD_STATES.IN_PROGRESS);

  let page = getIntValue(getConfig('TICKET_LOAD_PAGE'), 0);
  if (page < 0) page = 0;

  const range = getSchoolYearRange();
  const filters = buildTicketFilters(range.startDate, range.endDate);

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    const response = searchTicketsPage(filters, page, pageSize);
    const items = response && response.Items ? response.Items : [];

    if (items.length === 0) {
      rebuildTicketIndex(sheet);
      setLoadState(DATA_LOAD_TYPES.TICKETS, LOAD_STATES.COMPLETE);
      writeConfigValueDirect('TICKET_LOAD_PAGE', '');
      writeConfigValueDirect('TICKET_LOAD_TOTAL_PAGES',
        response && response.Paging ? String(response.Paging.PageCount) : '');
      refreshLaborTypesFromTickets();
      return;
    }

    const rows = items.map(mapTicketRow);
    writeTicketPage(sheet, page, pageSize, rows);

    page++;
    writeConfigValueDirect('TICKET_LOAD_PAGE', String(page));
    if (response && response.Paging && response.Paging.PageCount !== undefined) {
      writeConfigValueDirect('TICKET_LOAD_TOTAL_PAGES', String(response.Paging.PageCount));
      if (page >= response.Paging.PageCount) {
        rebuildTicketIndex(sheet);
        setLoadState(DATA_LOAD_TYPES.TICKETS, LOAD_STATES.COMPLETE);
        writeConfigValueDirect('TICKET_LOAD_PAGE', '');
        refreshLaborTypesFromTickets();
        return;
      }
    }
  }

  logOperation('TICKET_LOAD', 'INFO', 'Paused due to time limit. Page ' + page);
}

function refreshOpenTickets() {
  const startTime = Date.now();
  const pageSize = getPageSize();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RawData');
  if (!sheet) throw new Error('RawData sheet not found.');

  if (getLoadState(DATA_LOAD_TYPES.TICKETS) !== LOAD_STATES.COMPLETE ||
      getLoadState(DATA_LOAD_TYPES.ACTIVITIES) !== LOAD_STATES.COMPLETE) {
    logOperation('OPEN_REFRESH', 'INFO', 'Initial load incomplete - skipping');
    return;
  }

  ensureSchoolYearLocked();
  assertSchoolYearUnchanged();
  cacheConfigRowPositions();

  setLoadState(DATA_LOAD_TYPES.OPEN_REFRESH, LOAD_STATES.IN_PROGRESS);

  const range = getSchoolYearRange();
  const days = getIntValue(getConfig('OPEN_REFRESH_DAYS'), 14);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  if (cutoff < range.startDate) {
    cutoff.setTime(range.startDate.getTime());
  }

  const ticketIndex = buildTicketIndexMap(sheet);

  let stage = getStringValue(getConfig('OPEN_REFRESH_STAGE')) || 'OPEN';
  if (!getConfig('OPEN_REFRESH_STAGE')) {
    setConfig('OPEN_REFRESH_STAGE', stage);
  }
  let page = 0;

  while (Date.now() - startTime < MAX_RUNTIME_MS) {
    if (stage === 'OPEN') {
      page = getIntValue(getConfig('OPEN_REFRESH_OPEN_PAGE'), 0);
      if (page < 0) page = 0;

      const filters = buildOpenTicketFilters(range.startDate, range.endDate, cutoff);
      const response = searchTicketsPage(filters, page, pageSize);
      const items = response && response.Items ? response.Items : [];

      if (items.length === 0) {
        setConfig('OPEN_REFRESH_OPEN_PAGE', '');
        stage = 'CLOSED';
        setConfig('OPEN_REFRESH_STAGE', stage);
        continue;
      }

      upsertTickets(items, sheet, ticketIndex);
      refreshActivitiesForTickets(items.map(t => t.TicketId));

      page++;
      setConfig('OPEN_REFRESH_OPEN_PAGE', String(page));

      if (response && response.Paging && page >= response.Paging.PageCount) {
        setConfig('OPEN_REFRESH_OPEN_PAGE', '');
        stage = 'CLOSED';
        setConfig('OPEN_REFRESH_STAGE', stage);
      }
    } else {
      page = getIntValue(getConfig('OPEN_REFRESH_CLOSED_PAGE'), 0);
      if (page < 0) page = 0;

      const filters = buildClosedTicketFilters(range.startDate, range.endDate, cutoff);
      const response = searchTicketsPage(filters, page, pageSize);
      const items = response && response.Items ? response.Items : [];

      if (items.length === 0) {
        finishOpenRefresh();
        return;
      }

      upsertTickets(items, sheet, ticketIndex);
      refreshActivitiesForTickets(items.map(t => t.TicketId));

      page++;
      setConfig('OPEN_REFRESH_CLOSED_PAGE', String(page));

      if (response && response.Paging && page >= response.Paging.PageCount) {
        finishOpenRefresh();
        return;
      }
    }
  }

  logOperation('OPEN_REFRESH', 'INFO', 'Paused due to time limit.');
}

function finishOpenRefresh() {
  setConfig('OPEN_REFRESH_STAGE', '');
  setConfig('OPEN_REFRESH_OPEN_PAGE', '');
  setConfig('OPEN_REFRESH_CLOSED_PAGE', '');
  setLoadState(DATA_LOAD_TYPES.OPEN_REFRESH, LOAD_STATES.COMPLETE);
  updateLastSync();
}

function searchTicketsPage(filters, page, pageSize) {
  const sortExpr = encodeURIComponent('TicketCreatedDate asc');
  const payload = {
    ProductId: PRODUCT_ID,
    Filters: filters || [],
    FilterByProduct: true,
    IncludeDeleted: false
  };

  const endpoint = '/v1.0/tickets?$p=' + page + '&$s=' + pageSize + '&$o=' + sortExpr;
  return apiRequest('POST', endpoint, payload);
}

function buildTicketFilters(startDate, endDate) {
  return [
    {
      Facet: 'totallabortime',
      Value: 'numoperator:greaterthan:0',
      Negative: false
    },
    {
      Facet: 'createddate',
      Value: 'daterange:' + formatDateForApi(startDate) + '-' + formatDateForApi(endDate),
      Negative: false
    }
  ];
}

function buildOpenTicketFilters(startDate, endDate, cutoffDate) {
  const filters = buildTicketFilters(startDate, endDate);
  if (cutoffDate) {
    filters.push({
      Facet: 'modifieddate',
      Value: 'daterange:' + formatDateForApi(cutoffDate) + '-' + formatDateForApi(new Date()),
      Negative: false
    });
  }
  filters.push({
    Facet: 'isclosed',
    Value: 'false',
    Negative: false,
    GroupIndex: 2
  });
  return filters;
}

function buildClosedTicketFilters(startDate, endDate, cutoffDate) {
  const filters = buildTicketFilters(startDate, endDate);
  filters.push({
    Facet: 'closeddate',
    Value: 'daterange:' + formatDateForApi(cutoffDate) + '-' + formatDateForApi(new Date()),
    Negative: false,
    GroupIndex: 2
  });
  return filters;
}

function mapTicketRow(t) {
  return [
    t.TicketId || '',
    t.TicketNumber || '',
    t.Subject || '',
    parseApiDate(t.CreatedDate || ''),
    parseApiDate(t.ClosedDate || ''),
    t.TotalLaborTime || 0,
    t.TotalLaborTime ? (t.TotalLaborTime / 60) : 0,
    t.TotalLaborCost || 0,
    t.LaborType ? t.LaborType.LaborTypeId || '' : '',
    t.LaborType ? t.LaborType.Name || '' : '',
    t.AssignedToUser ? t.AssignedToUser.Name || '' : '',
    t.AssignedToUser ? t.AssignedToUser.Email || '' : '',
    t.AssignedToTeam ? t.AssignedToTeam.TeamName || '' : '',
    t.Location ? t.Location.Name || '' : '',
    t.For ? t.For.Name || '' : '',
    t.WorkflowStep ? t.WorkflowStep.StatusName || '' : '',
    t.ResolutionAction || '',
    t.IsClosed === true ? 'Closed' : 'Open',
    t.AssignedToUserId || '',
    t.AssignedToTeamId || '',
    t.LocationId || ''
  ];
}

function writeTicketPage(sheet, page, pageSize, rows) {
  const startRow = page * pageSize + 2;
  const maxRows = pageSize;
  if (maxRows > 0) {
    sheet.getRange(startRow, 1, maxRows, sheet.getLastColumn()).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function buildTicketIndexMap(sheet) {
  const indexSheet = getTicketIndexSheet();
  let map = loadIndexMap(indexSheet);
  if (Object.keys(map).length === 0 && sheet.getLastRow() > 1) {
    rebuildTicketIndex(sheet);
    map = loadIndexMap(indexSheet);
  }
  return map;
}

function upsertTickets(items, sheet, indexMap) {
  const rowsToAppend = [];
  const appendIds = [];
  const updates = {};

  items.forEach(item => {
    const row = mapTicketRow(item);
    const ticketId = row[0];
    const existingRow = indexMap[ticketId];

    if (existingRow) {
      updates[existingRow] = row;
    } else {
      rowsToAppend.push(row);
      appendIds.push(ticketId);
    }
  });

  writeBatchedUpdates(sheet, updates);

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    const indexSheet = getTicketIndexSheet();
    const timestamp = new Date().toISOString();
    const indexRows = appendIds.map((id, idx) => [id, startRow + idx, timestamp]);
    appendIndexRows(indexSheet, indexRows);
    appendIds.forEach((id, idx) => {
      indexMap[id] = startRow + idx;
    });
  }
}

function writeBatchedUpdates(sheet, rowMap) {
  const rowIndexes = Object.keys(rowMap)
    .map(value => parseInt(value, 10))
    .filter(value => !isNaN(value))
    .sort((a, b) => a - b);

  if (rowIndexes.length === 0) return;

  let batchStart = rowIndexes[0];
  let batchRows = [rowMap[batchStart]];
  let previous = batchStart;

  for (let i = 1; i < rowIndexes.length; i++) {
    const rowIndex = rowIndexes[i];
    if (rowIndex === previous + 1) {
      batchRows.push(rowMap[rowIndex]);
    } else {
      sheet.getRange(batchStart, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
      batchStart = rowIndex;
      batchRows = [rowMap[rowIndex]];
    }
    previous = rowIndex;
  }

  if (batchRows.length > 0) {
    sheet.getRange(batchStart, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
  }
}

function refreshLaborTypesFromTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RawData');
  const laborSheet = ss.getSheetByName('LaborTypes');
  if (!rawSheet || !laborSheet) return;

  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) return;

  const values = rawSheet.getRange(2, 9, lastRow - 1, 2).getValues();
  const map = {};
  values.forEach(row => {
    const id = row[0];
    const name = row[1];
    if (id && name) {
      map[id] = name;
    }
  });

  const rows = Object.keys(map).map(id => [id, map[id]]);
  clearSheetData(laborSheet);
  if (rows.length > 0) {
    laborSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function formatDateForApi(date) {
  const d = new Date(date);
  var month = String(d.getMonth() + 1);
  var day = String(d.getDate());
  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;
  return month + '/' + day + '/' + d.getFullYear();
}
