/**
 * ActivityLog.gs - Activity loader and refresh
 */

function loadActivitiesInitial() {
  ensureSchoolYearLocked();
  assertSchoolYearUnchanged();
  cacheConfigRowPositions();
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RawData');
  const actSheet = ss.getSheetByName('ActivityLog');
  if (!rawSheet || !actSheet) throw new Error('RawData or ActivityLog sheet not found.');

  setLoadState(DATA_LOAD_TYPES.ACTIVITIES, LOAD_STATES.IN_PROGRESS);

  const ticketIds = getTicketIds(rawSheet);
  if (ticketIds.length === 0) {
    setLoadState(DATA_LOAD_TYPES.ACTIVITIES, LOAD_STATES.COMPLETE);
    return;
  }

  const ticketMap = buildTicketContextMap(rawSheet);
  const userMap = buildUserMap();
  const laborMap = buildLaborTypeMap();
  const resolutionMap = buildResolutionActionMap();

  const activityIndex = buildActivityIndexMap(actSheet);
  const batchSize = getActivityBatchSize();

  // Retry failed tickets first
  if (!processFailedActivityTickets(startTime, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap)) {
    return;
  }

  let index = getIntValue(getConfig('ACTIVITY_TICKET_INDEX'), 0);
  if (index < 0) index = 0;

  let startPage = getIntValue(getConfig('ACTIVITY_BATCH_PAGE'), 0);
  if (startPage < 0) startPage = 0;

  var totalBatches = Math.ceil(ticketIds.length / batchSize);

  for (let i = index; i < ticketIds.length; i += batchSize) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      writeConfigValueDirect('ACTIVITY_TICKET_INDEX', String(i));
      var batchNum = Math.floor(i / batchSize) + 1;
      logOperation('ACTIVITY_LOAD', 'INFO',
        'Paused — processed ' + i + '/' + ticketIds.length + ' tickets (' +
        batchNum + '/' + totalBatches + ' batches). Will resume on next trigger.');
      return;
    }

    const batchIds = ticketIds.slice(i, i + batchSize);
    const batchEnd = Math.min(i + batchSize, ticketIds.length);
    const batchLabel = 'tickets ' + (i + 1) + '-' + batchEnd + ' of ' + ticketIds.length;
    const pageStart = i === index ? startPage : 0;

    logOperation('ACTIVITY_LOAD', 'INFO',
      'Fetching activities for ' + batchLabel +
      (pageStart > 0 ? ' (resuming at activity page ' + pageStart + ')' : ''));

    const result = loadActivitiesBatch(batchIds, pageStart, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap, startTime);
    if (!result.completed) {
      writeConfigValueDirect('ACTIVITY_TICKET_INDEX', String(i));
      writeConfigValueDirect('ACTIVITY_BATCH_PAGE', String(result.nextPage));
      logOperation('ACTIVITY_LOAD', 'INFO',
        'Paused mid-batch at ' + batchLabel + ', activity page ' + result.nextPage +
        '. Will resume on next trigger.');
      return;
    }

    logOperation('ACTIVITY_LOAD', 'INFO',
      'Completed ' + batchLabel + ' — ' + result.activitiesWritten + ' activities written');

    writeConfigValueDirect('ACTIVITY_LAST_TICKET_ID', String(batchIds[batchIds.length - 1]));
    writeConfigValueDirect('ACTIVITY_TICKET_INDEX', String(i + batchSize));
    writeConfigValueDirect('ACTIVITY_BATCH_PAGE', '');
  }

  // If failures remain, keep in progress for retries
  if (hasActivityFailures()) {
    logOperation('ACTIVITY_LOAD', 'WARNING', 'Failures remain. Will retry on next run.');
    return;
  }

  setLoadState(DATA_LOAD_TYPES.ACTIVITIES, LOAD_STATES.COMPLETE);
  writeConfigValueDirect('ACTIVITY_TICKET_INDEX', '');
  updateLastSync();

  // De-duplicate action/labor effort into NetHours now that every row is loaded.
  try {
    computeNetHours();
  } catch (e) {
    logOperation('NET_HOURS', 'ERROR', 'NetHours computation failed: ' + e.message);
  }

  // Run post-load validation and trigger reconcile if needed
  try {
    validateAndTriggerReconcile();
  } catch (e) {
    logOperation('VALIDATION', 'ERROR', 'Post-load validation failed: ' + e.message);
  }
}

function refreshActivitiesForTickets(ticketIds) {
  ensureSchoolYearLocked();
  assertSchoolYearUnchanged();
  if (!ticketIds || ticketIds.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RawData');
  const actSheet = ss.getSheetByName('ActivityLog');
  if (!rawSheet || !actSheet) return;

  const ticketMap = buildTicketContextMap(rawSheet);
  const userMap = buildUserMap();
  const laborMap = buildLaborTypeMap();
  const resolutionMap = buildResolutionActionMap();
  const activityIndex = buildActivityIndexMap(actSheet);
  const batchSize = getActivityBatchSize();

  for (let i = 0; i < ticketIds.length; i += batchSize) {
    const batchIds = ticketIds.slice(i, i + batchSize);
    loadActivitiesBatch(batchIds, 0, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap, Date.now());
  }

  // Refreshed rows may have changed a ticket's action/labor mix — recompute NetHours.
  try {
    computeNetHours();
  } catch (e) {
    logOperation('NET_HOURS', 'ERROR', 'NetHours computation failed: ' + e.message);
  }
}

function loadActivitiesBatch(ticketIds, startPage, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap, startTime) {
  let page = startPage || 0;
  const pageSize = getPageSize();
  let hasMore = true;
  let activitiesWritten = 0;
  const responseTicketIds = {};
  const ticketsWithActivities = {};

  while (hasMore) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      return { completed: false, nextPage: page, activitiesWritten: activitiesWritten };
    }

    try {
      const endpoint = '/v1.0/tickets/activities?$p=' + page + '&$s=' + pageSize;
      const response = apiRequest('POST', endpoint, ticketIds);
      const entries = extractActivityEntries(response);
      const rows = mapActivityEntries(entries, ticketMap, userMap, laborMap, resolutionMap, responseTicketIds, ticketsWithActivities);

      writeActivities(rows, actSheet, activityIndex);
      activitiesWritten += rows.length;

      if (response && response.Paging && response.Paging.PageCount !== undefined) {
        page += 1;
        hasMore = page < response.Paging.PageCount;
      } else {
        hasMore = entries.length >= pageSize;
        page += 1;
      }
    } catch (error) {
      ticketIds.forEach(ticketId => recordActivityFailure(ticketId, error.message));
      return { completed: false, nextPage: page, activitiesWritten: activitiesWritten };
    }
  }

  reconcileBatchTicketFailures(ticketIds, responseTicketIds, ticketsWithActivities, startPage || 0);
  return { completed: true, nextPage: page, activitiesWritten: activitiesWritten };
}

function reconcileBatchTicketFailures(ticketIds, responseTicketIds, ticketsWithActivities, startPage) {
  const skipMissingTicketFailures = startPage > 0;
  if (skipMissingTicketFailures) {
    logOperation('ACTIVITY_LOAD', 'INFO', 'Batch resumed at page ' + startPage + '; skipping missing-ticket failure checks for this run.');
  }

  var zeroActivityCount = 0;
  var missingFromResponseCount = 0;

  ticketIds.forEach(ticketId => {
    if (responseTicketIds[ticketId]) {
      clearActivityFailure(ticketId);
      if (!ticketsWithActivities[ticketId]) {
        zeroActivityCount++;
      }
      return;
    }

    if (skipMissingTicketFailures) return;
    missingFromResponseCount++;
    recordActivityFailure(ticketId, 'No activity response returned in successful batch');
  });

  if (zeroActivityCount > 0) {
    logOperation('ACTIVITY_LOAD', 'INFO',
      zeroActivityCount + ' of ' + ticketIds.length + ' tickets had zero activities in batch response.');
  }
  if (missingFromResponseCount > 0) {
    logOperation('ACTIVITY_LOAD', 'WARNING',
      missingFromResponseCount + ' of ' + ticketIds.length + ' tickets missing from batch response. Recorded as failures.');
  }
}

function loadActivitiesForTicket(ticketId, ticketMap, userMap, laborMap, resolutionMap) {
  const response = apiRequest('GET', '/v1.0/tickets/' + ticketId + '/timeline', null);
  const entries = extractTimelineEntries(response);
  const rows = [];
  const ticketContext = ticketMap[ticketId] || {};

  entries.forEach(entry => {
    const items = entry.ActivityItems || entry.activityItems || [];
    items.forEach(item => {
      if (isEffortMissing(item)) return;
      if (item.$type && item.$type.indexOf('TicketActivityAction') === -1) return;

      const row = buildActivityRow(item, entry, ticketId, ticketContext, userMap, laborMap, resolutionMap);
      if (row) {
        rows.push(row);
      }
    });
  });

  return rows;
}

function extractTimelineEntries(response) {
  if (!response) return [];
  if (Array.isArray(response)) return response;
  if (response.Items) return response.Items;
  if (response.TicketActivityLogEntry) return [response.TicketActivityLogEntry];
  if (response.TicketActivityLogEntries) return response.TicketActivityLogEntries;
  return [];
}

function extractActivityEntries(response) {
  if (!response) return [];
  if (Array.isArray(response)) return response;
  if (response.Items) return response.Items;
  return [];
}

function mapActivityEntries(entries, ticketMap, userMap, laborMap, resolutionMap, responseTicketIds, ticketsWithActivities) {
  const rows = [];

  entries.forEach(entry => {
    const ticketId = entry.TicketId || entry.ticketId || '';
    if (!ticketId) return;
    responseTicketIds[ticketId] = true;
    const ticketContext = ticketMap[ticketId] || {};
    const items = entry.ActivityItems || entry.activityItems || [];
    let hasActivities = false;

    items.forEach(item => {
      if (isEffortMissing(item)) return;
      if (item.$type && item.$type.indexOf('TicketActivityAction') === -1) return;
      const row = buildActivityRow(item, entry, ticketId, ticketContext, userMap, laborMap, resolutionMap);
      if (row) {
        rows.push(row);
        hasActivities = true;
      }
    });

    if (hasActivities) {
      ticketsWithActivities[ticketId] = true;
    }
  });

  return rows;
}

function isEffortMissing(item) {
  return !item || item.Effort === null || item.Effort === undefined;
}

function buildActivityRow(item, entry, ticketId, ticketContext, userMap, laborMap, resolutionMap) {
  const activityId = normalizeActivityId(item.TicketActivityActionId || item.TicketActivityId);
  if (!activityId) return null;

  const effortMins = item.Effort === null || item.Effort === undefined ? 0 : item.Effort;
  const userId = item.ByUserId || '';
  const laborTypeId = item.LaborTypeId || '';
  const resolutionId = item.ResolutionActionId || '';

  const user = userMap[userId] || {};
  const laborName = item.LaborTypeName || (item.LaborType ? item.LaborType.Name : '') || laborMap[laborTypeId] || '';
  const resolutionName = item.ResolutionAction || resolutionMap[resolutionId] || '';

  let userName = user.name || item.ByUserName || item.ByUser || '';
  let effectiveUserId = userId;

  // Attribution: rows logged under the "System Service" automation account carry
  // no real technician. Re-attribute them to the ticket's assigned user so the
  // work counts toward a person rather than the service account.
  if (isSystemServiceName_(userName) && ticketContext.assignedUserName) {
    userName = ticketContext.assignedUserName;
    effectiveUserId = ticketContext.assignedUserId || userId;
  }

  // EntryType distinguishes the two records iiQ returns for the same work:
  // a labor entry (has LaborType) vs. a resolution action (has ResolutionAction).
  const entryType = laborTypeId ? 'Labor' : 'Action';

  // Team: prefer the ticket's assigned team, then fall back to the performing
  // user's team. Only ~17% of tickets carry an assigned team but ~99% have a
  // user, so a ticket-only rule would drop most hours from ByTeam. Coalesce a
  // missing team to UNASSIGNED so the row is bucketed, never silently dropped.
  const teamSource = userMap[effectiveUserId] || user || {};
  const teamId = ticketContext.assignedTeamId || teamSource.teamId || '';
  const teamName = ticketContext.assignedTeamName || teamSource.teamName || UNASSIGNED;

  return [
    activityId,
    ticketId,
    ticketContext.ticketNumber || '',
    parseApiDate(item.ActivityDate || entry.CreatedDate || ''),
    effortMins,
    effortMins / 60,
    item.HourlyRate || 0,
    item.LaborCost || 0,
    laborTypeId,
    laborName,
    resolutionId,
    resolutionName,
    effectiveUserId,
    userName,
    item.Notes || '',
    item.IsPublic === true ? 1 : 0,
    teamId,
    teamName,
    ticketContext.locationId || '',
    // Coalesce dimension labels so blanks bucket under UNASSIGNED and every
    // full breakdown (ByTeam/ByDepartment/ByIssueCategory/ByIssueType) reconciles
    // to the same total. (PerformedByUser is intentionally NOT coalesced — blank
    // names are repaired by the Backfill Missing User Names tool instead.)
    ticketContext.locationName || UNASSIGNED,
    ticketContext.issueCategoryId || '',
    ticketContext.issueCategoryName || UNASSIGNED,
    ticketContext.issueTypeId || '',
    ticketContext.issueTypeName || UNASSIGNED,
    entryType,
    0  // NetHours placeholder — filled by computeNetHours() after the full load.
  ];
}

// Bucket label for rows whose team/location/category/issue-type is blank, so the
// row still appears in (and counts toward the total of) every full breakdown.
const UNASSIGNED = '(Unassigned)';

// Names treated as the iiQ system/automation account (not a real technician).
// Activity rows under these names are re-attributed to the ticket's assigned user.
const SYSTEM_SERVICE_NAMES = ['System Service'];

function isSystemServiceName_(name) {
  if (!name) return false;
  return SYSTEM_SERVICE_NAMES.indexOf(String(name).trim()) !== -1;
}

// Recomputes EntryType (col Y) and NetHours (col Z) for every ActivityLog row.
//
// iiQ records the same work twice for many tickets: once as a resolution Action
// (carries effort) and once as a Labor entry (carries the same effort). Summing
// raw EffortHours double-counts. NetHours de-duplicates with a per-ticket rule:
//   - If a ticket has ANY labor entry, only its Labor rows count (Action rows -> 0).
//   - If a ticket has no labor entry, its Action rows count (legacy behaviour
//     before the district started logging labor mid-year).
// This counts every ticket's work exactly once and loses no tickets. All hours
// rollups sum NetHours; the labor-type / resolution breakdowns keep using raw
// EffortHours because they are already scoped to a single record type.
//
// Runs as a full-sheet pass AFTER the activity load completes, because a ticket's
// Action and Labor rows can arrive on different pages of the paginated load.
function computeNetHours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ActivityLog');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const n = lastRow - 1;

  const ticketIds = sheet.getRange(2, 2, n, 1).getValues();     // col B  TicketId
  const effortHours = sheet.getRange(2, 6, n, 1).getValues();   // col F  EffortHours
  const laborTypeIds = sheet.getRange(2, 9, n, 1).getValues();  // col I  LaborTypeId

  // Pass 1: flag tickets that have at least one labor row.
  const ticketHasLabor = {};
  for (let r = 0; r < n; r++) {
    if (String(laborTypeIds[r][0] || '').trim()) {
      ticketHasLabor[ticketIds[r][0]] = true;
    }
  }

  // Pass 2: derive EntryType + NetHours per row.
  const out = new Array(n);
  for (let r = 0; r < n; r++) {
    const isLaborRow = !!String(laborTypeIds[r][0] || '').trim();
    const eff = Number(effortHours[r][0]) || 0;
    const net = ticketHasLabor[ticketIds[r][0]] ? (isLaborRow ? eff : 0) : eff;
    out[r] = [isLaborRow ? 'Labor' : 'Action', net];
  }

  // Write EntryType (col 25 = Y) and NetHours (col 26 = Z) together.
  sheet.getRange(2, 25, n, 2).setValues(out);
  SpreadsheetApp.flush();
  logOperation('NET_HOURS', 'INFO', 'Recomputed EntryType/NetHours for ' + n + ' rows.');
  return n;
}

// Rewrites ActivityLog Team (cols Q/R) for every row in place, applying the same
// hybrid rule as buildActivityRow: ticket's AssignedTeam (RawData) else the
// performing user's team (Users sheet), coalescing blanks to UNASSIGNED. No API
// calls — lets an already-loaded sheet pick up the team-attribution fix without a
// full reload. Returns the number of rows updated.
function repairTeamAttribution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ActivityLog');
  const rawSheet = ss.getSheetByName('RawData');
  if (!sheet || sheet.getLastRow() < 2) return 0;

  // ticketId -> {teamId, teamName} from RawData
  const ticketTeam = {};
  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rv = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, RAWDATA_HEADERS.length).getValues();
    rv.forEach(r => { if (r[0]) ticketTeam[r[0]] = { teamId: r[19] || '', teamName: r[12] || '' }; });
  }

  // userId -> {teamId, teamName} from Users sheet (cols F/G)
  const userTeam = {};
  const usersSheet = ss.getSheetByName('Users');
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const uv = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 7).getValues();
    uv.forEach(r => { if (r[0] && !userTeam[r[0]]) userTeam[r[0]] = { teamId: r[5] || '', teamName: r[6] || '' }; });
  }

  const n = sheet.getLastRow() - 1;
  const ticketIds = sheet.getRange(2, 2, n, 1).getValues();   // col B  TicketId
  const userIds = sheet.getRange(2, 13, n, 1).getValues();    // col M  PerformedByUserId
  const out = new Array(n);
  for (let r = 0; r < n; r++) {
    const t = ticketTeam[ticketIds[r][0]] || {};
    const u = userTeam[userIds[r][0]] || {};
    const teamId = t.teamId || u.teamId || '';
    const teamName = t.teamName || u.teamName || UNASSIGNED;
    out[r] = [teamId, teamName];
  }
  sheet.getRange(2, 17, n, 2).setValues(out);  // cols Q/R
  SpreadsheetApp.flush();
  logOperation('TEAM_REPAIR', 'INFO', 'Repaired team attribution for ' + n + ' rows.');
  return n;
}

function normalizeActivityId(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function getTicketIds(rawSheet) {
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) return [];
  return rawSheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .map(r => r[0])
    .filter(Boolean);
}

function buildTicketContextMap(rawSheet) {
  const lastRow = rawSheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;

  const values = rawSheet.getRange(2, 1, lastRow - 1, RAWDATA_HEADERS.length).getValues();
  values.forEach(row => {
    const ticketId = row[0];
    if (!ticketId) return;
    map[ticketId] = {
      ticketNumber: row[1],
      // Ticket's assigned user/team — used for attribution so reports reflect the
      // ticket assignment, not the performing user's home team. (RawData cols:
      // 10=AssignedUser, 12=AssignedTeam, 18=AssignedUserId, 19=AssignedTeamId)
      assignedUserName: row[10] || '',
      assignedUserId: row[18] || '',
      assignedTeamName: row[12] || '',
      assignedTeamId: row[19] || '',
      locationName: row[13],
      locationId: row[20],
      issueCategoryId: row[21] || '',
      issueCategoryName: row[22] || '',
      issueTypeId: row[23] || '',
      issueTypeName: row[24] || ''
    };
  });

  return map;
}

function buildUserMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  values.forEach(row => {
    const userId = row[0];
    if (!userId) return;
    if (!map[userId]) {
      map[userId] = {
        name: row[1],
        teamId: row[5],
        teamName: row[6]
      };
    }
  });

  return map;
}

// Repairs blank PerformedByUser (col N) values in existing ActivityLog rows
// by looking up PerformedByUserId (col M) in the Users sheet. Also fills
// blank TeamId/TeamName (cols Q/R) when the matched user has a team.
// Returns { repaired, stillMissing } where stillMissing is the count of rows
// whose user ID is still not in the Users sheet (caller should backfill first).
function repairBlankPerformedByUserNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSheet = ss.getSheetByName('ActivityLog');
  if (!actSheet) return { repaired: 0, stillMissing: 0 };

  var lastRow = actSheet.getLastRow();
  if (lastRow < 2) return { repaired: 0, stillMissing: 0 };

  var userMap = buildUserMap();

  // Read columns M:R (PerformedByUserId, PerformedByUser, Notes, IsPublic, TeamId, TeamName)
  // in a single batch. We only edit M (id is preserved), N (name), Q (teamId), R (teamName);
  // O (Notes) and P (IsPublic) are passed through unchanged.
  var range = actSheet.getRange(2, 13, lastRow - 1, 6);
  var values = range.getValues();

  var repaired = 0;
  var stillMissing = 0;
  for (var i = 0; i < values.length; i++) {
    var userId = values[i][0];
    var name = values[i][1];
    if (!userId || name) continue;

    var user = userMap[userId];
    if (!user || !user.name) {
      stillMissing++;
      continue;
    }

    values[i][1] = user.name;
    // Only fill team cells if they are currently blank
    if (!values[i][4] && user.teamId) values[i][4] = user.teamId;
    if (!values[i][5] && user.teamName) values[i][5] = user.teamName;
    repaired++;
  }

  if (repaired > 0) {
    range.setValues(values);
  }

  return { repaired: repaired, stillMissing: stillMissing };
}

function buildLaborTypeMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LaborTypes');
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  values.forEach(row => {
    const id = row[0];
    if (id) map[id] = row[1];
  });

  return map;
}

function buildResolutionActionMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ResolutionActions');
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  values.forEach(row => {
    const id = row[0];
    if (id) map[id] = row[1];
  });

  return map;
}

function buildActivityIndexMap(sheet) {
  const indexSheet = getActivityIndexSheet();
  let map = loadIndexMap(indexSheet);
  if (Object.keys(map).length === 0 && sheet.getLastRow() > 1) {
    rebuildActivityIndex(sheet);
    map = loadIndexMap(indexSheet);
  }
  return map;
}

function writeActivities(rows, sheet, indexMap) {
  if (!rows || rows.length === 0) return;

  const toAppend = [];
  const appendIds = [];
  const updates = {};

  rows.forEach(row => {
    const id = normalizeActivityId(row[0]);
    if (!id) return;

    row[0] = id;
    const existingRow = indexMap[id];
    if (existingRow) {
      updates[existingRow] = row;
    } else {
      toAppend.push(row);
      appendIds.push(id);
    }
  });

  writeBatchedUpdates(sheet, updates);

  if (toAppend.length > 0) {
    // Flush pending writes before appending to ensure getLastRow() is accurate
    // and to reduce transient "Service Spreadsheets failed" errors.
    SpreadsheetApp.flush();
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    const indexSheet = getActivityIndexSheet();
    const timestamp = new Date().toISOString();
    const indexRows = appendIds.map((id, idx) => [id, startRow + idx, timestamp]);
    appendIndexRows(indexSheet, indexRows);
    appendIds.forEach((id, idx) => {
      indexMap[id] = startRow + idx;
    });
  }
}

function getActivityFailuresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ActivityFailures');
  if (!sheet) {
    sheet = ss.insertSheet('ActivityFailures');
    sheet.getRange(1, 1, 1, 4).setValues([['TicketId', 'FailCount', 'LastError', 'LastAttempt']]);
    sheet.hideSheet();
  }
  return sheet;
}

function getActivityFailures() {
  const sheet = getActivityFailuresSheet();
  const lastRow = sheet.getLastRow();
  const failures = {};
  if (lastRow < 2) return failures;

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  values.forEach((row, idx) => {
    const ticketId = row[0];
    if (ticketId) {
      failures[ticketId] = {
        row: idx + 2,
        count: row[1] || 0
      };
    }
  });

  return failures;
}

function recordActivityFailure(ticketId, message) {
  const sheet = getActivityFailuresSheet();
  const failures = getActivityFailures();
  const now = new Date().toISOString();

  if (failures[ticketId]) {
    const row = failures[ticketId].row;
    const count = failures[ticketId].count + 1;
    sheet.getRange(row, 2, 1, 3).setValues([[count, message || '', now]]);
  } else {
    sheet.appendRow([ticketId, 1, message || '', now]);
  }

  logOperation('ACTIVITY_FAIL', 'WARNING', ticketId + ' - ' + (message || 'unknown error'));
}

function clearActivityFailure(ticketId) {
  const sheet = getActivityFailuresSheet();
  const failures = getActivityFailures();
  if (!failures[ticketId]) return;
  sheet.getRange(failures[ticketId].row, 1, 1, 4).clearContent();
}

function hasActivityFailures() {
  const failures = getActivityFailures();
  return Object.keys(failures).length > 0;
}

function processFailedActivityTickets(startTime, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap) {
  const failures = getActivityFailures();
  const failedIds = Object.keys(failures);
  if (failedIds.length === 0) return true;

  const batchSize = getActivityBatchSize();
  for (let i = 0; i < failedIds.length; i += batchSize) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      return false;
    }

    const batchIds = failedIds.slice(i, i + batchSize);
    const result = loadActivitiesBatch(batchIds, 0, actSheet, activityIndex, ticketMap, userMap, laborMap, resolutionMap, startTime);
    if (!result.completed) {
      return false;
    }
  }

  return true;
}
