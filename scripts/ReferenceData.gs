/**
 * ReferenceData.gs - Teams, Users, and Resolution Actions
 */

function refreshReferenceData() {
  loadTeams();
  loadUsers();
  loadLaborTypes();
  loadResolutionActions();
}

function loadTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteSheetIfExists(ss, 'Teams');
  const sheet = ss.insertSheet('Teams');
  const headers = ['TeamId', 'TeamName', 'MemberCount', 'OpenTickets'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const teams = fetchAllPagesWithQueryParams('/v1.0/teams', null, 'GET');

  // API does not return open ticket counts — column D is formula-driven (see setupTeamsSheet)
  const rows = teams.map(function(t) {
    var memberCount = [t.MembersCount, t.MemberCount].find(function(v) { return v != null; });
    return [
      t.TeamId || t.TeamID || t.Id || '',
      t.TeamName || t.Name || '',
      memberCount != null ? memberCount : ''
    ];
  });

  clearSheetData(sheet);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    // OpenTickets formula: count RawData rows where AssignedTeamId matches and IsClosed="Open"
    var formulas = rows.map(function(_, i) {
      var r = i + 2;
      return ['=COUNTIFS(RawData!T:T,A' + r + ',RawData!R:R,"Open")'];
    });
    sheet.getRange(2, 4, formulas.length, 1).setFormulas(formulas);
  }

  logOperation('REF_TEAMS', 'SUCCESS', 'Loaded ' + rows.length + ' teams');
}

function loadUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const usersSheet = ss.getSheetByName('Users');
  if (!teamsSheet || !usersSheet) throw new Error('Teams or Users sheet not found.');

  cacheConfigRowPositions();
  const startTime = Date.now();
  const teamRows = teamsSheet.getDataRange().getValues().slice(1).filter(r => r[0]);

  var teamIndex = getIntValue(getConfigValueDirect('USER_LOAD_TEAM_INDEX'), 0);

  // Clear sheet only on fresh start
  if (teamIndex === 0) {
    clearSheetData(usersSheet);
  }

  setLoadState(DATA_LOAD_TYPES.USERS, LOAD_STATES.IN_PROGRESS);

  // On resume, remove any partial data from the team we're about to process
  // (handles force-kill between sheet write and index update)
  if (teamIndex > 0 && teamIndex < teamRows.length) {
    var resumeTeamId = String(teamRows[teamIndex][0]);
    var lastRow = usersSheet.getLastRow();
    if (lastRow > 1) {
      var teamIdCol = usersSheet.getRange(2, 6, lastRow - 1, 1).getValues();
      var firstDupRow = 0;
      for (var r = teamIdCol.length - 1; r >= 0; r--) {
        if (String(teamIdCol[r][0]) === resumeTeamId) {
          firstDupRow = r + 2;
        } else {
          break;
        }
      }
      if (firstDupRow > 0) {
        usersSheet.getRange(firstDupRow, 1, lastRow - firstDupRow + 1, usersSheet.getLastColumn()).clear();
      }
    }
  }

  while (teamIndex < teamRows.length) {
    if (Date.now() - startTime >= MAX_RUNTIME_MS) {
      writeConfigValueDirect('USER_LOAD_TEAM_INDEX', String(teamIndex));
      logOperation('REF_USERS', 'INFO', 'Paused at team ' + teamIndex + ' of ' + teamRows.length);
      return;
    }

    var row = teamRows[teamIndex];
    var teamId = row[0];
    var teamName = row[1];

    var members;
    try {
      members = fetchAllPagesWithQueryParams('/v1.0/teams/' + teamId + '/members', null, 'GET');
    } catch (e) {
      logOperation('REF_USERS', 'WARNING',
        'Skipping team ' + teamName + ' (' + teamId + '): ' + e.message);
      teamIndex++;
      writeConfigValueDirect('USER_LOAD_TEAM_INDEX', String(teamIndex));
      continue;
    }

    if (members.length > 0) {
      var userRows = members.map(function(member) {
        return [
          member.UserId || member.Id || '',
          member.UserName || member.Name || '',
          member.Email || '',
          member.LocationName || member.Location || '',
          (member.Role && member.Role.Name) || '',
          teamId,
          teamName
        ];
      });

      var startRow = usersSheet.getLastRow() + 1;
      usersSheet.getRange(startRow, 1, userRows.length, userRows[0].length).setValues(userRows);
    }

    teamIndex++;
    writeConfigValueDirect('USER_LOAD_TEAM_INDEX', String(teamIndex));
  }

  // All teams processed — now augment with agents not on any team
  // (e.g. iiQ Administrators, system service accounts, agents with no team assignment).
  // Activities can be logged by these users, but the team-walk above misses them.
  var augmented = augmentUsersWithAgents(usersSheet);

  writeConfigValueDirect('USER_LOAD_TEAM_INDEX', '');
  setLoadState(DATA_LOAD_TYPES.USERS, LOAD_STATES.COMPLETE);
  logOperation('REF_USERS', 'SUCCESS', 'Loaded users from ' + teamRows.length + ' teams' +
    (augmented > 0 ? ' + ' + augmented + ' team-less agents' : ''));
}

// Fetch /v1.0/users/agents and append any agents not already in the Users sheet.
// Returns the number of new rows appended.
function augmentUsersWithAgents(usersSheet) {
  var existingIds = collectUserIds(usersSheet);
  var agents;
  try {
    agents = fetchAllPagesWithQueryParams('/v1.0/users/agents', null, 'GET');
  } catch (e) {
    logOperation('REF_USERS', 'WARNING', 'Failed to fetch /users/agents: ' + e.message);
    return 0;
  }

  var newRows = [];
  agents.forEach(function(a) {
    var userId = a.UserId || a.Id || '';
    if (!userId || existingIds[userId]) return;
    existingIds[userId] = true;
    newRows.push([
      userId,
      a.Name || ((a.FirstName || '') + ' ' + (a.LastName || '')).trim(),
      a.Email || a.Username || '',
      (a.Location && a.Location.Name) || a.LocationName || '',
      (a.Role && a.Role.Name) || '',
      '',
      ''
    ]);
  });

  if (newRows.length === 0) return 0;
  var startRow = usersSheet.getLastRow() + 1;
  usersSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  return newRows.length;
}

// Scan ActivityLog for PerformedByUserIds not in the Users sheet, fetch each
// individually via /v1.0/users/{id}, and append. Catches users whose role has
// since changed (e.g. demoted agents) and so are absent from /users/agents.
// Returns the number of new rows appended.
function backfillMissingUsersFromActivities() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName('Users');
  var actSheet = ss.getSheetByName('ActivityLog');
  if (!usersSheet || !actSheet) return 0;

  var existingIds = collectUserIds(usersSheet);
  var lastRow = actSheet.getLastRow();
  if (lastRow < 2) return 0;

  // Column M (PerformedByUserId) is the 13th column in ActivityLog
  var ids = actSheet.getRange(2, 13, lastRow - 1, 1).getValues();
  var missing = {};
  ids.forEach(function(r) {
    var id = r[0];
    if (id && !existingIds[id]) missing[id] = true;
  });

  var missingIds = Object.keys(missing);
  if (missingIds.length === 0) return 0;

  logOperation('REF_USERS', 'INFO',
    'Backfilling ' + missingIds.length + ' user(s) referenced by ActivityLog but missing from Users sheet');

  var newRows = [];
  missingIds.forEach(function(userId) {
    try {
      var response = apiRequest('GET', '/v1.0/users/' + userId, null);
      var u = (response && response.Item) || response;
      if (!u || !(u.UserId || u.Id)) {
        logOperation('REF_USERS', 'WARNING', 'No user returned for ' + userId);
        return;
      }
      newRows.push([
        u.UserId || u.Id,
        u.Name || ((u.FirstName || '') + ' ' + (u.LastName || '')).trim(),
        u.Email || u.Username || '',
        (u.Location && u.Location.Name) || u.LocationName || '',
        (u.Role && u.Role.Name) || '',
        '',
        ''
      ]);
    } catch (e) {
      logOperation('REF_USERS', 'WARNING', 'Failed to fetch user ' + userId + ': ' + e.message);
    }
  });

  if (newRows.length === 0) return 0;
  var startRow = usersSheet.getLastRow() + 1;
  usersSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  logOperation('REF_USERS', 'SUCCESS', 'Backfilled ' + newRows.length + ' user(s) from ActivityLog references');
  return newRows.length;
}

function collectUserIds(usersSheet) {
  var ids = {};
  var lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return ids;
  var values = usersSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(function(r) {
    if (r[0]) ids[r[0]] = true;
  });
  return ids;
}

function loadLaborTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LaborTypes');
  if (!sheet) throw new Error('LaborTypes sheet not found.');

  const productId = getProductId();
  const items = fetchAllPagesWithQueryParams(
    '/v1.0/labor/types?ProductId=' + productId, null, 'GET'
  );

  const rows = items.map(function(lt) {
    return [
      lt.LaborTypeId || lt.Id || '',
      lt.Name || '',
      lt.IsOvertime === true ? 'TRUE' : 'FALSE',
      lt.OTMultiplier || 0
    ];
  });

  clearSheetData(sheet);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  logOperation('REF_LABOR_TYPES', 'SUCCESS', 'Loaded ' + rows.length + ' labor types');
}

function loadResolutionActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ResolutionActions');
  if (!sheet) throw new Error('ResolutionActions sheet not found.');

  cacheConfigRowPositions();
  const startTime = Date.now();
  const pageSize = getPageSize();

  var page = getIntValue(getConfigValueDirect('RESOLUTION_LOAD_PAGE'), 0);

  // Clear sheet only on fresh start
  if (page === 0) {
    clearSheetData(sheet);
  }

  setLoadState(DATA_LOAD_TYPES.RESOLUTION_ACTIONS, LOAD_STATES.IN_PROGRESS);

  while (true) {
    if (Date.now() - startTime >= MAX_RUNTIME_MS) {
      writeConfigValueDirect('RESOLUTION_LOAD_PAGE', String(page));
      logOperation('REF_RESOLUTIONS', 'INFO', 'Paused at page ' + page);
      return;
    }

    var response = apiRequest('GET', '/v1.0/resolutions/actions?$p=' + page + '&$s=' + pageSize, null);
    var items = response && response.Items ? response.Items : (Array.isArray(response) ? response : []);

    if (items.length === 0) break;

    var rows = items.map(function(a) {
      return [
        a.ResolutionActionId || a.ActionId || '',
        a.Name || a.ActionName || '',
        a.ActionCategoryName || a.Category || '',
        a.Scope || ''
      ];
    });

    // Position-based write — re-processing a page overwrites in place, preventing duplicates
    var startRow = page * pageSize + 2;
    sheet.getRange(startRow, 1, pageSize, 4).clearContent();
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    page++;
    writeConfigValueDirect('RESOLUTION_LOAD_PAGE', String(page));

    if (response && response.Paging && typeof response.Paging.TotalRows === 'number') {
      if (page * pageSize >= response.Paging.TotalRows) break;
    } else {
      if (items.length < pageSize) break;
    }
  }

  writeConfigValueDirect('RESOLUTION_LOAD_PAGE', '');
  setLoadState(DATA_LOAD_TYPES.RESOLUTION_ACTIONS, LOAD_STATES.COMPLETE);
  logOperation('REF_RESOLUTIONS', 'SUCCESS', 'Loaded resolution actions');
}

function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
}
