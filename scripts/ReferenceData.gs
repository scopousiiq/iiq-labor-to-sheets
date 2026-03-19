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
    var members = fetchAllPagesWithQueryParams('/v1.0/teams/' + teamId + '/members', null, 'GET');

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

  // All teams processed
  writeConfigValueDirect('USER_LOAD_TEAM_INDEX', '');
  setLoadState(DATA_LOAD_TYPES.USERS, LOAD_STATES.COMPLETE);
  logOperation('REF_USERS', 'SUCCESS', 'Loaded users from ' + teamRows.length + ' teams');
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
