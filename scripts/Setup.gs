/**
 * Setup.gs - Spreadsheet setup for iiQ Labor Tracker
 */

// --- Header constants (single source of truth for column counts) ---

const RAWDATA_HEADERS = [
  'TicketId', 'TicketNumber', 'Subject', 'CreatedDate', 'ClosedDate',
  'TotalLaborMins', 'TotalLaborHours', 'TotalLaborCost', 'LaborTypeId', 'LaborTypeName',
  'AssignedUser', 'AssignedUserEmail', 'AssignedTeam', 'Location', 'Requester',
  'Status', 'ResolutionAction', 'IsClosed', 'AssignedUserId', 'AssignedTeamId', 'LocationId'
];

const ACTIVITY_HEADERS = [
  'ActivityId', 'TicketId', 'TicketNumber', 'ActivityDate', 'EffortMins', 'EffortHours',
  'HourlyRate', 'LaborCost', 'LaborTypeId', 'LaborTypeName', 'ResolutionActionId',
  'ResolutionActionName', 'PerformedByUserId', 'PerformedByUser', 'Notes', 'IsPublic',
  'TeamId', 'TeamName', 'LocationId', 'LocationName'
];

// --- Helper for idempotent analytics sheet recreation ---

function deleteSheetIfExists(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    ss.deleteSheet(existing);
  }
}

function setupLaborTrackerDashboard() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Setup iiQ Labor Tracker',
    'This will create all required sheets, headers, and formulas.\n\n' +
    'Existing sheets will not be overwritten.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  const skipped = [];

  if (setupInstructionsSheet(ss)) created.push('Instructions'); else skipped.push('Instructions');
  if (setupConfigSheet(ss)) created.push('Config'); else skipped.push('Config');
  if (setupDateFiltersSheet(ss)) created.push('DateFilters'); else skipped.push('DateFilters');
  if (setupRawDataSheet(ss)) created.push('RawData'); else skipped.push('RawData');
  if (setupActivityLogSheet(ss)) created.push('ActivityLog'); else skipped.push('ActivityLog');
  if (setupTeamsSheet(ss)) created.push('Teams'); else skipped.push('Teams');
  if (setupUsersSheet(ss)) created.push('Users'); else skipped.push('Users');
  if (setupLaborTypesSheet(ss)) created.push('LaborTypes'); else skipped.push('LaborTypes');
  if (setupResolutionActionsSheet(ss)) created.push('ResolutionActions'); else skipped.push('ResolutionActions');
  if (setupTicketIndexSheet(ss)) created.push('TicketIndex'); else skipped.push('TicketIndex');
  if (setupActivityIndexSheet(ss)) created.push('ActivityIndex'); else skipped.push('ActivityIndex');
  if (setupByTeamSheet(ss)) created.push('ByTeam'); else skipped.push('ByTeam');
  if (setupByIndividualSheet(ss)) created.push('ByIndividual'); else skipped.push('ByIndividual');
  if (setupByDepartmentSheet(ss)) created.push('ByDepartment'); else skipped.push('ByDepartment');
  if (setupByLaborTypeSheet(ss)) created.push('ByLaborType'); else skipped.push('ByLaborType');
  if (setupByResolutionSheet(ss)) created.push('ByResolution'); else skipped.push('ByResolution');
  if (setupDashboardSheet(ss)) created.push('Dashboard'); else skipped.push('Dashboard');
  if (setupYearSummarySheet(ss)) created.push('YearSummary'); else skipped.push('YearSummary');
  if (setupLogsSheet(ss)) created.push('Logs'); else skipped.push('Logs');
  if (setupActivityFailuresSheet(ss)) created.push('ActivityFailures'); else skipped.push('ActivityFailures');
  ensureDataSheetsProtected();
  reorderSheets(ss);

  const message = [];
  if (created.length > 0) message.push('Created: ' + created.join(', '));
  if (skipped.length > 0) message.push('Already existed: ' + skipped.join(', '));
  message.push('\nNext steps:');
  message.push('1. Fill in Config sheet with API credentials');
  message.push('2. Run iiQ Data > Setup > Test API Connection');
  message.push('3. Run iiQ Data > Load Data > Start Initial Load');

  ui.alert('Setup Complete', message.join('\n'), ui.ButtonSet.OK);
}

// Desired tab order — matches the Sheet Reference section in Instructions.
// Hidden/internal sheets go at the end.
var SHEET_ORDER = [
  'Instructions', 'Config', 'DateFilters',
  'RawData', 'ActivityLog',
  'Teams', 'Users', 'LaborTypes', 'ResolutionActions',
  'ByTeam', 'ByIndividual', 'ByDepartment', 'ByLaborType', 'ByResolution',
  'YearSummary', 'Dashboard', 'Logs',
  // Hidden / internal (end of tab bar)
  'TicketIndex', 'ActivityIndex', 'ActivityFailures'
];

function reorderSheets(ss) {
  var position = 1;
  SHEET_ORDER.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.activate();
      ss.moveActiveSheet(position);
      position++;
    }
  });
  // Any sheets not in the list (e.g. user-created) stay after the ordered ones
}

function setupInstructionsSheet(ss) {
  if (ss.getSheetByName('Instructions')) return false;

  const sheet = ss.insertSheet('Instructions');
  sheet.setColumnWidth(1, 700);
  sheet.setColumnWidth(2, 500);

  // --- Section builders ---
  var row = 1;

  function writeHeader(text) {
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(14);
    row++;
  }

  function writeSectionHeader(text) {
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(11);
    row++;
  }

  function writeLine(text) {
    sheet.getRange(row, 1).setValue(text);
    row++;
  }

  function writePair(col1, col2) {
    sheet.getRange(row, 1).setValue(col1);
    sheet.getRange(row, 2).setValue(col2);
    row++;
  }

  function writePairBold(col1, col2) {
    sheet.getRange(row, 1).setValue(col1).setFontWeight('bold');
    sheet.getRange(row, 2).setValue(col2);
    row++;
  }

  function blankRow() { row++; }

  // ===== TITLE =====
  writeHeader('iiQ LABOR TRACKER');
  writeLine('This spreadsheet pulls labor hours and resolution action data from IncidentIQ into');
  writeLine('Google Sheets, providing rollup reports by team, individual, department, and labor type.');
  writeLine('All analytics update automatically via formulas — no manual calculation needed.');
  blankRow();

  // ===== QUICK START =====
  writeSectionHeader('QUICK START');
  writeLine('1. Run  iiQ Data > Setup > Run Complete Setup');
  writeLine('2. Go to the Config sheet and fill in:');
  writeLine('     API_BASE_URL  —  your district\'s IncidentIQ URL (e.g. https://district.incidentiq.com)');
  writeLine('     BEARER_TOKEN  —  your API bearer token (JWT)');
  writeLine('     SITE_ID  —  your site UUID');
  writeLine('     SCHOOL_YEAR_START / SCHOOL_YEAR_END  —  the date range for data');
  writeLine('3. Run  iiQ Data > Setup > Test API Connection  to verify credentials');
  writeLine('4. Run  iiQ Data > Load Data > Start Initial Load  to begin pulling data');
  writeLine('5. Wait for loading to complete (large datasets load in batches across multiple runs)');
  writeLine('6. (Optional) Run  iiQ Data > Setup > Setup Automated Triggers  for daily refresh');
  blankRow();

  // ===== HOW DATA LOADING WORKS =====
  writeSectionHeader('HOW DATA LOADING WORKS');
  writeLine('Data loads in three sequential phases:');
  writeLine('  Phase 1:  Reference data (Teams, Users, Resolution Actions)');
  writeLine('  Phase 2:  Ticket data (paginated, fetches all tickets in the school year)');
  writeLine('  Phase 3:  Activity log (fetches time entries for each ticket)');
  blankRow();
  writeLine('Google Apps Script has a 6-minute execution limit. Large loads automatically pause');
  writeLine('and resume. You can resume manually (iiQ Data > Load Data > Continue Loading) or');
  writeLine('let the automated monitor trigger pick it up every 10 minutes.');
  blankRow();
  writeLine('Progress is tracked in the Config sheet (TICKET_LOAD_PAGE, ACTIVITY_TICKET_INDEX, etc.).');
  writeLine('Use  iiQ Data > Check Status  to see current progress at any time.');
  blankRow();

  // ===== SHEET REFERENCE =====
  writeSectionHeader('SHEET REFERENCE');
  blankRow();
  writePairBold('Sheet', 'Description');
  writePair('Instructions', 'This sheet — setup guide and reference');
  writePair('Config', 'All settings, credentials, and load state (key-value pairs)');
  writePair('DateFilters', 'Date range selector for analytics (MTD, QTD, YTD, School YTD, etc.)');
  writePair('RawData', 'All tickets with labor fields (' + RAWDATA_HEADERS.length + ' columns)');
  writePair('ActivityLog', 'All resolution action time entries (' + ACTIVITY_HEADERS.length + ' columns) — primary source for rollups');
  writePair('Teams', 'Reference: team names and IDs');
  writePair('Users', 'Reference: user names, emails, team assignments');
  writePair('LaborTypes', 'Reference: labor type names and IDs');
  writePair('ResolutionActions', 'Reference: resolution action names, categories, scopes');
  writePair('ByTeam', 'Rollup: hours, cost, entry count, ticket count per team');
  writePair('ByIndividual', 'Rollup: hours, cost, entry count per individual (with team)');
  writePair('ByDepartment', 'Rollup: hours, cost, entry count per location/department');
  writePair('ByLaborType', 'Rollup: hours, cost, entry count per labor type');
  writePair('ByResolution', 'Rollup: hours, cost, entry count per resolution action');
  writePair('YearSummary', 'Monthly aggregation by team, agent, labor type, and resolution');
  writePair('Dashboard', 'KPI summary referencing the rollup sheets');
  writePair('Logs', 'Operation log (newest first, auto-trimmed to 1000 rows)');
  blankRow();

  // ===== DATE FILTERS =====
  writeSectionHeader('USING DATE FILTERS');
  writeLine('All rollup sheets (ByTeam, ByIndividual, etc.) filter data by the DateFilters sheet.');
  writeLine('Change the Filter Mode dropdown in DateFilters!B2 to adjust the date range:');
  blankRow();
  writePairBold('Filter Mode', 'Date Range');
  writePair('School YTD', 'School year start through today (default)');
  writePair('MTD', 'First of current month through today');
  writePair('QTD', 'First of current quarter through today');
  writePair('YTD', 'January 1 through today');
  writePair('Last 7 Days', 'Past 7 days');
  writePair('Last 30 Days', 'Past 30 days');
  writePair('Manual', 'Custom start/end dates (fill in rows 3 and 4)');
  blankRow();
  writeLine('Rollup formulas recalculate automatically when you change the filter mode.');
  blankRow();

  // ===== MENU REFERENCE =====
  writeSectionHeader('MENU REFERENCE  (iiQ Data)');
  blankRow();
  writePairBold('Menu Item', 'What It Does');
  writePair('Check Status', 'Shows current load progress and data counts');
  writePair('View Dashboard', 'Navigates to the Dashboard sheet');
  blankRow();
  writeLine('  Setup submenu:');
  writePair('  Run Complete Setup', 'Creates all sheets, headers, formulas, and protections');
  writePair('  Regenerate Analytics Sheets', 'Rebuilds all formula-based sheets (ByTeam, etc.) from scratch');
  writePair('  Test API Connection', 'Verifies API credentials work');
  writePair('  Verify Configuration', 'Checks all required Config settings are filled in');
  writePair('  Setup Automated Triggers', 'Installs monitor (10 min) and daily refresh (2 AM) triggers');
  writePair('  Remove Automated Triggers', 'Removes all time-based triggers');
  writePair('  View Trigger Status', 'Shows which triggers are currently installed');
  blankRow();
  writeLine('  Load Data submenu:');
  writePair('  Start Initial Load', 'Begins loading all data for the configured school year');
  writePair('  Continue Loading', 'Resumes a paused load from where it left off');
  writePair('  Refresh Reference Data', 'Reloads Teams, Users, and Resolution Actions');
  writePair('  Open Ticket Refresh', 'Updates open tickets and recently closed tickets');
  blankRow();
  writeLine('  Troubleshooting submenu:');
  writePair('  View Logs', 'Navigates to the Logs sheet');
  writePair('  Reset Load States', 'Resets all load progress (does not delete data)');
  writePair('  Full Reload (Clear Data)', 'Deletes all data and unlocks school year — requires triggers removed first');
  blankRow();

  // ===== AUTOMATION =====
  writeSectionHeader('AUTOMATION');
  writeLine('Two automated triggers are available (install via iiQ Data > Setup > Setup Automated Triggers):');
  blankRow();
  writePairBold('Trigger', 'Schedule & Purpose');
  writePair('Data Load Monitor', 'Every 10 minutes — resumes any paused loads automatically');
  writePair('Daily Open Refresh', 'Daily at 2 AM — refreshes open and recently closed tickets');
  blankRow();
  writeLine('Triggers skip gracefully if another operation is already running (no conflicts).');
  writeLine('The daily refresh only runs after the initial load is complete.');
  blankRow();

  // ===== SCHOOL YEAR & DATA SCOPE =====
  writeSectionHeader('SCHOOL YEAR & DATA SCOPE');
  writeLine('Each spreadsheet holds one school year of data. The date range is set in Config:');
  writeLine('  SCHOOL_YEAR_START  and  SCHOOL_YEAR_END');
  blankRow();
  writeLine('Once data loading begins, the school year dates are LOCKED to prevent accidental changes.');
  writeLine('To load a different school year:');
  writeLine('  1. Remove triggers  (iiQ Data > Setup > Remove Automated Triggers)');
  writeLine('  2. Full Reload  (iiQ Data > Troubleshooting > Full Reload) — this clears all data and unlocks dates');
  writeLine('  3. Update SCHOOL_YEAR_START and SCHOOL_YEAR_END in Config');
  writeLine('  4. Start Initial Load');
  blankRow();
  writeLine('For multiple school years, make a copy of the spreadsheet and configure each with different dates.');
  blankRow();

  // ===== TROUBLESHOOTING =====
  writeSectionHeader('TROUBLESHOOTING');
  blankRow();
  writePairBold('Problem', 'Solution');
  writePair('Load seems stuck', 'Check Status. If paused, run Continue Loading or wait for the monitor trigger.');
  writePair('API connection fails', 'Verify API_BASE_URL (no /api suffix), BEARER_TOKEN, and SITE_ID in Config.');
  writePair('"Another operation is running"', 'Wait a few minutes. Locks auto-expire after 6 minutes.');
  writePair('Analytics show wrong data', 'Check DateFilters date range. Run Regenerate Analytics Sheets.');
  writePair('Need to change school year', 'Remove triggers first, then Full Reload to unlock and clear data.');
  writePair('Rollup sheets are empty', 'ActivityLog must have data. Ensure all 3 load phases completed.');
  writePair('Duplicate or stale data', 'Data uses upsert (update-or-insert). Run Open Ticket Refresh for latest.');
  writePair('Activity failures', 'Check the ActivityFailures sheet. Failed tickets retry on next load.');
  blankRow();

  // ===== TIPS =====
  writeSectionHeader('TIPS');
  writeLine('- The Logs sheet records every operation — check it first when debugging.');
  writeLine('- Config values are all strings. Don\'t change auto-managed keys manually.');
  writeLine('- Rollup sheets are formula-driven and update instantly when ActivityLog data changes.');
  writeLine('- Hidden sheets (TicketIndex, ActivityIndex, ActivityFailures) support fast lookups — don\'t modify.');
  writeLine('- THROTTLE_MS (default 1000) controls delay between API calls. Lower = faster but may hit rate limits.');
  writeLine('- PAGE_SIZE (default 2000) controls records per API call. Larger = fewer calls but more memory.');

  // Freeze row 1 for the title
  sheet.setFrozenRows(1);

  return true;
}

function setupConfigSheet(ss) {
  if (ss.getSheetByName('Config')) return false;

  const sheet = ss.insertSheet('Config');
  sheet.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  const range = getSchoolYearRange();
  const defaultStart = range.startDate;
  const defaultEnd = range.endDate;

  const rows = [
    ['API_BASE_URL', 'https://your-district.incidentiq.com', 'Base URL only (no /api)'],
    ['BEARER_TOKEN', '', 'Paste your JWT token'],
    ['SITE_ID', '', 'Site UUID'],
    ['SCHOOL_YEAR_START', defaultStart, 'School year start (YYYY-MM-DD)'],
    ['SCHOOL_YEAR_END', defaultEnd, 'School year end (YYYY-MM-DD)'],
    ['PAGE_SIZE', '2000', 'Records per API call'],
    ['THROTTLE_MS', '1000', 'Delay between API calls (ms)'],
    ['OPEN_REFRESH_DAYS', '14', 'Days of closed tickets to refresh'],
    ['LAST_SYNC', '', 'Auto-updated timestamp'],
    ['SCHOOL_YEAR_LABEL', '', 'Auto-generated label'],
    ['ACTIVITY_TICKET_INDEX', '', 'Managed automatically'],
    ['ACTIVITY_BATCH_PAGE', '', 'Managed automatically'],
    ['SCHOOL_YEAR_LOCKED', 'FALSE', 'Managed automatically'],
    ['SCHOOL_YEAR_LOCKED_AT', '', 'Managed automatically'],
    ['SCHOOL_YEAR_LOCKED_START', '', 'Managed automatically'],
    ['SCHOOL_YEAR_LOCKED_END', '', 'Managed automatically'],
    ['PAGE_SIZE_LOCKED', '', 'Managed automatically']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Formula for SCHOOL_YEAR_LABEL (last row in the list)
  const labelIndex = rows.findIndex(row => row[0] === 'SCHOOL_YEAR_LABEL');
  if (labelIndex >= 0) {
    const labelRow = 2 + labelIndex;
    sheet.getRange(labelRow, 2).setFormula('=TEXT(B5,"YYYY")&"-"&TEXT(B6,"YYYY")');
  }

  sheet.setColumnWidths(1, 3, 240);
  return true;
}

function setupDateFiltersSheet(ss) {
  deleteSheetIfExists(ss, 'DateFilters');
  const sheet = ss.insertSheet('DateFilters');
  sheet.getRange(1, 1, 1, 3).setValues([['Filter Control', 'Value', 'Notes']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  const rows = [
    ['Filter Mode', 'School YTD', 'Manual, MTD, QTD, YTD, School YTD, Last 7 Days, Last 30 Days'],
    ['Start Date', '', 'Used only when Filter Mode = Manual'],
    ['End Date', '', 'Used only when Filter Mode = Manual'],
    ['Calculated Start', '', 'Auto-calculated start'],
    ['Calculated End', '', 'Auto-calculated end']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  sheet.getRange('B5').setFormula(
    '=SWITCH(B2,' +
    '"Manual",B3,' +
    '"MTD",EOMONTH(TODAY(),-1)+1,' +
    '"QTD",DATE(YEAR(TODAY()),CEILING(MONTH(TODAY())/3,1)*3-2,1),' +
    '"YTD",DATE(YEAR(TODAY()),1,1),' +
    '"School YTD",VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE),' +
    '"Last 7 Days",TODAY()-6,' +
    '"Last 30 Days",TODAY()-29,' +
    'VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE))'
  );

  sheet.getRange('B6').setFormula(
    '=SWITCH(B2,' +
    '"Manual",B4,' +
    '"School YTD",MIN(TODAY(),VLOOKUP("SCHOOL_YEAR_END",Config!A:B,2,FALSE)),' +
    'TODAY())'
  );

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Manual', 'MTD', 'QTD', 'YTD', 'School YTD', 'Last 7 Days', 'Last 30 Days'], true)
    .build();
  sheet.getRange('B2').setDataValidation(rule);

  sheet.setColumnWidths(1, 3, 240);
  return true;
}

function setupRawDataSheet(ss) {
  if (ss.getSheetByName('RawData')) return false;

  const sheet = ss.insertSheet('RawData');
  sheet.getRange(1, 1, 1, RAWDATA_HEADERS.length).setValues([RAWDATA_HEADERS]);
  sheet.getRange(1, 1, 1, RAWDATA_HEADERS.length).setFontWeight('bold');
  return true;
}

function setupActivityLogSheet(ss) {
  if (ss.getSheetByName('ActivityLog')) return false;

  const sheet = ss.insertSheet('ActivityLog');
  sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS]);
  sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setFontWeight('bold');
  return true;
}

function setupTeamsSheet(ss) {
  if (ss.getSheetByName('Teams')) return false;
  const sheet = ss.insertSheet('Teams');
  const headers = ['TeamId', 'TeamName', 'MemberCount', 'OpenTickets'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function setupUsersSheet(ss) {
  if (ss.getSheetByName('Users')) return false;
  const sheet = ss.insertSheet('Users');
  const headers = ['UserId', 'Name', 'Email', 'Location', 'Role', 'TeamId', 'TeamName'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function setupLaborTypesSheet(ss) {
  if (ss.getSheetByName('LaborTypes')) return false;
  const sheet = ss.insertSheet('LaborTypes');
  const headers = ['LaborTypeId', 'LaborTypeName'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function setupResolutionActionsSheet(ss) {
  if (ss.getSheetByName('ResolutionActions')) return false;
  const sheet = ss.insertSheet('ResolutionActions');
  const headers = ['ActionId', 'ActionName', 'Category', 'Scope'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function setupByTeamSheet(ss) {
  deleteSheetIfExists(ss, 'ByTeam');
  const sheet = ss.insertSheet('ByTeam');
  const headers = ['Team', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'teams,UNIQUE(FILTER(ActivityLog!R2:R,ActivityLog!R2:R<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'hours,BYROW(teams,LAMBDA(t,SUMIFS(ActivityLog!F:F,ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(teams,LAMBDA(t,SUMIFS(ActivityLog!H:H,ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(teams,LAMBDA(t,COUNTIFS(ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(teams,LAMBDA(t,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!R2:R=t,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(teams,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupByIndividualSheet(ss) {
  deleteSheetIfExists(ss, 'ByIndividual');
  const sheet = ss.insertSheet('ByIndividual');
  const headers = ['Individual', 'Team', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'users,UNIQUE(FILTER(ActivityLog!N2:N,ActivityLog!N2:N<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'teams,BYROW(users,LAMBDA(u,IFERROR(INDEX(ActivityLog!R:R,MATCH(u,ActivityLog!N:N,0)),""))),' +
    'hours,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!F:F,ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!H:H,ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(users,LAMBDA(u,COUNTIFS(ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(users,LAMBDA(u,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!N2:N=u,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(users,teams,hours,cost,entries,tickets,avg),0),3,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupByDepartmentSheet(ss) {
  deleteSheetIfExists(ss, 'ByDepartment');
  const sheet = ss.insertSheet('ByDepartment');
  const headers = ['Department', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'locs,UNIQUE(FILTER(ActivityLog!T2:T,ActivityLog!T2:T<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'hours,BYROW(locs,LAMBDA(l,SUMIFS(ActivityLog!F:F,ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(locs,LAMBDA(l,SUMIFS(ActivityLog!H:H,ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(locs,LAMBDA(l,COUNTIFS(ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(locs,LAMBDA(l,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!T2:T=l,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(locs,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupByLaborTypeSheet(ss) {
  deleteSheetIfExists(ss, 'ByLaborType');
  const sheet = ss.insertSheet('ByLaborType');
  const headers = ['Labor Type', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'types,UNIQUE(FILTER(ActivityLog!J2:J,ActivityLog!J2:J<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'hours,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!F:F,ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!H:H,ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(types,LAMBDA(t,COUNTIFS(ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(types,LAMBDA(t,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!J2:J=t,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(types,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupByResolutionSheet(ss) {
  deleteSheetIfExists(ss, 'ByResolution');
  const sheet = ss.insertSheet('ByResolution');
  const headers = ['Resolution Action', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'actions,UNIQUE(FILTER(ActivityLog!L2:L,ActivityLog!L2:L<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'hours,BYROW(actions,LAMBDA(a,SUMIFS(ActivityLog!F:F,ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(actions,LAMBDA(a,SUMIFS(ActivityLog!H:H,ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(actions,LAMBDA(a,COUNTIFS(ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(actions,LAMBDA(a,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!L2:L=a,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(actions,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupDashboardSheet(ss) {
  deleteSheetIfExists(ss, 'Dashboard');
  const sheet = ss.insertSheet('Dashboard');

  const rows = [
    ['Labor Dashboard', ''],
    ['Period', ''],
    ['Total Hours', ''],
    ['Total Cost', ''],
    ['Tickets with Labor', ''],
    ['Avg Hours per Ticket', ''],
    ['Top Performer', ''],
    ['Top Team', '']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');

  sheet.getRange('B2').setFormula('=TEXT(DateFilters!B5,"MMM D, YYYY")&" - "&TEXT(DateFilters!B6,"MMM D, YYYY")');
  sheet.getRange('B3').setFormula('=SUMIFS(ActivityLog!F:F,ActivityLog!D:D,">="&DateFilters!$B$5,ActivityLog!D:D,"<="&DateFilters!$B$6)');
  sheet.getRange('B4').setFormula('=SUMIFS(ActivityLog!H:H,ActivityLog!D:D,">="&DateFilters!$B$5,ActivityLog!D:D,"<="&DateFilters!$B$6)');
  sheet.getRange('B5').setFormula('=IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!D2:D>=DateFilters!$B$5,ActivityLog!D2:D<=DateFilters!$B$6)),0)');
  sheet.getRange('B6').setFormula('=IF(B5>0,B3/B5,0)');
  sheet.getRange('B7').setFormula('=IFERROR(INDEX(SORT(ByIndividual!A2:G,3,FALSE),1,1),"")');
  sheet.getRange('B8').setFormula('=IFERROR(INDEX(SORT(ByTeam!A2:F,2,FALSE),1,1),"")');

  sheet.setColumnWidths(1, 2, 220);
  return true;
}

function setupYearSummarySheet(ss) {
  deleteSheetIfExists(ss, 'YearSummary');
  const sheet = ss.insertSheet('YearSummary');
  const headers = ['SchoolYear', 'GroupType', 'Month', 'GroupName', 'Hours', 'Cost', 'ActivityCount', 'TicketCount'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=IFERROR(LET(' +
    'syStart,VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE),' +
    'syEnd,VLOOKUP("SCHOOL_YEAR_END",Config!A:B,2,FALSE),' +
    'label,TEXT(syStart,"YYYY")&"-"&TEXT(syEnd,"YYYY"),' +
    'data,FILTER({TEXT(ActivityLog!D2:D,"YYYY-MM"),ActivityLog!B2:B,ActivityLog!F2:F,ActivityLog!H2:H,ActivityLog!R2:R,ActivityLog!N2:N,ActivityLog!J2:J,ActivityLog!L2:L},ActivityLog!D2:D>=syStart,ActivityLog!D2:D<=syEnd,ActivityLog!D2:D<>""),' +
    'teamAgg,QUERY(data,"select Col1,Col5,sum(Col3),sum(Col4),count(Col3),count(distinct Col2) where Col5 is not null group by Col1,Col5",0),' +
    'teamOut,ARRAYFORMULA({label,"Team",INDEX(teamAgg,,1),INDEX(teamAgg,,2),INDEX(teamAgg,,3),INDEX(teamAgg,,4),INDEX(teamAgg,,5),INDEX(teamAgg,,6)}),' +
    'agentAgg,QUERY(data,"select Col1,Col6,sum(Col3),sum(Col4),count(Col3),count(distinct Col2) where Col6 is not null group by Col1,Col6",0),' +
    'agentOut,ARRAYFORMULA({label,"Agent",INDEX(agentAgg,,1),INDEX(agentAgg,,2),INDEX(agentAgg,,3),INDEX(agentAgg,,4),INDEX(agentAgg,,5),INDEX(agentAgg,,6)}),' +
    'laborAgg,QUERY(data,"select Col1,Col7,sum(Col3),sum(Col4),count(Col3),count(distinct Col2) where Col7 is not null group by Col1,Col7",0),' +
    'laborOut,ARRAYFORMULA({label,"LaborType",INDEX(laborAgg,,1),INDEX(laborAgg,,2),INDEX(laborAgg,,3),INDEX(laborAgg,,4),INDEX(laborAgg,,5),INDEX(laborAgg,,6)}),' +
    'resAgg,QUERY(data,"select Col1,Col8,sum(Col3),sum(Col4),count(Col3),count(distinct Col2) where Col8 is not null group by Col1,Col8",0),' +
    'resOut,ARRAYFORMULA({label,"Resolution",INDEX(resAgg,,1),INDEX(resAgg,,2),INDEX(resAgg,,3),INDEX(resAgg,,4),INDEX(resAgg,,5),INDEX(resAgg,,6)}),' +
    'VSTACK(teamOut,agentOut,laborOut,resOut)' +
    '),"")';

  sheet.getRange(2, 1).setFormula(formula);
  return true;
}

function setupLogsSheet(ss) {
  if (ss.getSheetByName('Logs')) return false;
  const sheet = ss.insertSheet('Logs');
  const headers = ['Timestamp', 'Operation', 'Status', 'Details'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return true;
}

function setupActivityFailuresSheet(ss) {
  if (ss.getSheetByName('ActivityFailures')) return false;
  const sheet = ss.insertSheet('ActivityFailures');
  const headers = ['TicketId', 'FailCount', 'LastError', 'LastAttempt'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.hideSheet();
  return true;
}

function setupTicketIndexSheet(ss) {
  if (ss.getSheetByName('TicketIndex')) return false;
  const sheet = ss.insertSheet('TicketIndex');
  const headers = ['TicketId', 'RowNumber', 'UpdatedAt'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.hideSheet();
  return true;
}

function setupActivityIndexSheet(ss) {
  if (ss.getSheetByName('ActivityIndex')) return false;
  const sheet = ss.insertSheet('ActivityIndex');
  const headers = ['ActivityId', 'RowNumber', 'UpdatedAt'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.hideSheet();
  return true;
}

// --- Regenerate Analytics Sheets (idempotent) ---

function regenerateAnalyticsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDateFiltersSheet(ss);
  setupByTeamSheet(ss);
  setupByIndividualSheet(ss);
  setupByDepartmentSheet(ss);
  setupByLaborTypeSheet(ss);
  setupByResolutionSheet(ss);
  setupDashboardSheet(ss);
  setupYearSummarySheet(ss);
  reorderSheets(ss);
  logOperation('SETUP', 'SUCCESS', 'Regenerated all analytics sheets');
}

function regenerateAnalyticsSheetsWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Regenerate Analytics Sheets',
    'This will delete and recreate all formula-driven analytics sheets ' +
    '(DateFilters, ByTeam, ByIndividual, ByDepartment, ByLaborType, ByResolution, Dashboard, YearSummary).\n\n' +
    'Data sheets (RawData, ActivityLog, etc.) will NOT be affected.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  regenerateAnalyticsSheets();
  ui.alert('Done', 'Analytics sheets have been regenerated.', ui.ButtonSet.OK);
}
