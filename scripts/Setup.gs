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

  const message = [];
  if (created.length > 0) message.push('Created: ' + created.join(', '));
  if (skipped.length > 0) message.push('Already existed: ' + skipped.join(', '));
  message.push('\nNext steps:');
  message.push('1. Fill in Config sheet with API credentials');
  message.push('2. Run iiQ Data > Setup > Test API Connection');
  message.push('3. Run iiQ Data > Load Data > Start Initial Load');

  ui.alert('Setup Complete', message.join('\n'), ui.ButtonSet.OK);
}

function setupInstructionsSheet(ss) {
  if (ss.getSheetByName('Instructions')) return false;

  const sheet = ss.insertSheet('Instructions');
  sheet.setColumnWidth(1, 800);

  const content = [
    ['iiQ LABOR TRACKER - SETUP AND USAGE'],
    [''],
    ['This spreadsheet pulls labor time from Incident IQ into Google Sheets.'],
    ['Use the iiQ Data menu to set up, load data, and configure automation.'],
    [''],
    ['Quick Start'],
    ['1. Run iiQ Data > Setup > Run Complete Setup'],
    ['2. Fill in Config sheet (API_BASE_URL, BEARER_TOKEN, SITE_ID)'],
    ['3. Run iiQ Data > Setup > Test API Connection'],
    ['4. Run iiQ Data > Load Data > Start Initial Load'],
    ['5. (Optional) Enable automation for daily refresh'],
    [''],
    ['Notes'],
    ['- Data is limited to the configured school year'],
    ['- School year dates lock after data loading begins'],
    ['- ActivityLog drives rollups and the YearSummary output'],
    ['- For historical years, make a copy and adjust school year dates']
  ];

  sheet.getRange(1, 1, content.length, 1).setValues(content);
  sheet.getRange(1, 1, 1, 1).setFontWeight('bold');
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
    'avg,IF(tickets>0,hours/tickets,0),' +
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
    'avg,IF(tickets>0,hours/tickets,0),' +
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
    'avg,IF(tickets>0,hours/tickets,0),' +
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
    'avg,IF(tickets>0,hours/tickets,0),' +
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
    'avg,IF(tickets>0,hours/tickets,0),' +
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
