/**
 * Setup.gs - Spreadsheet setup for iiQ Labor Tracker
 */

// --- Header constants (single source of truth for column counts) ---

const RAWDATA_HEADERS = [
  'TicketId', 'TicketNumber', 'Subject', 'CreatedDate', 'ClosedDate',
  'TotalLaborMins', 'TotalLaborHours', 'TotalLaborCost', 'LaborTypeId', 'LaborTypeName',
  'AssignedUser', 'AssignedUserEmail', 'AssignedTeam', 'Location', 'Requester',
  'Status', 'ResolutionAction', 'IsClosed', 'AssignedUserId', 'AssignedTeamId', 'LocationId',
  'IssueCategoryId', 'IssueCategoryName', 'IssueTypeId', 'IssueTypeName'
];

const ACTIVITY_HEADERS = [
  'ActivityId', 'TicketId', 'TicketNumber', 'ActivityDate', 'EffortMins', 'EffortHours',
  'HourlyRate', 'LaborCost', 'LaborTypeId', 'LaborTypeName', 'ResolutionActionId',
  'ResolutionActionName', 'PerformedByUserId', 'PerformedByUser', 'Notes', 'IsPublic',
  'TeamId', 'TeamName', 'LocationId', 'LocationName',
  'IssueCategoryId', 'IssueCategoryName', 'IssueTypeId', 'IssueTypeName',
  // EntryType: 'Action' (resolution-action row) or 'Labor' (labor-entry row).
  // NetHours: de-duplicated effort used by all hours rollups (see computeNetHours).
  'EntryType', 'NetHours'
];

// --- Helpers ---

var FMT_DECIMAL = '#,##0.00';
var FMT_INTEGER = '#,##0';
var FMT_CURRENCY = '$#,##0.00';

// --- iiQ Brand Theme (CANONICAL_PATTERNS.md "Visual Styling") ---
// Core palette is brand blue + gold + white; Labor has no assigned domain accent
// in the brand table, so the hero/Dashboard tab uses gold and data/analytics tabs
// use brand blue. Never hardcode these hex values inline.
var BRAND = {
  blue:      '#365C96',  // primary — headers, banners, data/analytics tabs
  blueDeep:  '#162F67',  // section-header / title text
  blueLight: '#E0F3FF',  // section-band fill
  bluePale:  '#F3F8FF',  // zebra alternating row
  gold:      '#FEBB12',  // accent — header underline rule, hero tab
  white:     '#FFFFFF',
  line:      '#D7DDE6'   // neutral tab for Config/Logs/internal sheets
};

// Header row: brand-blue fill, white bold, thin gold bottom rule, 28px tall.
// `row` defaults to 1 (some sheets keep their header on a different row).
function styleHeaderRow_(sheet, cols, row) {
  row = row || 1;
  var r = sheet.getRange(row, 1, 1, Math.max(cols, 1));
  r.setBackground(BRAND.blue).setFontColor(BRAND.white).setFontWeight('bold')
   .setFontSize(10).setVerticalAlignment('middle');
  r.setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 28);
}

// Zebra banding (white / pale-blue) with a brand-blue header. Idempotent.
// The banding theme (and setFirstRowColor) governs font colors we don't set
// explicitly, and can render text white — including on the white first row.
// So we pin data rows to a dark font here, and callers must re-assert the
// header style AFTER banding so its white-on-blue is the final word.
function applyBanding_(sheet, cols, rows, headerRow) {
  headerRow = headerRow || 1;
  var width = Math.max(cols, 1);
  sheet.getBandings().forEach(function(b) { b.remove(); });
  var total = Math.max(rows - headerRow + 1, 1);
  var banding = sheet.getRange(headerRow, 1, total, width).applyRowBanding();
  banding.setHeaderRowColor(BRAND.blue).setFirstRowColor(BRAND.white).setSecondRowColor(BRAND.bluePale);
  // Force data rows to a dark font so no banding theme can leave white-on-white.
  var bodyRows = rows - headerRow;
  if (bodyRows > 0) {
    sheet.getRange(headerRow + 1, 1, bodyRows, width).setFontColor(BRAND.blueDeep);
  }
}

// Hero banner across row 1: brand-blue fill, white 16pt bold, gold underline,
// 40px tall, frozen. Used on Dashboard and Instructions.
function heroBanner_(sheet, cols) {
  var r = sheet.getRange(1, 1, 1, Math.max(cols, 1));
  if (cols > 1) { try { r.merge(); } catch (e) { /* already merged */ } }
  r.setBackground(BRAND.blue).setFontColor(BRAND.white).setFontWeight('bold')
   .setFontSize(16).setVerticalAlignment('middle');
  r.setBorder(null, null, true, null, null, null, BRAND.gold, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
}

// Per-sheet theming config. header = header row to style (omit for none);
// band = apply zebra banding; gridlines = hide gridlines; hero = title banner.
var THEME_SHEETS = [
  { name: 'Instructions',       tab: 'gold', hero: true,  gridlines: true },
  { name: 'Dashboard',          tab: 'gold', hero: true,  gridlines: true },
  { name: 'RawData',            tab: 'blue', header: 1, band: true },
  { name: 'ActivityLog',        tab: 'blue', header: 1, band: true },
  { name: 'Teams',              tab: 'blue', header: 1, band: true },
  { name: 'Users',              tab: 'blue', header: 1, band: true },
  { name: 'LaborTypes',         tab: 'blue', header: 1, band: true },
  { name: 'ResolutionActions',  tab: 'blue', header: 1, band: true },
  { name: 'ByTeam',             tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByIndividual',       tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByDepartment',       tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByLaborType',        tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByResolution',       tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByIssueCategory',    tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ByIssueType',        tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'AgentByCategory',    tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'TeamByCategory',     tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'CategoryByLaborType',tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'LocationByCategory', tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'ZeroLabor',          tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'YearSummary',        tab: 'blue', header: 1, band: true, gridlines: true },
  { name: 'AgentPivot',         tab: 'blue', header: 2, gridlines: true },
  // IndividualLookup has an irregular layout (controls + side-by-side summary),
  // so only tab color + hidden gridlines are applied — no generic header band.
  { name: 'IndividualLookup',   tab: 'blue', gridlines: true },
  { name: 'Config',             tab: 'line', header: 1 },
  { name: 'DateFilters',        tab: 'line' },
  { name: 'Logs',               tab: 'line', header: 1 }
];

// Applies the iiQ brand theme to every sheet by name. Safe to re-run; styles
// whatever sheets currently exist (so it also re-themes an already-built sheet).
function applyBrandTheme_(ss) {
  SpreadsheetApp.flush();
  var failed = [];
  THEME_SHEETS.forEach(function(cfg) {
    var sh = ss.getSheetByName(cfg.name);
    if (!sh) return;
    var tabColor = cfg.tab === 'gold' ? BRAND.gold : (cfg.tab === 'line' ? BRAND.line : BRAND.blue);
    try { sh.setTabColor(tabColor); } catch (e) { /* ignore */ }
    if (cfg.gridlines) { try { sh.setHiddenGridlines(true); } catch (e) { /* ignore */ } }
    var lastCol = Math.max(sh.getLastColumn(), 1);
    // Per-sheet guard: banding and the body font pass cover every data row, so
    // on the biggest sheets they are the most likely thing here to fail. One
    // sheet failing must not leave every later tab unstyled.
    try {
      if (cfg.hero) {
        heroBanner_(sh, Math.max(lastCol, 2));
      } else if (cfg.header) {
        // Band first, then style the header — so the header's white-on-blue is
        // applied last and can't be clobbered by the banding's white first row.
        if (cfg.band) {
          applyBanding_(sh, lastCol, Math.max(sh.getLastRow(), cfg.header), cfg.header);
        }
        styleHeaderRow_(sh, lastCol, cfg.header);
      }
    } catch (e) {
      failed.push(cfg.name);
    }
  });
  if (failed.length) {
    logOperation('SETUP', 'WARNING', 'Brand theme skipped ' + failed.length +
      ' sheet(s): ' + failed.join(', ') + '. Re-run to retry.');
  }
  logOperation('SETUP', 'INFO', 'Applied iiQ brand theme to all sheets.');
}

// Apply number formats to data columns. `formats` is an array of
// [colLetter, formatString] pairs. Formats from `dataStartRow` down to 1000.
function formatSheetColumns(sheet, dataStartRow, formats) {
  formats.forEach(function(pair) {
    var col = pair[0];
    var fmt = pair[1];
    sheet.getRange(col + dataStartRow + ':' + col).setNumberFormat(fmt);
  });
}

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
  if (setupByIssueCategorySheet(ss)) created.push('ByIssueCategory'); else skipped.push('ByIssueCategory');
  if (setupByIssueTypeSheet(ss)) created.push('ByIssueType'); else skipped.push('ByIssueType');
  if (setupIndividualLookupSheet(ss)) created.push('IndividualLookup'); else skipped.push('IndividualLookup');
  if (setupAgentPivotSheet(ss)) created.push('AgentPivot'); else skipped.push('AgentPivot');
  if (setupZeroLaborSheet(ss)) created.push('ZeroLabor'); else skipped.push('ZeroLabor');
  if (setupAgentByCategorySheet(ss)) created.push('AgentByCategory'); else skipped.push('AgentByCategory');
  if (setupTeamByCategorySheet(ss)) created.push('TeamByCategory'); else skipped.push('TeamByCategory');
  if (setupCategoryByLaborTypeSheet(ss)) created.push('CategoryByLaborType'); else skipped.push('CategoryByLaborType');
  if (setupLocationByCategorySheet(ss)) created.push('LocationByCategory'); else skipped.push('LocationByCategory');
  if (setupDashboardSheet(ss)) created.push('Dashboard'); else skipped.push('Dashboard');
  if (setupYearSummarySheet(ss)) created.push('YearSummary'); else skipped.push('YearSummary');
  if (setupLogsSheet(ss)) created.push('Logs'); else skipped.push('Logs');
  if (setupActivityFailuresSheet(ss)) created.push('ActivityFailures'); else skipped.push('ActivityFailures');

  // Idempotent Config sheet migration: adds telemetry + version rows to
  // existing installs whose Config sheet pre-dates those keys, and always
  // stamps SCRIPT_VERSION to the running code version.
  const configMigration = ensureConfigSheetUpToDate();

  ensureDataSheetsProtected();
  reorderSheets(ss);
  applyBrandTheme_(ss);

  const message = [];
  if (created.length > 0) message.push('Created: ' + created.join(', '));
  if (skipped.length > 0) message.push('Already existed: ' + skipped.join(', '));
  if (configMigration.added.length > 0) {
    message.push('Config rows added: ' + configMigration.added.join(', '));
  }
  message.push('Config SCRIPT_VERSION stamped: v' + SCRIPT_VERSION);
  message.push('\nNext steps:');
  message.push('1. Fill in Config sheet with API credentials');
  message.push('2. Run iiQ Labor > Setup > Test API Connection');
  message.push('3. Run iiQ Labor > Setup > Check for Updates (populates LATEST_VERSION)');
  message.push('4. Run iiQ Labor > Labor Data > Start Initial Load');

  ui.alert('Setup Complete', message.join('\n'), ui.ButtonSet.OK);
}

// Desired tab order — matches the Sheet Reference section in Instructions.
// Hidden/internal sheets go at the end.
var SHEET_ORDER = [
  'Instructions', 'Config', 'DateFilters',
  'RawData', 'ActivityLog',
  'Teams', 'Users', 'LaborTypes', 'ResolutionActions',
  'ByTeam', 'ByIndividual', 'ByDepartment', 'ByLaborType', 'ByResolution',
  'ByIssueCategory', 'ByIssueType',
  'IndividualLookup', 'AgentPivot', 'ZeroLabor',
  'AgentByCategory', 'TeamByCategory', 'CategoryByLaborType', 'LocationByCategory',
  'YearSummary', 'Dashboard', 'Logs',
  // Hidden / internal (end of tab bar)
  'TicketIndex', 'ActivityIndex', 'ActivityFailures'
];

// Returns the number of tabs it could not move. Never throws: tab order is
// cosmetic and must not cost the caller a completed rebuild.
function reorderSheets(ss) {
  // Flush pending sheet create/delete operations before reordering —
  // without this, activate()+moveActiveSheet() can fail after rapid
  // delete-and-recreate cycles during regenerateAnalyticsSheets().
  SpreadsheetApp.flush();

  var position = 1;
  var failed = 0;
  SHEET_ORDER.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var target = position;
    position++;
    // Skip tabs already in place. Each activate()+moveActiveSheet() pair is a
    // document-structure mutation, and on a large spreadsheet those are what
    // provoke "Service Spreadsheets failed while accessing document"; most
    // re-runs only need to move a handful of tabs.
    if (sheet.getIndex() === target) return;
    try {
      sheet.activate();
      ss.moveActiveSheet(target);
    } catch (e) {
      // One retry after a flush — the failure is usually a stale view of the
      // sheet list right after a delete-and-recreate, not a permanent error.
      try {
        SpreadsheetApp.flush();
        sheet.activate();
        ss.moveActiveSheet(target);
      } catch (e2) {
        failed++;
        logOperation('SETUP', 'WARNING',
          'Could not move tab "' + name + '" to position ' + target + ': ' + e2.message);
      }
    }
  });
  // Any sheets not in the list (e.g. user-created) stay after the ordered ones
  return failed;
}

function setupInstructionsSheet(ss) {
  // Idempotent (canonical Instructions pattern) — recreate so content/branding
  // changes propagate on re-run. (Config/Logs stay non-idempotent to preserve data.)
  deleteSheetIfExists(ss, 'Instructions');

  const sheet = ss.insertSheet('Instructions');
  sheet.setColumnWidth(1, 700);
  sheet.setColumnWidth(2, 500);
  sheet.setHiddenGridlines(true);

  // --- Section builders ---
  var row = 1;

  function writeHeader(text) {
    // Hero banner is finalized by applyBrandTheme_(); set the title text here.
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(14);
    row++;
  }

  function writeSectionHeader(text) {
    // Brand section band: deep-blue bold text on a light-blue fill (both columns).
    sheet.getRange(row, 1, 1, 2).setBackground(BRAND.blueLight);
    sheet.getRange(row, 1).setValue(text).setFontWeight('bold').setFontSize(11).setFontColor(BRAND.blueDeep);
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
  writeLine('1. Run  iiQ Labor > Setup > Run Complete Setup');
  writeLine('2. Go to the Config sheet and fill in:');
  writeLine('     API_BASE_URL  —  your district\'s IncidentIQ URL (e.g. https://district.incidentiq.com)');
  writeLine('     BEARER_TOKEN  —  your API bearer token (JWT) — obtain from iiQ: Admin > Developer Tools');
  writeLine('     SITE_ID  —  your site UUID');
  writeLine('     MODULE  —  Ticketing or Facilities (selects the IncidentIQ module)');
  writeLine('     SCHOOL_YEAR_START / SCHOOL_YEAR_END  —  the date range for data');
  writeLine('3. Run  iiQ Labor > Setup > Test API Connection  to verify credentials');
  writeLine('4. Run  iiQ Labor > Labor Data > Start Initial Load  to begin pulling data');
  writeLine('5. Wait for loading to complete (large datasets load in batches across multiple runs)');
  writeLine('6. (Optional) Run  iiQ Labor > Setup > Setup Automated Triggers  for daily refresh');
  blankRow();

  // ===== HOW DATA LOADING WORKS =====
  writeSectionHeader('HOW DATA LOADING WORKS');
  writeLine('Data loads in three sequential phases:');
  writeLine('  Phase 1:  Reference data (Teams, Users, Resolution Actions)');
  writeLine('  Phase 2:  Ticket data (paginated, fetches all tickets in the school year)');
  writeLine('  Phase 3:  Activity log (fetches time entries for each ticket)');
  blankRow();
  writeLine('Google Apps Script has a 6-minute execution limit. Large loads automatically pause');
  writeLine('and resume. You can resume manually (iiQ Labor > Labor Data > Continue Loading) or');
  writeLine('let the automated monitor trigger pick it up every 10 minutes.');
  blankRow();
  writeLine('Progress is tracked in the Config sheet (TICKET_LOAD_PAGE, ACTIVITY_TICKET_INDEX, etc.).');
  writeLine('Use  iiQ Labor > Check Status  to see current progress at any time.');
  blankRow();

  // ===== SHEET REFERENCE =====
  writeSectionHeader('SHEET REFERENCE');
  blankRow();
  writePairBold('Sheet', 'Description');
  writePair('Instructions', 'This sheet — setup guide and reference');
  writePair('Config', 'All settings, credentials, and load state (key-value pairs)');
  writePair('DateFilters', 'Date range selector for analytics (This/Last Month, Week, Quarter, etc.)');
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
  writePair('ByIssueCategory', 'Rollup: hours, cost, entry count per issue category');
  writePair('ByIssueType', 'Rollup: hours, cost, entry count per issue type');
  writePair('IndividualLookup', 'Select an individual to see their tickets and activity detail');
  writePair('AgentPivot', 'Per-agent hours broken down by labor type (Standard, Travel, Overtime, Weekend)');
  writePair('ZeroLabor', 'Closed tickets with zero labor hours logged (flag for review)');
  writePair('AgentByCategory', 'Cross-dimension: hours per agent broken down by issue category');
  writePair('TeamByCategory', 'Cross-dimension: hours per team broken down by issue category');
  writePair('CategoryByLaborType', 'Cross-dimension: hours per issue category broken down by labor type');
  writePair('LocationByCategory', 'Cross-dimension: hours per location broken down by issue category');
  writePair('YearSummary', 'Monthly aggregation by team, agent, labor type, and resolution');
  writePair('Dashboard', 'KPI summary referencing the rollup sheets');
  writePair('Logs', 'Operation log (newest first, auto-trimmed to 1000 rows)');
  blankRow();

  // ===== DATE FILTERS =====
  writeSectionHeader('USING DATE FILTERS');
  writeLine('All rollup sheets (ByTeam, ByIndividual, AgentPivot, etc.) filter by the DateFilters sheet.');
  writeLine('Change the Filter Mode dropdown in DateFilters!B2 to adjust the date range:');
  blankRow();
  writePairBold('Filter Mode', 'Date Range');
  writePair('This Month', 'First of current month through today');
  writePair('Last Month', 'Full previous month');
  writePair('This Week', 'Monday of current week through today');
  writePair('Last Week', 'Full previous week (Monday–Sunday)');
  writePair('This Quarter', 'First of current quarter through today');
  writePair('Last Quarter', 'Full previous quarter');
  writePair('This Calendar Year', 'January 1 through today');
  writePair('Last Calendar Year', 'Full previous calendar year');
  writePair('This School Year', 'School year start through today (default)');
  writePair('Last School Year', 'Full previous school year');
  writePair('Manual', 'Custom start/end dates (fill in rows 3 and 4)');
  blankRow();
  writeLine('Rollup formulas recalculate automatically when you change the filter mode.');
  blankRow();

  // ===== MENU REFERENCE =====
  writeSectionHeader('MENU REFERENCE  (iiQ Labor)');
  blankRow();
  writePairBold('Menu Item', 'What It Does');
  writePair('> Continue Loading', 'Primary action — resumes a paused load from where it left off');
  writePair('Check Status', 'Shows current load progress and data counts');
  writePair('Open Dashboard', 'Navigates to the Dashboard sheet');
  blankRow();
  writeLine('  Setup submenu:');
  writePair('  Run Complete Setup', 'Creates all sheets, headers, formulas, protections, and brand theme');
  writePair('  Regenerate Analytics Sheets', 'Rebuilds all formula-based sheets (ByTeam, etc.) from scratch');
  writePair('  Verify Configuration', 'Checks all required Config settings are filled in');
  writePair('  Test API Connection', 'Verifies API credentials work');
  writePair('  Setup Automated Triggers', 'Installs monitor (10 min) and daily refresh (2 AM) triggers');
  writePair('  View Trigger Status', 'Shows which triggers are currently installed');
  writePair('  Remove Automated Triggers', 'Removes all time-based triggers');
  writePair('  Check for Updates', 'Compares installed version against the latest release');
  blankRow();
  writeLine('  Labor Data submenu:');
  writePair('  Start Initial Load', 'Begins loading all data for the configured school year');
  writePair('  Open Ticket Refresh', 'Updates open tickets and recently closed tickets');
  writePair('  Refresh Reference Data', 'Reloads Teams, Users, and Resolution Actions');
  writePair('  Refresh Labor Types', 'Reloads the LaborTypes reference from the API');
  writePair('  Show Status', 'Shows current load progress and data counts');
  blankRow();
  writeLine('  Troubleshooting submenu:');
  writePair('  Validate Data', 'Reconciles ticket/labor counts and flags discrepancies');
  writePair('  Backfill Missing User Names', 'Resolves blank PerformedByUser names from the API');
  writePair('  Recompute Net Hours', 'Rebuilds the de-duplicated NetHours column (no reload)');
  writePair('  Repair Team Attribution', 'Recomputes ActivityLog team (ticket team, else user team)');
  writePair('  View Logs', 'Navigates to the Logs sheet');
  writePair('  Reset Load States', 'Resets all load progress (does not delete data)');
  writePair('  Full Reload (Clear Data)', 'Deletes all data and unlocks school year — requires triggers removed first');
  writePair('  Send Telemetry Ping (Debug)', 'Sends a one-off anonymous usage ping (see telemetry section)');
  blankRow();

  // ===== AUTOMATION =====
  writeSectionHeader('AUTOMATION');
  writeLine('Two automated triggers are available (install via iiQ Labor > Setup > Setup Automated Triggers):');
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
  writeLine('Once data loading begins, the school year dates and Module are LOCKED to prevent accidental changes.');
  writeLine('To load a different school year:');
  writeLine('  1. Remove triggers  (iiQ Labor > Setup > Remove Automated Triggers)');
  writeLine('  2. Full Reload  (iiQ Labor > Troubleshooting > Full Reload) — this clears all data and unlocks dates');
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
  blankRow();

  // ===== TELEMETRY =====
  writeSectionHeader('ANONYMOUS USAGE TELEMETRY');
  writeLine('Anonymous usage telemetry pings the iiQ team once per successful refresh, so');
  writeLine('we can see which districts run which version, on which iiQ instance, with');
  writeLine('how many activity rows — to prioritize features and catch regressions early.');
  blankRow();
  writeLine('What is sent (per ping):');
  writeLine('  • Stable install ID (UUID — generated locally, contains no PII)');
  writeLine('  • Project name (iiq-labor-to-sheets) and script version');
  writeLine('  • iiQ instance hostname (e.g. demo.incidentiq.com)');
  writeLine('  • ActivityLog row count');
  writeLine('  • Names of analytics sheets present (only the canonical set this project ships)');
  writeLine('  • Script time zone, install timestamp, send timestamp');
  blankRow();
  writeLine('What is NOT sent: ticket/labor data, API tokens, user names or emails,');
  writeLine('custom sheet names you add yourself, anything from row contents.');
  blankRow();
  writeLine('POLICY: Automated polling requires telemetry opt-in.');
  writeLine('  • To opt out, set TELEMETRY_ENABLED to FALSE in the Config sheet.');
  writeLine('  • This DISABLES automated polling: time-based triggers uninstall on next fire.');
  writeLine('  • Manual menu actions (iiQ Labor → Labor Data → ...) continue to work.');
  blankRow();
  writeLine('To re-enable: set TELEMETRY_ENABLED back to TRUE, then run');
  writeLine('iiQ Labor → Setup → Setup Automated Triggers to reinstall the triggers.');
  blankRow();

  // ===== DASHBOARD INTEGRATION =====
  writeSectionHeader('DASHBOARD INTEGRATION (Looker Studio / Power BI)');
  writeLine('The rollup sheets are BI-ready. To build an external dashboard:');
  blankRow();
  writeLine('Looker Studio:');
  writeLine('  1. In Looker Studio, Create > Data Source > Google Sheets connector.');
  writeLine('  2. Pick this spreadsheet, then a rollup tab (e.g. ByTeam, ByIndividual, YearSummary).');
  writeLine('  3. Use "Use first row as headers"; set Total Hours / Total Cost as numeric metrics.');
  writeLine('  4. Recommended charts: time-series of YearSummary hours by month; bar chart of');
  writeLine('     ByTeam / ByIndividual hours; scorecard of Dashboard Total Hours.');
  writeLine('  5. Add a date-range control bound to DateFilters, or filter on the month column.');
  blankRow();
  writeLine('Power BI:');
  writeLine('  1. Publish the sheet to the web (File > Share > Publish to web) or use the Google');
  writeLine('     Sheets / web connector with Get Data > Web.');
  writeLine('  2. Point at a rollup tab; set Total Hours / Total Cost as decimal columns.');
  writeLine('  3. Build matrix visuals (Team x Issue Category) from the cross-dimension tabs.');
  blankRow();
  writeLine('Tip: connect to the rollup tabs (already aggregated), not the raw ActivityLog, for');
  writeLine('fast dashboards. Hours come from NetHours (de-duplicated), so totals match the sheet.');
  blankRow();

  // ===== SUPPORT =====
  writeSectionHeader('SUPPORT');
  writeLine('1. Check the Logs sheet — it records every operation, newest first.');
  writeLine('2. Re-read the TROUBLESHOOTING section above for the most common fixes.');
  writeLine('3. Run iiQ Labor → Setup → Verify Configuration and Test API Connection.');
  writeLine('4. Run iiQ Labor → Setup → Check for Updates to confirm you are on the latest version.');
  writeLine('5. Project home: github.com/scopousiiq/iiq-labor-to-sheets');
  blankRow();
  writeLine('Last updated: ' + new Date().toISOString().split('T')[0]);

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
    ['MODULE', 'Ticketing', 'Ticketing or Facilities'],
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
    ['PAGE_SIZE_LOCKED', '', 'Managed automatically'],
    ['MODULE_LOCKED', '', 'Managed automatically'],
    ['SCRIPT_VERSION', SCRIPT_VERSION, 'Installed version (auto-stamped)'],
    ['LATEST_VERSION', '', 'Filled by iiQ Labor → Setup → Check for Updates'],
    ['VERSION_CHECK_DATE', '', 'Last update-check date'],
    ['TELEMETRY_ENABLED', 'TRUE', 'Set FALSE to opt out (also disables automated polling)']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Formula for SCHOOL_YEAR_LABEL (last row in the list)
  const labelIndex = rows.findIndex(row => row[0] === 'SCHOOL_YEAR_LABEL');
  if (labelIndex >= 0) {
    const labelRow = 2 + labelIndex;
    sheet.getRange(labelRow, 2).setFormula('=TEXT(B6,"YYYY")&"-"&TEXT(B7,"YYYY")');
  }

  // Data validation dropdown for MODULE
  const moduleIndex = rows.findIndex(row => row[0] === 'MODULE');
  if (moduleIndex >= 0) {
    const moduleRow = 2 + moduleIndex;
    const moduleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ticketing', 'Facilities'], true)
      .build();
    sheet.getRange(moduleRow, 2).setDataValidation(moduleRule);
  }

  sheet.setColumnWidths(1, 3, 240);
  return true;
}

// Idempotent migration: adds telemetry + version-notification rows that the
// current code expects but that an older Config sheet may lack. Safe to call
// on a fresh install (setupConfigSheet already wrote them — no-op) or an
// existing install (appends missing rows). Always stamps SCRIPT_VERSION
// to the running code version so districts can see the live build at a
// glance without running checkForUpdates.
//
// Returns { added: [keys appended] }.
function ensureConfigSheetUpToDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) {
    // Defensive: setupConfigSheet should have created it. Bail clean.
    return { added: [] };
  }

  // Rows the current code expects. Order matches setupConfigSheet so that
  // a fresh install + a migrated install end up structurally identical.
  const required = [
    ['SCRIPT_VERSION', String(SCRIPT_VERSION), 'Installed version (auto-stamped)'],
    ['LATEST_VERSION', '', 'Filled by iiQ Labor → Setup → Check for Updates'],
    ['VERSION_CHECK_DATE', '', 'Last update-check date'],
    ['TELEMETRY_ENABLED', 'TRUE', 'Set FALSE to opt out (also disables automated polling)']
  ];

  // Scan existing keys
  const lastRow = sheet.getLastRow();
  const existingKeys = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r) {
      if (r[0]) existingKeys[r[0]] = true;
    });
  }

  const toAppend = required.filter(function(r) { return !existingKeys[r[0]]; });
  if (toAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, 3).setValues(toAppend);
  }

  // Always stamp SCRIPT_VERSION to the running code version (overwrite if
  // existed). setConfig handles both update-existing and append-if-missing.
  setConfig('SCRIPT_VERSION', String(SCRIPT_VERSION));

  return { added: toAppend.map(function(r) { return r[0]; }) };
}

function setupDateFiltersSheet(ss) {
  deleteSheetIfExists(ss, 'DateFilters');
  const sheet = ss.insertSheet('DateFilters');
  sheet.getRange(1, 1, 1, 3).setValues([['Filter Control', 'Value', 'Notes']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  const rows = [
    ['Filter Mode', 'This School Year', 'Select a predefined range or Manual for custom dates'],
    ['Start Date', '', 'Used only when Filter Mode = Manual'],
    ['End Date', '', 'Used only when Filter Mode = Manual'],
    ['Calculated Start', '', 'Auto-calculated start'],
    ['Calculated End', '', 'Auto-calculated end'],
    ['School Year Start', '', 'Full school year start (used by YearSummary)'],
    ['School Year End', '', 'Full school year end (used by YearSummary)']
  ];

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Calculated Start — SWITCH on filter mode
  // Week starts Monday (WEEKDAY type 2: 1=Mon..7=Sun)
  sheet.getRange('B5').setFormula(
    '=SWITCH(B2,' +
    '"Manual",B3,' +
    '"This Month",EOMONTH(TODAY(),-1)+1,' +
    '"Last Month",EOMONTH(TODAY(),-2)+1,' +
    '"This Week",TODAY()-WEEKDAY(TODAY(),2)+1,' +
    '"Last Week",TODAY()-WEEKDAY(TODAY(),2)+1-7,' +
    '"This Quarter",DATE(YEAR(TODAY()),CEILING(MONTH(TODAY())/3,1)*3-2,1),' +
    '"Last Quarter",EDATE(DATE(YEAR(TODAY()),CEILING(MONTH(TODAY())/3,1)*3-2,1),-3),' +
    '"This Calendar Year",DATE(YEAR(TODAY()),1,1),' +
    '"Last Calendar Year",DATE(YEAR(TODAY())-1,1,1),' +
    '"This School Year",VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE),' +
    '"Last School Year",EDATE(VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE),-12),' +
    'VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE))'
  );

  // Calculated End — modes that end before today get explicit end dates;
  // all "This" modes default to TODAY()
  sheet.getRange('B6').setFormula(
    '=SWITCH(B2,' +
    '"Manual",B4,' +
    '"Last Month",EOMONTH(TODAY(),-1),' +
    '"Last Week",TODAY()-WEEKDAY(TODAY(),2),' +
    '"Last Quarter",DATE(YEAR(TODAY()),CEILING(MONTH(TODAY())/3,1)*3-2,1)-1,' +
    '"Last Calendar Year",DATE(YEAR(TODAY())-1,12,31),' +
    '"This School Year",MIN(TODAY(),VLOOKUP("SCHOOL_YEAR_END",Config!A:B,2,FALSE)),' +
    '"Last School Year",VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE)-1,' +
    'TODAY())'
  );

  // Fixed school year dates — always the full school year, independent of filter mode.
  sheet.getRange('B7').setFormula('=VLOOKUP("SCHOOL_YEAR_START",Config!A:B,2,FALSE)');
  sheet.getRange('B8').setFormula('=VLOOKUP("SCHOOL_YEAR_END",Config!A:B,2,FALSE)');

  const filterModes = [
    'Manual',
    'This Month', 'Last Month',
    'This Week', 'Last Week',
    'This Quarter', 'Last Quarter',
    'This Calendar Year', 'Last Calendar Year',
    'This School Year', 'Last School Year'
  ];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(filterModes, true)
    .build();
  sheet.getRange('B2').setDataValidation(rule);

  sheet.setColumnWidths(1, 3, 240);
  return true;
}

function setupRawDataSheet(ss) {
  var existed = true;
  var sheet = ss.getSheetByName('RawData');
  if (!sheet) {
    sheet = ss.insertSheet('RawData');
    existed = false;
  }
  // Always ensure headers are current (columns may have been added)
  sheet.getRange(1, 1, 1, RAWDATA_HEADERS.length).setValues([RAWDATA_HEADERS]);
  sheet.getRange(1, 1, 1, RAWDATA_HEADERS.length).setFontWeight('bold');
  return !existed;
}

function setupActivityLogSheet(ss) {
  var existed = true;
  var sheet = ss.getSheetByName('ActivityLog');
  if (!sheet) {
    sheet = ss.insertSheet('ActivityLog');
    existed = false;
  }
  // Always ensure headers are current (columns may have been added)
  sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setValues([ACTIVITY_HEADERS]);
  sheet.getRange(1, 1, 1, ACTIVITY_HEADERS.length).setFontWeight('bold');
  return !existed;
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
  deleteSheetIfExists(ss, 'LaborTypes');
  const sheet = ss.insertSheet('LaborTypes');
  const headers = ['LaborTypeId', 'LaborTypeName', 'IsOvertime', 'OTMultiplier'];
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
    'hours,BYROW(teams,LAMBDA(t,SUMIFS(ActivityLog!Z:Z,ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(teams,LAMBDA(t,SUMIFS(ActivityLog!H:H,ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(teams,LAMBDA(t,COUNTIFS(ActivityLog!R:R,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(teams,LAMBDA(t,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!R2:R=t,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(teams,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
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
    'hours,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!Z:Z,ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!H:H,ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(users,LAMBDA(u,COUNTIFS(ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(users,LAMBDA(u,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!N2:N=u,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(users,teams,hours,cost,entries,tickets,avg),0),3,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['C', FMT_DECIMAL], ['D', FMT_CURRENCY], ['E', FMT_INTEGER], ['F', FMT_INTEGER], ['G', FMT_DECIMAL]]);
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
    'hours,BYROW(locs,LAMBDA(l,SUMIFS(ActivityLog!Z:Z,ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(locs,LAMBDA(l,SUMIFS(ActivityLog!H:H,ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(locs,LAMBDA(l,COUNTIFS(ActivityLog!T:T,l,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(locs,LAMBDA(l,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!T2:T=l,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(locs,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
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
    // EffortHours is right here, unlike the other rollups: a labor type name only
    // exists on labor rows, and NetHours equals EffortHours on every labor row.
    'hours,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!F:F,ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!H:H,ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(types,LAMBDA(t,COUNTIFS(ActivityLog!J:J,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(types,LAMBDA(t,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!J2:J=t,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(types,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
  return true;
}

function setupIndividualLookupSheet(ss) {
  deleteSheetIfExists(ss, 'IndividualLookup');
  const sheet = ss.insertSheet('IndividualLookup');

  // --- Layout ---
  // Row 1: title
  // Row 2: "Select Individual" label + dropdown in B2
  // Row 3: "Filter Issue Type" label + dropdown in B3 (blank = every issue type)
  // Row 4: period
  // Rows 5-6: summary (hours, cost, entries, tickets)
  // Row 7: overtime summary
  // Row 8: detail headers (A-I) and "Ticket Summary" heading (K)
  // Row 9+: detail formula; ticket summary headers in K9, its formula in K10
  // Columns P-S: hidden helpers (two dropdown source lists + the split selection)

  sheet.getRange('A1').setValue('Individual Lookup').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2').setValue('Select Individual').setFontWeight('bold');

  // Dynamic dropdown: individuals active in the filtered date range
  const dropdownFormula =
    '=SORT(UNIQUE(FILTER(ActivityLog!N2:N,ActivityLog!N2:N<>"",ActivityLog!D2:D>=DateFilters!$B$5,ActivityLog!D2:D<=DateFilters!$B$6)))';
  sheet.getRange('P2').setFormula(dropdownFormula);

  // Data validation referencing the dynamic list
  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('P2:P'), true)
    .build();
  sheet.getRange('B2').setDataValidation(dropdownRule);

  // --- Issue type filter ---
  // Listed as "Category > Type" rather than the bare type name: issue type names
  // are not unique on their own — the same name recurs under several categories —
  // so a bare name would silently merge unrelated work into one total. Subject is
  // deliberately not the filter key either: it is free text the module generates
  // in more than one format, so a single issue type shows up under several
  // different subject strings.
  sheet.getRange('A3').setValue('Filter Issue Type').setFontWeight('bold');
  sheet.getRange('C3').setValue('Leave blank for all issue types')
    .setFontStyle('italic').setFontColor('#666666');

  sheet.getRange('Q2').setFormula(
    '=SORT(UNIQUE(FILTER(ActivityLog!V2:V&" > "&ActivityLog!X2:X,' +
    'ActivityLog!X2:X<>"",ActivityLog!D2:D>=DateFilters!$B$5,ActivityLog!D2:D<=DateFilters!$B$6)))'
  );

  const issueTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange('Q2:Q'), true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('B3').setDataValidation(issueTypeRule);

  // Split the selection back into category and type so every rollup below can
  // match ActivityLog!V and ActivityLog!X as separate criteria. Matching the
  // concatenation instead would push array arithmetic into each formula, which
  // is the documented way to poison a LET/HSTACK sheet with #N/A.
  sheet.getRange('R2').setFormula('=IFERROR(TRIM(LEFT($B$3,FIND(" > ",$B$3)-1)),"")');
  sheet.getRange('S2').setFormula('=IFERROR(TRIM(MID($B$3,FIND(" > ",$B$3)+3,500)),"")');

  // Wraps a metric so an empty filter reproduces the pre-filter formula exactly:
  // {I} is where the two extra SUMIFS/COUNTIFS criteria get spliced in.
  function issueAware(template) {
    return '=IF($B$2="","",IF($B$3="",' +
      template.replace('{I}', '') + ',' +
      template.replace('{I}', ',ActivityLog!V:V,$R$2,ActivityLog!X:X,$S$2') + '))';
  }

  // FILTER form of the same idea — collapses to an always-true array when no
  // issue type is selected.
  const ISSUE_COND =
    'IF($B$3="",ActivityLog!D2:D<>"",ActivityLog!V2:V=$R$2),' +
    'IF($B$3="",ActivityLog!D2:D<>"",ActivityLog!X2:X=$S$2)';

  // SUMPRODUCT form, for the overtime row.
  const ISSUE_TERM =
    '*IF($B$3="",1,(ActivityLog!V$2:V=$R$2)*(ActivityLog!X$2:X=$S$2))';

  const DATE_CRIT =
    'ActivityLog!D:D,">="&DateFilters!$B$5,ActivityLog!D:D,"<="&DateFilters!$B$6';

  // Period (row 4)
  sheet.getRange('A4').setValue('Period').setFontWeight('bold');
  sheet.getRange('B4').setFormula('=TEXT(DateFilters!B5,"MMM D, YYYY")&" - "&TEXT(DateFilters!B6,"MMM D, YYYY")');

  // Summary (rows 5-6). Hours come from NetHours (col Z), not raw EffortHours,
  // so a ticket carrying both labor rows and action rows is counted once.
  sheet.getRange('A5').setValue('Total Hours').setFontWeight('bold');
  sheet.getRange('B5').setFormula(issueAware(
    'SUMIFS(ActivityLog!Z:Z,ActivityLog!N:N,$B$2,' + DATE_CRIT + '{I})'
  ));
  sheet.getRange('C5').setValue('Total Cost').setFontWeight('bold');
  sheet.getRange('D5').setFormula(issueAware(
    'SUMIFS(ActivityLog!H:H,ActivityLog!N:N,$B$2,' + DATE_CRIT + '{I})'
  ));

  sheet.getRange('A6').setValue('Entries').setFontWeight('bold');
  sheet.getRange('B6').setFormula(issueAware(
    'COUNTIFS(ActivityLog!N:N,$B$2,' + DATE_CRIT + '{I})'
  ));
  sheet.getRange('C6').setValue('Tickets').setFontWeight('bold');
  // ROWS(UNIQUE(FILTER(...))) rather than COUNTUNIQUE: an empty FILTER yields
  // #N/A and the COUNTA family counts that error as 1. Once a filter can be
  // applied an empty result is routine, so this would report 1 ticket for none.
  sheet.getRange('D6').setFormula(
    '=IF($B$2="","",IFERROR(ROWS(UNIQUE(FILTER(ActivityLog!B2:B,' +
    'ActivityLog!N2:N=$B$2,ActivityLog!D2:D>=DateFilters!$B$5,ActivityLog!D2:D<=DateFilters!$B$6,' +
    ISSUE_COND + '))),0))'
  );

  // Overtime row (row 7) — auto-detected from LaborTypes IsOvertime column.
  // These sum raw EffortHours (col F) deliberately: the COUNTIFS term only
  // matches rows carrying a LaborTypeId, and on those rows NetHours always
  // equals EffortHours, so there is nothing to de-duplicate.
  sheet.getRange('A7').setValue('Overtime').setFontWeight('bold').setFontSize(11);

  const OT_BASE =
    '(ActivityLog!N$2:N=$B$2)' +
    '*(ActivityLog!D$2:D>=DateFilters!$B$5)' +
    '*(ActivityLog!D$2:D<=DateFilters!$B$6)' +
    '*(COUNTIFS(LaborTypes!A$2:A,ActivityLog!I$2:I,LaborTypes!C$2:C,"TRUE"))' +
    ISSUE_TERM;

  function otMetric(measure) {
    return '=IF($B$2="","",IFERROR(SUMPRODUCT(' + OT_BASE + measure + '),0))';
  }

  sheet.getRange('B7').setValue('OT Hours').setFontWeight('bold');
  sheet.getRange('C7').setFormula(otMetric('*ActivityLog!F$2:F'));

  sheet.getRange('D7').setValue('OT Cost').setFontWeight('bold');
  sheet.getRange('E7').setFormula(otMetric('*ActivityLog!H$2:H'));

  sheet.getRange('F7').setValue('OT Entries').setFontWeight('bold');
  sheet.getRange('G7').setFormula(otMetric(''));

  // Detail headers (row 8)
  const detailHeaders = ['TicketNumber', 'Subject', 'ActivityDate', 'NetHours', 'LaborCost',
                         'LaborType', 'ResolutionAction', 'EntryType', 'Notes'];
  sheet.getRange(8, 1, 1, detailHeaders.length).setValues([detailHeaders]);
  sheet.getRange(8, 1, 1, detailHeaders.length).setFontWeight('bold');

  // Detail formula — filtered activity rows for the selected individual
  // Uses LET to filter first, then MAP to look up Subject from RawData and TEXT for date formatting.
  // NetHours is shown instead of EffortHours so the column ties to Total Hours
  // above, and EntryType rides along so a superseded action row reading 0.00
  // is explicable rather than looking like missing data.
  const detailFormula = '=IF($B$2="","Select an individual from the dropdown above.",' +
    'IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'raw,SORT(FILTER(' +
    'HSTACK(ActivityLog!C2:C,ActivityLog!B2:B,ActivityLog!D2:D,ActivityLog!Z2:Z,ActivityLog!H2:H,ActivityLog!J2:J,ActivityLog!L2:L,ActivityLog!Y2:Y,ActivityLog!O2:O),' +
    'ActivityLog!N2:N=$B$2,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD,' + ISSUE_COND +
    '),3,FALSE),' +
    'tNums,INDEX(raw,,1),' +
    'tIds,INDEX(raw,,2),' +
    'subjs,MAP(tIds,LAMBDA(id,IFERROR(INDEX(RawData!C:C,MATCH(id,RawData!A:A,0)),""))),'+
    // MAP so TEXT() applies per row and the date column spills (a bare
    // TEXT(INDEX(raw,,3),...) returns only a scalar -> dates appeared on row 1 only).
    'dates,MAP(INDEX(raw,,3),LAMBDA(d,IF(d="","",TEXT(d,"M/D/YYYY")))),' +
    // CHOOSECOLS, not INDEX(raw,,{4,5,...}): an array column argument makes INDEX
    // return only its first column, so the trailing detail columns came back blank.
    'HSTACK(tNums,subjs,dates,CHOOSECOLS(raw,4,5,6,7,8,9))' +
    '),"No activity entries found for this individual in the selected date range."))';

  sheet.getRange(9, 1).setFormula(detailFormula);

  // Ticket summary beside the detail list, in column K with J left as a spacer.
  // Issue category and type are ticket-level fields denormalized onto every
  // activity row, so restricting the id list is enough — the per-ticket sums
  // below cannot pick up hours belonging to a different issue type.
  sheet.getRange(8, 11).setValue('Ticket Summary').setFontWeight('bold').setFontSize(11);
  const ticketSummaryHeaders = ['TicketNumber', 'Subject', 'Total Hours', 'Total Cost', 'Entries'];
  sheet.getRange(9, 11, 1, ticketSummaryHeaders.length).setValues([ticketSummaryHeaders]);
  sheet.getRange(9, 11, 1, ticketSummaryHeaders.length).setFontWeight('bold');

  const ticketSummaryFormula = '=IF($B$2="","",' +
    'IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'ids,UNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!N2:N=$B$2,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD,' + ISSUE_COND + ')),' +
    'nums,BYROW(ids,LAMBDA(id,IFERROR(INDEX(ActivityLog!C:C,MATCH(id,ActivityLog!B:B,0)),"?"))),' +
    'subjs,BYROW(ids,LAMBDA(id,IFERROR(INDEX(RawData!C:C,MATCH(id,RawData!A:A,0)),"?"))),' +
    'hours,BYROW(ids,LAMBDA(id,SUMIFS(ActivityLog!Z:Z,ActivityLog!B:B,id,ActivityLog!N:N,$B$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(ids,LAMBDA(id,SUMIFS(ActivityLog!H:H,ActivityLog!B:B,id,ActivityLog!N:N,$B$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(ids,LAMBDA(id,COUNTIFS(ActivityLog!B:B,id,ActivityLog!N:N,$B$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'IFERROR(SORT(HSTACK(nums,subjs,hours,cost,entries),3,FALSE),HSTACK(nums,subjs,hours,cost,entries))' +
    '),"No tickets found."))';

  sheet.getRange(10, 11).setFormula(ticketSummaryFormula);

  // Hide the helper columns P-S (two dropdown source lists + split selection)
  sheet.hideColumns(16, 4);

  // Column widths
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 240);  // holds both dropdowns; issue type values are long
  sheet.setColumnWidth(3, 200);  // "Total Cost" label / the filter hint
  sheet.setColumnWidth(9, 250);  // Notes
  sheet.setColumnWidth(10, 24);  // spacer between detail and ticket summary
  sheet.setColumnWidth(11, 140);
  sheet.setColumnWidth(12, 250); // Subject
  sheet.setColumnWidth(13, 100);
  sheet.setColumnWidth(14, 100);
  sheet.setColumnWidth(15, 80);

  formatSheetColumns(sheet, 9, [['D', FMT_DECIMAL], ['E', FMT_CURRENCY]]);
  formatSheetColumns(sheet, 10, [['M', FMT_DECIMAL], ['N', FMT_CURRENCY], ['O', FMT_INTEGER]]);
  sheet.getRange('B5').setNumberFormat(FMT_DECIMAL);
  sheet.getRange('D5').setNumberFormat(FMT_CURRENCY);
  sheet.getRange('B6').setNumberFormat(FMT_INTEGER);
  sheet.getRange('D6').setNumberFormat(FMT_INTEGER);
  sheet.getRange('C7').setNumberFormat(FMT_DECIMAL);
  sheet.getRange('E7').setNumberFormat(FMT_CURRENCY);
  sheet.getRange('G7').setNumberFormat(FMT_INTEGER);

  sheet.setFrozenRows(8);
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
    // EffortHours, not NetHours: only action rows carry a resolution action, so
    // this is the action-only view of recorded time and is not meant to tie to
    // the de-duplicated total. Switching to NetHours would zero out any action
    // whose ticket also had a labor entry, making each action's hours depend on
    // something about the ticket rather than on the action.
    'hours,BYROW(actions,LAMBDA(a,SUMIFS(ActivityLog!F:F,ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(actions,LAMBDA(a,SUMIFS(ActivityLog!H:H,ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(actions,LAMBDA(a,COUNTIFS(ActivityLog!L:L,a,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(actions,LAMBDA(a,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!L2:L=a,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(actions,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
  return true;
}

function setupByIssueCategorySheet(ss) {
  deleteSheetIfExists(ss, 'ByIssueCategory');
  const sheet = ss.insertSheet('ByIssueCategory');
  const headers = ['Issue Category', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    // Case-insensitive de-dup: UNIQUE/FILTER are case-SENSITIVE (so "Issue Not
    // Listed" and "Issue not listed" would list as two rows) but SUMIFS is
    // case-INSENSITIVE (sums both into each), which double-counted the total.
    // Uppercase each category (via MAP — UPPER() alone is NOT array-aware and
    // returns only the first cell), then UNIQUE so case-variants collapse to one
    // row; SUMIFS matches all variants into it. (Labels show uppercase — the real
    // fix is normalizing category capitalization in iiQ.)
    'cats,UNIQUE(MAP(FILTER(ActivityLog!V2:V,ActivityLog!V2:V<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD),LAMBDA(x,UPPER(x)))),' +
    'hours,BYROW(cats,LAMBDA(c,SUMIFS(ActivityLog!Z:Z,ActivityLog!V:V,c,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(cats,LAMBDA(c,SUMIFS(ActivityLog!H:H,ActivityLog!V:V,c,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(cats,LAMBDA(c,COUNTIFS(ActivityLog!V:V,c,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(cats,LAMBDA(c,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!V2:V=c,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(cats,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 250);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
  return true;
}

function setupByIssueTypeSheet(ss) {
  deleteSheetIfExists(ss, 'ByIssueType');
  const sheet = ss.insertSheet('ByIssueType');
  const headers = ['Issue Type', 'Total Hours', 'Total Cost', 'Entry Count', 'Ticket Count', 'Avg Hours/Ticket'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'types,UNIQUE(FILTER(ActivityLog!X2:X,ActivityLog!X2:X<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'hours,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!Z:Z,ActivityLog!X:X,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'cost,BYROW(types,LAMBDA(t,SUMIFS(ActivityLog!H:H,ActivityLog!X:X,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'entries,BYROW(types,LAMBDA(t,COUNTIFS(ActivityLog!X:X,t,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'tickets,BYROW(types,LAMBDA(t,IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!X2:X=t,ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),0))),' +
    'avg,MAP(hours,tickets,LAMBDA(h,t,IF(t>0,h/t,0))),' +
    'SORT(IFERROR(HSTACK(types,hours,cost,entries,tickets,avg),0),2,FALSE)' +
    ')';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 250);
  formatSheetColumns(sheet, 2, [['B', FMT_DECIMAL], ['C', FMT_CURRENCY], ['D', FMT_INTEGER], ['E', FMT_INTEGER], ['F', FMT_DECIMAL]]);
  return true;
}

// --- Cross-dimension sheets (QUERY-based) ---
// These use QUERY's GROUP BY for multi-column pivots, with date filtering
// via date literal syntax (same pattern as YearSummary).

function buildDateClause_() {
  return '" Col4>=date \'"&TEXT(startD,"yyyy-MM-dd")&"\' and Col4<=date \'"&TEXT(endD,"yyyy-MM-dd")&"\'"';
}

function setupAgentByCategorySheet(ss) {
  deleteSheetIfExists(ss, 'AgentByCategory');
  const sheet = ss.insertSheet('AgentByCategory');
  const headers = ['Agent', 'Issue Category', 'Hours', 'Cost', 'Entries'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Col14=PerformedByUser(N), Col22=IssueCategoryName(V), Col26=NetHours(Z), Col8=LaborCost(H), Col1=ActivityId(A)
  const formula = '=IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'dateClause,' + buildDateClause_() + ',' +
    'QUERY(ActivityLog!A2:Z,' +
    '"select Col14,Col22,sum(Col26),sum(Col8),count(Col1) where Col14 is not null and Col22 is not null and"&dateClause&' +
    '" group by Col14,Col22 order by Col14,sum(Col26) desc label sum(Col26) \'\',sum(Col8) \'\',count(Col1) \'\'"' +
    ',0)),"No data found for the selected date range.")';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  formatSheetColumns(sheet, 2, [['C', FMT_DECIMAL], ['D', FMT_CURRENCY], ['E', FMT_INTEGER]]);
  return true;
}

function setupTeamByCategorySheet(ss) {
  deleteSheetIfExists(ss, 'TeamByCategory');
  const sheet = ss.insertSheet('TeamByCategory');
  const headers = ['Team', 'Issue Category', 'Hours', 'Cost', 'Entries'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Col18=TeamName(R), Col22=IssueCategoryName(V), Col26=NetHours(Z)
  const formula = '=IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'dateClause,' + buildDateClause_() + ',' +
    'QUERY(ActivityLog!A2:Z,' +
    '"select Col18,Col22,sum(Col26),sum(Col8),count(Col1) where Col18 is not null and Col22 is not null and"&dateClause&' +
    '" group by Col18,Col22 order by Col18,sum(Col26) desc label sum(Col26) \'\',sum(Col8) \'\',count(Col1) \'\'"' +
    ',0)),"No data found for the selected date range.")';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  formatSheetColumns(sheet, 2, [['C', FMT_DECIMAL], ['D', FMT_CURRENCY], ['E', FMT_INTEGER]]);
  return true;
}

function setupCategoryByLaborTypeSheet(ss) {
  deleteSheetIfExists(ss, 'CategoryByLaborType');
  const sheet = ss.insertSheet('CategoryByLaborType');
  const headers = ['Issue Category', 'Labor Type', 'Hours', 'Cost', 'Entries'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Col22=IssueCategoryName(V), Col10=LaborTypeName(J)
  const formula = '=IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'dateClause,' + buildDateClause_() + ',' +
    'QUERY(ActivityLog!A2:X,' +
    '"select Col22,Col10,sum(Col6),sum(Col8),count(Col1) where Col22 is not null and Col10 is not null and"&dateClause&' +
    '" group by Col22,Col10 order by Col22,sum(Col6) desc label sum(Col6) \'\',sum(Col8) \'\',count(Col1) \'\'"' +
    ',0)),"No data found for the selected date range.")';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 200);
  formatSheetColumns(sheet, 2, [['C', FMT_DECIMAL], ['D', FMT_CURRENCY], ['E', FMT_INTEGER]]);
  return true;
}

function setupLocationByCategorySheet(ss) {
  deleteSheetIfExists(ss, 'LocationByCategory');
  const sheet = ss.insertSheet('LocationByCategory');
  const headers = ['Location', 'Issue Category', 'Hours', 'Cost', 'Entries'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Col20=LocationName(T), Col22=IssueCategoryName(V), Col26=NetHours(Z)
  const formula = '=IFERROR(LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'dateClause,' + buildDateClause_() + ',' +
    'QUERY(ActivityLog!A2:Z,' +
    '"select Col20,Col22,sum(Col26),sum(Col8),count(Col1) where Col20 is not null and Col22 is not null and"&dateClause&' +
    '" group by Col20,Col22 order by Col20,sum(Col26) desc label sum(Col26) \'\',sum(Col8) \'\',count(Col1) \'\'"' +
    ',0)),"No data found for the selected date range.")';

  sheet.getRange(2, 1).setFormula(formula);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  formatSheetColumns(sheet, 2, [['C', FMT_DECIMAL], ['D', FMT_CURRENCY], ['E', FMT_INTEGER]]);
  return true;
}

function setupAgentPivotSheet(ss) {
  deleteSheetIfExists(ss, 'AgentPivot');
  const sheet = ss.insertSheet('AgentPivot');

  // Row 1: Sort controls
  sheet.getRange('A1').setValue('Sort By').setFontWeight('bold');
  sheet.getRange('B1').setValue('Total Hours');
  sheet.getRange('C1').setValue('Order').setFontWeight('bold');
  sheet.getRange('D1').setValue('Descending');

  // Sort By dropdown — column names matching row 2 headers
  const sortByRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Agent', 'Team', 'Standard', 'Travel', 'Overtime', 'Weekend', 'Other', 'Total Hours'], true)
    .build();
  sheet.getRange('B1').setDataValidation(sortByRule);

  // Order dropdown
  const orderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ascending', 'Descending'], true)
    .build();
  sheet.getRange('D1').setDataValidation(orderRule);

  // Row 2: Headers — labor type names reference ActivityLog!J (LaborTypeName).
  // Adjust C2:F2 to match your district's labor type names. "Other" captures
  // de-duplicated hours not attributed to those four labor types (e.g. action-only
  // tickets that carry no labor type) so C+D+E+F+G reconcile to the Total.
  const headers = ['Agent', 'Team', 'Standard', 'Travel', 'Overtime', 'Weekend', 'Other', 'Total Hours'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, headers.length).setFontWeight('bold');

  // Formula references header cells ($C$2..$F$2) so users can rename labor types
  // without editing the formula. Each labor-type column SUMIFs raw EffortHours
  // (labor rows == NetHours). Total sums NetHours (de-duplicated). "Other" is the
  // remainder so the type columns + Other always equal the Total.
  // Sort column determined by SWITCH on $B$1, sort direction by $D$1.
  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'users,UNIQUE(FILTER(ActivityLog!N2:N,ActivityLog!N2:N<>"",ActivityLog!D2:D>=startD,ActivityLog!D2:D<=endD)),' +
    'teams,BYROW(users,LAMBDA(u,IFERROR(INDEX(ActivityLog!R:R,MATCH(u,ActivityLog!N:N,0)),"---"))),' +
    // The per-type columns sum EffortHours because a labor type name only appears
    // on labor rows, where it equals NetHours. Total uses NetHours across all
    // rows, so "Other" absorbs the action-row hours and the row still adds up.
    'stdH,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!F:F,ActivityLog!N:N,u,ActivityLog!J:J,$C$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'trvH,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!F:F,ActivityLog!N:N,u,ActivityLog!J:J,$D$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'otH,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!F:F,ActivityLog!N:N,u,ActivityLog!J:J,$E$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'wkndH,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!F:F,ActivityLog!N:N,u,ActivityLog!J:J,$F$2,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'total,BYROW(users,LAMBDA(u,SUMIFS(ActivityLog!Z:Z,ActivityLog!N:N,u,ActivityLog!D:D,">="&startD,ActivityLog!D:D,"<="&endD))),' +
    'other,MAP(total,stdH,trvH,otH,wkndH,LAMBDA(t,s,r,o,w,t-s-r-o-w)),' +
    'sCol,SWITCH($B$1,"Agent",1,"Team",2,$C$2,3,$D$2,4,$E$2,5,$F$2,6,"Other",7,8),' +
    'sAsc,$D$1="Ascending",' +
    'SORT(IFERROR(HSTACK(users,teams,stdH,trvH,otH,wkndH,other,total),0),sCol,sAsc)' +
    ')';

  sheet.getRange(3, 1).setFormula(formula);

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 180);
  sheet.setFrozenRows(2);
  formatSheetColumns(sheet, 3, [['C', FMT_DECIMAL], ['D', FMT_DECIMAL], ['E', FMT_DECIMAL], ['F', FMT_DECIMAL], ['G', FMT_DECIMAL], ['H', FMT_DECIMAL]]);
  return true;
}

function setupZeroLaborSheet(ss) {
  deleteSheetIfExists(ss, 'ZeroLabor');
  const sheet = ss.insertSheet('ZeroLabor');

  // Closed tickets with zero labor hours in the filtered date range.
  // Helps managers find "Work Complete" tickets where no time was logged.
  const headers = ['TicketNumber', 'Subject', 'CreatedDate', 'ClosedDate', 'AssignedUser', 'AssignedTeam', 'Location', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Wrap dates in TEXT() to prevent serial number display in HSTACK output
  const formula = '=LET(' +
    'startD,DateFilters!$B$5,' +
    'endD,DateFilters!$B$6,' +
    'IFERROR(SORT(FILTER(' +
    'HSTACK(RawData!B2:B,RawData!C2:C,TEXT(RawData!D2:D,"M/D/YYYY"),TEXT(RawData!E2:E,"M/D/YYYY"),RawData!K2:K,RawData!M2:M,RawData!N2:N,RawData!P2:P),' +
    'RawData!R2:R="Closed",' +
    'RawData!F2:F=0,' +
    'RawData!D2:D>=startD,' +
    'RawData!D2:D<=endD' +
    '),4,FALSE),' +
    '"No closed tickets with zero labor found in the selected date range."))';

  sheet.getRange(2, 1).setFormula(formula);

  sheet.setColumnWidth(2, 300);  // Subject
  sheet.setColumnWidth(7, 200);  // Location
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
  sheet.getRange('B3').setFormula('=SUMIFS(ActivityLog!Z:Z,ActivityLog!D:D,">="&DateFilters!$B$5,ActivityLog!D:D,"<="&DateFilters!$B$6)');
  sheet.getRange('B4').setFormula('=SUMIFS(ActivityLog!H:H,ActivityLog!D:D,">="&DateFilters!$B$5,ActivityLog!D:D,"<="&DateFilters!$B$6)');
  sheet.getRange('B5').setFormula('=IFERROR(COUNTUNIQUE(FILTER(ActivityLog!B2:B,ActivityLog!D2:D>=DateFilters!$B$5,ActivityLog!D2:D<=DateFilters!$B$6)),0)');
  sheet.getRange('B6').setFormula('=IF(B5>0,B3/B5,0)');
  sheet.getRange('B7').setFormula('=IFERROR(INDEX(SORT(ByIndividual!A2:G,3,FALSE),1,1),"")');
  sheet.getRange('B8').setFormula('=IFERROR(INDEX(SORT(ByTeam!A2:F,2,FALSE),1,1),"")');

  sheet.getRange('B3').setNumberFormat(FMT_DECIMAL);
  sheet.getRange('B4').setNumberFormat(FMT_CURRENCY);
  sheet.getRange('B5').setNumberFormat(FMT_INTEGER);
  sheet.getRange('B6').setNumberFormat(FMT_DECIMAL);

  sheet.setColumnWidths(1, 2, 220);
  return true;
}

function setupYearSummarySheet(ss) {
  deleteSheetIfExists(ss, 'YearSummary');
  const sheet = ss.insertSheet('YearSummary');
  const headers = ['SchoolYear', 'GroupType', 'Month', 'GroupName', 'Hours', 'Cost', 'ActivityCount'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Queries ActivityLog range directly (not a constructed array) so QUERY
  // correctly detects numeric column types for sum(). Uses QUERY's built-in
  // year()/month() functions and date literal syntax for filtering.
  // Each group wrapped in IFERROR; empty groups produce placeholder rows
  // that are filtered out after VSTACK.
  const formula = '=IFERROR(LET(' +
    'syStart,DateFilters!$B$7,' +
    'syEnd,DateFilters!$B$8,' +
    'label,TEXT(syStart,"YYYY")&"-"&TEXT(syEnd,"YYYY"),' +
    'dateClause," Col4>=date \'"&TEXT(syStart,"yyyy-MM-dd")&"\' and Col4<=date \'"&TEXT(syEnd,"yyyy-MM-dd")&"\'",' +
    'empty,CHOOSE({1,2,3,4,5,6,7},"","","","","","",""),' +

    'teamResult,IFERROR(LET(' +
      'agg,QUERY(ActivityLog!A2:Z,"select year(Col4),month(Col4)+1,Col18,sum(Col26),sum(Col8),count(Col6) where Col18 is not null and"&dateClause&" group by year(Col4),month(Col4)+1,Col18",0),' +
      'mo,MAP(INDEX(agg,,1),INDEX(agg,,2),LAMBDA(y,m,TEXT(y,"0")&"-"&TEXT(m,"00"))),' +
      'HSTACK(MAP(mo,LAMBDA(x,IF(x<>"",label,""))),MAP(mo,LAMBDA(x,IF(x<>"","Team",""))),mo,INDEX(agg,,3),INDEX(agg,,4),INDEX(agg,,5),INDEX(agg,,6))),empty),' +

    'agentResult,IFERROR(LET(' +
      'agg,QUERY(ActivityLog!A2:Z,"select year(Col4),month(Col4)+1,Col14,sum(Col26),sum(Col8),count(Col6) where Col14 is not null and"&dateClause&" group by year(Col4),month(Col4)+1,Col14",0),' +
      'mo,MAP(INDEX(agg,,1),INDEX(agg,,2),LAMBDA(y,m,TEXT(y,"0")&"-"&TEXT(m,"00"))),' +
      'HSTACK(MAP(mo,LAMBDA(x,IF(x<>"",label,""))),MAP(mo,LAMBDA(x,IF(x<>"","Agent",""))),mo,INDEX(agg,,3),INDEX(agg,,4),INDEX(agg,,5),INDEX(agg,,6))),empty),' +

    'laborResult,IFERROR(LET(' +
      'agg,QUERY(ActivityLog!A2:Z,"select year(Col4),month(Col4)+1,Col10,sum(Col26),sum(Col8),count(Col6) where Col10 is not null and"&dateClause&" group by year(Col4),month(Col4)+1,Col10",0),' +
      'mo,MAP(INDEX(agg,,1),INDEX(agg,,2),LAMBDA(y,m,TEXT(y,"0")&"-"&TEXT(m,"00"))),' +
      'HSTACK(MAP(mo,LAMBDA(x,IF(x<>"",label,""))),MAP(mo,LAMBDA(x,IF(x<>"","LaborType",""))),mo,INDEX(agg,,3),INDEX(agg,,4),INDEX(agg,,5),INDEX(agg,,6))),empty),' +

    'resResult,IFERROR(LET(' +
      'agg,QUERY(ActivityLog!A2:T,"select year(Col4),month(Col4)+1,Col12,sum(Col6),sum(Col8),count(Col6) where Col12 is not null and"&dateClause&" group by year(Col4),month(Col4)+1,Col12",0),' +
      'mo,MAP(INDEX(agg,,1),INDEX(agg,,2),LAMBDA(y,m,TEXT(y,"0")&"-"&TEXT(m,"00"))),' +
      'HSTACK(MAP(mo,LAMBDA(x,IF(x<>"",label,""))),MAP(mo,LAMBDA(x,IF(x<>"","Resolution",""))),mo,INDEX(agg,,3),INDEX(agg,,4),INDEX(agg,,5),INDEX(agg,,6))),empty),' +

    'combined,VSTACK(teamResult,agentResult,laborResult,resResult),' +
    'FILTER(combined,INDEX(combined,,4)<>"")' +
    '),"No activity data found for the school year.")';

  sheet.getRange(2, 1).setFormula(formula);
  formatSheetColumns(sheet, 2, [['E', FMT_DECIMAL], ['F', FMT_CURRENCY], ['G', FMT_INTEGER]]);
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
  setupByIssueCategorySheet(ss);
  setupByIssueTypeSheet(ss);
  setupIndividualLookupSheet(ss);
  setupAgentPivotSheet(ss);
  setupZeroLaborSheet(ss);
  setupAgentByCategorySheet(ss);
  setupTeamByCategorySheet(ss);
  setupCategoryByLaborTypeSheet(ss);
  setupLocationByCategorySheet(ss);
  setupDashboardSheet(ss);
  setupYearSummarySheet(ss);

  // Tab order and theming are cosmetic and run last, so by this point every
  // sheet has already been rebuilt and committed. Neither is allowed to throw
  // away that work: on a large spreadsheet these two passes are the ones that
  // hit "Service Spreadsheets failed while accessing document", and an
  // exception here previously surfaced as a total failure even though the
  // rebuild had succeeded. Both are safe to re-run.
  try {
    reorderSheets(ss);
  } catch (e) {
    logOperation('SETUP', 'WARNING', 'Tab reorder failed: ' + e.message);
  }
  try {
    applyBrandTheme_(ss);
  } catch (e) {
    logOperation('SETUP', 'WARNING', 'Brand theme failed: ' + e.message);
  }
  logOperation('SETUP', 'SUCCESS', 'Regenerated all analytics sheets');
}

function regenerateAnalyticsSheetsWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Regenerate Analytics Sheets',
    'This will delete and recreate all formula-driven analytics sheets ' +
    '(LaborTypes, DateFilters, all By* rollups, IndividualLookup, AgentPivot, ZeroLabor, cross-dimension sheets, Dashboard, YearSummary).\n\n' +
    'Data sheets (RawData, ActivityLog, LaborTypes, etc.) will NOT be affected.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  regenerateAnalyticsSheets();
  ui.alert('Done', 'Analytics sheets have been regenerated.', ui.ButtonSet.OK);
}
