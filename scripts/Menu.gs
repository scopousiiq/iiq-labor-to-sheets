/**
 * Menu.gs - iiQ Data menu
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('iiQ Labor')
    .addItem('▶ Continue Loading', 'executeNextLoad')
    .addItem('Check Status', 'showLoadStatus')
    .addItem('Open Dashboard', 'openDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Run Complete Setup', 'setupLaborTrackerDashboard')
      .addItem('Regenerate Analytics Sheets', 'regenerateAnalyticsSheetsWithConfirm')
      .addItem('Verify Configuration', 'showConfigStatus')
      .addItem('Test API Connection', 'showApiTestResult')
      .addSeparator()
      .addItem('Setup Automated Triggers', 'setupDefaultTriggers')
      .addItem('View Trigger Status', 'showAutomationStatus')
      .addItem('Remove Automated Triggers', 'removeAllTriggers')
      .addSeparator()
      .addItem('Check for Updates', 'menuCheckForUpdates'))
    .addSubMenu(ui.createMenu('Labor Data')
      .addItem('Start Initial Load', 'startInitialLoad')
      .addItem('Open Ticket Refresh', 'startOpenRefresh')
      .addItem('Refresh Reference Data', 'menuRefreshReferenceData')
      .addItem('Refresh Labor Types', 'menuRefreshLaborTypes')
      .addSeparator()
      .addItem('Show Status', 'showLoadStatus'))
    .addSubMenu(ui.createMenu('Troubleshooting')
      .addItem('Validate Data', 'showValidationResults')
      .addItem('Backfill Missing User Names', 'menuBackfillMissingUserNames')
      .addItem('Recompute Net Hours', 'menuRecomputeNetHours')
      .addItem('Repair Team Attribution', 'menuRepairTeamAttribution')
      .addSeparator()
      .addItem('View Logs', 'showLogs')
      .addItem('Reset Load States', 'resetLoadStatesWithConfirm')
      .addItem('Full Reload (Clear Data)', 'startFullReloadWithConfirm')
      .addSeparator()
      .addItem('Send Telemetry Ping (Debug)', 'menuSendTelemetryPing'))
    .addToUi();
}

function openDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function showLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Logs');
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Logs sheet not found.');
  }
}

function showConfigStatus() {
  const validation = validateConfig();
  const ui = SpreadsheetApp.getUi();

  if (validation.isValid) {
    ui.alert('Configuration Valid',
      'All required settings are configured.\n\n' +
      'API URL: ' + getConfig('API_BASE_URL') + '\n' +
      'Site ID: ' + getConfig('SITE_ID') + '\n' +
      'School Year: ' + getConfig('SCHOOL_YEAR_START') + ' to ' + getConfig('SCHOOL_YEAR_END') + '\n' +
      'School Year Locked: ' + (isSchoolYearLocked() ? 'Yes' : 'No') + '\n' +
      'PAGE_SIZE Locked: ' + (getConfig('PAGE_SIZE_LOCKED') ? 'Yes' : 'No'),
      ui.ButtonSet.OK);
  } else {
    ui.alert('Configuration Incomplete',
      'Missing required settings:\n\n' + validation.missing.join('\n') +
      '\n\nPlease update the Config sheet.',
      ui.ButtonSet.OK);
  }
}

function menuRefreshReferenceData() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Reference Data'); return; }
  try {
    refreshReferenceData();
    SpreadsheetApp.getUi().alert('Done', 'Reference data refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    releaseScriptLock(lock);
  }
}

function menuCheckForUpdates() {
  const ui = SpreadsheetApp.getUi();
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Check for Updates'); return; }
  try {
    const result = checkForUpdates();
    if (!result || !result.ok) {
      ui.alert('Update Check Failed',
        'Could not reach GitHub. See the Logs sheet for details.',
        ui.ButtonSet.OK);
      return;
    }
    if (result.updateAvailable) {
      ui.alert('Update Available',
        'Installed: v' + result.current + '\n' +
        'Available: v' + result.latest + '\n\n' +
        (result.message ? result.message + '\n\n' : '') +
        (result.releaseUrl ? 'Release: ' + result.releaseUrl : ''),
        ui.ButtonSet.OK);
    } else {
      ui.alert('Up to Date',
        'v' + result.current + ' is the latest version.',
        ui.ButtonSet.OK);
    }
  } finally {
    releaseScriptLock(lock);
  }
}

function menuRefreshLaborTypes() {
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Refresh Labor Types'); return; }
  try {
    loadLaborTypes();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('LaborTypes');
    const count = sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
    SpreadsheetApp.getUi().alert('Done', 'Labor types refreshed from API (' + count + ' types loaded).', SpreadsheetApp.getUi().ButtonSet.OK);
  } finally {
    releaseScriptLock(lock);
  }
}

// Repairs blank PerformedByUser entries in ActivityLog:
//   1. Pull team-less agents from /v1.0/users/agents into the Users sheet.
//   2. Backfill any remaining unknown user IDs (e.g. demoted agents) by
//      fetching /v1.0/users/{id} individually.
//   3. Rewrite blank ActivityLog!N values in place from the now-complete Users sheet.
function menuBackfillMissingUserNames() {
  const ui = SpreadsheetApp.getUi();
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Backfill Missing User Names'); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
      ui.alert('Backfill Failed', 'Users sheet not found. Run Setup first.', ui.ButtonSet.OK);
      return;
    }

    const addedAgents = augmentUsersWithAgents(usersSheet);
    const addedFromActivity = backfillMissingUsersFromActivities();
    const result = repairBlankPerformedByUserNames();

    logOperation('USER_BACKFILL', 'SUCCESS',
      'Added ' + addedAgents + ' agent(s) + ' + addedFromActivity + ' historical user(s); ' +
      'repaired ' + result.repaired + ' ActivityLog row(s); ' + result.stillMissing + ' still missing');

    ui.alert('Backfill Complete',
      'Users sheet:\n' +
      '  +' + addedAgents + ' team-less agents\n' +
      '  +' + addedFromActivity + ' historical users\n\n' +
      'ActivityLog:\n' +
      '  ' + result.repaired + ' row(s) repaired\n' +
      '  ' + result.stillMissing + ' row(s) still missing (user not retrievable)',
      ui.ButtonSet.OK);
  } catch (e) {
    logOperation('USER_BACKFILL', 'ERROR', e.message);
    ui.alert('Backfill Failed', e.message, ui.ButtonSet.OK);
  } finally {
    releaseScriptLock(lock);
  }
}

// Recomputes the de-duplicated NetHours column (and EntryType) for every
// ActivityLog row. Useful after manually editing data, or to repair an older
// sheet loaded before NetHours existed, without a full reload.
function menuRecomputeNetHours() {
  const ui = SpreadsheetApp.getUi();
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Recompute Net Hours'); return; }
  try {
    const rows = computeNetHours();
    ui.alert('Done', 'Recomputed EntryType/NetHours for ' + rows + ' ActivityLog row(s).', ui.ButtonSet.OK);
  } catch (e) {
    logOperation('NET_HOURS', 'ERROR', e.message);
    ui.alert('Recompute Failed', e.message, ui.ButtonSet.OK);
  } finally {
    releaseScriptLock(lock);
  }
}

// Recomputes ActivityLog Team (cols Q/R) in place using the hybrid rule
// (ticket's assigned team, else the performing user's team, else "(Unassigned)").
// Lets an already-loaded sheet adopt the team-attribution fix without a reload.
function menuRepairTeamAttribution() {
  const ui = SpreadsheetApp.getUi();
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Repair Team Attribution'); return; }
  try {
    const rows = repairTeamAttribution();
    ui.alert('Done', 'Repaired team attribution for ' + rows + ' ActivityLog row(s).', ui.ButtonSet.OK);
  } catch (e) {
    logOperation('TEAM_REPAIR', 'ERROR', e.message);
    ui.alert('Repair Failed', e.message, ui.ButtonSet.OK);
  } finally {
    releaseScriptLock(lock);
  }
}

// Sends a one-off anonymous usage ping for debugging telemetry connectivity.
// Honors the TELEMETRY_ENABLED opt-out (no-op if FALSE); bypasses only the
// trigger-presence check so it works before triggers are installed.
function menuSendTelemetryPing() {
  const ui = SpreadsheetApp.getUi();
  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Send Telemetry Ping'); return; }
  try {
    reportTelemetry(true);
    ui.alert('Telemetry Ping Sent',
      'A debug ping was attempted. It only sends if TELEMETRY_ENABLED is TRUE in the Config sheet. ' +
      'No ticket/labor data, tokens, or user names are included — see the Anonymous Usage Telemetry ' +
      'section of the Instructions sheet. Check Logs for details.',
      ui.ButtonSet.OK);
  } finally { releaseScriptLock(lock); }
}

function resetLoadStatesWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Load States',
    'This will reset all load states and progress.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Reset Load States'); return; }
  try {
    resetLoadStates();
  } finally {
    releaseScriptLock(lock);
  }

  ui.alert('Reset Complete', 'Load states have been reset.', ui.ButtonSet.OK);
}

function startFullReloadWithConfirm() {
  if (!requireNoTriggers('Full Reload')) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Full Reload (Clear Data)',
    'This will delete all data from RawData, ActivityLog, and reference sheets.\n' +
    'It will also unlock the school year configuration so you can change dates.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const confirm = ui.alert('Confirm', 'This cannot be undone. Continue?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Full Reload'); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    unlockSchoolYearConfig();

    const sheetsToClear = [
      'RawData',
      'ActivityLog',
      'Teams',
      'Users',
      'LaborTypes',
      'ResolutionActions',
      'ActivityFailures',
      'TicketIndex',
      'ActivityIndex'
    ];
    sheetsToClear.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
      }
    });

    setConfig('LAST_SYNC', '');
    resetLoadStates();
    ensureDataSheetsProtected();
  } finally {
    releaseScriptLock(lock);
  }

  ui.alert('Data Cleared', 'All data has been cleared and the school year is unlocked.\n\nUpdate the school year dates, then run Start Initial Load.', ui.ButtonSet.OK);
}

function resetAllDataWithConfirm() {
  startFullReloadWithConfirm();
}

function resetLoadStates() {
  clearLoadStates();
  setConfig('USER_LOAD_TEAM_INDEX', '');
  setConfig('RESOLUTION_LOAD_PAGE', '');
  setConfig('TICKET_LOAD_PAGE', '');
  setConfig('TICKET_LOAD_TOTAL_PAGES', '');
  setConfig('ACTIVITY_TICKET_INDEX', '');
  setConfig('ACTIVITY_LAST_TICKET_ID', '');
  setConfig('ACTIVITY_BATCH_PAGE', '');
  setConfig('OPEN_REFRESH_STAGE', '');
  setConfig('OPEN_REFRESH_OPEN_PAGE', '');
  setConfig('OPEN_REFRESH_CLOSED_PAGE', '');
  setConfig('TICKET_LOAD_EXPECTED_COUNT', '');
  setConfig('TICKET_LOAD_FIRST_TOTAL_ROWS', '');
  setConfig('TICKET_LOAD_TOTAL_ROWS_DRIFT', '');
  setConfig('TICKET_RECONCILE_PAGE', '');
  setConfig('TICKET_RECONCILE_ATTEMPTS', '');
}
