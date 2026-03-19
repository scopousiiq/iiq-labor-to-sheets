/**
 * Menu.gs - iiQ Data menu
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('iiQ Data')
    .addItem('Check Status', 'showLoadStatus')
    .addItem('View Dashboard', 'openDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('Run Complete Setup', 'setupLaborTrackerDashboard')
      .addItem('Regenerate Analytics Sheets', 'regenerateAnalyticsSheetsWithConfirm')
      .addItem('Test API Connection', 'showApiTestResult')
      .addItem('Verify Configuration', 'showConfigStatus')
      .addSeparator()
      .addItem('Setup Automated Triggers', 'setupDefaultTriggers')
      .addItem('Remove Automated Triggers', 'removeAllTriggers')
      .addItem('View Trigger Status', 'showAutomationStatus'))
    .addSubMenu(ui.createMenu('Load Data')
      .addItem('Start Initial Load', 'startInitialLoad')
      .addItem('Continue Loading', 'executeNextLoad')
      .addItem('Refresh Reference Data', 'menuRefreshReferenceData')
      .addItem('Refresh Labor Types', 'menuRefreshLaborTypes')
      .addItem('Open Ticket Refresh', 'startOpenRefresh'))
    .addSubMenu(ui.createMenu('Troubleshooting')
      .addItem('View Logs', 'showLogs')
      .addItem('Reset Load States', 'resetLoadStatesWithConfirm')
      .addItem('Full Reload (Clear Data)', 'startFullReloadWithConfirm'))
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
}
