/**
 * Triggers.gs - Time-driven triggers
 */

function ensureMonitorTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasMonitor = triggers.some(function(t) {
    return t.getHandlerFunction() === 'triggerDataLoadMonitor';
  });
  if (!hasMonitor) {
    // Policy: installing a time-based trigger requires telemetry opt-in.
    assertTelemetryEnabledForTriggers();
    ScriptApp.newTrigger('triggerDataLoadMonitor')
      .timeBased()
      .everyMinutes(10)
      .create();
    logOperation('TRIGGERS', 'INFO', 'Auto-installed monitor trigger for load continuation');
  }
}

function setupDefaultTriggers() {
  const ui = SpreadsheetApp.getUi();

  // Policy: Automated polling requires telemetry opt-in. Throws if
  // TELEMETRY_ENABLED is not TRUE in the Config sheet.
  try {
    assertTelemetryEnabledForTriggers();
  } catch (e) {
    ui.alert(
      'Telemetry Required',
      e.message + '\n\n' +
      'To enable automated triggers:\n' +
      '1. Open the Config sheet\n' +
      '2. Set TELEMETRY_ENABLED to TRUE\n' +
      '3. Re-run Setup Automated Triggers',
      ui.ButtonSet.OK
    );
    logOperation('TRIGGERS', 'BLOCKED', e.message);
    return;
  }

  removeAllTriggers();

  ScriptApp.newTrigger('triggerDataLoadMonitor')
    .timeBased()
    .everyMinutes(10)
    .create();

  ScriptApp.newTrigger('triggerDailyOpenRefresh')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  logOperation('TRIGGERS', 'SUCCESS', 'Installed default triggers');
  ui.alert('Triggers Installed', 'Monitor (10 min) and daily open refresh installed.', ui.ButtonSet.OK);
}

function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  logOperation('TRIGGERS', 'SUCCESS', 'Removed all triggers');
}

function showAutomationStatus() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    ui.alert('Automation Status', 'No triggers installed.', ui.ButtonSet.OK);
    return;
  }
  const lines = triggers.map(t => '• ' + t.getHandlerFunction()).join('\n');
  ui.alert('Automation Status', triggers.length + ' trigger(s) active:\n' + lines, ui.ButtonSet.OK);
}

function triggerDailyOpenRefresh() {
  // Policy: Automated polling requires telemetry opt-in.
  if (!enforceTelemetryGate()) return;

  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('TRIGGER_OPEN_REFRESH', 'SKIP', 'Another operation is running');
    return;
  }
  try {
    logOperation('TRIGGER_OPEN_REFRESH', 'INFO', 'Daily open refresh triggered');

    // Skip if initial load is still running
    if (getLoadState(DATA_LOAD_TYPES.TICKETS) !== LOAD_STATES.COMPLETE ||
        getLoadState(DATA_LOAD_TYPES.ACTIVITIES) !== LOAD_STATES.COMPLETE) {
      logOperation('TRIGGER_OPEN_REFRESH', 'INFO', 'Initial load incomplete - skipping');
      return;
    }

    setLoadState(DATA_LOAD_TYPES.OPEN_REFRESH, LOAD_STATES.PENDING);
    refreshOpenTickets();
  } finally {
    releaseScriptLock(lock);
  }

  reportTelemetry();
}
