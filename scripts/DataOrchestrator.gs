/**
 * DataOrchestrator.gs - Load orchestration and monitoring
 */

const DATA_LOAD_TYPES = {
  TEAMS: 'TEAMS',
  USERS: 'USERS',
  LABOR_TYPES: 'LABOR_TYPES',
  RESOLUTION_ACTIONS: 'RESOLUTION_ACTIONS',
  TICKETS: 'TICKETS',
  ACTIVITIES: 'ACTIVITIES',
  TICKET_RECONCILE: 'TICKET_RECONCILE',
  OPEN_REFRESH: 'OPEN_REFRESH'
};

const LOAD_GROUPS = {
  1: [DATA_LOAD_TYPES.TEAMS, DATA_LOAD_TYPES.USERS, DATA_LOAD_TYPES.LABOR_TYPES, DATA_LOAD_TYPES.RESOLUTION_ACTIONS],
  2: [DATA_LOAD_TYPES.TICKETS],
  3: [DATA_LOAD_TYPES.ACTIVITIES],
  4: [DATA_LOAD_TYPES.TICKET_RECONCILE]
};

const LOAD_STATES = {
  IDLE: 'idle',
  PENDING: 'pending',
  IN_PROGRESS: 'in_progress',
  COMPLETE: 'complete',
  ERROR: 'error'
};

function getLoadState(dataType) {
  return getConfig('LOAD_STATE_' + dataType) || LOAD_STATES.IDLE;
}

function setLoadState(dataType, state, message) {
  setConfig('LOAD_STATE_' + dataType, state);
  if (message) {
    setConfig('LOAD_STATE_' + dataType + '_MESSAGE', message);
  }
  logOperation('ORCHESTRATOR', 'INFO', dataType + ' -> ' + state + (message ? ': ' + message : ''));
}

function clearLoadStates() {
  Object.keys(DATA_LOAD_TYPES).forEach(key => {
    setLoadState(DATA_LOAD_TYPES[key], LOAD_STATES.IDLE);
  });
}

function initializeAllLoads() {
  Object.keys(DATA_LOAD_TYPES).forEach(key => {
    const type = DATA_LOAD_TYPES[key];
    if (type === DATA_LOAD_TYPES.OPEN_REFRESH || type === DATA_LOAD_TYPES.TICKET_RECONCILE) return;
    setLoadState(type, LOAD_STATES.PENDING);
  });

  setConfig('USER_LOAD_TEAM_INDEX', '');
  setConfig('RESOLUTION_LOAD_PAGE', '');
  setConfig('TICKET_LOAD_PAGE', '');
  setConfig('TICKET_LOAD_TOTAL_PAGES', '');
  setConfig('ACTIVITY_TICKET_INDEX', '');
  setConfig('ACTIVITY_LAST_TICKET_ID', '');
  setConfig('ACTIVITY_BATCH_PAGE', '');
  setConfig('ACTIVITY_FAIL_QUEUE_READY', '');
  setConfig('OPEN_REFRESH_STAGE', '');
  setConfig('OPEN_REFRESH_OPEN_PAGE', '');
  setConfig('OPEN_REFRESH_CLOSED_PAGE', '');
  setConfig('TICKET_LOAD_EXPECTED_COUNT', '');
  setConfig('TICKET_LOAD_FIRST_TOTAL_ROWS', '');
  setConfig('TICKET_LOAD_TOTAL_ROWS_DRIFT', '');
  setConfig('TICKET_RECONCILE_PAGE', '');
  setConfig('TICKET_RECONCILE_ATTEMPTS', '');
}

function isGroupComplete(groupNum) {
  const types = LOAD_GROUPS[groupNum] || [];
  return types.every(type => getLoadState(type) === LOAD_STATES.COMPLETE);
}

function findPendingInGroup(groupNum) {
  const types = LOAD_GROUPS[groupNum] || [];
  for (const type of types) {
    if (getLoadState(type) === LOAD_STATES.IN_PROGRESS) return type;
  }
  for (const type of types) {
    if (getLoadState(type) === LOAD_STATES.PENDING) return type;
  }
  for (const type of types) {
    if (getLoadState(type) === LOAD_STATES.ERROR) return type;
  }
  return null;
}

function getNextPendingLoad() {
  for (const groupNum in LOAD_GROUPS) {
    const group = parseInt(groupNum, 10);
    for (let g = 1; g < group; g++) {
      if (!isGroupComplete(g)) {
        return findPendingInGroup(g);
      }
    }
    const pending = findPendingInGroup(group);
    if (pending) return pending;
  }
  return null;
}

function executeNextLoad() {
  const lock = acquireScriptLock();
  if (!lock) {
    showOperationBusyMessage('Continue Loading');
    return false;
  }
  try {
    return executeNextLoadInternal_();
  } finally {
    releaseScriptLock(lock);
  }
}

function executeNextLoadInternal_() {
  const validation = validateConfig();
  if (!validation.isValid) {
    logOperation('ORCHESTRATOR', 'ERROR', 'Missing config: ' + validation.missing.join(', '));
    return false;
  }

  const nextLoad = getNextPendingLoad();
  if (!nextLoad) {
    const openState = getLoadState(DATA_LOAD_TYPES.OPEN_REFRESH);
    if (openState === LOAD_STATES.PENDING || openState === LOAD_STATES.IN_PROGRESS) {
      refreshOpenTickets();
      return true;
    }
    logOperation('ORCHESTRATOR', 'INFO', 'No pending loads');
    return false;
  }

  try {
    switch (nextLoad) {
      case DATA_LOAD_TYPES.TEAMS:
        setLoadState(nextLoad, LOAD_STATES.IN_PROGRESS);
        loadTeams();
        setLoadState(nextLoad, LOAD_STATES.COMPLETE);
        break;
      case DATA_LOAD_TYPES.USERS:
        loadUsers();
        break;
      case DATA_LOAD_TYPES.LABOR_TYPES:
        setLoadState(nextLoad, LOAD_STATES.IN_PROGRESS);
        loadLaborTypes();
        setLoadState(nextLoad, LOAD_STATES.COMPLETE);
        break;
      case DATA_LOAD_TYPES.RESOLUTION_ACTIONS:
        loadResolutionActions();
        break;
      case DATA_LOAD_TYPES.TICKETS:
        loadTicketsPaginated();
        break;
      case DATA_LOAD_TYPES.ACTIVITIES:
        loadActivitiesInitial();
        break;
      case DATA_LOAD_TYPES.TICKET_RECONCILE:
        reconcileTickets();
        break;
      default:
        throw new Error('Unknown load type: ' + nextLoad);
    }

    return true;
  } catch (error) {
    setLoadState(nextLoad, LOAD_STATES.ERROR, error.message);
    logOperation('ORCHESTRATOR', 'ERROR', 'Load failed: ' + nextLoad + ' - ' + error.message);
    return false;
  }
}

function triggerDataLoadMonitor() {
  const lock = tryAcquireScriptLock();
  if (!lock) {
    logOperation('TRIGGER_MONITOR', 'SKIP', 'Another operation is running');
    return;
  }
  try {
    logOperation('TRIGGER_MONITOR', 'INFO', 'Monitor triggered');
    const executed = executeNextLoadInternal_();
    if (executed) {
      logOperation('TRIGGER_MONITOR', 'SUCCESS', 'Load executed');
    }
  } catch (error) {
    logOperation('TRIGGER_MONITOR', 'ERROR', error.message);
  } finally {
    releaseScriptLock(lock);
  }
}

function startInitialLoad() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Start Initial Load',
    'This will load teams, users, resolution actions, tickets, and activities for the configured school year.\n\n' +
    'The process is resumable and may take hours for large datasets.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Start Initial Load'); return; }
  try {
    lockSchoolYearConfig();
    assertSchoolYearUnchanged();
    ensureDataSheetsProtected();
    initializeAllLoads();

    // Loop through loads — quick loads (reference data) run back-to-back,
    // long loads (tickets/activities) run once then the trigger handles continuation
    while (true) {
      var previousPending = getNextPendingLoad();
      var executed = executeNextLoadInternal_();
      if (!executed) break;
      var currentPending = getNextPendingLoad();
      // If the same type is still pending, it paused (time limit) — stop looping
      if (currentPending && currentPending === previousPending) break;
    }

    // Auto-install monitor trigger if there are remaining loads
    if (getNextPendingLoad()) {
      ensureMonitorTrigger();
    }
  } finally {
    releaseScriptLock(lock);
  }
}

function startOpenRefresh() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Start Open Ticket Refresh',
    'This will refresh open tickets and recently closed tickets (last 14 days) within the configured school year.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Open Ticket Refresh'); return; }
  try {
    setLoadState(DATA_LOAD_TYPES.OPEN_REFRESH, LOAD_STATES.PENDING);
    refreshOpenTickets();
  } finally {
    releaseScriptLock(lock);
  }
}

function showLoadStatus() {
  const ui = SpreadsheetApp.getUi();
  const lines = [];
  Object.keys(DATA_LOAD_TYPES).forEach(key => {
    const type = DATA_LOAD_TYPES[key];
    lines.push(type + ': ' + getLoadState(type));
  });
  ui.alert('Load Status', lines.join('\n'), ui.ButtonSet.OK);
}
