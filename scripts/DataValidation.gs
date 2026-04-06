/**
 * DataValidation.gs - Post-load validation and ticket reconciliation
 */

const VALIDATION_SAMPLE_CAP = 20;
const MAX_RECONCILE_ATTEMPTS = 2;

/**
 * Runs all validation checks and returns a structured result object.
 * Called automatically after activities complete, or manually via menu.
 */
function validateLoadedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    ticketCount: validateTicketCount_(ss),
    laborMinutes: validateLaborMinutes_(ss),
    timestamp: new Date().toISOString()
  };

  logValidationResults_(result);
  return result;
}

/**
 * Runs validation and triggers reconcile if needed.
 * Called after initial load completes. Separated from validateLoadedData()
 * so reconcileTickets() can call validateLoadedData() without re-triggering itself.
 */
function validateAndTriggerReconcile() {
  var result = validateLoadedData();

  // Trigger ticket reconcile if actual < expected
  if (result.ticketCount.expected > 0 && result.ticketCount.delta > 0) {
    setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.PENDING);
    logOperation('VALIDATION', 'WARNING',
      'Ticket count shortfall detected (' + result.ticketCount.delta +
      ' missing). TICKET_RECONCILE set to pending.');
  }

  return result;
}

// --- Ticket Count Validation ---

function validateTicketCount_(ss) {
  const expected = getIntValue(getConfig('TICKET_LOAD_EXPECTED_COUNT'), -1);
  const drift = getStringValue(getConfig('TICKET_LOAD_TOTAL_ROWS_DRIFT'));

  // Count unique non-blank IDs from TicketIndex
  const indexSheet = ss.getSheetByName('TicketIndex');
  let indexCount = 0;
  if (indexSheet && indexSheet.getLastRow() > 1) {
    const indexIds = indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 1).getValues();
    const indexSet = {};
    indexIds.forEach(function(row) {
      if (row[0]) indexSet[row[0]] = true;
    });
    indexCount = Object.keys(indexSet).length;
  }

  // Cross-check against RawData non-blank col A
  const rawSheet = ss.getSheetByName('RawData');
  let rawCount = 0;
  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawIds = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 1).getValues();
    const rawSet = {};
    rawIds.forEach(function(row) {
      if (row[0]) rawSet[row[0]] = true;
    });
    rawCount = Object.keys(rawSet).length;
  }

  let indexCorruption = false;
  if (indexCount !== rawCount) {
    indexCorruption = true;
    logOperation('VALIDATION', 'WARNING',
      'TicketIndex count (' + indexCount + ') != RawData unique count (' + rawCount +
      '). Index may be corrupt.');
  }

  // Use the higher of the two as actual (index corruption could undercount)
  const actual = Math.max(indexCount, rawCount);
  const delta = expected >= 0 ? expected - actual : 0;

  return {
    expected: expected,
    actual: actual,
    indexCount: indexCount,
    rawCount: rawCount,
    delta: delta,
    drift: drift || '',
    indexCorruption: indexCorruption
  };
}

// --- Labor Minute Reconciliation ---

function validateLaborMinutes_(ss) {
  const rawSheet = ss.getSheetByName('RawData');
  const actSheet = ss.getSheetByName('ActivityLog');

  if (!rawSheet || rawSheet.getLastRow() < 2) {
    return { checked: false, reason: 'No RawData' };
  }
  if (!actSheet || actSheet.getLastRow() < 2) {
    return { checked: false, reason: 'No ActivityLog data' };
  }

  // Build per-ticket sum from ActivityLog: col B (TicketId), col E (EffortMins)
  const actData = actSheet.getRange(2, 2, actSheet.getLastRow() - 1, 4).getValues();
  const actSums = {};
  actData.forEach(function(row) {
    var ticketId = row[0]; // col B
    var effortMins = row[3]; // col E (offset: B=0, C=1, D=2, E=3)
    if (!ticketId) return;
    if (!actSums[ticketId]) actSums[ticketId] = 0;
    actSums[ticketId] += (Number(effortMins) || 0);
  });

  // Read RawData: col A (TicketId), col F (TotalLaborMins)
  var rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 6).getValues();

  var missingActivities = [];
  var underReported = [];
  var overReported = [];
  var matched = 0;

  rawData.forEach(function(row) {
    var ticketId = row[0]; // col A
    var ticketLaborMins = Number(row[5]) || 0; // col F
    if (!ticketId) return;

    var activitySum = actSums[ticketId] || 0;

    // Skip tickets where both values are 0 — no labor expected or found
    if (ticketLaborMins === 0 && activitySum === 0) {
      matched++;
    } else if (ticketLaborMins > 0 && activitySum === 0) {
      missingActivities.push(ticketId);
    } else if (activitySum < ticketLaborMins) {
      underReported.push(ticketId);
    } else if (activitySum > ticketLaborMins) {
      overReported.push(ticketId);
    } else {
      matched++;
    }
  });

  return {
    checked: true,
    matched: matched,
    missingActivities: missingActivities.length,
    missingActivitiesSample: missingActivities.slice(0, VALIDATION_SAMPLE_CAP),
    underReported: underReported.length,
    underReportedSample: underReported.slice(0, VALIDATION_SAMPLE_CAP),
    overReported: overReported.length,
    overReportedSample: overReported.slice(0, VALIDATION_SAMPLE_CAP)
  };
}

// --- Structured Logging ---

function logValidationResults_(result) {
  var tc = result.ticketCount;
  var lm = result.laborMinutes;

  var lines = [
    'expected_tickets=' + tc.expected,
    'actual_unique_tickets=' + tc.actual,
    'index_count=' + tc.indexCount,
    'raw_count=' + tc.rawCount,
    'mismatch_delta=' + tc.delta,
    'drift=' + (tc.drift || 'none'),
    'index_corruption=' + tc.indexCorruption
  ];

  if (lm.checked) {
    lines.push('matched=' + lm.matched);
    lines.push('missing_activities=' + lm.missingActivities);
    lines.push('under_reported=' + lm.underReported);
    lines.push('over_reported=' + lm.overReported);

    if (lm.missingActivities > 0) {
      lines.push('missing_sample=[' + lm.missingActivitiesSample.join(',') + ']');
    }
    if (lm.underReported > 0) {
      lines.push('under_sample=[' + lm.underReportedSample.join(',') + ']');
    }
    if (lm.overReported > 0) {
      lines.push('over_sample=[' + lm.overReportedSample.join(',') + ']');
    }
  } else {
    lines.push('labor_check_skipped=' + lm.reason);
  }

  var status = (tc.delta > 0 || (lm.checked && lm.missingActivities > 0))
    ? 'WARNING' : 'SUCCESS';

  logOperation('VALIDATION', status, lines.join(' | '));
}

// --- Manual Menu Entry Point ---

function showValidationResults() {
  var lock = acquireScriptLock();
  if (!lock) { showOperationBusyMessage('Validate Data'); return; }

  try {
    var result = validateLoadedData();
  } finally {
    releaseScriptLock(lock);
  }

  var tc = result.ticketCount;
  var lm = result.laborMinutes;

  var msg = 'TICKET COUNT\n';
  msg += 'Expected: ' + tc.expected + '\n';
  msg += 'Actual (unique): ' + tc.actual + '\n';
  if (tc.delta > 0) {
    msg += 'MISSING: ' + tc.delta + ' tickets\n';
  } else {
    msg += 'Status: OK\n';
  }
  if (tc.drift) {
    msg += 'Dataset drift during load: ' + tc.drift + '\n';
  }
  if (tc.indexCorruption) {
    msg += 'WARNING: TicketIndex/RawData count mismatch\n';
  }

  msg += '\nLABOR MINUTES\n';
  if (lm.checked) {
    msg += 'Matched: ' + lm.matched + '\n';
    msg += 'Missing all activities: ' + lm.missingActivities + '\n';
    msg += 'Under-reported (partial): ' + lm.underReported + '\n';
    msg += 'Over-reported (timing): ' + lm.overReported + '\n';
  } else {
    msg += 'Skipped: ' + lm.reason + '\n';
  }

  SpreadsheetApp.getUi().alert('Data Validation Results', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- Ticket Reconciliation ---

function reconcileTickets() {
  var attempts = getIntValue(getConfig('TICKET_RECONCILE_ATTEMPTS'), 0);
  if (attempts >= MAX_RECONCILE_ATTEMPTS) {
    logOperation('TICKET_RECONCILE', 'SKIP', 'Max attempts (' + MAX_RECONCILE_ATTEMPTS + ') already reached.');
    return;
  }

  var expectedCount = getIntValue(getConfig('TICKET_LOAD_EXPECTED_COUNT'), -1);
  if (expectedCount < 0) {
    logOperation('TICKET_RECONCILE', 'SKIP', 'No expected count stored. Nothing to reconcile against.');
    setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.COMPLETE);
    return;
  }

  ensureSchoolYearLocked();
  assertSchoolYearUnchanged();
  cacheConfigRowPositions();
  const startTime = Date.now();
  const pageSize = getPageSize();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RawData');
  if (!sheet) throw new Error('RawData sheet not found.');

  setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.IN_PROGRESS);

  const ticketIndex = buildTicketIndexMap(sheet);
  const range = getSchoolYearRange();
  const filters = buildTicketFilters(range.startDate, range.endDate);

  let page = getIntValue(getConfig('TICKET_RECONCILE_PAGE'), 0);
  if (page < 0) page = 0;

  var appendedTicketIds = [];

  while (true) {
    if (Date.now() - startTime >= MAX_RUNTIME_MS) {
      writeConfigValueDirect('TICKET_RECONCILE_PAGE', String(page));
      logOperation('TICKET_RECONCILE', 'INFO', 'Paused at page ' + page);
      return;
    }

    var response = searchTicketsPage(filters, page, pageSize);
    var items = response && response.Items ? response.Items : [];

    if (items.length === 0) break;

    // Identify tickets not already in the index (these are the missing ones)
    items.forEach(function(item) {
      var tid = item.TicketId || '';
      if (tid && !ticketIndex[tid]) {
        appendedTicketIds.push(tid);
      }
    });

    upsertTickets(items, sheet, ticketIndex);

    page++;
    writeConfigValueDirect('TICKET_RECONCILE_PAGE', String(page));

    if (response && response.Paging && response.Paging.PageCount !== undefined) {
      if (page >= response.Paging.PageCount) break;
    }
  }

  // Fetch activities for any newly recovered tickets
  if (appendedTicketIds.length > 0) {
    logOperation('TICKET_RECONCILE', 'INFO',
      'Recovered ' + appendedTicketIds.length + ' tickets. Fetching their activities.');
    refreshActivitiesForTickets(appendedTicketIds);
  }

  // Re-validate after reconcile pass
  writeConfigValueDirect('TICKET_RECONCILE_PAGE', '');
  rebuildTicketIndex(sheet);

  var postResult = validateLoadedData();
  if (postResult.ticketCount.delta > 0) {
    attempts++;
    writeConfigValueDirect('TICKET_RECONCILE_ATTEMPTS', String(attempts));

    if (attempts >= MAX_RECONCILE_ATTEMPTS) {
      setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.ERROR,
        'Persistent shortfall after ' + attempts + ' attempts. Delta: ' + postResult.ticketCount.delta);
      writeConfigValueDirect('TICKET_LOAD_EXPECTED_COUNT', '');
      logOperation('TICKET_RECONCILE', 'ERROR',
        'Max attempts reached. Persistent delta: ' + postResult.ticketCount.delta +
        '. Clearing expected count to stop retries.');
      return;
    }

    logOperation('TICKET_RECONCILE', 'WARNING',
      'Still short by ' + postResult.ticketCount.delta + ' after attempt ' + attempts +
      '. Will retry on next trigger cycle.');
    setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.PENDING);
    return;
  }

  setLoadState(DATA_LOAD_TYPES.TICKET_RECONCILE, LOAD_STATES.COMPLETE);
  writeConfigValueDirect('TICKET_RECONCILE_ATTEMPTS', '');
  logOperation('TICKET_RECONCILE', 'SUCCESS',
    'Reconciliation complete. Ticket counts match.' +
    (appendedTicketIds.length > 0 ? ' Recovered ' + appendedTicketIds.length + ' tickets.' : ''));
}
