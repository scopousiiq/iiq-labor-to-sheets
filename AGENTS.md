# AGENTS.md — iiQ Labor Tracker

## Project Structure

All Google Apps Script source files are in `scripts/`. There is no build system, package manager, or test framework. Functions are tested by running them directly in the Apps Script editor.

## Script Files and Key Functions

### Config.gs — Foundation layer
- `getConfig(key)` / `setConfig(key, value)` — Config sheet key-value store (all values stored as strings)
- `getStringValue()`, `getIntValue()`, `getBoolValue()` — Type coercion helpers for Sheets values
- `cacheConfigRowPositions()` / `writeConfigValueDirect()` / `getConfigValueDirect()` — Cached config access for tight loops
- `acquireScriptLock()` / `tryAcquireScriptLock()` / `releaseScriptLock()` — LockService concurrency control
- `requireNoTriggers(operationName)` — Guards destructive operations
- `logOperation(operation, status, details)` — Logs to Logs sheet (newest first, auto-pruned)
- `getPageSize()`, `getThrottleMs()`, `getSchoolYearRange()` — Config accessors
- `lockSchoolYearConfig()` / `unlockSchoolYearConfig()` — School year locking

### ApiClient.gs — HTTP layer
- `apiRequest(method, endpoint, payload, retryCount)` — HTTP client with exponential backoff (max 5 retries)
- `fetchAllPagesWithQueryParams(endpoint, payload, method)` — Paginated fetch collecting all results
- `testApiConnection()` / `showApiTestResult()` — Connection test

### DataOrchestrator.gs — Load state machine
- `executeNextLoad()` — Finds and executes next pending load (acquires lock)
- `executeNextLoadInternal_()` — Core logic without lock (used by triggers)
- `triggerDataLoadMonitor()` — Trigger entry point (tryAcquire lock, skip if busy)
- `startInitialLoad()` — Menu entry: initializes all loads and starts first
- `startOpenRefresh()` — Menu entry: triggers open ticket refresh
- `initializeAllLoads()` — Sets all load states to pending
- Load groups: Group 1 (Teams, Users, ResolutionActions) → Group 2 (Tickets) → Group 3 (Activities)

### TicketData.gs — Ticket loading
- `loadTicketsPaginated()` — Paginated ticket loading with 5.5-min runtime guard and config caching
- `refreshOpenTickets()` — Two-stage (open then recently closed) ticket refresh
- `mapTicketRow(t)` — Maps API response to 21-column row (`IsClosed` → `'Closed'`/`'Open'`)
- `upsertTickets()` — Insert-or-update via index map
- `writeBatchedUpdates(sheet, rowMap)` — Shared batched write function (also used by ActivityLog)
- `buildTicketFilters()`, `buildOpenTicketFilters()`, `buildClosedTicketFilters()` — IIQ API filter builders
- `formatDateForApi(date)` — `M/D/YYYY` format for IIQ filters

### ActivityLog.gs — Activity loading
- `loadActivitiesInitial()` — Batch activity loading with config caching, failure retry
- `refreshActivitiesForTickets(ticketIds)` — Refresh activities for specific tickets
- `loadActivitiesBatch()` — Load activities for a batch of ticket IDs
- `buildActivityRow()` — Maps API response to 20-column row (`IsPublic` → `1`/`0`)
- `writeActivities()` — Upsert via shared `writeBatchedUpdates()`
- `buildTicketContextMap(rawSheet)` — Reads RawData using `RAWDATA_HEADERS.length` for column count
- `buildUserMap()`, `buildLaborTypeMap()`, `buildResolutionActionMap()` — Reference data lookups
- `recordActivityFailure()` / `clearActivityFailure()` / `processFailedActivityTickets()` — Failure tracking

### ReferenceData.gs — Reference data
- `refreshReferenceData()` — Loads teams, users, and resolution actions
- `loadTeams()`, `loadUsers()`, `loadResolutionActions()` — Individual reference loaders
- `clearSheetData(sheet)` — Clears data rows below header

### Index.gs — Fast lookups
- `rebuildTicketIndex(rawSheet)` / `rebuildActivityIndex(activitySheet)` — Rebuild index sheets
- `loadIndexMap(sheet)` — Load ID→row mapping from index sheet
- `ensureDataSheetsProtected()` — Protect RawData, ActivityLog, and index sheets

### Setup.gs — Sheet creation
- `RAWDATA_HEADERS` (21 items) / `ACTIVITY_HEADERS` (20 items) — Column layout constants
- `deleteSheetIfExists(ss, name)` — Helper for idempotent analytics sheet recreation
- `setupLaborTrackerDashboard()` — Main setup: creates all sheets
- `setupConfigSheet()`, `setupRawDataSheet()`, `setupActivityLogSheet()` — Data sheets (skip-if-exists)
- `setupByTeamSheet()`, `setupByIndividualSheet()`, etc. — Analytics sheets (delete-and-recreate)
- `regenerateAnalyticsSheets()` / `regenerateAnalyticsSheetsWithConfirm()` — Rebuild all formula sheets

### Triggers.gs — Automation
- `setupDefaultTriggers()` — Installs 10-min monitor + daily 2 AM refresh
- `removeAllTriggers()` — Removes all project triggers
- `triggerDailyOpenRefresh()` — Trigger entry point (tryAcquire lock, skip if busy)

### Menu.gs — UI layer
- `onOpen()` — Creates iiQ Data menu with Setup, Load Data, Troubleshooting submenus
- `menuRefreshReferenceData()` — Lock-wrapped reference data refresh
- `startFullReloadWithConfirm()` — Requires no triggers, clears all data
- `resetLoadStatesWithConfirm()` — Lock-wrapped state reset

## Coding Conventions
- 2-space indentation
- All config values stored as strings (never bare numbers)
- Boolean data values: `'Closed'`/`'Open'` or `1`/`0` (never `true`/`false`)
- Entry points wrapped with LockService (`acquireScriptLock` for menu, `tryAcquireScriptLock` for triggers)
- Column counts derived from header array constants, never hardcoded
