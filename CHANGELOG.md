# Changelog

## 2026-09-30 — Unreleased

### Fixed
- **Every report left out work logged on the last day of the selected date range.** Entries carry a time of day, but the range end is a plain date — midnight at the *start* of that day — and the formulas kept only entries at or before it. So for "Last Month" the final day of the month was missing, and for the "This …" ranges today's work didn't appear until tomorrow. Hours, cost, entry and ticket counts were all low by whatever was logged that day. Every date filter now runs through the end of the last day, on every rollup, both lookups, AgentPivot, ZeroLabor, the cross-dimension sheets, the Dashboard and YearSummary (whose school-year total also lost the final day). Existing sheets pick this up with **iiQ Labor → Setup → Regenerate Analytics Sheets** — no data reload needed. (`Setup.gs`)

## 2026-09-14 — Unreleased

### Added
- **LocationLookup sheet — labor by location.** Pick a location from a dropdown and see its total hours, cost, entry count, ticket count and overtime for the selected date range, plus three side-by-side breakdowns: the agents who worked there (with team), the issue categories the time went to, and every ticket with its hours, cost and entry count. A second dropdown narrows all of it to one issue type, listed as `Category > Type`; leave it blank for every type. The pairing matters in both directions: a bare type name is not unique (`Equipment Repair` and `Issue not listed` each recur under several categories, and merging them would combine unrelated work), while category alone is too coarse to reach a specific task — a type like Kitchen Tasks is invisible inside its parent category. The "By Issue Type" block groups at the same `Category > Type` grain for the same reason, and collapses capitalization variants (`Issue Not Listed` / `Issue not listed`) into one row showing the first spelling seen — without that step the block would list each variant separately and sum all of them into every row, totalling more than the location's real hours. Both dropdowns follow the DateFilters range, so they only offer locations and categories with activity in the period. This is the location counterpart to IndividualLookup. Note that per-location *totals* were already available on the `ByDepartment` sheet, which groups by location despite its name — LocationLookup is the drill-down into a single one.
- **Issue type filter on IndividualLookup.** A second dropdown narrows an individual's hours, cost, entry count, ticket count, overtime figures, activity detail and ticket summary to a single issue type. Leave it blank for every issue type, which reproduces the previous numbers exactly. Values are listed as `Category > Type` because issue type names are not unique on their own — the same type name recurs under several categories, and a bare name would silently merge unrelated work into one total. The list follows the selected date range, so it only offers types with activity in the period.
- **EntryType column on the IndividualLookup activity detail**, so a row showing 0.00 hours is visibly a resolution action that was superseded by a labor entry on the same ticket, rather than looking like missing data.

### Fixed
- **The IndividualLookup activity detail only ever filled its first four columns.** LaborCost, LaborType, ResolutionAction and Notes were always blank. The formula sliced its columns with `INDEX(range,,{4,5,6,7,8})`, and an array of column numbers makes `INDEX` return only the first of them — with no error to show anything was wrong. It now uses `CHOOSECOLS`, and every detail column populates.
- **IndividualLookup showed "1 ticket" when a selection had no tickets at all.** The ticket count wrapped `COUNTUNIQUE` around a filter that yields `#N/A` when empty, and the COUNTA family counts that error as one item. Now uses `ROWS(UNIQUE(FILTER(...)))`, which reports 0.
- **The IndividualLookup detail hours column did not tie to the Total Hours above it.** The detail listed raw EffortHours while the total used de-duplicated NetHours; the column now shows NetHours and the two agree.

- **"Regenerate Analytics Sheets" reported total failure on large spreadsheets even though every sheet had been rebuilt.** It failed with `Service Spreadsheets failed while accessing document with id ...`, thrown by the tab-reordering step that runs *after* all the sheets are built — so the rebuild had actually succeeded, but the tabs were left in creation order and unstyled, and nothing said so. Reordering and theming are now non-fatal: a failure in either is logged as a WARNING and the run completes. Reordering also skips tabs already in the right place and retries once, which cuts the document-structure operations that provoke the error in the first place, and the brand theme now guards each sheet individually so one large sheet can't leave every later tab unstyled. Both passes are safe to re-run.

### Changed
- IndividualLookup gained a row for the new filter, so its summary block and detail table each moved down one row, and the ticket summary moved one column right to leave a spacer.
- Number formats applied to the IndividualLookup hours, cost and count cells.

## 2026-07-01 — v1.1.1

### Fixed
- **White font on a white background on data tabs (ActivityLog and other banded sheets)** (`Setup.gs`) — The brand theme styled the header row (white font) *before* applying zebra banding. The banding's default theme and its white first-row color govern any font color not set explicitly, so it could render text white — including on the white banded rows — leaving data unreadable. `applyBrandTheme_` now bands first and styles the header last (so the header's white-on-blue is the final word), and `applyBanding_` pins data rows to a dark (`blueDeep`) font so no banding theme can produce white-on-white. Existing installs can repair the display via **iiQ Labor → Setup → Regenerate Analytics Sheets** (or Run Complete Setup).

## 2026-06-23 — v1.1.0

### Fixed
- **Labor time was double-counted in the hours reports.** For many tickets, iiQ stores the same work twice — once as a resolution *action* and once as a *labor* entry, each carrying the same hours — so the rollups added it up twice. The ActivityLog sheet now has a de-duplicated **NetHours** column plus an **EntryType** column marking each row `Action` or `Labor`. When a ticket has any labor entry, only its labor hours count; otherwise its action hours count — so each ticket's time is counted exactly once. Every hours rollup (ByTeam, ByIndividual, ByDepartment, ByIssueCategory, ByIssueType, AgentByCategory, TeamByCategory, LocationByCategory, AgentPivot total, Dashboard, IndividualLookup, YearSummary) now uses NetHours. ByLaborType, CategoryByLaborType, and ByResolution stay single-type views (labor-only / action-only) by design and are not expected to equal the de-duplicated total.
- **Team was attributed to the technician's home team instead of the ticket's team.** Activity rows now use the ticket's assigned team, falling back to the performing user's team when the ticket has no team assigned.
- **"System Service" time was not credited to a technician.** Time logged under the iiQ System Service automation account is re-attributed to the ticket's assigned user.
- **Report totals did not reconcile across breakdowns.** Rows with a blank team, location, category, or issue type were dropped from their breakdown; they now appear under a visible **(Unassigned)** bucket, so every full breakdown sums to the same total.
- **ByIssueCategory over-counted, and could collapse to a single row,** when a category's capitalization varied (e.g. "Issue Not Listed" vs "Issue not listed"). Capitalization variants are now merged into one row with a correct total. (Standardizing category names in iiQ is the cleaner long-term fix.)
- **IndividualLookup** — the Activity Date column only filled the first detail row; it now fills every row.
- **AgentPivot** — the per-labor-type columns did not add up to the Total; a new **Other** column captures hours that have no labor type, so the columns reconcile to the Total.

### Added
- **EntryType and NetHours columns** on the ActivityLog sheet.
- **Recompute Net Hours** and **Repair Team Attribution** (Troubleshooting menu) — apply the de-duplication and team rules to an already-loaded sheet without a full reload.
- **Send Telemetry Ping (Debug)** (Troubleshooting menu) — sends a single anonymous usage ping; honors the `TELEMETRY_ENABLED` opt-out.

### Changed
- **iiQ brand styling** applied across all tabs — colored header rows, zebra banding, tab colors, hero banners on the Dashboard and Instructions, and hidden gridlines on the report tabs.
- **Menu renamed to "iiQ Labor"** (from "iiQ Data") and reorganized — a top-level **Continue Loading**, **Open Dashboard**, a **Labor Data** submenu, and a tidier Setup and Troubleshooting layout.
- **Instructions sheet** updated for the new menu and styling, with added **Dashboard Integration** (Looker Studio / Power BI) and **Support** sections.
- **YearSummary** monthly rollup reworked for more reliable formula expansion.

## 2026-05-26

### Changed
- **BEARER_TOKEN documentation** (`README.md`, `Setup.gs`) — Updated "where to find your API token" guidance from "Admin > Integrations > API" to "Admin > Developer Tools" to match the current iiQ admin UI.

### Added
- **Anonymous usage telemetry** (`scripts/Telemetry.gs`, `Config.gs`, `Triggers.gs`, `DataOrchestrator.gs`, `Setup.gs`) — One ping per successful trigger refresh to the iiQ-owned aggregator; payload is install ID (UUID), project slug, version, iiQ hostname, ActivityLog row count, and the names of the standard analytics sheets present. No labor data, tokens, user names, or custom sheet names are sent. `enforceTelemetryGate()` runs at the head of every trigger-fired function (`triggerDataLoadMonitor`, `triggerDailyOpenRefresh`) and auto-uninstalls all CLOCK triggers if `TELEMETRY_ENABLED` is not TRUE. `assertTelemetryEnabledForTriggers()` blocks `setupDefaultTriggers` / `ensureMonitorTrigger` when telemetry is off. `reportTelemetry()` runs at the tail of each successful trigger.
- **Remote version check** (`Config.gs`, `Menu.gs`, `version.json`) — Added `SCRIPT_VERSION` constant (1.0.0), `version.json` at repo root, and `checkForUpdates()` that fetches the remote `version.json` from GitHub, compares semver, and writes `SCRIPT_VERSION`/`LATEST_VERSION`/`VERSION_CHECK_DATE` to the Config sheet (with yellow/green background to signal update available vs. up to date). New menu item: `iiQ Data → Setup → Check for Updates`.
- **Config sheet seeding** (`Setup.gs`) — `setupConfigSheet` now writes `SCRIPT_VERSION`, `LATEST_VERSION`, `VERSION_CHECK_DATE`, and `TELEMETRY_ENABLED=TRUE` rows on fresh install. Instructions sheet gets a new "Anonymous Usage Telemetry" section describing exactly what is and isn't sent, plus the opt-out instructions.

### Migration note
For installs created before this release, the `TELEMETRY_ENABLED` row will not yet exist in the Config sheet. Re-running `iiQ Data → Setup → Run Complete Setup` calls `ensureConfigSheetUpToDate()`, which idempotently appends `SCRIPT_VERSION`, `LATEST_VERSION`, `VERSION_CHECK_DATE`, and `TELEMETRY_ENABLED` rows (with the proper Notes column) to any existing Config sheet — and always stamps `SCRIPT_VERSION` to the running code version. Districts can flip `TELEMETRY_ENABLED` to FALSE at any time. (`assertTelemetryEnabledForTriggers` will also seed `TELEMETRY_ENABLED=TRUE` on first trigger install if the row is still missing, as a defensive fallback.)

### Fixed
- **Blank `PerformedByUser` in ActivityLog** (`ReferenceData.gs`, `ActivityLog.gs`, `Menu.gs`) — Users not on any team (admins, system service accounts, agents with no team assignment, and demoted/role-changed agents) were never loaded into the Users sheet because `loadUsers()` only walked `/v1.0/teams/{teamId}/members`. As a result, the `userMap` lookup in `buildActivityRow` returned blank for those users, leaving 1,612 of 13,900 ActivityLog rows with a `PerformedByUserId` but no `PerformedByUser` name.

### Added
- **Team-less agent augmentation** (`ReferenceData.gs`) — After the team-walk completes, `loadUsers()` now calls `augmentUsersWithAgents()`, which pulls `/v1.0/users/agents` and appends any agents missing from the Users sheet (catches admins, system service accounts, and agents with no team).
- **Activity-driven user backfill** (`ReferenceData.gs`) — New `backfillMissingUsersFromActivities()` scans ActivityLog for unique `PerformedByUserId` values not in the Users sheet and fetches each via `GET /v1.0/users/{id}`. Catches users whose role has since changed (e.g. demoted agents who logged historical labor).
- **In-place ActivityLog repair** (`ActivityLog.gs`) — New `repairBlankPerformedByUserNames()` reads ActivityLog cols M–R in a single batch, fills blank `PerformedByUser` names (and any blank Team cells) from the now-complete Users sheet, and writes back in one `setValues` call.
- **Troubleshooting → Backfill Missing User Names menu item** (`Menu.gs`) — Wires the three steps above into a single user-facing action with a summary dialog (agents added, historical users added, rows repaired, rows still missing).

## 2026-04-06

### Fixed
- **Missing labor from pre-school-year tickets** (`TicketData.gs`) — Initial ticket load only filtered by `createddate` within the school year, missing tickets created before the school year that had labor logged during it. Added `modifieddate` OR filter using IIQ API `GroupIndex` grouping so the initial load captures tickets created during the school year OR modified during the school year. Fixes 540+ missing labor hours across 17 users from 431 tickets. Refresh filters are unchanged.

## 2026-03-20

### Added
- **Post-load data validation** (`DataValidation.gs`) — New file with sheet-only validation that runs automatically after activities complete. Checks ticket count against a load-time snapshot (`TICKET_LOAD_EXPECTED_COUNT`) and performs per-ticket labor minute reconciliation comparing `RawData.TotalLaborMins` to summed `ActivityLog.EffortMins`. Classifies mismatches directionally: `missing_activities`, `under_reported` (likely missing some), `over_reported` (likely timing gap). Logs structured results with capped sample ticket IDs.
- **Automatic ticket reconciliation** (`DataValidation.gs`, `DataOrchestrator.gs`) — New `TICKET_RECONCILE` load type (Group 4) triggered when validation detects ticket count shortfall. Re-paginates with `upsertTickets()` (resumable, respects `MAX_RUNTIME_MS`), fetches activities for recovered tickets, revalidates afterward. Capped at 2 automatic attempts; errors out and clears expected count on persistent failure.
- **Ticket count snapshot and drift detection** (`TicketData.gs`) — Stores first-seen and last-observed `Paging.TotalRows` during paginated ticket load. Logs drift if TotalRows changes across pages, distinguishing "dataset moved during load" from "pagination lost a ticket."
- **Validate Data menu item** (`Menu.gs`) — New item under Troubleshooting for manual spot-check validation with UI dialog showing ticket count and labor minute results.

### Fixed
- **Activity batch JSON truncation** (`ApiClient.gs`, `ActivityLog.gs`, `Config.gs`) — Batch activity requests with 2000 ticket IDs produced responses too large for `UrlFetchApp`, causing truncated JSON ("Unterminated string at position 32517"). Added JSON parse error retry with backoff in `apiRequest()`. Introduced `ACTIVITY_BATCH_SIZE` config (default 100) separate from `PAGE_SIZE` to keep activity batch responses small. All activity batch callers (`loadActivitiesInitial`, `refreshActivitiesForTickets`, `processFailedActivityTickets`) now use `getActivityBatchSize()`.
- **USERS load blocked by single bad team** (`ReferenceData.gs`) — If one team's `/members` endpoint returned an error (e.g., "Address unavailable"), the entire USERS load failed and blocked Group 1, preventing Tickets and Activities from ever starting. Now wraps per-team fetch in try/catch, logs a warning with team name/ID, skips the bad team, and continues.

## 2026-03-19

### Added
- **IssueCategory + IssueType fields** (`TicketData.gs`, `ActivityLog.gs`, `Setup.gs`) — RawData expanded to 25 columns (V-Y: IssueCategoryId, IssueCategoryName, IssueTypeId, IssueTypeName). ActivityLog expanded to 24 columns (U-X) via `buildTicketContextMap` denormalization. Enables "what kind of work" analysis alongside existing who/where/how dimensions.
- **ByIssueCategory sheet** (`Setup.gs`) — Rollup: hours, cost, entry/ticket count per issue category. LET/BYROW pattern matching existing rollup sheets.
- **ByIssueType sheet** (`Setup.gs`) — Rollup: hours, cost, entry/ticket count per issue type.
- **AgentByCategory sheet** (`Setup.gs`) — Cross-dimension: agent × issue category. QUERY-based with date filtering.
- **TeamByCategory sheet** (`Setup.gs`) — Cross-dimension: team × issue category.
- **CategoryByLaborType sheet** (`Setup.gs`) — Cross-dimension: issue category × labor type.
- **LocationByCategory sheet** (`Setup.gs`) — Cross-dimension: location × issue category.
- **AgentPivot sheet** (`Setup.gs`) — Per-agent pivot table with one row per technician and hours broken down by labor type (Standard, Travel, Overtime, Weekend, Total). Column headers reference cell values so districts can rename labor types without editing the formula. User-selectable sort via dropdowns. Filtered by DateFilters.
- **ZeroLabor sheet** (`Setup.gs`) — Lists closed tickets with zero labor hours logged in the filtered date range. Helps managers flag "Work Complete" tickets where no time was entered.

### Changed
- **DateFilters expanded** (`Setup.gs`) — Replaced old filter modes (MTD, QTD, YTD, School YTD, Last 7 Days, Last 30 Days) with configurable set: This Month, Last Month, This Week, Last Week, This Quarter, Last Quarter, This Calendar Year, Last Calendar Year, This School Year, Last School Year, Manual. Default changed from "School YTD" to "This School Year". Week filters use Monday start (WEEKDAY type 2).
- **Data sheet headers now self-heal** (`Setup.gs`) — `setupRawDataSheet` and `setupActivityLogSheet` always overwrite header row 1, even when the sheet exists. Ensures new columns appear without needing a Full Reload.
- **LaborTypes removed from Regenerate Analytics** (`Setup.gs`) — LaborTypes is only created during Complete Setup; Regenerate Analytics no longer deletes and recreates it.
- **Instructions sheet updated** (`Setup.gs`) — Added all new sheets to sheet reference. Updated date filter documentation to reflect new filter mode names.
- **SHEET_ORDER** (`Setup.gs`) — All new sheets added in logical groupings.

### Fixed
- **Activity loader silent data loss** (`ActivityLog.gs`) — `loadActivitiesBatch` previously cleared failure records for ALL ticket IDs on HTTP 200, even if the API returned activities for only some of them. Added `reconcileBatchTicketFailures` that only clears failures for tickets actually present in the response, records failures for missing tickets so they retry, and logs warnings for tickets with zero qualifying activities. Skips missing-ticket failure checks on resumed partial pagination to avoid false positives.
- **Zero-effort activities dropped** (`ActivityLog.gs`) — `!item.Effort` filter incorrectly skipped entries where `Effort = 0`. Replaced with `isEffortMissing()` that only skips `null`/`undefined` effort, preserving zero-effort entries.
- **Blank ActivityId phantom duplicates** (`ActivityLog.gs`) — Added `normalizeActivityId()` validation in `buildActivityRow` (returns null for blank IDs) and `writeActivities` (skips blank IDs before dedup/index writes). Prevents accumulation of unindexed duplicate rows.
- **QUERY-based sheets showing duplicate headers** (`Setup.gs`) — Cross-dimension sheets (AgentByCategory, etc.) had QUERY `label` clause outputting a header row that duplicated the script-written headers. Fixed by using empty label strings.
- **IssueCategory/IssueType fields empty** (`TicketData.gs`) — Used wrong API field paths (`t.IssueCategory`, `t.IssueType`). Corrected to `t.Issue.IssueCategoryId`, `t.Issue.IssueCategoryName`, `t.Issue.IssueTypeId`, `t.Issue.Name`.
- **AgentPivot LET variable name collision** (`Setup.gs`) — Variables `c1`–`c4` are cell references in Sheets; LET rejected them. Renamed to `stdH`, `trvH`, `otH`, `wkndH`.
- **ZeroLabor dates showing serial numbers** (`Setup.gs`) — Wrapped date columns in `TEXT(date,"M/D/YYYY")` inside HSTACK.
- **reorderSheets crash after regenerate** (`Setup.gs`) — Added `SpreadsheetApp.flush()` before reorder loop to commit pending delete/create operations.

## 2026-02-23

### Fixed
- **Teams OpenTickets column empty** (`ReferenceData.gs`) — The IIQ `/v1.0/teams` endpoint does not return an open ticket count. `loadTeams()` now writes only the 3 API-sourced columns (TeamId, TeamName, MemberCount) and populates column D (OpenTickets) with a `COUNTIFS` formula that counts open tickets per team from RawData (`AssignedTeamId` match + `IsClosed="Open"`). Also changed `loadTeams()` to delete-and-recreate the Teams sheet (matching the analytics idempotency pattern) instead of requiring it to already exist, so Refresh Reference Data self-heals a missing Teams sheet.
- **YearSummary sheet completely blank** (`Setup.gs`) — Root cause: QUERY on a constructed array (`{TEXT(...), F2:F, H2:H, ...}`) misidentifies numeric columns as text, so `sum()` silently returns nothing under IFERROR. Rewrote to QUERY `ActivityLog!A2:T` directly (actual sheet range) so QUERY detects column types correctly. Uses QUERY's built-in `year()`/`month()` functions and `date` literal syntax for date filtering, with dates sourced from `DateFilters!$B$7/$B$8`. Month grouping post-processed to YYYY-MM via TEXT. HSTACK replaces ARRAYFORMULA for array construction. Each group still wrapped in IFERROR with placeholder rows filtered out after VSTACK.

## 2026-02-20

### Added
- **IndividualLookup sheet** (`Setup.gs`) — New analytics sheet with a dropdown to select an individual (dynamically populated from ActivityLog for the active date range). Shows summary KPIs (total hours, cost, entries, tickets), a detail list of all activity entries, and a per-ticket summary with aggregated hours/cost. Integrated into `setupLaborTrackerDashboard()`, `regenerateAnalyticsSheets()`, `SHEET_ORDER`, and Instructions sheet reference.
- **`loadLaborTypes()` via API** (`ReferenceData.gs`) — New function that fetches labor types from `GET /v1.0/labor/types` endpoint. Populates LaborTypes sheet with `LaborTypeId`, `LaborTypeName`, `IsOvertime`, and `OTMultiplier` columns.
- **LABOR_TYPES in load orchestrator** (`DataOrchestrator.gs`) — Added `LABOR_TYPES` to `DATA_LOAD_TYPES` and Group 1 of `LOAD_GROUPS`. Labor types now load from the API alongside Teams, Users, and ResolutionActions during initial load.
- **Refresh Labor Types menu item** (`Menu.gs`) — New menu item at iiQ Data > Load Data > Refresh Labor Types. Calls `loadLaborTypes()` (API) and shows the count of types loaded.

### Changed
- **LaborTypes sheet schema** (`Setup.gs`) — Headers expanded from `[LaborTypeId, LaborTypeName]` to `[LaborTypeId, LaborTypeName, IsOvertime, OTMultiplier]`.
- **IndividualLookup overtime auto-detection** (`Setup.gs`) — Replaced manual overtime labor type dropdown with automatic detection. OT Hours, OT Cost, and OT Entries now use SUMPRODUCT formulas that match activity LaborTypeIds against the `IsOvertime=TRUE` flag in LaborTypes. No user selection needed.
- **API call timing in logs** (`ApiClient.gs`) — All `API_REQUEST` log entries now include elapsed time in milliseconds (e.g. `GET /v1.0/teams -> 200 (342ms)`). Applies to success, retry, and error log lines.
- **`refreshLaborTypes()` is now offline fallback** (`TicketData.gs`) — Extracts labor types from ActivityLog/RawData when API is unavailable. Writes 4 columns (IsOvertime/OTMultiplier left blank). `loadLaborTypes()` is the primary source via API. Removed `refreshLaborTypesFromTickets()` alias.
- **`refreshReferenceData()`** (`ReferenceData.gs`) — Now calls `loadLaborTypes()` (API) instead of `refreshLaborTypes()` (extraction fallback).
- **`finishOpenRefresh()`** (`TicketData.gs`) — Now calls `loadLaborTypes()` (API) to refresh labor types after daily open refresh.
- **`menuRefreshLaborTypes()`** (`Menu.gs`) — Now calls `loadLaborTypes()` (API) instead of extraction fallback.
- Removed `refreshLaborTypes()` calls from ticket load completion in `loadTicketsPaginated()` (`TicketData.gs`) — labor types are now loaded in Group 1 before tickets start.

### Fixed
- **LaborTypes not populating** — Labor types are now loaded directly from the IIQ API (`GET /v1.0/labor/types`) as a first-class reference data type, rather than being extracted from ticket/activity data after the fact.
- **Pause log not written before timeout** (`TicketData.gs`) — `loadTicketsPaginated()` and `refreshOpenTickets()` had the time check as the `while` loop condition with the pause log written *after* loop exit. If the last iteration's API call + sheet write pushed past the 6-minute Apps Script limit, the log was never written. Restructured both loops to check time *inside* the loop at the top of each iteration, logging the pause and returning *before* starting more work.
- **IndividualLookup detail list showing UUIDs and raw dates** (`Setup.gs`) — Column B showed TicketId (UUID) instead of Subject because ActivityLog has no Subject column. Rewrote the detail formula using LET/MAP to look up Subject from RawData via MATCH on TicketId. Also wrapped ActivityDate in TEXT("M/D/YYYY") to fix serial number display.
- **LaborTypes sheet missing new column headers** (`Setup.gs`) — Changed `setupLaborTypesSheet()` from skip-if-exists to delete-and-recreate (idempotent) pattern. Existing sheets with old 2-column headers now get rebuilt with all 4 columns (`LaborTypeId`, `LaborTypeName`, `IsOvertime`, `OTMultiplier`). Added to `regenerateAnalyticsSheets()` so it rebuilds on that menu action too.
- **YearSummary sheet empty** (`Setup.gs`) — QUERY formulas used `count(distinct Col2)` which is not valid Google Sheets QUERY syntax. The entire LET chain errored silently via IFERROR. Replaced with `count(Col3)` (activity count). Dropped TicketCount column since QUERY cannot compute distinct counts. Headers reduced to 7 columns.
