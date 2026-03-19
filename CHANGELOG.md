# Changelog

## 2026-03-19

### Added
- **AgentPivot sheet** (`Setup.gs`) — Per-agent pivot table with one row per technician and hours broken down by labor type (Standard, Travel, Overtime, Weekend, Total). Column headers reference cell values so districts can rename labor types without editing the formula. Filtered by DateFilters.
- **ZeroLabor sheet** (`Setup.gs`) — Lists closed tickets with zero labor hours logged in the filtered date range. Helps managers flag "Work Complete" tickets where no time was entered.

### Changed
- **DateFilters expanded** (`Setup.gs`) — Replaced old filter modes (MTD, QTD, YTD, School YTD, Last 7 Days, Last 30 Days) with DeKalb-requested set: This Month, Last Month, This Week, Last Week, This Quarter, Last Quarter, This Calendar Year, Last Calendar Year, This School Year, Last School Year, Manual. Default changed from "School YTD" to "This School Year". Week filters use Monday start (WEEKDAY type 2).
- **Instructions sheet updated** (`Setup.gs`) — Added AgentPivot and ZeroLabor to sheet reference. Updated date filter documentation to reflect new filter mode names.
- **Regenerate Analytics** (`Setup.gs`) — `regenerateAnalyticsSheets()` now includes AgentPivot and ZeroLabor. Confirmation dialog lists all sheets.
- **SHEET_ORDER** (`Setup.gs`) — AgentPivot and ZeroLabor added after IndividualLookup.

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
