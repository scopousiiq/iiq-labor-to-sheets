# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script solution that pulls labor hours and resolution action data from IncidentIQ's API into Google Sheets, providing rollup reports by team, individual, department, and labor/resolution type. The source-of-truth specification is `2026-02-04-labor-tracking-google-sheets-spec.md`.

## Development Environment

This is a **Google Apps Script** project. All `.gs` files in `scripts/` are deployed to the Apps Script editor attached to a Google Sheet. There is no local build system, package manager, or test framework.

- **Testing:** Run functions directly from the Google Apps Script editor (e.g., `testApiConnection()`, `startInitialLoad()`, `executeNextLoad()`)
- **Deployment:** Copy `.gs` file contents into the Apps Script editor bound to the target spreadsheet
- **Style:** 2-space indentation for `.gs` and `.html` files

## Architecture

### Script Files (`scripts/`)

| File | Role |
|------|------|
| `Config.gs` | Configuration read/write, type coercion helpers, config caching, LockService concurrency, `requireNoTriggers` guard, school year locking, date utilities, operation logging |
| `ApiClient.gs` | HTTP client with exponential backoff retry (max 5 retries), paginated fetch via `fetchAllPagesWithQueryParams()` |
| `DataOrchestrator.gs` | Load state machine: groups loads into 4 phases (reference data → tickets → activities → reconciliation), tracks per-type state (idle/pending/in_progress/complete/error) |
| `DataValidation.gs` | Post-load validation (ticket count snapshot check, per-ticket labor minute reconciliation) and auto-recovery via resumable ticket reconciliation |
| `TicketData.gs` | Paginated ticket loading with 5.5-minute runtime guard (`MAX_RUNTIME_MS`), TotalRows snapshot/drift tracking, open ticket refresh with upsert, filter builders for IIQ API |
| `ActivityLog.gs` | Activity loading per-ticket with batch processing, failure tracking/retry via `ActivityFailures` sheet, upsert via index maps |
| `ReferenceData.gs` | Loads teams, users (by iterating team members), and resolution actions |
| `Index.gs` | Hidden `TicketIndex`/`ActivityIndex` sheets for O(1) row lookups enabling upsert without full-sheet scans |
| `Setup.gs` | Creates all sheets with headers, formulas (LET/BYROW/QUERY-based rollups), data validation, and protections |
| `Triggers.gs` | Time-driven triggers: 10-minute monitor for resuming loads, daily 2 AM open refresh |
| `Menu.gs` | `onOpen()` menu: iiQ Data → Setup, Load Data, Troubleshooting submenus |

### Data Flow

1. **Load groups execute sequentially:** Group 1 (Teams, Users, ResolutionActions) → Group 2 (Tickets) → Group 3 (Activities) → Group 4 (Ticket Reconcile, only if validation detects shortfall)
2. **Resumable pagination:** Long loads pause at `MAX_RUNTIME_MS` (5.5 min) and resume via `triggerDataLoadMonitor` trigger or manual "Continue Loading"
3. **Upsert pattern:** Index sheets map entity IDs to row numbers. New records append; existing records update in-place via `writeBatchedUpdates()`
4. **School year locking:** Once data loading begins, `SCHOOL_YEAR_START`/`SCHOOL_YEAR_END`/`PAGE_SIZE`/`MODULE` are locked in Config and cell-protected. A full reload is required to change them.

### Google Sheets Structure

- **Config** — key-value settings (API_BASE_URL, BEARER_TOKEN, SITE_ID, school year dates, load state tracking)
- **DateFilters** — SWITCH-based formulas for date range selection (This/Last Month, This/Last Week, This/Last Quarter, This/Last Calendar Year, This/Last School Year, Manual)
- **RawData** (21 cols) — ticket data with labor fields
- **ActivityLog** (20 cols) — resolution action time entries with denormalized team/location
- **Teams, Users, ResolutionActions** — reference lookup tables
- **LaborTypes** (4 cols: LaborTypeId, LaborTypeName, IsOvertime, OTMultiplier) — loaded from `GET /v1.0/labor/types` API; `IsOvertime` flag drives automatic overtime detection in IndividualLookup
- **ByTeam, ByIndividual, ByDepartment, ByLaborType, ByResolution** — formula-driven rollup sheets using LET/BYROW/SUMIFS against ActivityLog, filtered by DateFilters. `ByDepartment` groups on LocationName (ActivityLog col T) — it is the per-location rollup; the tab and its column header are kept as-is because district dashboards reference both by name.
- **IndividualLookup** — dropdown to select an individual, plus a second dropdown to narrow to one issue type (listed as `Category > Type`; blank means every type); shows their activity detail and per-ticket summary for the filtered date range
- **LocationLookup** — dropdown to select a location, plus a second dropdown to narrow to one issue type (listed as `Category > Type`, the same pairing IndividualLookup uses; blank means every type); shows that location's totals and overtime, and three side-by-side blocks: agents who worked there, hours by `Category > Type`, and per-ticket detail. Each block filters ActivityLog to the location once and aggregates over that array rather than rescanning the full columns per output row.
- **ByIssueCategory** — hours, cost, entry/ticket count per issue category
- **ByIssueType** — hours, cost, entry/ticket count per issue type
- **AgentPivot** (7 cols: Agent, Team, Standard, Travel, Overtime, Weekend, Total Hours) — per-agent pivot with hours by labor type; column headers reference cell values so labor type names are adjustable
- **ZeroLabor** (8 cols) — closed tickets with zero labor hours in the filtered date range; flags tickets where no time was logged
- **AgentByCategory** — QUERY-based cross-dimension: agent × issue category (hours, cost, entries)
- **TeamByCategory** — QUERY-based cross-dimension: team × issue category
- **CategoryByLaborType** — QUERY-based cross-dimension: issue category × labor type
- **LocationByCategory** — QUERY-based cross-dimension: location × issue category
- **YearSummary** — QUERY-based monthly aggregation by group type (Team, Agent, LaborType, Resolution)
- **Dashboard** — KPI formulas referencing rollup sheets
- **TicketIndex, ActivityIndex** — hidden index sheets for fast lookups
- **ActivityFailures** — hidden sheet tracking failed activity loads for retry
- **Logs** — operation log (auto-trimmed to 1000 rows)

### IncidentIQ API

- Auth: Bearer token + SiteId header + ProductId header + `Client: ApiClient`
- Base URL is stored without `/api` suffix; `Config.gs:normalizeBaseUrl()` appends it
- Pagination: `$p` (zero-based page), `$s` (page size, default 2000)
- Ticket search: `POST /v1.0/tickets` with filter facets (`totallabortime`, `createddate`, `isclosed`, `closeddate`, `modifieddate`)
- Batch activities: `POST /v1.0/tickets/activities` with array of ticket IDs as payload
- Rate limiting: configurable `THROTTLE_MS` (default 1000ms) between requests, plus exponential backoff on 429/5xx

## Key Patterns

- **Config as key-value store:** `getConfig(key)` / `setConfig(key, value)` scan the Config sheet. All values stored as strings. Defaults are in `CONFIG_DEFAULTS`. Required keys validated via `CONFIG_REQUIRED`.
- **Type coercion:** `getStringValue()`, `getIntValue()`, `getBoolValue()` handle Sheets' unpredictable return types safely.
- **Config caching:** `cacheConfigRowPositions()` + `writeConfigValueDirect()` / `getConfigValueDirect()` avoid full-sheet scans in tight loading loops.
- **LockService concurrency:** `acquireScriptLock()` for menu items (wait briefly, show busy). `tryAcquireScriptLock()` for triggers (skip if busy). All entry points wrapped.
- **Destructive op safety:** `requireNoTriggers()` gates `startFullReloadWithConfirm()` to prevent trigger interference during data clearing.
- **Load state machine:** Each data type has `LOAD_STATE_<TYPE>` in Config. The orchestrator finds the next pending load respecting group ordering.
- **Batched sheet writes:** Both ticket and activity upserts collect updates into `{rowNumber: rowData}` maps and write consecutive rows in single `setValues()` calls via shared `writeBatchedUpdates()`.
- **Header constants:** `RAWDATA_HEADERS` (25 cols) and `ACTIVITY_HEADERS` (26 cols) in `Setup.gs` are the single source of truth for column counts.
- **Analytics idempotency:** Analytics sheets use `deleteSheetIfExists()` + recreate pattern. Data sheets use skip-if-exists. "Regenerate Analytics Sheets" menu item rebuilds all formula sheets.
- **Date formatting for API:** `formatDateForApi()` produces `M/D/YYYY` format required by IIQ filter facets.
- **BI-safe values:** `IsClosed` stored as `'Closed'`/`'Open'` (not boolean). `IsPublic` stored as `1`/`0` (not boolean).
- **Logging:** `logOperation()` inserts newest entries at row 2 (after header), auto-prunes beyond 1000 entries.

## RawData Column Layout (25 columns)

| Col | Letter | Header | Source |
|-----|--------|--------|--------|
| 1 | A | TicketId | `t.TicketId` |
| 2 | B | TicketNumber | `t.TicketNumber` |
| 3 | C | Subject | `t.Subject` |
| 4 | D | CreatedDate | `t.CreatedDate` |
| 5 | E | ClosedDate | `t.ClosedDate` |
| 6 | F | TotalLaborMins | `t.TotalLaborTime` |
| 7 | G | TotalLaborHours | `TotalLaborTime / 60` |
| 8 | H | TotalLaborCost | `t.TotalLaborCost` |
| 9 | I | LaborTypeId | `t.LaborType.LaborTypeId` |
| 10 | J | LaborTypeName | `t.LaborType.Name` |
| 11 | K | AssignedUser | `t.AssignedToUser.Name` |
| 12 | L | AssignedUserEmail | `t.AssignedToUser.Email` |
| 13 | M | AssignedTeam | `t.AssignedToTeam.TeamName` |
| 14 | N | Location | `t.Location.Name` |
| 15 | O | Requester | `t.For.Name` |
| 16 | P | Status | `t.WorkflowStep.StatusName` |
| 17 | Q | ResolutionAction | `t.ResolutionAction` |
| 18 | R | IsClosed | `'Closed'` or `'Open'` |
| 19 | S | AssignedUserId | `t.AssignedToUserId` |
| 20 | T | AssignedTeamId | `t.AssignedToTeamId` |
| 21 | U | LocationId | `t.LocationId` |
| 22 | V | IssueCategoryId | `t.IssueCategory.IssueCategoryId` |
| 23 | W | IssueCategoryName | `t.IssueCategory.Name` |
| 24 | X | IssueTypeId | `t.IssueType.IssueTypeId` |
| 25 | Y | IssueTypeName | `t.IssueType.Name` |

## ActivityLog Column Layout (26 columns)

| Col | Letter | Header | Source |
|-----|--------|--------|--------|
| 1 | A | ActivityId | `item.TicketActivityActionId` |
| 2 | B | TicketId | from batch context |
| 3 | C | TicketNumber | from `ticketMap` |
| 4 | D | ActivityDate | `item.ActivityDate` |
| 5 | E | EffortMins | `item.Effort` |
| 6 | F | EffortHours | `Effort / 60` |
| 7 | G | HourlyRate | `item.HourlyRate` |
| 8 | H | LaborCost | `item.LaborCost` |
| 9 | I | LaborTypeId | `item.LaborTypeId` |
| 10 | J | LaborTypeName | `item.LaborTypeName` / lookup |
| 11 | K | ResolutionActionId | `item.ResolutionActionId` |
| 12 | L | ResolutionActionName | `item.ResolutionAction` / lookup |
| 13 | M | PerformedByUserId | `item.ByUserId` |
| 14 | N | PerformedByUser | `userMap` lookup |
| 15 | O | Notes | `item.Notes` |
| 16 | P | IsPublic | `1` or `0` |
| 17 | Q | TeamId | `userMap` lookup |
| 18 | R | TeamName | `userMap` lookup |
| 19 | S | LocationId | from `ticketMap` |
| 20 | T | LocationName | from `ticketMap` |
| 21 | U | IssueCategoryId | from `ticketMap` |
| 22 | V | IssueCategoryName | from `ticketMap` |
| 23 | W | IssueTypeId | from `ticketMap` |
| 24 | X | IssueTypeName | from `ticketMap` |
| 25 | Y | EntryType | `computeNetHours()` — `'Labor'` or `'Action'` |
| 26 | Z | NetHours | `computeNetHours()` — de-duplicated effort |

`NetHours` is the hours column every rollup should sum. When a ticket carries at
least one labor row, its resolution-action rows are set to 0 so the same effort
is not counted twice; a ticket with no labor rows keeps its action-row effort.
`EffortHours` (col F) is only correct where the formula already restricts itself
to labor rows — a labor type name or a LaborTypeId criterion does that, a
resolution action name does not.

Columns U–X are ticket-level fields denormalized onto every activity row, so all
activity rows for one ticket share the same issue category and type.

## Config Key Reference

### Required (user must provide)

| Key | Example | Notes |
|-----|---------|-------|
| `API_BASE_URL` | `https://district.incidentiq.com` | No `/api` suffix |
| `BEARER_TOKEN` | JWT string | Auth token |
| `SITE_ID` | UUID | Site identifier |
| `SCHOOL_YEAR_START` | Date | School year start |
| `SCHOOL_YEAR_END` | Date | School year end |
| `MODULE` | `Ticketing` | `Ticketing` or `Facilities` — selects IIQ module |

### Optional (have defaults)

| Key | Default | Notes |
|-----|---------|-------|
| `PAGE_SIZE` | `'2000'` | Records per API call (string) |
| `THROTTLE_MS` | `'1000'` | Delay between calls (string) |
| `OPEN_REFRESH_DAYS` | `'14'` | Days of closed tickets to refresh (string) |

### Auto-managed (set by scripts)

| Key | Purpose |
|-----|---------|
| `LOAD_STATE_*` | Per-type load states (idle/pending/in_progress/complete/error) |
| `TICKET_LOAD_PAGE` | Current ticket pagination page |
| `TICKET_LOAD_TOTAL_PAGES` | Total ticket pages |
| `ACTIVITY_TICKET_INDEX` | Current activity batch index |
| `ACTIVITY_BATCH_PAGE` | Current activity page within batch |
| `ACTIVITY_LAST_TICKET_ID` | Last processed ticket for activities |
| `SCHOOL_YEAR_LOCKED` | `TRUE`/`FALSE` lock status |
| `SCHOOL_YEAR_LOCKED_AT` | Lock timestamp |
| `SCHOOL_YEAR_LOCKED_START` | Locked start date |
| `SCHOOL_YEAR_LOCKED_END` | Locked end date |
| `PAGE_SIZE_LOCKED` | Locked page size |
| `MODULE_LOCKED` | Locked module selection |
| `LAST_SYNC` | Last successful sync timestamp |
| `OPEN_REFRESH_STAGE` | `OPEN`/`CLOSED` stage during refresh |
| `OPEN_REFRESH_OPEN_PAGE` | Current page for open ticket refresh |
| `OPEN_REFRESH_CLOSED_PAGE` | Current page for closed ticket refresh |
| `TICKET_LOAD_EXPECTED_COUNT` | Last-observed `Paging.TotalRows` during ticket load (validation baseline) |
| `TICKET_LOAD_FIRST_TOTAL_ROWS` | First-observed `Paging.TotalRows` (drift detection baseline) |
| `TICKET_LOAD_TOTAL_ROWS_DRIFT` | Delta between first and last TotalRows if dataset changed during load |
| `TICKET_RECONCILE_PAGE` | Current reconciliation pagination page |
| `TICKET_RECONCILE_ATTEMPTS` | Number of reconciliation passes attempted (max 2) |

## Post-Load Validation & Reconciliation

After activities complete, `validateLoadedData()` runs automatically:

1. **Ticket count check** — Compares `TICKET_LOAD_EXPECTED_COUNT` (snapshot from ticket load) against unique non-blank TicketIds in TicketIndex, cross-checked against RawData. No API calls.
2. **Labor minute reconciliation** — Sums `ActivityLog.EffortMins` per ticket and compares to `RawData.TotalLaborMins`. Classifies directionally: `missing_activities` (sum=0, ticket>0), `under_reported` (sum<ticket), `over_reported` (sum>ticket, likely timing gap). No API calls.
3. **Reconcile trigger** — If ticket count shortfall detected, sets `LOAD_STATE_TICKET_RECONCILE` to pending. Activity mismatches are log-only (no auto-recovery in v1).

`TICKET_RECONCILE` is Group 4 in the load state machine, picked up by the 10-minute monitor trigger. It re-paginates with `upsertTickets()`, fetches activities for any recovered tickets, revalidates, and caps at 2 automatic attempts before error-state-and-stop.

Available manually via: iiQ Data → Troubleshooting → Validate Data.

## Formula Patterns

Analytics sheets use **ActivityLog** as the data source with **DateFilters** for date range control:
- `UNIQUE(FILTER(...))` to get distinct entities
- `BYROW(entities, LAMBDA(e, SUMIFS/COUNTIFS(...)))` for per-entity metrics
- `HSTACK(...)` to combine columns, `SORT(...)` to order results
- All formulas reference **Name columns** (not ID columns) for `UNIQUE`/`COUNTIFS`
- Key column references in formulas: `R` = TeamName, `N` = PerformedByUser, `T` = LocationName, `J` = LaborTypeName, `L` = ResolutionActionName, `V` = IssueCategoryName, `X` = IssueTypeName, `F` = EffortHours, `H` = LaborCost, `D` = ActivityDate, `B` = TicketId, `Y` = EntryType, `Z` = NetHours
- Sum `Z` (NetHours) for hours, not `F` — see the ActivityLog column layout below
- Bound the end of a date range with `< endD+1` (SUMIFS: `"<"&(endD+1)`; QUERY: `Col4<date` of `endD+1`), never `<= endD`. `DateFilters!B6` is a bare date at midnight while `ActivityDate` carries a time, so `<=` silently drops the whole last day
- To slice a multi-column LET variable use `CHOOSECOLS(raw,4,5,6)`. `INDEX(raw,,{4,5,6})` returns only the first of those columns, with no error to signal it

## Changelog

After every code change (new features, bug fixes, refactors), update `CHANGELOG.md` in the project root. Follow the format already established in that file:
- Group entries under a date heading (`## YYYY-MM-DD`)
- Use `### Added`, `### Changed`, `### Fixed` sub-headings as appropriate
- Each entry is a concise bullet describing the change and which file(s) were affected
- Add to the existing date section if multiple changes happen on the same day; otherwise create a new date heading at the top (newest first)
