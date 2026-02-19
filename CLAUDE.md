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
| `DataOrchestrator.gs` | Load state machine: groups loads into 3 phases (reference data → tickets → activities), tracks per-type state (idle/pending/in_progress/complete/error) |
| `TicketData.gs` | Paginated ticket loading with 5.5-minute runtime guard (`MAX_RUNTIME_MS`), open ticket refresh with upsert, filter builders for IIQ API |
| `ActivityLog.gs` | Activity loading per-ticket with batch processing, failure tracking/retry via `ActivityFailures` sheet, upsert via index maps |
| `ReferenceData.gs` | Loads teams, users (by iterating team members), and resolution actions |
| `Index.gs` | Hidden `TicketIndex`/`ActivityIndex` sheets for O(1) row lookups enabling upsert without full-sheet scans |
| `Setup.gs` | Creates all sheets with headers, formulas (LET/BYROW/QUERY-based rollups), data validation, and protections |
| `Triggers.gs` | Time-driven triggers: 10-minute monitor for resuming loads, daily 2 AM open refresh |
| `Menu.gs` | `onOpen()` menu: iiQ Data → Setup, Load Data, Troubleshooting submenus |

### Data Flow

1. **Load groups execute sequentially:** Group 1 (Teams, Users, ResolutionActions) → Group 2 (Tickets) → Group 3 (Activities)
2. **Resumable pagination:** Long loads pause at `MAX_RUNTIME_MS` (5.5 min) and resume via `triggerDataLoadMonitor` trigger or manual "Continue Loading"
3. **Upsert pattern:** Index sheets map entity IDs to row numbers. New records append; existing records update in-place via `writeBatchedUpdates()`
4. **School year locking:** Once data loading begins, `SCHOOL_YEAR_START`/`SCHOOL_YEAR_END`/`PAGE_SIZE` are locked in Config and cell-protected. A full reload is required to change them.

### Google Sheets Structure

- **Config** — key-value settings (API_BASE_URL, BEARER_TOKEN, SITE_ID, school year dates, load state tracking)
- **DateFilters** — SWITCH-based formulas for date range selection (MTD, QTD, YTD, School YTD, etc.)
- **RawData** (21 cols) — ticket data with labor fields
- **ActivityLog** (20 cols) — resolution action time entries with denormalized team/location
- **Teams, Users, LaborTypes, ResolutionActions** — reference lookup tables
- **ByTeam, ByIndividual, ByDepartment, ByLaborType, ByResolution** — formula-driven rollup sheets using LET/BYROW/SUMIFS against ActivityLog, filtered by DateFilters
- **YearSummary** — QUERY-based monthly aggregation by group type (Team, Agent, LaborType, Resolution)
- **Dashboard** — KPI formulas referencing rollup sheets
- **TicketIndex, ActivityIndex** — hidden index sheets for fast lookups
- **ActivityFailures** — hidden sheet tracking failed activity loads for retry
- **Logs** — operation log (auto-trimmed to 1000 rows)

### IncidentIQ API

- Auth: Bearer token + SiteId header + `Client: ApiClient`
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
- **Header constants:** `RAWDATA_HEADERS` (21 cols) and `ACTIVITY_HEADERS` (20 cols) in `Setup.gs` are the single source of truth for column counts.
- **Analytics idempotency:** Analytics sheets use `deleteSheetIfExists()` + recreate pattern. Data sheets use skip-if-exists. "Regenerate Analytics Sheets" menu item rebuilds all formula sheets.
- **Date formatting for API:** `formatDateForApi()` produces `M/D/YYYY` format required by IIQ filter facets.
- **BI-safe values:** `IsClosed` stored as `'Closed'`/`'Open'` (not boolean). `IsPublic` stored as `1`/`0` (not boolean).
- **Logging:** `logOperation()` inserts newest entries at row 2 (after header), auto-prunes beyond 1000 entries.

## RawData Column Layout (21 columns)

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

## ActivityLog Column Layout (20 columns)

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

## Config Key Reference

### Required (user must provide)

| Key | Example | Notes |
|-----|---------|-------|
| `API_BASE_URL` | `https://district.incidentiq.com` | No `/api` suffix |
| `BEARER_TOKEN` | JWT string | Auth token |
| `SITE_ID` | UUID | Site identifier |
| `SCHOOL_YEAR_START` | Date | School year start |
| `SCHOOL_YEAR_END` | Date | School year end |

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
| `LAST_SYNC` | Last successful sync timestamp |
| `OPEN_REFRESH_STAGE` | `OPEN`/`CLOSED` stage during refresh |
| `OPEN_REFRESH_OPEN_PAGE` | Current page for open ticket refresh |
| `OPEN_REFRESH_CLOSED_PAGE` | Current page for closed ticket refresh |

## Formula Patterns

Analytics sheets use **ActivityLog** as the data source with **DateFilters** for date range control:
- `UNIQUE(FILTER(...))` to get distinct entities
- `BYROW(entities, LAMBDA(e, SUMIFS/COUNTIFS(...)))` for per-entity metrics
- `HSTACK(...)` to combine columns, `SORT(...)` to order results
- All formulas reference **Name columns** (not ID columns) for `UNIQUE`/`COUNTIFS`
- Key column references in formulas: `R` = TeamName, `N` = PerformedByUser, `T` = LocationName, `J` = LaborTypeName, `L` = ResolutionActionName, `F` = EffortHours, `H` = LaborCost, `D` = ActivityDate, `B` = TicketId
