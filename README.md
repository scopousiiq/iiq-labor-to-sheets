# iiQ Labor to Sheets

Extract labor hours and resolution action data from Incident IQ into Google Sheets, providing rollup reports by team, individual, department, labor type, and resolution action. Built for K-12 school district IT departments.

## Quick Start

### 1. Create and Set Up Your Spreadsheet

**Option A: Copy the Template (Fastest)**

[**Make a copy of the Google Sheets template**](https://docs.google.com/spreadsheets/d/16wMrY3Dg4wjL0asl6CpnaZ70KI523cmgA3J4Rbmv6Gg/edit?usp=sharing)

This template includes all sheets and formulas pre-configured. Skip to step 2.

**Option B: Build from Scratch**

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` folder in this repository
4. Save and reload the spreadsheet
5. Click **iiQ Data > Setup > Run Complete Setup**

### 2. Configure API Access

In the `Config` sheet, enter your Incident IQ credentials:

| Setting | Value | Where to Find It |
|---------|-------|------------------|
| `API_BASE_URL` | `https://yourdistrict.incidentiq.com` | Your iiQ URL (the `/api` is added automatically) |
| `BEARER_TOKEN` | Your API token | iiQ Admin > Developer Tools |
| `SITE_ID` | Your site UUID | Only needed for multi-site districts |
| `MODULE` | `Ticketing` or `Facilities` | Dropdown in Config sheet (defaults to Ticketing) |
| `SCHOOL_YEAR_START` | School year start date | e.g., 2025-06-01 |
| `SCHOOL_YEAR_END` | School year end date | e.g., 2026-05-31 |

Then click **iiQ Data > Setup > Test API Connection** to verify credentials.

### 3. Load Your Data

1. Click **iiQ Data > Load Data > Start Initial Load** to begin importing
2. The script runs for ~5.5 minutes then pauses automatically (Apps Script limit)
3. Click **iiQ Data > Load Data > Continue Loading** to resume, or set up triggers (step 4)
4. Data loads in three phases: Reference Data (teams, users) > Tickets > Activities

> **Tip for large districts:** Set up automated triggers so you don't have to manually click "Continue Loading" repeatedly.

### 4. Set Up Automated Triggers (Recommended)

Click **iiQ Data > Setup > Setup Automated Triggers** to create all triggers automatically:

| Trigger | Schedule | What It Does |
|---------|----------|--------------|
| Data Load Monitor | Every 10 minutes | Resumes any paused loads automatically |
| Daily Open Refresh | Daily at 2 AM | Refreshes open and recently closed tickets |

> Triggers skip gracefully if another operation is running — no conflicts.

## What You Get

### Data Sheets

| Sheet | What It Shows |
|-------|---------------|
| `RawData` | All tickets with labor fields (25 columns, includes issue category/type) |
| `ActivityLog` | All time entries with labor type, cost, team, and issue info (24 columns) |
| `Teams` | Team directory |
| `Users` | User directory with team assignments |
| `LaborTypes` | Labor type names with overtime flags |
| `ResolutionActions` | Resolution action categories |

### Analytics Sheets (formula-driven, update automatically)

| Sheet | What It Shows |
|-------|---------------|
| `ByTeam` | Hours, cost, and ticket count per team |
| `ByIndividual` | Hours, cost, and ticket count per person (with team) |
| `ByDepartment` | Hours, cost, and ticket count per location (the tab name predates the field it groups on) |
| `ByLaborType` | Hours, cost, and ticket count per labor type |
| `ByResolution` | Hours, cost, and ticket count per resolution action |
| `ByIssueCategory` | Hours, cost, and ticket count per issue category |
| `ByIssueType` | Hours, cost, and ticket count per issue type |
| `AgentPivot` | Per-technician hours broken down by labor type — verify 40-hour work weeks |
| `ZeroLabor` | Closed tickets with zero labor logged — flag for manager review |
| `AgentByCategory` | Cross-dimension: agent hours broken down by issue category |
| `TeamByCategory` | Cross-dimension: team hours broken down by issue category |
| `CategoryByLaborType` | Cross-dimension: issue category hours broken down by labor type |
| `LocationByCategory` | Cross-dimension: location hours broken down by issue category |
| `IndividualLookup` | Select a person to see their activity detail and per-ticket summary |
| `LocationLookup` | Select a location to see its agents, issue types, and tickets — filterable to one `Category > Type` |
| `YearSummary` | Monthly aggregation by team, agent, labor type, and resolution |
| `Dashboard` | KPI summary (total hours, cost, top performer, top team) |

### Date Filters

All analytics sheets filter by the `DateFilters` sheet. Change the dropdown to switch ranges:

| Filter | Date Range |
|--------|------------|
| This Month / Last Month | Current or previous calendar month |
| This Week / Last Week | Current or previous week (Monday start) |
| This Quarter / Last Quarter | Current or previous calendar quarter |
| This Calendar Year / Last Calendar Year | Current or previous calendar year |
| This School Year / Last School Year | Based on configured school year dates |
| Manual | Custom start/end dates |

## Common Issues

| Problem | Solution |
|---------|----------|
| "API configuration missing" | Fill in `API_BASE_URL` and `BEARER_TOKEN` in the Config sheet |
| HTTP 401 error | Your Bearer token expired — get a new one from iiQ |
| Script timeout | Normal for large districts — run "Continue Loading" or let the trigger resume it |
| Analytics sheets empty | ActivityLog must have data — ensure all 3 load phases completed |
| "Another operation is running" | Wait a few minutes — locks auto-expire after 6 minutes |
| Need to change school year | Remove triggers first, then Full Reload to unlock and clear data |
| AgentPivot shows 0 for a labor type | Rename the column headers (C2:F2) to match your district's labor type names |

## Documentation

- [**CLAUDE.md**](CLAUDE.md) — Technical reference for developers (architecture, column layouts, formula patterns)

## Disclaimer

This is an independent utility, not a supported product feature. No SLA or maintenance commitment is implied. Use as-is and customize to fit your district's needs.

## License

MIT — Free to use and modify for your district. See [LICENSE](LICENSE).
