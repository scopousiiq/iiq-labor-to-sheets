# IncidentIQ Labor Tracking & Rollup Reports - Google Sheets Solution

> **Purpose**: Specification for building a Google Sheets + Apps Script solution to export labor hours and resolution action data from IncidentIQ, providing rollup reports by team, individual, department, and labor/resolution type with flexible date range filters.
>
> **Created**: 2026-02-04
> **For**: District time tracking reporting needs

---

## Executive Summary

Build a Google Sheets + Apps Script solution that pulls labor hours and resolution action data from IncidentIQ's API, providing rollup reports by team, individual, department, and labor/resolution type with flexible date range filters.

---

## Problem Statement

The district wants to:
1. Track what everyone on their team is actually working on
2. Log time for both ticket work AND non-ticket activities
3. Generate rollup reports by team/individual/department
4. Filter by various date ranges (week, month, quarter, school year)

**Current Limitations in IncidentIQ:**
- Labor time is tied to tickets (no standalone time entries)
- No built-in rollup reports by team/department
- No pre-built date range aggregations

**Solution:** Extract data via API and aggregate in Google Sheets with pivot tables and custom reports.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    Google Sheets Workbook                       │
├─────────────────────────────────────────────────────────────────┤
│  📋 Config          │ API credentials, date ranges, settings    │
│  📊 Raw Data        │ Ticket data with labor fields (API pull)  │
│  📊 Activity Log    │ Resolution actions with effort/time       │
│  📊 Teams           │ Team roster lookup table                  │
│  📊 Users           │ User → Team/Dept mapping                  │
│  📊 Labor Types     │ Labor type lookup table                   │
│  📈 By Team         │ Pivot: Hours by Team                      │
│  📈 By Individual   │ Pivot: Hours by Person                    │
│  📈 By Department   │ Pivot: Hours by Location/Dept             │
│  📈 By Labor Type   │ Pivot: Hours by Labor Category            │
│  📈 By Resolution   │ Pivot: Hours by Resolution Action         │
│  🎛️ Dashboard       │ Summary KPIs + charts                     │
└─────────────────────────────────────────────────────────────────┘
                              ▲
                              │ Apps Script
                              │
┌─────────────────────────────────────────────────────────────────┐
│                    IncidentIQ API                               │
├─────────────────────────────────────────────────────────────────┤
│  POST /api/v1.0/tickets        → Search tickets with labor      │
│  GET  /api/v1.0/tickets/{id}/timeline → Activity details        │
│  GET  /api/v1.0/teams          → Team list                      │
│  GET  /api/v1.0/teams/{id}/members → Team membership            │
│  GET  /api/v1.0/resolutions/actions → Resolution action types   │
└─────────────────────────────────────────────────────────────────┘
```

---

## API Reference

### Authentication

All requests require these headers:

```
Authorization: Bearer <JWT_TOKEN>
SiteId: <your-site-uuid>
Client: ApiClient
Content-Type: application/json
```

### Key Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/api/v1.0/tickets` | POST | Search tickets with labor filters |
| `/api/v1.0/tickets/{ticketId}/timeline` | GET | Get activity log with labor entries |
| `/api/v1.0/teams` | GET | List all teams |
| `/api/v1.0/teams/{teamId}/members` | GET | Get team members |
| `/api/v1.0/resolutions/actions` | GET | List resolution action types |

### Pagination

All list endpoints support:
- `$p` - Page index (zero-based)
- `$s` - Page size (default: 100, max: 100)

Response includes `Paging` object:
```json
{
  "Paging": {
    "TotalRows": 1500,
    "PageCount": 15,
    "PageSize": 100,
    "PageIndex": 0
  }
}
```

### Ticket Search with Labor Filters

```json
POST /api/v1.0/tickets?$p=0&$s=50

{
  "ProductId": "88df910c-91aa-e711-80c2-0004ffa00010",
  "Schema": "OpenWithModify",
  "Filters": [
    {
      "Facet": "totallabortime",
      "Value": "numoperator:greaterthan:0",
      "Negative": false,
      "GroupIndex": 0
    },
    {
      "Facet": "createddate",
      "Value": "date>=01/01/2026",
      "Negative": false,
      "GroupIndex": 1
    },
    {
      "Facet": "createddate",
      "Value": "date<=01/31/2026",
      "Negative": false,
      "GroupIndex": 1
    }
  ],
  "FilterByProduct": true,
  "IncludeDeleted": false
}
```

### Available Filter Facets for Labor

| Facet | Type | Description | Example Value |
|-------|------|-------------|---------------|
| `labortype` | UUID | Filter by labor type | `{labor-type-uuid}` |
| `totallabortime` | Numeric | Total minutes | `numoperator:greaterthan:60` |
| `totallaborcost` | Numeric | Total cost | `numoperator:greaterthan:100` |
| `agent` | UUID | Assigned technician | `{user-uuid}` |
| `team` | UUID | Assigned team | `{team-uuid}` |
| `createddate` | Date | Ticket creation date | `date>=01/01/2026` |
| `modifieddate` | Date | Last modified | `date<=01/31/2026` |

### Numeric Operator Syntax

```
numoperator:<operator>:<value>
```

| Operator | Meaning |
|----------|---------|
| `equals` | Exactly equal |
| `lessthan` | Less than |
| `lessthanequal` | Less than or equal |
| `greaterthan` | Greater than |
| `greaterthanequal` | Greater than or equal |

### Ticket Response - Labor Fields

```json
{
  "TicketId": "uuid",
  "TicketNumber": "12345",
  "Subject": "Ticket subject",
  "CreatedDate": "2026-01-15T10:30:00Z",
  "ClosedDate": "2026-01-16T14:00:00Z",
  "TotalLaborTime": 120,
  "TotalLaborCost": 60.00,
  "LaborType": {
    "LaborTypeId": "uuid",
    "Name": "On-Site Support"
  },
  "AssignedToUserId": "uuid",
  "AssignedToUser": {
    "UserId": "uuid",
    "Name": "John Smith",
    "Email": "john.smith@district.edu"
  },
  "AssignedToTeamId": "uuid",
  "AssignedToTeam": {
    "TeamId": "uuid",
    "TeamName": "IT Support"
  },
  "Location": {
    "LocationId": "uuid",
    "Name": "Main Campus"
  },
  "ResolutionAction": "Password Reset"
}
```

### Timeline Activity Entry with Labor

```json
{
  "TicketActivityLogEntry": {
    "TicketActivityId": "uuid",
    "CreatedDate": "2026-01-15T11:00:00Z",
    "ActivityItems": [
      {
        "$type": "Spark.Shared.Models.TicketActivityAction, Spark.Shared",
        "TicketActivityActionId": "uuid",
        "TicketActivityTypeId": 8,
        "ResolutionActionId": "uuid",
        "ResolutionAction": "First Contact Resolution",
        "ActivityDate": "2026-01-15T11:00:00Z",
        "ByUserId": "uuid",
        "Effort": 30,
        "HourlyRate": 50.00,
        "LaborCost": 25.00,
        "LaborTypeId": "uuid",
        "Notes": "Initial diagnosis completed",
        "IsPublic": true
      }
    ]
  }
}
```

---

## Google Sheets Structure

### Sheet 1: `Config`
Settings and credentials for the solution.

| Row | A (Setting) | B (Value) | Notes |
|-----|-------------|-----------|-------|
| 1 | **Setting** | **Value** | |
| 2 | API Base URL | `https://yoursite.incidentiq.com` | |
| 3 | Site ID | `{uuid}` | From IIQ admin |
| 4 | API Token | `{jwt}` | Stored securely (see security section) |
| 5 | Product ID | `88df910c-91aa-e711-80c2-0004ffa00010` | Default: Help Desk |
| 6 | Last Sync | `2026-02-04T10:30:00Z` | Auto-updated |
| 7 | Sync Status | `Success` | |
| 8 | --- | --- | |
| 9 | **Date Filters** | | |
| 10 | Filter Type | `This Month` | Dropdown |
| 11 | Start Date | `2026-02-01` | Auto-calculated |
| 12 | End Date | `2026-02-28` | Auto-calculated |

**Date Filter Dropdown Options:**
- This Week
- Last Week
- This Month
- Last Month
- This Quarter
- Last Quarter
- This School Year (Aug 1 - Jul 31)
- Last School Year
- This Calendar Year
- Last Calendar Year
- Custom Range

---

### Sheet 2: `RawData` (Ticket Data)
Raw ticket export with labor fields.

| Column | Header | API Source | Type |
|--------|--------|------------|------|
| A | TicketId | `TicketId` | UUID |
| B | TicketNumber | `TicketNumber` | Text |
| C | Subject | `Subject` | Text |
| D | CreatedDate | `CreatedDate` | DateTime |
| E | ClosedDate | `ClosedDate` | DateTime |
| F | TotalLaborMins | `TotalLaborTime` | Number |
| G | TotalLaborHours | `=F/60` | Formula |
| H | TotalLaborCost | `TotalLaborCost` | Currency |
| I | LaborType | `LaborType.Name` | Text |
| J | AssignedUser | `AssignedToUser.Name` | Text |
| K | AssignedUserEmail | `AssignedToUser.Email` | Email |
| L | AssignedTeam | `AssignedToTeam.TeamName` | Text |
| M | Location | `Location.Name` | Text |
| N | Requester | `For.Name` | Text |
| O | Status | `WorkflowStep.StatusName` | Text |
| P | ResolutionAction | `ResolutionAction` | Text |
| Q | IsClosed | `IsClosed` | Boolean |
| R | AssignedUserId | `AssignedToUserId` | UUID (hidden) |
| S | AssignedTeamId | `AssignedToTeamId` | UUID (hidden) |
| T | LocationId | `LocationId` | UUID (hidden) |

---

### Sheet 3: `ActivityLog` (Resolution Action Time Entries)
Detailed time entries from ticket timelines.

| Column | Header | API Source | Type |
|--------|--------|------------|------|
| A | ActivityId | `TicketActivityActionId` | UUID |
| B | TicketId | `TicketId` | UUID |
| C | TicketNumber | (lookup from RawData) | Text |
| D | ActivityDate | `ActivityDate` | DateTime |
| E | EffortMins | `Effort` | Number |
| F | EffortHours | `=E/60` | Formula |
| G | HourlyRate | `HourlyRate` | Currency |
| H | LaborCost | `LaborCost` | Currency |
| I | LaborTypeId | `LaborTypeId` | UUID |
| J | LaborTypeName | (lookup from LaborTypes) | Text |
| K | ResolutionActionId | `ResolutionActionId` | UUID |
| L | ResolutionActionName | `ResolutionAction` | Text |
| M | PerformedByUserId | `ByUserId` | UUID |
| N | PerformedByUser | (lookup from Users) | Text |
| O | Notes | `Notes` | Text |
| P | IsPublic | `IsPublic` | Boolean |

---

### Sheet 4: `Teams`
Team reference data.

| Column | Header | API Source |
|--------|--------|------------|
| A | TeamId | `TeamId` |
| B | TeamName | `TeamName` |
| C | MemberCount | `MembersCount` |
| D | OpenTickets | `Tickets` |

---

### Sheet 5: `Users`
User reference data with team membership.

| Column | Header | API Source |
|--------|--------|------------|
| A | UserId | `UserId` |
| B | Name | `Name` |
| C | Email | `Email` |
| D | Location | `LocationName` |
| E | Role | `Role.Name` |
| F | TeamId | (from team membership lookup) |
| G | TeamName | (from team membership lookup) |

---

### Sheet 6: `LaborTypes`
Labor type reference data.

| Column | Header | Source |
|--------|--------|--------|
| A | LaborTypeId | Extracted from ticket data |
| B | LaborTypeName | Extracted from ticket data |

---

### Sheet 7: `ResolutionActions`
Resolution action reference data.

| Column | Header | API Source |
|--------|--------|------------|
| A | ActionId | `ResolutionActionId` |
| B | ActionName | `Name` |
| C | Category | `ActionCategoryName` |
| D | Scope | `Scope` |

---

### Sheets 8-12: Pivot Report Sheets

#### `ByTeam` - Hours by Team
| Row Labels | Sum of Hours | Sum of Cost | Ticket Count | Avg Hours/Ticket |
|------------|--------------|-------------|--------------|------------------|
| IT Support | 245.5 | $12,275 | 156 | 1.57 |
| Network Ops | 189.0 | $9,450 | 89 | 2.12 |
| **Total** | **434.5** | **$21,725** | **245** | **1.77** |

#### `ByIndividual` - Hours by Person
| Row Labels | Team | Sum of Hours | Sum of Cost | Ticket Count |
|------------|------|--------------|-------------|--------------|
| John Smith | IT Support | 85.5 | $4,275 | 52 |
| Jane Doe | IT Support | 67.0 | $3,350 | 41 |

#### `ByDepartment` - Hours by Location/Department
| Location | Sum of Hours | Sum of Cost | Unique Users |
|----------|--------------|-------------|--------------|
| Main Campus | 312.0 | $15,600 | 8 |
| East Building | 122.5 | $6,125 | 4 |

#### `ByLaborType` - Hours by Labor Category
| Labor Type | Sum of Hours | Sum of Cost | Usage Count |
|------------|--------------|-------------|-------------|
| On-Site Support | 189.5 | $9,475 | 127 |
| Remote Support | 156.0 | $7,800 | 203 |
| Training | 45.0 | $2,250 | 12 |

#### `ByResolution` - Hours by Resolution Action
| Resolution Action | Category | Sum of Hours | Avg Time | Count |
|-------------------|----------|--------------|----------|-------|
| Password Reset | Account Mgmt | 23.5 | 0.25 | 94 |
| Hardware Repair | Repairs | 156.0 | 2.5 | 62 |

---

### Sheet 13: `Dashboard`
Summary view with KPIs and charts.

**KPI Cards:**
- Total Hours (period)
- Total Cost (period)
- Tickets Completed
- Avg Hours/Ticket
- Top Performer
- Busiest Team

**Charts:**
- Hours by Team (bar chart)
- Hours Trend (line chart, by week)
- Labor Type Distribution (pie chart)
- Top 10 Resolution Actions (horizontal bar)

---

## Google Apps Script Architecture

### File Structure

```
Code.gs              - Main entry point, menu creation
ApiClient.gs         - IncidentIQ API wrapper functions
DataSync.gs          - Data extraction and sheet population
DateUtils.gs         - Date range calculations
Triggers.gs          - Scheduled refresh setup
UI.gs                - Sidebar and dialog components
DatePicker.html      - Date range selection dialog
Settings.html        - Settings sidebar
```

---

### Code.gs - Main Entry Point

```javascript
// Menu creation
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 IIQ Labor Reports')
    .addItem('🔄 Sync All Data', 'syncAllData')
    .addItem('📅 Change Date Range', 'showDateRangePicker')
    .addSeparator()
    .addItem('🔄 Sync Tickets Only', 'syncTickets')
    .addItem('🔄 Sync Activities Only', 'syncActivities')
    .addItem('🔄 Sync Reference Data', 'syncReferenceData')
    .addSeparator()
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('📋 View Sync Log', 'showSyncLog')
    .addToUi();
}

function syncAllData() {
  syncReferenceData();
  syncTickets();
  syncActivities();
  refreshPivots();
  updateLastSync();
}
```

---

### ApiClient.gs - API Wrapper

```javascript
// API configuration
function getApiConfig() {
  const config = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Config');
  return {
    baseUrl: config.getRange('B2').getValue(),
    siteId: config.getRange('B3').getValue(),
    token: config.getRange('B4').getValue(),
    productId: config.getRange('B5').getValue()
  };
}

// Base API request
function apiRequest(endpoint, method, payload) {
  const config = getApiConfig();
  const options = {
    method: method || 'GET',
    headers: {
      'Authorization': 'Bearer ' + config.token,
      'SiteId': config.siteId,
      'Client': 'ApiClient',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(config.baseUrl + endpoint, options);
  const code = response.getResponseCode();

  if (code === 429) {
    // Rate limited - wait and retry
    const retryAfter = JSON.parse(response.getContentText()).RetryAfter || 60;
    Utilities.sleep(retryAfter * 1000);
    return apiRequest(endpoint, method, payload);
  }

  if (code !== 200) {
    throw new Error('API Error: ' + code + ' - ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

// Paginated fetch (gets ALL pages)
function fetchAllPages(endpoint, method, payload, pageSize) {
  const size = pageSize || 100;
  let allItems = [];
  let page = 0;
  let totalPages = 1;

  do {
    const url = endpoint + (endpoint.includes('?') ? '&' : '?') +
                '$p=' + page + '&$s=' + size;
    const response = apiRequest(url, method, payload);

    if (response.Items) {
      allItems = allItems.concat(response.Items);
    }

    if (response.Paging) {
      totalPages = response.Paging.PageCount;
    }

    page++;

    // Rate limiting protection
    Utilities.sleep(100);

  } while (page < totalPages);

  return allItems;
}

// Search tickets with labor data
function searchTicketsWithLabor(startDate, endDate) {
  const config = getApiConfig();
  const payload = {
    ProductId: config.productId,
    Schema: 'OpenWithModify',
    Filters: [
      {
        Facet: 'totallabortime',
        Value: 'numoperator:greaterthan:0',
        Negative: false,
        GroupIndex: 0
      },
      {
        Facet: 'createddate',
        Value: 'date>=' + formatDateForApi(startDate),
        Negative: false,
        GroupIndex: 1
      },
      {
        Facet: 'createddate',
        Value: 'date<=' + formatDateForApi(endDate),
        Negative: false,
        GroupIndex: 1
      }
    ],
    FilterByProduct: true,
    IncludeDeleted: false
  };

  return fetchAllPages('/api/v1.0/tickets', 'POST', payload, 50);
}

// Get ticket timeline (activities with labor)
function getTicketTimeline(ticketId) {
  return fetchAllPages('/api/v1.0/tickets/' + ticketId + '/timeline', 'GET', null, 100);
}

// Get all teams
function getTeams() {
  return fetchAllPages('/api/v1.0/teams', 'GET', null, 100);
}

// Get team members
function getTeamMembers(teamId) {
  return fetchAllPages('/api/v1.0/teams/' + teamId + '/members', 'GET', null, 100);
}

// Get resolution actions
function getResolutionActions() {
  return fetchAllPages('/api/v1.0/resolutions/actions', 'GET', null, 100);
}
```

---

### DataSync.gs - Data Extraction

```javascript
// Sync ticket data to RawData sheet
function syncTickets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RawData');
  const config = ss.getSheetByName('Config');

  const startDate = config.getRange('B11').getValue();
  const endDate = config.getRange('B12').getValue();

  // Clear existing data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }

  // Fetch tickets with labor
  const tickets = searchTicketsWithLabor(startDate, endDate);

  if (tickets.length === 0) {
    SpreadsheetApp.getUi().alert('No tickets found for the selected date range.');
    return;
  }

  // Map to rows
  const rows = tickets.map(t => [
    t.TicketId,
    t.TicketNumber,
    t.Subject,
    t.CreatedDate,
    t.ClosedDate,
    t.TotalLaborTime || 0,
    null, // Formula column
    t.TotalLaborCost || 0,
    t.LaborType ? t.LaborType.Name : '',
    t.AssignedToUser ? t.AssignedToUser.Name : '',
    t.AssignedToUser ? t.AssignedToUser.Email : '',
    t.AssignedToTeam ? t.AssignedToTeam.TeamName : '',
    t.Location ? t.Location.Name : '',
    t.For ? t.For.Name : '',
    t.WorkflowStep ? t.WorkflowStep.StatusName : '',
    t.ResolutionAction || '',
    t.IsClosed,
    t.AssignedToUserId,
    t.AssignedToTeamId,
    t.LocationId
  ]);

  // Write data
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // Add formulas for TotalLaborHours (column G)
  const formulas = rows.map((_, i) => ['=F' + (i + 2) + '/60']);
  sheet.getRange(2, 7, rows.length, 1).setFormulas(formulas);

  Logger.log('Synced ' + tickets.length + ' tickets');
}

// Sync activity/timeline data
function syncActivities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RawData');
  const actSheet = ss.getSheetByName('ActivityLog');

  // Get ticket IDs from RawData
  const ticketIds = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(id => id);

  // Clear existing activities
  const lastRow = actSheet.getLastRow();
  if (lastRow > 1) {
    actSheet.getRange(2, 1, lastRow - 1, actSheet.getLastColumn()).clear();
  }

  // Fetch timelines for each ticket
  let allActivities = [];

  for (const ticketId of ticketIds) {
    const timeline = getTicketTimeline(ticketId);

    for (const entry of timeline) {
      if (!entry.ActivityItems) continue;

      for (const item of entry.ActivityItems) {
        // Only include resolution actions (type 8) with effort
        if (item.$type && item.$type.includes('TicketActivityAction') && item.Effort) {
          allActivities.push([
            item.TicketActivityActionId,
            ticketId,
            '', // Ticket number (lookup)
            item.ActivityDate,
            item.Effort,
            null, // EffortHours formula
            item.HourlyRate || 0,
            item.LaborCost || 0,
            item.LaborTypeId || '',
            '', // LaborTypeName (lookup)
            item.ResolutionActionId || '',
            item.ResolutionAction || '',
            item.ByUserId,
            '', // PerformedByUser (lookup)
            item.Notes || '',
            item.IsPublic
          ]);
        }
      }
    }

    // Rate limiting
    Utilities.sleep(50);
  }

  if (allActivities.length > 0) {
    actSheet.getRange(2, 1, allActivities.length, allActivities[0].length)
      .setValues(allActivities);

    // Add EffortHours formulas
    const formulas = allActivities.map((_, i) => ['=E' + (i + 2) + '/60']);
    actSheet.getRange(2, 6, allActivities.length, 1).setFormulas(formulas);
  }

  Logger.log('Synced ' + allActivities.length + ' activities');
}

// Sync reference data (teams, users, etc.)
function syncReferenceData() {
  syncTeams();
  syncUsers();
  syncResolutionActions();
}

function syncTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Teams');

  const teams = getTeams();

  // Clear and write
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clear();
  }

  const rows = teams.map(t => [
    t.TeamId,
    t.TeamName,
    t.MembersCount,
    t.Tickets
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

function syncUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName('Teams');
  const usersSheet = ss.getSheetByName('Users');

  // Get team IDs
  const teamIds = teamsSheet.getRange(2, 1, teamsSheet.getLastRow() - 1, 2)
    .getValues()
    .filter(row => row[0]);

  const teamMap = {};
  teamIds.forEach(row => teamMap[row[0]] = row[1]);

  // Fetch members for each team
  let allUsers = [];

  for (const [teamId, teamName] of Object.entries(teamMap)) {
    const members = getTeamMembers(teamId);

    for (const member of members) {
      allUsers.push([
        member.UserId,
        member.UserName || member.Name,
        member.Email || '',
        member.LocationName || '',
        member.Role || '',
        teamId,
        teamName
      ]);
    }
  }

  // Clear and write
  const lastRow = usersSheet.getLastRow();
  if (lastRow > 1) {
    usersSheet.getRange(2, 1, lastRow - 1, 7).clear();
  }

  if (allUsers.length > 0) {
    usersSheet.getRange(2, 1, allUsers.length, 7).setValues(allUsers);
  }
}

function syncResolutionActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ResolutionActions');

  const actions = getResolutionActions();

  // Clear and write
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clear();
  }

  const rows = actions.map(a => [
    a.ResolutionActionId,
    a.Name,
    a.ActionCategoryName,
    a.Scope
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}
```

---

### DateUtils.gs - Date Range Calculations

```javascript
// Calculate date ranges based on filter type
function calculateDateRange(filterType) {
  const now = new Date();
  let startDate, endDate;

  switch (filterType) {
    case 'This Week':
      startDate = getStartOfWeek(now);
      endDate = now;
      break;

    case 'Last Week':
      const lastWeek = new Date(now);
      lastWeek.setDate(lastWeek.getDate() - 7);
      startDate = getStartOfWeek(lastWeek);
      endDate = getEndOfWeek(lastWeek);
      break;

    case 'This Month':
      startDate = new Date(now.getFullYear(), now.getMonth(), 1);
      endDate = now;
      break;

    case 'Last Month':
      startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      endDate = new Date(now.getFullYear(), now.getMonth(), 0);
      break;

    case 'This Quarter':
      const quarter = Math.floor(now.getMonth() / 3);
      startDate = new Date(now.getFullYear(), quarter * 3, 1);
      endDate = now;
      break;

    case 'Last Quarter':
      const lastQ = Math.floor(now.getMonth() / 3) - 1;
      const year = lastQ < 0 ? now.getFullYear() - 1 : now.getFullYear();
      const q = lastQ < 0 ? 3 : lastQ;
      startDate = new Date(year, q * 3, 1);
      endDate = new Date(year, q * 3 + 3, 0);
      break;

    case 'This School Year':
      // School year: Aug 1 - Jul 31
      if (now.getMonth() >= 7) { // Aug-Dec
        startDate = new Date(now.getFullYear(), 7, 1);
      } else { // Jan-Jul
        startDate = new Date(now.getFullYear() - 1, 7, 1);
      }
      endDate = now;
      break;

    case 'Last School Year':
      if (now.getMonth() >= 7) {
        startDate = new Date(now.getFullYear() - 1, 7, 1);
        endDate = new Date(now.getFullYear(), 6, 31);
      } else {
        startDate = new Date(now.getFullYear() - 2, 7, 1);
        endDate = new Date(now.getFullYear() - 1, 6, 31);
      }
      break;

    case 'This Calendar Year':
      startDate = new Date(now.getFullYear(), 0, 1);
      endDate = now;
      break;

    case 'Last Calendar Year':
      startDate = new Date(now.getFullYear() - 1, 0, 1);
      endDate = new Date(now.getFullYear() - 1, 11, 31);
      break;

    default:
      // Custom - use values from Config sheet
      return null;
  }

  return { startDate, endDate };
}

function getStartOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getEndOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() + (6 - day));
  d.setHours(23, 59, 59, 999);
  return d;
}

function formatDateForApi(date) {
  const d = new Date(date);
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}

// Update Config sheet when filter type changes
function onFilterChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('Config');
  const filterType = config.getRange('B10').getValue();

  const range = calculateDateRange(filterType);
  if (range) {
    config.getRange('B11').setValue(range.startDate);
    config.getRange('B12').setValue(range.endDate);
  }
}
```

---

### Triggers.gs - Scheduled Sync

```javascript
// Set up scheduled sync (daily at 6 AM)
function setupDailySync() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduledSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger
  ScriptApp.newTrigger('scheduledSync')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Daily sync scheduled for 6:00 AM');
}

function scheduledSync() {
  try {
    syncAllData();
    Logger.log('Scheduled sync completed successfully');
  } catch (e) {
    Logger.log('Scheduled sync failed: ' + e.message);
    // Optionally send email notification
    // MailApp.sendEmail(adminEmail, 'IIQ Sync Failed', e.message);
  }
}

// Update last sync timestamp
function updateLastSync() {
  const config = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Config');
  config.getRange('B6').setValue(new Date().toISOString());
  config.getRange('B7').setValue('Success');
}
```

---

### UI.gs - User Interface

```javascript
// Show date range picker dialog
function showDateRangePicker() {
  const html = HtmlService.createHtmlOutputFromFile('DatePicker')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Date Range');
}

// Show settings sidebar
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setTitle('IIQ Labor Reports Settings')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Process date range selection from dialog
function setDateRange(filterType, customStart, customEnd) {
  const config = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Config');

  config.getRange('B10').setValue(filterType);

  if (filterType === 'Custom Range') {
    config.getRange('B11').setValue(new Date(customStart));
    config.getRange('B12').setValue(new Date(customEnd));
  } else {
    onFilterChange();
  }

  return 'Date range updated to: ' + filterType;
}
```

---

### DatePicker.html - Date Range Dialog

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; }
    .form-group { margin-bottom: 15px; }
    label { display: block; margin-bottom: 5px; font-weight: bold; }
    select, input { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; }
    .custom-dates { display: none; }
    .custom-dates.visible { display: block; }
    button { background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px; }
    button:hover { background: #3367d6; }
  </style>
</head>
<body>
  <div class="form-group">
    <label>Date Range</label>
    <select id="filterType" onchange="toggleCustomDates()">
      <option value="This Week">This Week</option>
      <option value="Last Week">Last Week</option>
      <option value="This Month" selected>This Month</option>
      <option value="Last Month">Last Month</option>
      <option value="This Quarter">This Quarter</option>
      <option value="Last Quarter">Last Quarter</option>
      <option value="This School Year">This School Year</option>
      <option value="Last School Year">Last School Year</option>
      <option value="This Calendar Year">This Calendar Year</option>
      <option value="Last Calendar Year">Last Calendar Year</option>
      <option value="Custom Range">Custom Range</option>
    </select>
  </div>

  <div class="custom-dates" id="customDates">
    <div class="form-group">
      <label>Start Date</label>
      <input type="date" id="startDate">
    </div>
    <div class="form-group">
      <label>End Date</label>
      <input type="date" id="endDate">
    </div>
  </div>

  <button onclick="apply()">Apply & Sync</button>

  <script>
    function toggleCustomDates() {
      const filterType = document.getElementById('filterType').value;
      const customDates = document.getElementById('customDates');
      customDates.className = filterType === 'Custom Range' ? 'custom-dates visible' : 'custom-dates';
    }

    function apply() {
      const filterType = document.getElementById('filterType').value;
      const startDate = document.getElementById('startDate').value;
      const endDate = document.getElementById('endDate').value;

      google.script.run
        .withSuccessHandler(() => {
          google.script.run.syncAllData();
          google.script.host.close();
        })
        .setDateRange(filterType, startDate, endDate);
    }
  </script>
</body>
</html>
```

---

## Security Considerations

### API Token Storage

**Option A: Script Properties (Recommended for simplicity)**
```javascript
// Store token securely
PropertiesService.getScriptProperties().setProperty('IIQ_TOKEN', token);

// Retrieve token
const token = PropertiesService.getScriptProperties().getProperty('IIQ_TOKEN');
```

**Option B: Config Sheet with Protection**
- Store in Config sheet (as shown in spec)
- Protect the Config sheet
- Hide sensitive rows

### Access Control

1. **Restrict sheet sharing** - Only share with authorized personnel
2. **Hide Config sheet** - Right-click > Hide sheet
3. **Protect sensitive ranges** - Data > Protected sheets and ranges
4. **Use service account** - For automated syncs without user credentials

---

## Implementation Phases

### Phase 1: Foundation (Day 1-2)
- [ ] Create Google Sheet with all sheets defined
- [ ] Implement `ApiClient.gs` with auth and pagination
- [ ] Test API connectivity with simple endpoints
- [ ] Implement `syncTeams()` and `syncResolutionActions()`

### Phase 2: Core Data (Day 3-4)
- [ ] Implement `syncTickets()` with labor filters
- [ ] Implement `syncActivities()` for timeline data
- [ ] Add date range calculation utilities
- [ ] Create Config sheet with data validation

### Phase 3: Reports (Day 5-6)
- [ ] Create pivot tables for each report type
- [ ] Build Dashboard sheet with KPIs
- [ ] Add charts (bar, line, pie)
- [ ] Test with real data

### Phase 4: Polish (Day 7)
- [ ] Add custom menu
- [ ] Create HTML dialogs for date picker
- [ ] Set up scheduled triggers
- [ ] Document usage instructions
- [ ] Train district staff

---

## Addressing the "Everything Else" Problem

For logging time on non-ticket work, the district has two options:

### Option A: Create Administrative Ticket Types (Recommended)
Create ticket issue types for non-ticket work:
- [ADMIN] Meetings
- [ADMIN] Training
- [ADMIN] Documentation
- [ADMIN] Professional Development
- [ADMIN] Administrative Tasks
- [ADMIN] Travel Time

**Pros:** Uses existing system, data flows through same reports
**Cons:** Inflates ticket counts, may feel bureaucratic

### Option B: Separate Time Sheet
Create a parallel `ManualTimeLog` sheet for non-ticket time:

| Date | User | Category | Hours | Notes |
|------|------|----------|-------|-------|
| 2026-02-04 | John Smith | Meeting | 1.5 | Staff meeting |
| 2026-02-04 | Jane Doe | Training | 2.0 | Security cert |

**Pros:** Clean separation, no fake tickets
**Cons:** Two data sources to manage, manual entry required

---

## Verification Steps

1. **API Connectivity Test**
   - Run `testApiConnection()` from Script Editor
   - Verify 200 response with valid data

2. **Data Sync Test**
   - Run `syncAllData()` manually
   - Verify row counts in each sheet
   - Spot-check data accuracy against IIQ UI

3. **Date Range Test**
   - Change filter type in Config
   - Verify dates calculate correctly
   - Re-sync and verify filtered results

4. **Report Accuracy Test**
   - Sum hours in RawData manually
   - Compare to pivot table totals
   - Cross-reference with IIQ reports

5. **Scheduled Sync Test**
   - Set up trigger
   - Wait for execution
   - Check sync log and data freshness

---

## Files Summary

| File | Type | Purpose |
|------|------|---------|
| `Code.gs` | Script | Main entry point |
| `ApiClient.gs` | Script | API wrapper |
| `DataSync.gs` | Script | Data extraction |
| `DateUtils.gs` | Script | Date calculations |
| `Triggers.gs` | Script | Scheduled jobs |
| `UI.gs` | Script | UI components |
| `DatePicker.html` | HTML | Date range dialog |
| `Settings.html` | HTML | Settings sidebar |

---

## Success Criteria

- [ ] All data syncs without errors
- [ ] Reports show accurate totals matching IIQ
- [ ] Date filters work correctly for all options
- [ ] Scheduled sync runs reliably
- [ ] District can generate reports independently
