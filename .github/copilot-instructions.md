# Surveillance Camera Access Request System

## Architecture Overview

Google Apps Script web application with a multi-tier approval workflow. Frontend HTML files communicate via `google.script.run` to a backend stored in `程式碼.js`. Data is persisted in Google Sheets with three main sheets:
- **申請人與攝影機資料** - Maps requesters to authorized cameras
- **申請紀錄** - Application log with workflow state (審核1中 → 審核2中 → 開通中 → 已開通)
- **審核開通** - Admin role definitions (columns B/D/F for Approver1/Approver2/Activator)
- **申請人/信箱** - User name/email mapping for display

## Workflow Engine

Three-stage approval process driven by status column (F):
1. **審核1中**: First approver reviews → updates to 審核2中
2. **審核2中**: Second approver reviews → updates to 開通中  
3. **開通中**: Activator grants access → updates to 已開通

Core function: `processApproval(rowNum)` - handles single-record state transitions and email notifications. Batch operations use `processBatchApproval(rowNumbers)` which wraps the single-record handler.

## Critical Patterns

### Column Index Constants
All sheet operations use 1-based constants like `LOG_STATUS_COLUMN_INDEX = 6`. When accessing array data (0-based), subtract 1: `row[LOG_STATUS_COLUMN_INDEX - 1]`.

### Data Formatting
`formatDataForFrontend(dataRows)` converts Date objects to Taiwan timezone strings. Always apply before sending data to HTML frontend to prevent timezone/serialization issues.

### User Lookup
`getUserNameByEmail(email)` queries the 申請人/信箱 sheet with 6-hour cache. Used to display human-readable names instead of emails.

### Permission Checks
- `isUserAnAdmin(email)` - central permission validator (checks all three admin lists)
- Always validate permissions in backend functions exposed to frontend
- `doGet(e)` routes to `unauthorized.html` for non-admins accessing review page

### Cache Strategy
Admin email lists cached for 10 minutes via `getAdminListFromSheet(column)`. User name map cached for 6 hours. Clear caches if roles/users don't update immediately.

## Frontend-Backend Communication

HTML files use scriptlets `<?!= ... ?>` for server-side templating (e.g., injecting Web App URL). All backend calls are async:

```javascript
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errorHandler)
  .backendFunction(params);
```

Key exposed functions: `getRequesterData()`, `processForm(formData)`, `getTasksForCurrentUser()`, `getMyApplications()`.

## Key Files

- **程式碼.js**: All backend logic, no code splitting
- **表單.html**: Application form with dynamic camera dropdown based on requester selection
- **review.html**: Admin review dashboard with batch approval checkboxes
- **myapply.html**: User's personal application history (filtered by email + name match)
- **index.html**: Portal landing page with Bootstrap cards linking to three main functions

## Google Sheets Menu

`onOpen()` creates custom menu "🖥️ 監視器調閱系統" with shortcuts that open HTML pages in modal dialogs via `HtmlService.createTemplateFromFile()`.

## Email Notifications

`sendNotificationEmail(recipients, subject, body, options)` sends HTML emails with clickable links. Default links to review page; pass `options.linkPage` to customize (e.g., 'myapply' for applicant notifications).

## Deployment

Web app config in `appsscript.json`:
- **executeAs**: USER_DEPLOYING (runs as deployer's account)
- **access**: DOMAIN (only organization members can access)
- Timezone: Asia/Taipei

Deploy as web app to get URL, which is dynamically injected into HTML pages via `ScriptApp.getService().getUrl()`.

## Common Edits

- **Add workflow stage**: Update status constants, modify `processApproval()` switch cases, adjust email templates
- **Change columns**: Update `LOG_*_COLUMN_INDEX` constants and verify `formatDataForFrontend()` cases
- **Add admin role**: Create new `getAdminListFromSheet()` wrapper, update permission checks
- **Modify frontend filters**: Edit `desiredColumnIndexes` arrays in `getTasksForCurrentUser()` or `getMyApplications()`
