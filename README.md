# 📋 Job Application Tracker — Google Apps Script

Automatically tracks job applications by scanning a dedicated Gmail label and populating a Google Sheet.

## Features

| Feature | Details |
|---|---|
| **Gmail Integration** | Scans only the `Job-Applications` label — no full inbox scan |
| **Incremental Scanning** | Remembers last scan time via `PropertiesService` |
| **Duplicate Prevention** | Uses Gmail thread IDs to avoid duplicate rows |
| **Smart Status Detection** | Keywords → `Interview`, `Offer`, `Rejected`, or `Applied` |
| **Auto-Timestamping** | Editing the Status column auto-updates Last Updated |
| **Time-Driven Trigger** | Configurable auto-scan every 6–12 hours |
| **Scan Log** | All scan activity logged to a dedicated "Scan Log" sheet |
| **Safe Updates** | Company & Role columns are never overwritten by the script |

---

## Sheet Columns

| Column | Name | Auto / Manual |
|--------|------|---------------|
| A | Company | Manual |
| B | Role | Manual |
| C | Applied Date | Auto |
| D | Status | Auto (editable) |
| E | Email Subject | Auto |
| F | Email Link | Auto |
| G | Last Updated | Auto |
| H | Follow-up | Manual |
| I | Notes | Manual |
| J | Interview Date | Manual |
| K | Thread ID | Auto (hidden) |

---

## Setup Instructions

### 1. Create the Gmail Label

In Gmail, create a label called **`Job-Applications`** and apply it to relevant emails (manually or with a Gmail filter).

**Recommended Gmail filter:**
```
Matches: from:(noreply@*.com OR careers@* OR recruiting@* OR talent@*)
Do this: Apply label "Job-Applications"
```

### 2. Create the Google Sheet

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet.
2. Name it something like **"Job Application Tracker"**.

### 3. Add the Script

1. In your spreadsheet, go to **Extensions → Apps Script**.
2. Delete any boilerplate code in the editor.
3. Copy the entire contents of [`Code.gs`](Code.gs) and paste it into the editor.
4. Click **💾 Save** (or `Ctrl+S`).

### 4. Initialize

1. Close and re-open the spreadsheet (or reload the page).
2. A new menu **"Job Tracker"** will appear in the menu bar.
3. Click **Job Tracker → Setup Sheet** to create headers and formatting.
4. Click **Job Tracker → Scan Emails** to run your first scan.
5. *(Optional)* Click **Job Tracker → Install Auto-Scan Trigger** for automatic scanning.

### 5. Authorize Permissions

On first run, Google will ask you to authorize the script. It needs:
- `Gmail (read-only)` — to read emails under the label
- `Spreadsheet` — to write data
- `Script properties` — to store the last scan timestamp
- `Triggers` — to install time-driven auto-scan

> [!NOTE]
> You may see a "This app isn't verified" warning. Click **Advanced → Go to (project name)** to proceed.

---

## Status Detection Keywords

| Status | Trigger Keywords |
|--------|-----------------|
| **Rejected** | `unfortunately`, `regret to inform`, `not moving forward`, `rejected`, `position has been filled`, etc. |
| **Offer** | `offer letter`, `pleased to offer`, `job offer`, `compensation package`, etc. |
| **Interview** | `interview`, `schedule a call`, `phone screen`, `technical assessment`, `coding challenge`, etc. |
| **Applied** | Default — when no other keywords match |

> Rejection is checked first to correctly handle emails that mention "interview" in a rejection context.

---

## Customization

| Setting | Location | Default |
|---------|----------|---------|
| Gmail label name | `GMAIL_LABEL` constant | `Job-Applications` |
| Trigger frequency | `TRIGGER_HOURS` constant | `6` hours |
| Sheet name | `SHEET_NAME` constant | `Applications` |

To adjust, edit the constants at the top of `Code.gs`.

---

## Architecture

```
onOpen()              → Adds "Job Tracker" menu
setupSheet()          → Creates sheet + headers + formatting
scanEmails()          → Main entry point for scanning
  ├─ fetchThreadsSince_()  → Paginated, incremental Gmail fetch
  ├─ buildThreadMap_()     → Thread ID → row lookup (batch read)
  └─ processThread_()     → Upsert logic per thread
       └─ detectStatus_() → Keyword-based status classification
onEdit()              → Auto-timestamp on Status column change
installTrigger()      → Sets up time-driven auto-scan
logMessage_()         → Writes to "Scan Log" sheet
```

---

## License

MIT — use freely.
