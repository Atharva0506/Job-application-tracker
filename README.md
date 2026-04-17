# 📋 Job Application Tracker — Google Apps Script

Automatically tracks job applications by scanning a dedicated Gmail label and populating a Google Sheet with company names, roles, status, and email links.

## ✨ Features

| Feature | Details |
|---|---|
| **Auto Company & Role Extraction** | Parses company name and role from email subjects automatically |
| **Gmail Integration** | Scans only the `Job-Applications` label — no full inbox scan |
| **Incremental Scanning** | Remembers last scan time via `PropertiesService` |
| **Duplicate Prevention** | Uses Gmail thread IDs to avoid duplicate rows |
| **Smart Status Detection** | Keywords → `Interview`, `Offer`, `Rejected`, or `Applied` |
| **Aggressive Junk Filtering** | Blocks 100+ patterns — newsletters, alerts, scam offers, promos |
| **Auto-Timestamping** | Editing the Status column auto-updates Last Updated |
| **Time-Driven Trigger** | Configurable auto-scan every 6–12 hours |
| **Scan Log** | All activity logged to a dedicated "Scan Log" sheet |
| **Safe Updates** | Company & Role are never overwritten once manually edited |

---

## 📊 Sheet Columns

| Column | Name | Auto / Manual |
|--------|------|---------------|
| A | Company | **Auto-extracted** (editable) |
| B | Role | **Auto-extracted** (editable) |
| C | Applied Date | Auto |
| D | Status | Auto (editable) |
| E | Email Subject | Auto |
| F | Email Link | Auto (clickable) |
| G | Last Updated | Auto |
| H | Follow-up | Manual |
| I | Notes | Manual |
| J | Interview Date | Manual |
| K | Thread ID | Auto (hidden) |

---

## 🚀 Setup Instructions

### Step 1 — Create the Gmail Label

1. Open [Gmail](https://mail.google.com)
2. In the left sidebar → scroll down → **"+ Create new label"**
3. Name it exactly: **`Job-Applications`**
4. Click **Create**

### Step 2 — Create a Gmail Filter (auto-label future emails)

1. Click the **⚙️ gear icon** (top right) → **See all settings**
2. Go to the **"Filters and Blocked Addresses"** tab
3. Click **"Create a new filter"**
4. In **"Has the words"** field, paste:

```
{("your application was sent") ("application received") ("thank you for applying") ("we received your application") ("application confirmed") ("application submitted") ("application status update") ("regarding your application") ("your application for") ("interview scheduled") ("interview invitation") ("schedule your interview") ("offer letter") ("pleased to offer") ("your candidacy") ("we regret to inform") ("not moving forward") ("coding challenge") ("technical assessment") ("online assessment") ("take-home assignment") ("application acknowledgment") ("job application received") ("new job application") ("application update") ("thank you for your application") ("confirmation of application") ("thank you for considering") ("following up on your") ("we've received your application") ("thank you from") ("next step")}
```

5. In **"Doesn't have"** field, paste:

```
"apply now" "jobs for you" "new jobs" "is hiring" "jobs picked" "job alert" "handpicked" "new opportunities" "walk-in" "webinar" "unsubscribe from" "off campus drive" "off campus hiring" "job guarantee" "tg prime" "batch arranged by"
```

6. Click **"Create filter"**
7. Check ✅ **Apply the label:** → select **`Job-Applications`**
8. Check ✅ **Also apply filter to matching conversations**
9. Click **"Create filter"**

> **⚠️ Note:** The script has its own built-in junk filter, so even if some unwanted emails slip through the Gmail filter, they'll be automatically skipped during scanning.

### Step 3 — Create the Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) → create a **new spreadsheet**
2. Name it: **"Job Application Tracker"** (or any name you like)

### Step 4 — Add the Script

1. In your spreadsheet → **Extensions → Apps Script**
2. Delete any boilerplate code in the editor
3. Copy the **entire contents** of [`Code.gs`](Code.gs)
4. Paste it into the Apps Script editor
5. Click **💾 Save** (or `Ctrl+S`)
6. Close the Apps Script tab

### Step 5 — Initialize & Run

1. **Reload** the spreadsheet (close and reopen, or refresh the page)
2. Wait a few seconds — a **"Job Tracker"** menu will appear in the menu bar
3. Click **Job Tracker → ⚙️ Setup Sheet** — creates headers and formatting
4. Click **Job Tracker → 📧 Scan Emails** — runs your first scan
5. *(Optional)* Click **Job Tracker → ⏰ Install Auto-Scan Trigger** for automatic scanning

### Step 6 — Authorize Permissions

On first run, Google will prompt you to authorize. The script needs:

| Permission | Why |
|---|---|
| Gmail (read-only) | Read emails under the label |
| Spreadsheet | Write application data |
| Script Properties | Store last scan timestamp |
| Triggers | Time-driven auto-scan |

> **💡 Tip:** You may see a "This app isn't verified" warning. Click **Advanced → Go to (project name)** to proceed. This is normal for custom scripts.

---

## 🧠 How It Works

### Company Name Extraction (20+ patterns)

The script automatically extracts company names from email subjects:

| Email Subject | Extracted Company |
|---|---|
| "Thank you for applying to **Barclays**" | Barclays |
| "**Capgemini Group** – New Job Application Received" | Capgemini |
| "Application Received for Web Developer at **Outscal**" | Outscal |
| "**DevOps Engineer** Role Application Update- **FlytBase**" | FlytBase |
| "Following up on your recent application to **Google**" | Google |
| "Thank you from **Emerson**!" | Emerson |
| "News on your **Accenture** Application" | Accenture |
| "Interview Invitation For Software Developer" | *(uses sender name)* |

If no pattern matches, it falls back to the **sender display name** (e.g., "FlytBase Hiring Team" → "FlytBase Hiring Team").

### Status Detection

| Status | Trigger Keywords |
|--------|-----------------|
| **Rejected** | `unfortunately`, `regret to inform`, `not moving forward`, `not been selected`, etc. |
| **Offer** | `pleased to offer`, `job offer`, `compensation package`, `extend an offer`, etc. |
| **Interview** | `interview invitation`, `interview scheduled`, `technical assessment`, `shortlisted`, etc. |
| **Applied** | Default — when no other keywords match |

### Junk Mail Filter (100+ patterns)

Blocks emails from:

| Category | Examples |
|---|---|
| **Job portals** | Glassdoor alerts, LinkedIn recommendations, Naukri, Indeed, Monster |
| **Campus drives** | Profound, Revature, Rupeek promos |
| **Newsletters** | TechGig (TG Prime), ZTM Monthly, Google Cloud |
| **Social media** | Twitter/X notifications, LinkedIn social |
| **Course promos** | xPRO, TestPro, GUVI, BITS Pilani, upGrad |
| **Scam offers** | "Dear Congratulations - Your Job Offer Letter @..." |
| **Spam internships** | Uptricks, Navodita, MeriSKILL, CORIZO, FrameX |
| **Banking** | ICICI, HDFC (accidentally labeled) |
| **Events** | Hackathons, GDG events, webinars |

---

## ⚙️ Customization

| Setting | Location | Default |
|---------|----------|---------|
| Gmail label name | `GMAIL_LABEL` constant | `Job-Applications` |
| Trigger frequency | `TRIGGER_HOURS` constant | `6` hours |
| Sheet name | `SHEET_NAME` constant | `Applications` |

Edit the constants at the top of `Code.gs` to customize.

---

## 🏗️ Architecture

```
onOpen()                    → Custom menu
setupSheet()                → Sheet creation + formatting
scanEmails()                → Incremental scan entry point
  ├─ fetchThreadsSince_()   → Paginated Gmail fetch
  ├─ buildThreadMap_()      → Thread ID → row lookup
  └─ processThread_()       → Per-thread processing
       ├─ isJunkEmail_()    → 3-layer junk detection
       ├─ extractCompanyAndRole_()  → 20+ regex patterns
       └─ detectStatus_()  → Keyword-based classification
fullRescan()                → Clear timestamp + rescan all
onEdit()                    → Auto-timestamp on Status change
installTrigger()            → Time-driven auto-scan setup
logMessage_()               → Scan Log sheet
```

---

## 📝 Menu Options

| Menu Item | Action |
|---|---|
| 📧 Scan Emails | Incremental scan (only new emails) |
| 🔄 Full Rescan | Clears timestamp, processes ALL labeled emails |
| ⚙️ Setup Sheet | Creates/formats the Applications sheet |
| ⏰ Install Auto-Scan Trigger | Enables automatic scanning every 6 hours |
| 🗑️ Remove Auto-Scan Trigger | Disables automatic scanning |

---

## License

MIT — use freely.
