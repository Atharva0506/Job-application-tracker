/**
 * ============================================================
 * Job Application Tracker — Google Apps Script
 * ============================================================
 * 
 * Scans the Gmail label "Job-Applications" incrementally,
 * populates / updates a Google Sheet with application data,
 * and detects status via keyword matching.
 *
 * Sheet columns (1-indexed):
 *   A  Company          (manual — never overwritten)
 *   B  Role             (manual — never overwritten)
 *   C  Applied Date
 *   D  Status
 *   E  Email Subject
 *   F  Email Link
 *   G  Last Updated
 *   H  Follow-up        (manual)
 *   I  Notes            (manual)
 *   J  Interview Date   (manual)
 *   K  Thread ID        (hidden helper column)
 *
 * @author  Auto-generated
 * @version 1.0.0
 */

/* ─────────────────────────────────────────────
 * Constants
 * ───────────────────────────────────────────── */

const SHEET_NAME        = 'Applications';
const GMAIL_LABEL       = 'Job-Applications';
const LAST_SCAN_KEY     = 'lastScanTime';
const TRIGGER_HOURS     = 6;          // run every N hours (6–12)
const THREAD_ID_COL     = 11;         // column K
const STATUS_COL        = 4;          // column D
const LAST_UPDATED_COL  = 7;          // column G
const EMAIL_SUBJECT_COL = 5;          // column E
const EMAIL_LINK_COL    = 6;          // column F
const APPLIED_DATE_COL  = 3;          // column C

const HEADERS = [
  'Company', 'Role', 'Applied Date', 'Status',
  'Email Subject', 'Email Link', 'Last Updated',
  'Follow-up', 'Notes', 'Interview Date', 'Thread ID'
];


/* ─────────────────────────────────────────────
 * Menu & Initialisation
 * ───────────────────────────────────────────── */

/**
 * Adds the custom "Job Tracker" menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Tracker')
    .addItem('Scan Emails', 'scanEmails')
    .addItem('Full Rescan (All Emails)', 'fullRescan')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addItem('Install Auto-Scan Trigger', 'installTrigger')
    .addItem('Remove Auto-Scan Trigger', 'removeTrigger')
    .addToUi();
}


/**
 * Clears the last scan timestamp and re-scans ALL emails
 * under the label from scratch.
 */
function fullRescan() {
  PropertiesService.getScriptProperties().deleteProperty(LAST_SCAN_KEY);
  logMessage_('🔄 Last scan timestamp cleared — running full rescan...');
  scanEmails();
}


/**
 * One-time sheet setup — creates the sheet + headers + formatting
 * if they don't already exist.
 */
function setupSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Write headers if row 1 is empty
  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // Bold header row
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Set reasonable column widths
    const widths = [160, 180, 110, 100, 280, 200, 140, 120, 200, 120, 160];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // Hide the Thread ID column (K)
    sheet.hideColumns(THREAD_ID_COL);

    // Auto-filter
    sheet.getRange(1, 1, 1, HEADERS.length).createFilter();
  }

  SpreadsheetApp.getUi().alert('✅ Sheet "' + SHEET_NAME + '" is ready.');
}


/* ─────────────────────────────────────────────
 * Core: Email Scanner
 * ───────────────────────────────────────────── */

/**
 * Main entry point — scans the Gmail label incrementally
 * and upserts rows based on Gmail thread IDs.
 */
function scanEmails() {
  const sheet = getOrCreateSheet_();

  // ── Incremental scan window ──
  const props        = PropertiesService.getScriptProperties();
  const lastScanRaw  = props.getProperty(LAST_SCAN_KEY);
  const lastScanDate = lastScanRaw ? new Date(lastScanRaw) : null;
  const now          = new Date();

  // ── Locate the Gmail label ──
  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  if (!label) {
    logMessage_('⚠️  Gmail label "' + GMAIL_LABEL + '" not found. Create it first.');
    showToast_('Label "' + GMAIL_LABEL + '" not found.');
    return;
  }

  // ── Fetch threads (newest first, paginated) ──
  const threads = fetchThreadsSince_(label, lastScanDate);
  if (threads.length === 0) {
    logMessage_('ℹ️  No new threads since last scan.');
    showToast_('No new emails found.');
    props.setProperty(LAST_SCAN_KEY, now.toISOString());
    return;
  }

  // ── Build a lookup map of existing Thread IDs → row numbers ──
  const threadMap = buildThreadMap_(sheet);

  let inserted = 0;
  let updated  = 0;
  let skipped  = 0;

  for (const thread of threads) {
    try {
      const result = processThread_(sheet, thread, threadMap);
      if (result === 'inserted') inserted++;
      else if (result === 'updated') updated++;
      else skipped++;
    } catch (err) {
      skipped++;
      logMessage_('❌ Error processing thread ' + thread.getId() + ': ' + err.message);
    }
  }

  // ── Persist scan timestamp ──
  props.setProperty(LAST_SCAN_KEY, now.toISOString());

  const summary = '✅ Scan complete — ' + inserted + ' new, ' + updated + ' updated, ' + skipped + ' skipped.';
  logMessage_(summary);
  showToast_(summary);
}


/* ─────────────────────────────────────────────
 * Thread Processing
 * ───────────────────────────────────────────── */

/**
 * Processes a single Gmail thread — inserts a new row or updates
 * an existing one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Gmail.GmailThread} thread
 * @param {Object<string, number>} threadMap  threadId → row number
 * @returns {'inserted'|'updated'|'skipped'}
 */
function processThread_(sheet, thread, threadMap) {
  const threadId = thread.getId();
  const messages = thread.getMessages();

  if (!messages || messages.length === 0) {
    logMessage_('⏭️  Skipped thread ' + threadId + ' — no messages.');
    return 'skipped';
  }

  const latestMsg    = messages[messages.length - 1];
  const firstMsg     = messages[0];
  const subject      = latestMsg.getSubject() || '(no subject)';
  const senderEmail  = firstMsg.getFrom() || '';
  const body         = (latestMsg.getPlainBody() || '').substring(0, 2000);

  // ── Skip job alerts / newsletters / promos ──
  if (isJobAlert_(senderEmail, subject, body)) {
    logMessage_('⏭️  Skipped alert/newsletter: "' + subject.substring(0, 80) + '" from ' + senderEmail);
    return 'skipped';
  }

  const emailLink    = buildGmailLink_(threadId);
  const status       = detectStatus_(subject + ' ' + body);
  const dateReceived = latestMsg.getDate();

  const existingRow = threadMap[threadId];

  if (existingRow) {
    // ── UPDATE existing row (preserve Company & Role) ──
    sheet.getRange(existingRow, STATUS_COL).setValue(status);
    sheet.getRange(existingRow, EMAIL_SUBJECT_COL).setValue(subject);
    sheet.getRange(existingRow, EMAIL_LINK_COL).setValue(emailLink);
    sheet.getRange(existingRow, LAST_UPDATED_COL).setValue(dateReceived);
    return 'updated';
  }

  // ── INSERT new row ──
  const newRow = [
    '',              // A — Company   (manual)
    '',              // B — Role      (manual)
    dateReceived,    // C — Applied Date
    status,          // D — Status
    subject,         // E — Email Subject
    emailLink,       // F — Email Link
    dateReceived,    // G — Last Updated
    '',              // H — Follow-up (manual)
    '',              // I — Notes     (manual)
    '',              // J — Interview Date (manual)
    threadId         // K — Thread ID (hidden)
  ];

  sheet.appendRow(newRow);

  // Track the newly added row so the same thread isn't inserted
  // again if it appears twice in the current batch.
  const lastRow = sheet.getLastRow();
  threadMap[threadId] = lastRow;

  return 'inserted';
}


/* ─────────────────────────────────────────────
 * Status Detection
 * ───────────────────────────────────────────── */

/**
 * Detects application status from email text using keyword matching.
 * Order matters — more specific statuses are checked first.
 *
 * @param {string} text  Combined subject + body text
 * @returns {string}     One of: 'Interview', 'Offer', 'Rejected', 'Applied'
 */
function detectStatus_(text) {
  const lower = text.toLowerCase();

  // Rejection keywords (check before interview — some rejections
  // mention "after the interview")
  const rejectionKeywords = [
    'unfortunately', 'regret to inform', 'not moving forward',
    'decided not to proceed', 'will not be moving',
    'not been selected', 'rejected', 'unable to offer',
    'we have decided to', 'pursue other candidates',
    'not a match', 'position has been filled'
  ];
  if (rejectionKeywords.some(kw => lower.includes(kw))) {
    return 'Rejected';
  }

  // Offer keywords
  const offerKeywords = [
    'offer letter', 'pleased to offer', 'job offer',
    'offer of employment', 'compensation package',
    'we are excited to offer', 'we would like to offer',
    'formal offer', 'extend an offer'
  ];
  if (offerKeywords.some(kw => lower.includes(kw))) {
    return 'Offer';
  }

  // Interview keywords
  const interviewKeywords = [
    'interview', 'schedule a call', 'phone screen',
    'technical assessment', 'coding challenge',
    'hiring manager', 'meet the team',
    'next steps', 'calendar invite',
    'would like to schedule', 'availability'
  ];
  if (interviewKeywords.some(kw => lower.includes(kw))) {
    return 'Interview';
  }

  return 'Applied';
}


/* ─────────────────────────────────────────────
 * Job Alert / Newsletter Detection
 * ───────────────────────────────────────────── */

/**
 * Determines if an email is a job alert, newsletter, or promotional
 * email rather than an actual application-related email.
 *
 * @param {string} sender   The "From" field
 * @param {string} subject  Email subject
 * @param {string} body     Plain text body (truncated)
 * @returns {boolean}       true = skip this email
 */
function isJobAlert_(sender, subject, body) {
  const senderLower  = sender.toLowerCase();
  const subjectLower = subject.toLowerCase();
  const bodyLower    = body.toLowerCase();

  // ── 1. Block known alert/newsletter senders ──
  const blockedSenders = [
    'noreply@glassdoor.com',
    'jobs-noreply@linkedin.com',
    'jobalerts-noreply@linkedin.com',
    'news@linkedin.com',
    'messages-noreply@linkedin.com',
    'naukri.com',
    'campus.naukri.com',
    'monster.com',
    'indeed.com',
    'jobalert@indeed.com',
    'shine.com',
    'timesjobs.com',
    'foundit.in',
    'hirist.com',
    'iimjobs.com',
    'instahyre.com',
    'hiration.com',
    'apna.co',
    'internshala.com',
    'wellfound.com',
    'ziprecruiter.com',
    'careerbuilder.com',
    'dice.com',
    'simplyhired.com',
    'gocloud',
    'hackerrank.com',       // contest promos, not application responses
    'no-reply@hackerearth.com'
  ];
  if (blockedSenders.some(s => senderLower.includes(s))) {
    return true;
  }

  // ── 2. Block by subject patterns (job alerts / recommendations) ──
  const alertSubjectPatterns = [
    'job alert',
    'jobs for you',
    'jobs picked for you',
    'jobs in your area',
    'new jobs in',
    'jobs matching',
    'jobs that match',
    'recommended jobs',
    'handpicked',
    'apply now',
    'is hiring',
    'are hiring',
    'see all jobs',
    'more jobs for you',
    'job listings for',
    'top jobs',
    'trending jobs',
    'jobs based on',
    'easy apply',
    'walk-in',
    'walk in drive',
    'webinar',
    'workshop',
    'bootcamp',
    'training program',
    'discover roles',
    'view jobs',
    'job openings',
    'who\'s hiring',
    'career fair',
    'hiring event',
    'salary guide',
    'weekly digest',
    'daily digest',
    'newsletter'
  ];
  if (alertSubjectPatterns.some(p => subjectLower.includes(p))) {
    return true;
  }

  // ── 3. Block by body patterns (newsletter/promo indicators) ──
  const alertBodyPatterns = [
    'unsubscribe from job alerts',
    'manage your job alerts',
    'job alert preferences',
    'see all jobs',
    'you received this email because you are subscribed',
    'your job listings for',
    'based on your job search',
    'based on your profile and search',
    'jobs you might like',
    'we found new jobs',
    'why am i seeing this',
    'update your preferences'
  ];
  if (alertBodyPatterns.some(p => bodyLower.includes(p))) {
    return true;
  }

  return false;
}


/* ─────────────────────────────────────────────
 * Gmail Helpers
 * ───────────────────────────────────────────── */

/**
 * Fetches threads under a label that have activity since `sinceDate`.
 * If `sinceDate` is null, fetches all threads under the label.
 * Uses pagination to avoid hitting the 500-thread cap.
 *
 * @param {GoogleAppsScript.Gmail.GmailLabel} label
 * @param {Date|null} sinceDate
 * @returns {GoogleAppsScript.Gmail.GmailThread[]}
 */
function fetchThreadsSince_(label, sinceDate) {
  const allThreads = [];
  const pageSize   = 100;
  let   start      = 0;

  while (true) {
    const batch = label.getThreads(start, pageSize);
    if (batch.length === 0) break;

    for (const thread of batch) {
      // If we have a cutoff date, skip older threads
      if (sinceDate && thread.getLastMessageDate() < sinceDate) {
        // Threads are ordered newest-first, so once we hit an old
        // one the rest are even older — stop early.
        return allThreads;
      }
      allThreads.push(thread);
    }

    if (batch.length < pageSize) break; // no more pages
    start += pageSize;
  }

  return allThreads;
}


/**
 * Builds a clickable Gmail web link from a thread ID.
 *
 * @param {string} threadId
 * @returns {string} Gmail URL
 */
function buildGmailLink_(threadId) {
  return 'https://mail.google.com/mail/u/0/#inbox/' + threadId;
}


/* ─────────────────────────────────────────────
 * Sheet Helpers
 * ───────────────────────────────────────────── */

/**
 * Returns the Applications sheet, creating & formatting it if needed.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  return sheet;
}


/**
 * Builds a map of Thread ID → row number for duplicate checking.
 * Reads all Thread IDs in column K in a single batch call.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>}
 */
function buildThreadMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map     = {};

  if (lastRow < 2) return map; // header only

  const ids = sheet.getRange(2, THREAD_ID_COL, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0]).trim();
    if (id) {
      map[id] = i + 2; // convert 0-index → sheet row (data starts at row 2)
    }
  }

  return map;
}


/* ─────────────────────────────────────────────
 * onEdit Trigger — Auto-timestamp on status change
 * ───────────────────────────────────────────── */

/**
 * Simple-trigger handler: when the Status column (D) is edited
 * manually, auto-update the Last Updated column (G) with the
 * current timestamp.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();

    // Only react to edits in the Status column, skipping the header
    if (col === STATUS_COL && row > 1) {
      sheet.getRange(row, LAST_UPDATED_COL).setValue(new Date());
    }
  } catch (err) {
    // Simple triggers have limited permissions; log quietly
    Logger.log('onEdit error: ' + err.message);
  }
}


/* ─────────────────────────────────────────────
 * Time-based Trigger Management
 * ───────────────────────────────────────────── */

/**
 * Installs a time-driven trigger to run scanEmails every N hours.
 * Removes any existing scan triggers first to prevent duplicates.
 */
function installTrigger() {
  removeTrigger(); // clean slate

  ScriptApp.newTrigger('scanEmails')
    .timeBased()
    .everyHours(TRIGGER_HOURS)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Auto-scan trigger installed — runs every ' + TRIGGER_HOURS + ' hours.'
  );
}


/**
 * Removes all time-driven triggers associated with scanEmails.
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'scanEmails') {
      ScriptApp.deleteTrigger(t);
    }
  }
}


/* ─────────────────────────────────────────────
 * Logging & Utilities
 * ───────────────────────────────────────────── */

/**
 * Appends a timestamped message to a "Scan Log" sheet.
 * Creates the log sheet on first use.
 *
 * @param {string} message
 */
function logMessage_(message) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   log   = ss.getSheetByName('Scan Log');

  if (!log) {
    log = ss.insertSheet('Scan Log');
    log.getRange('A1').setValue('Timestamp');
    log.getRange('B1').setValue('Message');
    log.getRange(1, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground('#f4b400')
      .setFontColor('#ffffff');
    log.setColumnWidth(1, 180);
    log.setColumnWidth(2, 600);
  }

  log.appendRow([new Date(), message]);
  Logger.log(message);
}


/**
 * Shows a non-blocking toast notification in the spreadsheet.
 *
 * @param {string} message
 * @param {number} [seconds=5]
 */
function showToast_(message, seconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(message, 'Job Tracker', seconds || 5);
  } catch (_) {
    // toast may fail in time-driven context — ignore
  }
}
