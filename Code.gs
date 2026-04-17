/**
 * ============================================================
 * Job Application Tracker — Google Apps Script  v2.0
 * ============================================================
 *
 * Scans the Gmail label "Job-Applications" incrementally,
 * populates / updates a Google Sheet with application data,
 * auto-extracts company names, and filters out junk mail.
 *
 * Sheet columns (1-indexed):
 *   A  Company          (auto-extracted, editable)
 *   B  Role             (auto-extracted, editable)
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
 * @version 2.0.0
 */

/* ─────────────────────────────────────────────
 * Constants
 * ───────────────────────────────────────────── */

const SHEET_NAME        = 'Applications';
const GMAIL_LABEL       = 'Job-Applications';
const LAST_SCAN_KEY     = 'lastScanTime';
const TRIGGER_HOURS     = 6;
const COMPANY_COL       = 1;          // column A
const ROLE_COL          = 2;          // column B
const APPLIED_DATE_COL  = 3;          // column C
const STATUS_COL        = 4;          // column D
const EMAIL_SUBJECT_COL = 5;          // column E
const EMAIL_LINK_COL    = 6;          // column F
const LAST_UPDATED_COL  = 7;          // column G
const THREAD_ID_COL     = 11;         // column K

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
    .addItem('📧 Scan Emails', 'scanEmails')
    .addItem('🔄 Full Rescan (All Emails)', 'fullRescan')
    .addSeparator()
    .addItem('⚙️ Setup Sheet', 'setupSheet')
    .addItem('⏰ Install Auto-Scan Trigger', 'installTrigger')
    .addItem('🗑️ Remove Auto-Scan Trigger', 'removeTrigger')
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
 * One-time sheet setup — creates the sheet + headers + formatting.
 */
function setupSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getRange('A1').getValue() === '') {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    const widths = [180, 200, 120, 100, 300, 200, 140, 120, 200, 120, 160];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    sheet.hideColumns(THREAD_ID_COL);
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

  const props        = PropertiesService.getScriptProperties();
  const lastScanRaw  = props.getProperty(LAST_SCAN_KEY);
  const lastScanDate = lastScanRaw ? new Date(lastScanRaw) : null;
  const now          = new Date();

  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  if (!label) {
    logMessage_('⚠️  Gmail label "' + GMAIL_LABEL + '" not found. Create it first.');
    showToast_('Label "' + GMAIL_LABEL + '" not found.');
    return;
  }

  const threads = fetchThreadsSince_(label, lastScanDate);
  if (threads.length === 0) {
    logMessage_('ℹ️  No new threads since last scan.');
    showToast_('No new emails found.');
    props.setProperty(LAST_SCAN_KEY, now.toISOString());
    return;
  }

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
 * an existing one. Skips junk/alert emails.
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
  const senderRaw    = firstMsg.getFrom() || '';
  const body         = (latestMsg.getPlainBody() || '').substring(0, 2000);

  // ── Skip junk mail ──
  if (isJunkEmail_(senderRaw, subject, body)) {
    logMessage_('⏭️  Skipped junk: "' + subject.substring(0, 80) + '"');
    return 'skipped';
  }

  const emailLink    = buildGmailLink_(threadId);
  const status       = detectStatus_(subject + ' ' + body);
  const dateReceived = latestMsg.getDate();

  const existingRow = threadMap[threadId];

  if (existingRow) {
    // ── UPDATE existing row (preserve Company & Role if manually edited) ──
    sheet.getRange(existingRow, STATUS_COL).setValue(status);
    sheet.getRange(existingRow, EMAIL_SUBJECT_COL).setValue(subject);
    sheet.getRange(existingRow, EMAIL_LINK_COL).setValue(emailLink);
    sheet.getRange(existingRow, LAST_UPDATED_COL).setValue(dateReceived);
    return 'updated';
  }

  // ── Extract company & role ──
  const senderName   = extractSenderName_(senderRaw);
  const extracted    = extractCompanyAndRole_(subject, senderName, body);

  // ── INSERT new row ──
  const newRow = [
    extracted.company,   // A — Company (auto-extracted)
    extracted.role,      // B — Role    (auto-extracted)
    dateReceived,        // C — Applied Date
    status,              // D — Status
    subject,             // E — Email Subject
    emailLink,           // F — Email Link
    dateReceived,        // G — Last Updated
    '',                  // H — Follow-up
    '',                  // I — Notes
    '',                  // J — Interview Date
    threadId             // K — Thread ID
  ];

  sheet.appendRow(newRow);
  const lastRow = sheet.getLastRow();
  threadMap[threadId] = lastRow;

  return 'inserted';
}


/* ─────────────────────────────────────────────
 * Company & Role Extraction
 * ───────────────────────────────────────────── */

/**
 * Extracts the sender display name from a "From" field.
 * e.g. "FlytBase Hiring Team <careers@flytbase.com>" → "FlytBase Hiring Team"
 */
function extractSenderName_(fromField) {
  const match = fromField.match(/^"?([^"<]+)"?\s*</);
  if (match) return match[1].trim();
  // If no angle brackets, return the whole thing minus the email
  return fromField.replace(/<[^>]+>/, '').trim();
}


/**
 * Extracts company name and role from email subject line
 * using pattern matching. Falls back to sender name.
 *
 * @returns {{ company: string, role: string }}
 */
function extractCompanyAndRole_(subject, senderName, body) {
  const s = subject.trim();
  let company = '';
  let role = '';
  let match;

  // ── Pattern: "Thank you for applying to COMPANY" ──
  match = s.match(/(?:thanks?(?:\s+you)?)\s+(?:for\s+)?applying\s+to\s+(.+?)(?:\s*[!.,;:\-–—]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "Thank you for your [job] application to COMPANY" ──
  match = s.match(/thank\s+you\s+for\s+your\s+(?:job\s+)?application\s+to\s+(.+?)(?:\s*[!.,;:\-–—]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "COMPANY: We've received your application" ──
  match = s.match(/^(.+?):\s*we'?ve?\s+received\s+your\s+application/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "We've received your application" (use sender) ──
  if (/we'?ve?\s+received\s+your\s+application/i.test(s)) {
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "COMPANY – New Job Application Received" ──
  match = s.match(/^(.+?)\s*[–\-—]\s*(?:new\s+)?(?:job\s+)?application\s+received/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "Application for ROLE received, Thank you!" ──
  match = s.match(/application\s+for\s+(.+?)\s+received/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "Application Received for ROLE at COMPANY" ──
  match = s.match(/application\s+received\s+(?:for\s+)?(.+?)\s+at\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(match[2]);
    return { company, role };
  }

  // ── Pattern: "Application received for ROLE ." ──
  match = s.match(/application\s+received\s+for\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "ROLE Role Application Update- COMPANY" ──
  match = s.match(/(.+?)\s+(?:role\s+)?application\s+update\s*[-–—]\s*(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(match[2]);
    return { company, role };
  }

  // ── Pattern: "Your ROLE Application at COMPANY" ──
  match = s.match(/your\s+(.+?)\s+application\s+at\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(match[2]);
    return { company, role };
  }

  // ── Pattern: "Thank you for your application to COMPANY - ID ROLE" ──
  match = s.match(/applying\s+to\s+(.+?)\s*-\s*\d*\s*(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    role = match[2].trim();
    return { company, role };
  }

  // ── Pattern: "Interview Invitation For ROLE" ──
  match = s.match(/interview\s+invitation\s+for\s+(.+?)(?:\s*[!.,;:\-–—]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "Interview Process for ROLE - COMPANY" ──
  match = s.match(/interview\s+(?:process|details)\s+for\s+(.+?)\s*[-–—]\s*(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(match[2]);
    return { company, role };
  }

  // ── Pattern: "You have successfully submitted your COMPANY job application" ──
  match = s.match(/submitted\s+your\s+(.+?)\s+job\s+application/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    // Try to get role from after the dash: "- ROLE"
    const roleMatch = s.match(/application\s*[-–—]\s*(.+?)$/i);
    if (roleMatch) role = roleMatch[1].trim();
    return { company, role };
  }

  // ── Pattern: "Thank you from COMPANY!" ──
  match = s.match(/thank\s+you\s+from\s+(.+?)(?:\s*[!.,;:\-–—]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    // Try to extract role from body
    const roleMatch = body.match(/(?:application\s+for|position\s+of|role\s+of)\s+(.+?)(?:\.|,|\n)/i);
    if (roleMatch) role = roleMatch[1].trim();
    return { company, role };
  }

  // ── Pattern: "Following up on your recent application to COMPANY" ──
  match = s.match(/(?:following up|update)\s+on\s+your\s+(?:recent\s+)?application\s+to\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "News on your COMPANY Application" ──
  match = s.match(/(?:news|update)\s+on\s+your\s+(.+?)\s+application/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "Thank you for considering COMPANY" ──
  match = s.match(/thank\s+you\s+for\s+considering\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "COMPANY Job Application Update" ──
  match = s.match(/^(.+?)\s+job\s+application\s+(?:update|status)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "Confirmation of Application Received for ROLE" ──
  match = s.match(/confirmation\s+of\s+application\s+received\s+for\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "Thank You For Applying at COMPANY" ──
  match = s.match(/thank\s+you\s+for\s+applying\s+(?:at|to)\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "COMPANY Online Assessment" or "Online Assessment by COMPANY" ──
  match = s.match(/(?:online\s+assessment|assessment\s+invitation)\s+(?:by|from)\s+(.+?)(?:\s*[!.,;:\-–—]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    return { company, role };
  }

  // ── Pattern: "[Y Combinator] Your application for COMPANY - ROLE" ──
  match = s.match(/your\s+application\s+for\s+(.+?)\s*[-–—]\s*(.+?)(?:\s*[!.,;(]|$)/i);
  if (match) {
    company = cleanCompanyName_(match[1]);
    role = match[2].trim();
    return { company, role };
  }

  // ── Pattern: "Shortlist for next stage - ROLE" ──
  match = s.match(/shortlist(?:ed)?\s+(?:for\s+)?(?:next\s+stage\s*[-–—]\s*)?(.+?)(?:\s*[!.,;]|$)/i);
  if (match) {
    role = match[1].trim();
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Pattern: "ROLE at COMPANY" ──
  match = s.match(/^(.+?)\s+at\s+(.+?)(?:\s*[!.,;]|$)/i);
  if (match && match[1].length < 60 && match[2].length < 60) {
    role = match[1].replace(/^(?:reminder:\s*|re:\s*)/i, '').trim();
    company = cleanCompanyName_(match[2]);
    return { company, role };
  }

  // ── Pattern: "Thank You For Applying!" (generic — use sender name) ──
  if (/thank\s+you\s+for\s+(?:applying|your\s+application)/i.test(s)) {
    company = cleanCompanyName_(senderName);
    return { company, role };
  }

  // ── Fallback: use sender name as company ──
  company = cleanCompanyName_(senderName);
  return { company, role };
}


/**
 * Cleans up extracted company name — removes common suffixes,
 * email noise, and trims.
 */
function cleanCompanyName_(name) {
  if (!name) return '';
  return name
    .replace(/\s*<[^>]+>/, '')                    // remove emails
    .replace(/\s*[-–—]\s*$/, '')                  // trailing dashes
    .replace(/\s*(Pvt\.?|Private|Ltd\.?|Limited|Inc\.?|LLC|Corp\.?|Group|Technologies|Solutions)\s*\.?$/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}


/* ─────────────────────────────────────────────
 * Status Detection
 * ───────────────────────────────────────────── */

/**
 * Detects application status from email text using keyword matching.
 * Order: Rejected → Offer → Interview → Applied
 */
function detectStatus_(text) {
  const lower = text.toLowerCase();

  const rejectionKeywords = [
    'unfortunately', 'regret to inform', 'not moving forward',
    'decided not to proceed', 'will not be moving',
    'not been selected', 'unable to offer',
    'we have decided to', 'pursue other candidates',
    'not a match', 'position has been filled',
    'not able to move forward', 'will not be proceeding',
    'application was not successful'
  ];
  if (rejectionKeywords.some(kw => lower.includes(kw))) {
    return 'Rejected';
  }

  const offerKeywords = [
    'pleased to offer', 'job offer',
    'offer of employment', 'compensation package',
    'we are excited to offer', 'we would like to offer',
    'formal offer', 'extend an offer'
  ];
  if (offerKeywords.some(kw => lower.includes(kw))) {
    return 'Offer';
  }

  const interviewKeywords = [
    'interview invitation', 'interview scheduled',
    'schedule a call', 'phone screen',
    'technical assessment', 'coding challenge',
    'meet the team', 'calendar invite',
    'would like to schedule', 'shortlisted',
    'online assessment', 'next stage',
    'been shortlisted'
  ];
  if (interviewKeywords.some(kw => lower.includes(kw))) {
    return 'Interview';
  }

  return 'Applied';
}


/* ─────────────────────────────────────────────
 * Junk Email Detection  (v2.0 — built from real scan data)
 * ───────────────────────────────────────────── */

/**
 * Returns true if the email is junk — job alerts, newsletters,
 * social notifications, course promos, scam offers, banking, etc.
 */
function isJunkEmail_(sender, subject, body) {
  const senderLower  = sender.toLowerCase();
  const subjectLower = subject.toLowerCase();
  const bodyLower    = body.toLowerCase();

  // ══════════════════════════════════════════
  // 1. BLOCKED SENDERS  (domain or address)
  // ══════════════════════════════════════════
  const blockedSenders = [
    // Job portals (alerts, not application confirmations)
    'noreply@glassdoor.com',
    'jobs-noreply@linkedin.com',
    'jobalerts-noreply@linkedin.com',
    'news@linkedin.com',
    'messages-noreply@linkedin.com',
    'naukri.com',
    'monster.com',
    'indeed.com',
    'shine.com',
    'timesjobs.com',
    'foundit.in',
    'iimjobs.com',
    'apna.co',
    'ziprecruiter.com',
    'careerbuilder.com',
    'dice.com',
    'simplyhired.com',

    // Campus drive / training promoters
    'profound',
    'guvi',
    'xpro',
    'testpro',

    // Spam internship mills
    'uptricks',
    'navodita',
    'meriskill',
    'corizo',
    'framex',
    'codsoft',

    // Tech news / newsletters
    'techgig.com',
    'zerotomastery',
    'ztm',
    'tg prime',

    // Social media notifications
    'notification@twitter.com',
    'notify@x.com',
    'notify@twitter.com',
    'postmaster@twitter.com',

    // Events & hackathons
    'gdg-noreply@google.com',
    'google-developer-groups',

    // Education promos
    'bits-pilani',
    'bitspilani',
    'spjimr',
    'greatlearning',
    'upgrad',

    // Banking / non-job
    'icici',
    'hdfcbank',
    'axisbank',

    // Cloud newsletters (not job responses)
    'cloudcommunity@google.com',
    'gocloud',

    // GitHub non-job notifications
    'notifications@github.com',

    // Coding platform promos (not application responses)
    'no-reply@hackerearth.com',
    'hackerrank.com',
    'info@codingninjas',
    'geeksforgeeks'
  ];

  if (blockedSenders.some(s => senderLower.includes(s))) {
    return true;
  }

  // ══════════════════════════════════════════
  // 2. BLOCKED SUBJECT PATTERNS
  // ══════════════════════════════════════════
  const blockedSubjects = [
    // Job alerts & recommendations
    'job alert',
    'jobs for you',
    'jobs picked for you',
    'jobs in your area',
    'new jobs in',
    'jobs matching',
    'jobs that match',
    'recommended jobs',
    'more jobs for you',
    'job listings',
    'top jobs',
    'trending jobs',
    'jobs based on',
    'job openings',
    'see all jobs',
    'view jobs',
    'discover roles',

    // Off campus drives & campus events
    'off campus drive',
    'off campus hiring',
    'mega campus',
    'virtual campus',
    'virtual open campus',
    'campus by',
    'hiring drive for',
    'batch arranged by',

    // Course / training promos
    '100% job guarantee',
    'get offer letter on day 1',
    'training program',
    'training and internship program',
    'bootcamp',
    'ready to launch your',
    'launch your software',
    'elevate your career',
    'full stack courses',
    'festive offer',
    'flat 10k off',
    'work integrated learning',

    // Scam offer letters
    'dear congratulations',
    'your job offer letter @',
    'your job offer letter from',

    // Social media notifications
    'posted:',
    'tweeted:',
    'reacted to this',
    'reacted to',

    // Newsletters & digests
    'tg prime:',
    'tg prime',
    'ztm monthly',
    'weekly digest',
    'daily digest',
    'newsletter',
    'monthly:',

    // Events & hackathons
    'hack-a-bit',
    'hackathon',
    'registration closes',
    'join us in our next event',

    // Education spam
    'master\'s application',
    'master\'s in',
    'scholarship opportunity',
    'secure your admission',
    'prepare early for your',
    'pgdm online',

    // Coding contest promos (not job applications)
    'coding challenge is live',
    'weekly coding challenge',
    'test your coding speed',

    // Banking / finance
    'credit card',
    'retail card',
    'a/c is now active',
    'demat',

    // Walk-ins / webinars
    'walk-in',
    'walk in drive',
    'webinar',

    // General promo patterns
    'handpicked',
    'apply now',
    'is hiring',
    'are hiring',
    'who\'s hiring',
    'career fair',
    'hiring event',
    'salary guide',
    'interesting job opportunity',

    // Specific known junk senders' subject patterns
    'abdul got the',
    'internship program for students',
    'github education',
    'next is you',

    // Spam internship offers
    'internship at navodita',
    'internship at uptricks',
    'marketing interns invitation',
    'data analyst internship offer letter & projects',

    // Non-job tech news
    'sbi to hire',
    'wfh is a stress',
    'ai is stealing',
    'global ai power',
    'nvidia\'s ai',
  ];

  if (blockedSubjects.some(p => subjectLower.includes(p))) {
    return true;
  }

  // ══════════════════════════════════════════
  // 3. BLOCKED BODY PATTERNS
  // ══════════════════════════════════════════
  const blockedBody = [
    'unsubscribe from job alerts',
    'manage your job alerts',
    'job alert preferences',
    'see all jobs',
    'your job listings for',
    'based on your job search',
    'based on your profile and search',
    'jobs you might like',
    'we found new jobs',
    'update your preferences',
    'based on your title and location',
    'you received this email because you are subscribed'
  ];

  if (blockedBody.some(p => bodyLower.includes(p))) {
    return true;
  }

  return false;
}


/* ─────────────────────────────────────────────
 * Gmail Helpers
 * ───────────────────────────────────────────── */

/**
 * Fetches threads under a label since a given date, paginated.
 */
function fetchThreadsSince_(label, sinceDate) {
  const allThreads = [];
  const pageSize   = 100;
  let   start      = 0;

  while (true) {
    const batch = label.getThreads(start, pageSize);
    if (batch.length === 0) break;

    for (const thread of batch) {
      if (sinceDate && thread.getLastMessageDate() < sinceDate) {
        return allThreads;
      }
      allThreads.push(thread);
    }

    if (batch.length < pageSize) break;
    start += pageSize;
  }

  return allThreads;
}


/**
 * Builds a clickable Gmail web link from a thread ID.
 */
function buildGmailLink_(threadId) {
  return 'https://mail.google.com/mail/u/0/#inbox/' + threadId;
}


/* ─────────────────────────────────────────────
 * Sheet Helpers
 * ───────────────────────────────────────────── */

/**
 * Returns the Applications sheet, creating & formatting it if needed.
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
 */
function buildThreadMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map     = {};

  if (lastRow < 2) return map;

  const ids = sheet.getRange(2, THREAD_ID_COL, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0]).trim();
    if (id) {
      map[id] = i + 2;
    }
  }

  return map;
}


/* ─────────────────────────────────────────────
 * onEdit Trigger — Auto-timestamp on status change
 * ───────────────────────────────────────────── */

/**
 * Auto-updates Last Updated when Status column is manually edited.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();

    if (col === STATUS_COL && row > 1) {
      sheet.getRange(row, LAST_UPDATED_COL).setValue(new Date());
    }
  } catch (err) {
    Logger.log('onEdit error: ' + err.message);
  }
}


/* ─────────────────────────────────────────────
 * Time-based Trigger Management
 * ───────────────────────────────────────────── */

/**
 * Installs a time-driven trigger to run scanEmails every N hours.
 */
function installTrigger() {
  removeTrigger();

  ScriptApp.newTrigger('scanEmails')
    .timeBased()
    .everyHours(TRIGGER_HOURS)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Auto-scan trigger installed — runs every ' + TRIGGER_HOURS + ' hours.'
  );
}


/**
 * Removes all time-driven triggers for scanEmails.
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
 */
function logMessage_(message) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let   log = ss.getSheetByName('Scan Log');

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
 * Shows a non-blocking toast notification.
 */
function showToast_(message, seconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(message, 'Job Tracker', seconds || 5);
  } catch (_) {
    // toast may fail in time-driven context
  }
}
