/**
 * Solutions team feedback: anonymous responses to Google Sheets, plus a protected summary feed.
 *
 * Setup
 * 1. Create a Google Sheet (e.g. "Solutions Feedback Responses").
 * 2. In the Sheet: Extensions > Apps Script. Paste this file over Code.gs.
 * 3. Project Settings (gear icon) > Script properties > Add script property:
 *      SUMMARY_KEY = a long random string (this is the access key for the summary page).
 * 4. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone   (or "Anyone within <your domain>")
 * 5. Copy the Web app URL (ends in /exec) into ENDPOINT in BOTH html files.
 * 6. Reload the Sheet and run Feedback > Build summary tab (approve permissions the first time).
 * 7. After any change to this script: Deploy > Manage deployments > Edit > New version.
 *
 * Anonymity: the script never reads the submitter's email, IP or user agent.
 * Only the date (no time), the quarter and the answers are stored.
 * Anyone can submit; only someone with SUMMARY_KEY can read results back.
 */

// Leave empty when the script is opened from the Sheet (container-bound).
// Otherwise paste the Sheet ID from its URL.
const SHEET_ID = '';

const SCALE_LABELS = {
  1: 'Needs Improvement',
  2: 'Inconsistent',
  3: 'Reliable',
  4: 'Trusted Partner',
  5: 'Strategic Partner'
};

const TABS = {
  partnership: {
    name: 'Sales Partnership Score',
    headers: ['Date', 'Quarter', 'Score', 'Score label', 'Comment'],
    pick: [0, 1, 2, 4]            // columns sent to the summary page
  },
  pulse: {
    name: 'Customer Perception Pulse',
    headers: ['Date', 'Quarter', 'Helped clarify value', 'Comment'],
    pick: [0, 1, 2, 3]
  }
};

const SUMMARY_TAB = 'Summary';
const MAX_ROWS = 5000;            // most recent rows returned per tab

/* ---------- Menu ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Feedback')
    .addItem('Build summary tab', 'buildSummary')
    .addToUi();
}

function buildSummary() {
  buildSummary_(getSs_());
}

/* ---------- Web app ---------- */

function doGet() {
  return json_({ ok: true, service: 'solutions-feedback' });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request' });
  }

  // Read path: protected by SUMMARY_KEY
  if (data.action === 'summary') {
    try {
      return json_(summary_(data.key));
    } catch (err) {
      return json_({ ok: false, error: 'Server error' });
    }
  }

  // Write path: open to anyone, anonymous
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // Honeypot: bots fill this in, people never see it. Pretend success, save nothing.
    if (data.website) return json_({ ok: true });

    const quarter = String(data.quarter || '');
    if (!/^Q[1-4] \d{4}$/.test(quarter)) return json_({ ok: false, error: 'Invalid quarter' });

    const ss = getSs_();
    const tz = ss.getSpreadsheetTimeZone();
    const ymd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const date = Utilities.parseDate(ymd, tz, 'yyyy-MM-dd');   // a real date, no time
    let saved = 0;

    if (data.partnership) {
      const score = Number(data.partnership.score);
      if (Number.isInteger(score) && score >= 1 && score <= 5) {
        append_(ss, TABS.partnership, [date, quarter, score, SCALE_LABELS[score], clean_(data.partnership.comment)]);
        saved++;
      }
    }

    if (data.pulse) {
      const answer = data.pulse.clarified;
      if (answer === 'Yes' || answer === 'No') {
        append_(ss, TABS.pulse, [date, quarter, answer, clean_(data.pulse.comment)]);
        saved++;
      }
    }

    if (!saved) return json_({ ok: false, error: 'Nothing valid to save' });
    if (!ss.getSheetByName(SUMMARY_TAB)) buildSummary_(ss);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'Server error' });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ---------- Summary feed ---------- */

function summary_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('SUMMARY_KEY');
  if (!expected) return { ok: false, error: 'SUMMARY_KEY is not set' };
  if (!safeEqual_(String(key || ''), expected)) return { ok: false, error: 'Unauthorized' };

  const ss = getSs_();
  const tz = ss.getSpreadsheetTimeZone();
  return {
    ok: true,
    partnership: read_(ss, TABS.partnership, tz),
    pulse: read_(ss, TABS.pulse, tz)
  };
}

function read_(ss, tab, tz) {
  const sheet = ss.getSheetByName(tab.name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const last = sheet.getLastRow();
  const start = Math.max(2, last - MAX_ROWS + 1);
  const values = sheet.getRange(start, 1, last - start + 1, tab.headers.length).getValues();
  return values.map(function (row) {
    return tab.pick.map(function (i) {
      const v = row[i];
      return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v;
    });
  });
}

/* ---------- Summary tab (in the Sheet) ---------- */

function buildSummary_(ss) {
  ensureTab_(ss, TABS.partnership);
  ensureTab_(ss, TABS.pulse);

  let sh = ss.getSheetByName(SUMMARY_TAB);
  if (sh) sh.clear(); else sh = ss.insertSheet(SUMMARY_TAB, 0);

  const P = "'" + TABS.partnership.name + "'";
  const U = "'" + TABS.pulse.name + "'";

  const periods = [
    ['This week',    '=TODAY()-WEEKDAY(TODAY(),3)'],
    ['This month',   '=EOMONTH(TODAY(),-1)+1'],
    ['This quarter', '=DATE(YEAR(TODAY()),INT((MONTH(TODAY())-1)/3)*3+1,1)'],
    ['This year',    '=DATE(YEAR(TODAY()),1,1)'],
    ['All time',     '=DATE(2000,1,1)']
  ];

  const rows = [['Period', 'From', 'Partnership responses', 'Average score (1-5)', 'Pulse responses', '% Yes']];
  periods.forEach(function (p, i) {
    const r = i + 2;
    rows.push([
      p[0],
      p[1],
      '=COUNTIFS(' + P + '!$A:$A,">="&$B' + r + ')',
      '=IFERROR(AVERAGEIFS(' + P + '!$C:$C,' + P + '!$A:$A,">="&$B' + r + '),"")',
      '=COUNTIFS(' + U + '!$A:$A,">="&$B' + r + ')',
      '=IFERROR(COUNTIFS(' + U + '!$A:$A,">="&$B' + r + ',' + U + '!$C:$C,"Yes")/E' + r + ',"")'
    ]);
  });

  sh.getRange(1, 1, rows.length, 6).setFormulas(rows);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  sh.getRange(2, 2, 5, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 4, 5, 1).setNumberFormat('0.00');
  sh.getRange(2, 6, 5, 1).setNumberFormat('0%');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 6);
}

/* ---------- Helpers ---------- */

function getSs_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function ensureTab_(ss, tab) {
  let sheet = ss.getSheetByName(tab.name);
  if (!sheet) {
    sheet = ss.insertSheet(tab.name);
    sheet.appendRow(tab.headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, tab.headers.length).setFontWeight('bold');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  }
  return sheet;
}

function append_(ss, tab, row) {
  ensureTab_(ss, tab).appendRow(row);
}

/** Trims, caps length and stops spreadsheet formula injection (=, +, -, @ at the start). */
function clean_(text) {
  let s = String(text || '').trim().slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
