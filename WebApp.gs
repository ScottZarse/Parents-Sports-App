/**
 * Web app front end for the TeamSnap skip list.
 *
 * Add this as a SECOND script file in the same Apps Script project as
 * TeamSnapSync.gs. It reuses that file's CONFIG, COLS, TAG_KEY, and helper
 * functions, so both must live in the same project.
 *
 * DEPLOY:
 *   1. Add this file (+ Index.html) to the project.
 *   2. Deploy -> New deployment -> gear icon -> Web app.
 *   3. Execute as: Me.  Who has access: Anyone.
 *   4. Copy the web app URL, add ?key=ACCESS_CODE to the end of it (see
 *      ACCESS_CODE below), open THAT full link on your phone, then
 *      Share -> Add to Home Screen. The saved icon remembers the key.
 *
 * Access is gated by ACCESS_CODE below via a URL parameter, not by Google
 * sign-in. Personal Gmail accounts don't reliably expose visitor identity
 * to a script the way a Workspace account does, so email-based gating
 * doesn't work here - a shared link code is the reliable alternative, and
 * it also sidesteps the Google-account sign-in handoff that Safari
 * sometimes gets stuck on entirely.
 */

// Change this to whatever you like. Anyone with the link AND this code
// (baked into the URL as ?key=...) can open the app.
const ACCESS_CODE = 'panthers2026';

// How many days ahead the app shows by default.
const DEFAULT_WINDOW_DAYS = 21;

// ============================================================================
// ENTRY POINT
// ============================================================================

function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || '';

  if (key !== ACCESS_CODE) {
    return HtmlService.createHtmlOutput(
      '<div style="font:16px system-ui;padding:40px;text-align:center;' +
      'background:#12161B;color:#E8EDF2;min-height:100vh;margin:0">' +
      '<p>This link is missing its access code.</p>' +
      '<p style="color:#8A97A6;font-size:14px">Ask for the full link, ' +
      'including the part after <code>?key=</code>.</p>' +
      '</div>');
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kids Soccer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// READ
// ============================================================================

/**
 * Returns upcoming events for the UI.
 * going = true means attending (Going? column is blank).
 */
function getSchedule(days) {
  const windowDays = days || DEFAULT_WINDOW_DAYS;
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const tz = CONFIG.TIME_ZONE;

  if (sheet.getLastRow() < 2) return { events: [], players: [] };

  const lastCol = Math.max(sheet.getLastColumn(), COLS.length);
  const data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });

  const idx = {};
  COLS.forEach(function (c) { idx[c] = header.indexOf(c); });

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const horizon = new Date(today.getTime());
  horizon.setDate(horizon.getDate() + windowDays);

  const out = [];
  const players = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const uid = String(row[idx['UID']]).trim();
    if (!uid) continue;

    const dateVal = row[idx['Date']];
    const d = (dateVal instanceof Date)
      ? new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate())
      : parseYmd_(String(dateVal));
    if (!d || d < today || d > horizon) continue;

    const notes = String(row[idx['Notes']] || '');
    if (notes.indexOf('REMOVED FROM TEAMSNAP') !== -1) continue;

    const player = String(row[idx['Player']] || '').trim();
    if (player) players[player] = true;

    const startStr = cellToTimeString_(row[idx['Start']], tz);

    out.push({
      uid: uid,
      dateKey: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      dateLabel: Utilities.formatDate(d, tz, 'EEE d MMM'),
      start: startStr,
      sortMins: timeToMinutes_(startStr),
      player: player,
      team: String(row[idx['Team']] || ''),
      type: String(row[idx['Type']] || ''),
      title: String(row[idx['Event']] || ''),
      location: String(row[idx['Location']] || ''),
      notes: notes,
      cancelled: notes.indexOf('CANCELLED') !== -1,
      going: !isSkipped_(row[idx['Going?']])
    });
  }

  out.sort(function (a, b) {
    if (a.dateKey !== b.dateKey) return a.dateKey < b.dateKey ? -1 : 1;
    if (a.sortMins !== b.sortMins) return a.sortMins - b.sortMins;
    return a.player < b.player ? -1 : 1;
  });

  return { events: out, players: Object.keys(players).sort() };
}

/** "6:45 PM" -> minutes past midnight. "All day" sorts first. */
/**
 * Reads a Start/End cell safely. If Sheets silently turned the text
 * "6:00 PM" into a Date value (it does this automatically), that Date is
 * anchored to Dec 30 1899 - this reformats it back to "6:00 PM" instead of
 * printing the whole garbled timestamp. Plain text cells pass through as-is.
 */
function cellToTimeString_(val, tz) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, 'h:mm a');
  }
  return String(val || '').trim();
}

function timeToMinutes_(s) {
  if (!s || /all day/i.test(s)) return -1;
  const m = s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
  if (!m) return 0;
  let h = parseInt(m[1], 10);
  const mins = parseInt(m[2], 10);
  const ap = (m[3] || '').toUpperCase();
  if (ap === 'PM' && h !== 12) h += 12;
  if (ap === 'AM' && h === 12) h = 0;
  return h * 60 + mins;
}

function parseYmd_(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

// ============================================================================
// WRITE
// ============================================================================

/**
 * Sets attendance for one event.
 * going = true  -> clear the Going? cell, event belongs on the calendar
 * going = false -> write "No", event gets pulled off the calendar
 *
 * Updates the sheet immediately, then adjusts that single calendar event so
 * the change shows up right away instead of waiting for the next 15-minute run.
 */
function setGoing(uid, going) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('The schedule is syncing right now. Try again in a moment.');
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];

    const lastCol = Math.max(sheet.getLastColumn(), COLS.length);
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                        .map(function (h) { return String(h).trim(); });
    const uidCol = header.indexOf('UID') + 1;
    const goingCol = header.indexOf('Going?') + 1;
    if (uidCol < 1 || goingCol < 1) throw new Error('Sheet is missing UID or Going? column.');

    const uids = sheet.getRange(2, uidCol, sheet.getLastRow() - 1, 1).getValues();
    let target = -1;
    for (let i = 0; i < uids.length; i++) {
      if (String(uids[i][0]).trim() === uid) { target = i + 2; break; }
    }
    if (target === -1) throw new Error('That event is no longer on the schedule.');

    sheet.getRange(target, goingCol).setValue(going ? '' : 'No');
    SpreadsheetApp.flush();

    let calendarUpdated = true;
    try {
      applyOneToCalendar_(uid, !going);
    } catch (err) {
      // The sheet is the source of truth; the next scheduled sync will catch up.
      calendarUpdated = false;
      log_('Immediate calendar update failed for ' + uid + ': ' + err);
    }

    return { uid: uid, going: going, calendarUpdated: calendarUpdated };
  } finally {
    lock.releaseLock();
  }
}

/** Adds or removes a single event on the calendar, without a full sync. */
function applyOneToCalendar_(uid, skip) {
  const feed = fetchAndParseFeed_();
  let e = null;
  for (let i = 0; i < feed.length; i++) {
    if (feed[i].uid === uid) { e = feed[i]; break; }
  }
  if (!e) return;

  const cal = getOrCreateCalendar_();
  const from = new Date(e.start.getTime() - 36 * 60 * 60 * 1000);
  const to = new Date(e.start.getTime() + 36 * 60 * 60 * 1000);

  let existing = null;
  cal.getEvents(from, to).forEach(function (ev) {
    if (ev.getTag(TAG_KEY) === uid) existing = ev;
  });

  if (skip || e.cancelled) {
    if (existing) existing.deleteEvent();
    return;
  }

  if (existing) return;  // already there, nothing to do

  const desc = buildDescription_(e);
  const ev = e.allDay
    ? cal.createAllDayEvent(e.title, e.start, { location: e.location, description: desc })
    : cal.createEvent(e.title, e.start, e.end, { location: e.location, description: desc });

  ev.setTag(TAG_KEY, e.uid);
  applyColor_(ev, e.player);
}

/** Called by the "Sync now" button. Runs the full reconcile. */
function runFullSync() {
  syncNow();
  return true;
}
