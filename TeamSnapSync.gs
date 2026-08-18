/**
 * TeamSnap ONE -> Google Calendar sync with a skip list.
 *
 * What it does, every time it runs:
 *   1. Fetches the TeamSnap ICS feed.
 *   2. Renames events by player: "Practice - Emily - Outdoor".
 *   3. Updates the Google Sheet so new/changed events show up automatically.
 *   4. Reads your "Going?" column.
 *   5. Fills any untouched carpool legs from your Defaults tab.
 *   6. Writes every event you have NOT marked "No" into a calendar you own,
 *      keeping the venue as the event location, the TeamSnap link and the
 *      carpool block in the description. Anything marked "No" (or cancelled
 *      in TeamSnap) gets removed and stays gone.
 *
 * Because the events live on a calendar you own, you get full edit rights,
 * and the sync refreshes every 15 minutes instead of Google's ~12-24 hours.
 *
 * SETUP (one time):
 *   1. script.google.com -> New project -> paste this file in.
 *   2. Fill in CONFIG below.
 *   3. Run setup() once. Authorize when prompted.
 *   4. Done. It now runs on its own.
 *
 * To force an immediate refresh at any point, run syncNow().
 *
 * Carpool.gs must be in the same project. Two hooks below call into it:
 * applyDefaultsToUpcoming() in syncNow(), and carpoolBlock_() in
 * buildDescription_(). Both are marked CARPOOL HOOK.
 */

// ============================================================================
// CONFIG - edit these values
// ============================================================================

const CONFIG = {
  // Your TeamSnap ICS feed URL.
  ICS_URL: 'https://calendar-api.teamsnap.com/v1/user.ics?token=UksarQu6FbPfvjj6UXMTOJ5dLsUMHrHSIqawUYnu9p2vcxdbx_tfxyADUvdpUN5M&v=1786026801077',

  // The ID of your skip-list Sheet (the long string in its URL between /d/ and /edit).
  SHEET_ID: '14PNkSfqlPQeMzId-l8BxaJJ7xUBCKZW0ePHXmI9uLV0',

  // Tab name inside that spreadsheet.
  SHEET_NAME: 'Kids Soccer Schedules',

  // Explicit timezone for every displayed date/time. Set this rather than
  // relying on the spreadsheet's own timezone setting, which can default
  // to something other than yours (that mismatch is what caused times to
  // display a few hours off).
  TIME_ZONE: 'America/Chicago',

  // The deployed web app URL, used in each calendar event's description
  // so anyone looking at an event can find their way to the attendance app.
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwj-W3oYmsl4fgswNJDvYatadDSYpNF5ECcJx9P34tdhAvG6i3BDVa33I2ZD5N75ntY/exec',

  // Name of the Google Calendar to write into. Created automatically if missing.
  CALENDAR_NAME: 'Panthers Schedule',

  // Which kid plays on which team. Team names are matched loosely, so
  // "Panthers U12G Black" matches whether or not TeamSnap adds extra spaces.
  // Add a new line here when a kid joins a new team, then re-run
  // setupCarpoolTabs() so the Teams tab picks it up.
  PLAYERS: {
    'Panthers U13G':       'Maddie',
    'Panthers U12G Gold':  'Emily',
    'Panthers U12G Black': 'Emily',
    'Panthers U9B Gold':   'Cooper'
  },

  // Optional: color-code events by player. Leave {} to disable.
  // Valid values and how they appear in Google Calendar:
  //   PALE_RED = Flamingo    RED = Tomato        ORANGE = Tangerine
  //   YELLOW = Banana        PALE_GREEN = Sage   GREEN = Basil
  //   CYAN = Peacock         PALE_BLUE = Lavender
  //   BLUE = Blueberry       MAUVE = Grape       GRAY = Graphite
  PLAYER_COLORS: {
    'Maddie': 'CYAN',    // Peacock
    'Emily':  'ORANGE',  // Tangerine
    'Cooper': 'GREEN'    // Basil
  },

  // How far back to look when reconciling the calendar (days).
  LOOKBACK_DAYS: 7,

  // How far forward to look (days).
  LOOKAHEAD_DAYS: 400,

  // Set true to see verbose output in the execution log.
  VERBOSE: true
};

// Column order in the sheet. Must match the header row.
const COLS = ['UID', 'Date', 'Day', 'Start', 'End', 'Player', 'Team', 'Type',
              'Event', 'Location', 'Notes', 'Link', 'Going?'];

// Key used to tag calendar events so we can match them back to feed UIDs.
const TAG_KEY = 'teamsnap_uid';

// Google Calendar's numeric color IDs. More reliable than the Enum lookup.
const COLOR_IDS = {
  PALE_BLUE: '1',   // Lavender
  PALE_GREEN: '2',  // Sage
  MAUVE: '3',       // Grape
  PALE_RED: '4',    // Flamingo
  YELLOW: '5',      // Banana
  ORANGE: '6',      // Tangerine
  CYAN: '7',        // Peacock
  GRAY: '8',        // Graphite
  BLUE: '9',        // Blueberry
  GREEN: '10',      // Basil
  RED: '11'         // Tomato
};

/** Sets an event's color from CONFIG.PLAYER_COLORS. Safe to call repeatedly. */
function applyColor_(ev, player) {
  const name = CONFIG.PLAYER_COLORS[player];
  if (!name) return false;
  const id = COLOR_IDS[name];
  if (!id) {
    log_('Unknown color name "' + name + '" for ' + player);
    return false;
  }
  try {
    if (ev.getColor() !== id) {
      ev.setColor(id);
      return true;
    }
  } catch (err) {
    log_('setColor failed for ' + player + ': ' + err);
  }
  return false;
}

// ============================================================================
// SETUP
// ============================================================================

/**
 * Run this ONCE. Creates the calendar, seeds the sheet, installs the trigger.
 */
function setup() {
  const cal = getOrCreateCalendar_();
  log_('Calendar ready: ' + cal.getName() + ' (' + cal.getId() + ')');

  // Remove any previously installed triggers so we don't stack them.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('syncNow').timeBased().everyMinutes(15).create();
  log_('Trigger installed: syncNow every 15 minutes.');

  syncNow();
  log_('Setup complete.');
}

/**
 * Removes the recurring trigger. The calendar and sheet are left alone.
 */
function stopSyncing() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });
  log_('Sync trigger removed.');
}

// ============================================================================
// MAIN SYNC
// ============================================================================

function syncNow() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    log_('Another sync is already running. Skipping this pass.');
    return;
  }

  try {
    const feedEvents = fetchAndParseFeed_();
    log_('Feed returned ' + feedEvents.length + ' events.');

    const unknown = feedEvents.filter(function (e) { return !e.player; });
    if (unknown.length) {
      const teams = {};
      unknown.forEach(function (e) { teams[e.team || e.originalSummary] = true; });
      log_('WARNING: no player mapped for: ' + Object.keys(teams).join(', ') +
           '. Add them to CONFIG.PLAYERS.');
    }

    const skipMap = syncSheet_(feedEvents);

    // --- CARPOOL HOOK -------------------------------------------------
    // Fill any untouched legs before the calendar is written, so a brand
    // new event picks up its default driver on the same pass that creates
    // it rather than on the one after. Wrapped so a carpool problem can
    // never stop the schedule sync itself.
    try {
      if (typeof applyDefaultsToUpcoming === 'function') applyDefaultsToUpcoming();
    } catch (err) {
      log_('Carpool defaults failed (schedule sync continuing): ' + err);
    }
    // ------------------------------------------------------------------

    const stats = syncCalendar_(feedEvents, skipMap);

    log_('Sync complete. Created: ' + stats.created +
         ', Updated: ' + stats.updated +
         ', Removed: ' + stats.removed +
         ', Skipped by you: ' + stats.skipped);
  } catch (err) {
    Logger.log('SYNC FAILED: ' + err + '\n' + (err.stack || ''));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// FEED PARSING
// ============================================================================

function fetchAndParseFeed_() {
  const res = UrlFetchApp.fetch(CONFIG.ICS_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Feed fetch failed with HTTP ' + res.getResponseCode() +
                    '. Check that ICS_URL is correct and the token is still valid.');
  }

  // Unfold RFC5545 continuation lines: a newline followed by a space or tab
  // is a continuation of the previous line, not a new one.
  const text = res.getContentText().replace(/\r?\n[ \t]/g, '');
  const lines = text.split(/\r?\n/);

  const events = [];
  let cur = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line === 'BEGIN:VEVENT') {
      cur = {};
    } else if (line === 'END:VEVENT') {
      if (cur && cur.uid) events.push(finalizeEvent_(cur));
      cur = null;
    } else if (cur) {
      const idx = line.indexOf(':');
      if (idx === -1) continue;

      const rawKey = line.substring(0, idx);
      const value = line.substring(idx + 1);
      const prop = rawKey.split(';')[0];

      switch (prop) {
        case 'UID':         cur.uid = value; break;
        case 'SUMMARY':     cur.summary = unescapeIcs_(value); break;
        case 'DESCRIPTION': cur.description = unescapeIcs_(value); break;
        case 'LOCATION':    cur.location = unescapeIcs_(value); break;
        case 'STATUS':      cur.status = value; break;
        case 'DURATION':    cur.duration = value; break;
        case 'DTSTART':
          cur.dtstartRaw = value;
          cur.allDay = rawKey.indexOf('VALUE=DATE') !== -1;
          break;
        case 'DTEND':
          cur.dtendRaw = value;
          break;
      }
    }
  }

  return events;
}

function finalizeEvent_(e) {
  e.start = parseIcsDate_(e.dtstartRaw, e.allDay);

  if (e.dtendRaw) {
    e.end = parseIcsDate_(e.dtendRaw, e.allDay);
  } else if (e.allDay) {
    e.end = new Date(e.start.getTime() + 24 * 60 * 60 * 1000);
  } else {
    e.end = new Date(e.start.getTime() + parseDurationMs_(e.duration));
  }

  e.originalSummary = e.summary || '(untitled)';

  // --- Location: keep the venue, strip TeamSnap's duplicated prefix ---
  // "Heights Park - Heights Park \n711 W Arapaho Rd..." -> "Heights Park, 711 W Arapaho Rd..."
  e.location = (e.location || '').replace(/\n/g, ', ').trim();
  e.location = e.location.replace(/^(.+?)\s*-\s*\1\s*,?\s*/, '$1, ')
                         .replace(/\s+,/g, ',')
                         .replace(/,\s*,/g, ',')
                         .replace(/^,|,$/g, '')
                         .trim();

  e.cancelled = (e.status === 'CANCELLED');

  // --- Team, then player ---
  e.team = extractTeam_(e.originalSummary);
  e.player = lookupPlayer_(e.team);
  e.squad = squadOf_(e.team);

  // --- Type ---
  if (e.originalSummary.indexOf('Game:') === 0) e.type = 'Game';
  else if (e.originalSummary.indexOf('Practice:') === 0) e.type = 'Practice';
  else if (/pick[\s-]?up/i.test(e.originalSummary)) e.type = 'Pickup';
  else e.type = 'Other';

  // --- Notes and TeamSnap deep link, both pulled from DESCRIPTION ---
  const desc = e.description || '';
  const noteMatch = desc.match(/Notes:\s*([^\n]+)/);
  e.feedNotes = noteMatch ? noteMatch[1].trim() : '';

  const linkMatch = desc.match(/Link:\s*(\S+)/);
  e.link = linkMatch ? linkMatch[1].trim() : '';

  const arriveMatch = desc.match(/Arrival:\s*([^\n]+)/);
  e.arrival = arriveMatch ? arriveMatch[1].trim() : '';

  const uniformMatch = desc.match(/Uniform:\s*([^\n]+)/i);
  e.uniform = uniformMatch ? uniformMatch[1].trim() : '';

  // --- Renamed title ---
  e.title = buildTitle_(e);

  // --- Sheet notes: squad first (so Emily's Gold vs Black is obvious) ---
  const noteParts = [];
  if (e.squad && playerHasMultipleTeams_(e.player)) noteParts.push(e.squad + ' team');
  if (e.uniform) noteParts.push('Wear ' + e.uniform);
  if (e.feedNotes) noteParts.push(e.feedNotes);
  if (e.arrival && e.feedNotes.indexOf('rrive') === -1) {
    noteParts.push('Arrive ' + e.arrival);
  }
  e.notes = noteParts.join(' | ');

  return e;
}

/** "Practice: Panthers U12G Black Indoor Practice" -> "Panthers U12G Black" */
function extractTeam_(summary) {
  const stripped = summary.replace(/^(Practice|Game|Event):\s*/, '');
  const m = stripped.match(/(Panthers\s+U\d+[BG]\s*(?:Black|Gold)?)/);
  return m ? m[1].replace(/\s+/g, ' ').trim() : '';
}

/** 'Panthers U12G Gold' -> 'Gold'. Empty when there's no Gold/Black suffix. */
function squadOf_(team) {
  if (!team) return '';
  const last = team.split(' ').pop();
  return (last === 'Gold' || last === 'Black') ? last : '';
}

function lookupPlayer_(team) {
  if (!team) return '';
  if (CONFIG.PLAYERS[team]) return CONFIG.PLAYERS[team];

  // Loose match: collapse whitespace and compare case-insensitively.
  const norm = team.replace(/\s+/g, ' ').toLowerCase();
  const keys = Object.keys(CONFIG.PLAYERS);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].replace(/\s+/g, ' ').toLowerCase() === norm) {
      return CONFIG.PLAYERS[keys[i]];
    }
  }
  return '';
}

function playerHasMultipleTeams_(player) {
  if (!player) return false;
  let n = 0;
  Object.keys(CONFIG.PLAYERS).forEach(function (t) {
    if (CONFIG.PLAYERS[t] === player) n++;
  });
  return n > 1;
}

/**
 * Builds the renamed title.
 *   "Practice: Panthers U13G Indoor"                  -> "Practice - Maddie - Indoor"
 *   "Practice: Panthers U12G Black Outdoor Practice"  -> "Practice - Emily - Outdoor"
 *   "Practice: Panthers U9B Gold Outdoor - Loft City" -> "Practice - Cooper - Outdoor"
 *   "Game: Panthers U13G vs Reign FC (Arrive 30...)"  -> "Game - Maddie - vs Reign FC"
 *   "Panthers U12G Black Panther Pick-Up"             -> "Panther Pick-Up - Emily"
 */
function buildTitle_(e) {
  // If we can't identify the player, leave the original name alone rather
  // than producing something confusing.
  if (!e.player) return e.originalSummary;

  // Pick-Ups don't have a descriptor beyond themselves (no Indoor/Outdoor,
  // no opponent) - just "Pick-Up - Player" rather than trying to strip the
  // original feed text down.
  if (e.type === 'Pickup') return 'Pick-Up - ' + e.player;

  let d = e.originalSummary.replace(/^(Practice|Game|Event):\s*/, '');

  // Drop the team name.
  if (e.team) d = d.replace(e.team, '');
  // Handle the case where the feed used different spacing than our match.
  d = d.replace(/Panthers\s+U\d+[BG]\s*(?:Black|Gold)?/, '');

  // Drop a parenthetical like "(Arrive 30 minutes early)" - it lives in Notes.
  d = d.replace(/\s*\([^)]*\)\s*$/, '');

  // Drop a trailing venue suffix, since the venue is already the location.
  // "Outdoor - Loft City Church" -> "Outdoor" (only when it matches the location)
  const loc = (e.location || '').toLowerCase();
  d = d.replace(/\s*-\s*(.+)$/, function (match, tail) {
    return loc.indexOf(tail.trim().toLowerCase()) !== -1 ? '' : match;
  });

  // "Indoor Practice" -> "Indoor" (the word Practice is already the prefix).
  d = d.replace(/\s*Practice\s*$/i, '');

  d = d.replace(/\s+/g, ' ').trim().replace(/^-\s*|\s*-$/g, '').trim();

  if (!d) d = (e.type === 'Game') ? 'Game' : 'Practice';

  if (e.type === 'Practice' || e.type === 'Game') {
    return e.type + ' - ' + e.player + ' - ' + d;
  }
  return d + ' - ' + e.player;
}

/** "20260812T234500Z" or "20260814" -> Date */
function parseIcsDate_(raw, allDay) {
  if (!raw) throw new Error('Event is missing DTSTART.');

  const y = parseInt(raw.substr(0, 4), 10);
  const mo = parseInt(raw.substr(4, 2), 10) - 1;
  const d = parseInt(raw.substr(6, 2), 10);

  if (allDay || raw.length === 8) {
    return new Date(y, mo, d);  // local midnight
  }

  const h = parseInt(raw.substr(9, 2), 10);
  const mi = parseInt(raw.substr(11, 2), 10);
  const s = parseInt(raw.substr(13, 2), 10) || 0;

  if (raw.charAt(raw.length - 1) === 'Z') {
    return new Date(Date.UTC(y, mo, d, h, mi, s));
  }
  return new Date(y, mo, d, h, mi, s);  // floating time -> treat as local
}

/** "PT1H45M" / "P1DT" / "PT30M" -> milliseconds */
function parseDurationMs_(dur) {
  if (!dur) return 60 * 60 * 1000;  // default 1 hour
  const m = dur.match(/^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?/);
  if (!m) return 60 * 60 * 1000;
  const days = parseInt(m[1] || 0, 10);
  const hrs = parseInt(m[2] || 0, 10);
  const mins = parseInt(m[3] || 0, 10);
  const secs = parseInt(m[4] || 0, 10);
  const ms = ((days * 24 + hrs) * 60 + mins) * 60 * 1000 + secs * 1000;
  return ms || 60 * 60 * 1000;
}

function unescapeIcs_(s) {
  return (s || '')
    .replace(/\\n/gi, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\');
}

// ============================================================================
// SHEET SYNC
// ============================================================================

/**
 * Reconciles the sheet against the feed and returns a map of UID -> true
 * for every event the user has marked as skipped.
 *
 * New feed events get appended. Existing rows get their date/time/location
 * refreshed (so TeamSnap changes surface), but the "Going?" column is never
 * touched. Rows whose events vanished from the feed are marked in Notes.
 */
function syncSheet_(feedEvents) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];

  if (ss.getSpreadsheetTimeZone() !== CONFIG.TIME_ZONE) {
    ss.setSpreadsheetTimeZone(CONFIG.TIME_ZONE);
    log_('Corrected spreadsheet timezone to ' + CONFIG.TIME_ZONE);
  }

  if (sheet.getLastRow() === 0) sheet.appendRow(COLS);

  // Add any missing columns rather than failing, so the schema can evolve.
  ensureColumns_(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), COLS.length);
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });

  const colIdx = {};
  COLS.forEach(function (name) { colIdx[name] = header.indexOf(name); });

  lockTextColumns_(sheet, colIdx);

  const rowByUid = {};
  for (let r = 1; r < data.length; r++) {
    const uid = String(data[r][colIdx['UID']]).trim();
    if (uid) rowByUid[uid] = r;
  }

  const tz = CONFIG.TIME_ZONE;
  const feedUids = {};
  const skipMap = {};
  const updates = [];
  const additions = [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 1);

  feedEvents.forEach(function (e) {
    feedUids[e.uid] = true;

    const dateStr = Utilities.formatDate(e.start, tz, 'yyyy-MM-dd');
    const dayStr = Utilities.formatDate(e.start, tz, 'EEE');
    const startStr = e.allDay ? 'All day' : Utilities.formatDate(e.start, tz, 'h:mm a');
    const endStr = e.allDay ? '' : Utilities.formatDate(e.end, tz, 'h:mm a');
    const notes = e.cancelled
      ? ('CANCELLED IN TEAMSNAP' + (e.notes ? ' | ' + e.notes : ''))
      : e.notes;

    if (rowByUid.hasOwnProperty(e.uid)) {
      const r = rowByUid[e.uid];
      const row = data[r];

      // Preserve the user's Going? value; refresh everything else.
      if (isSkipped_(row[colIdx['Going?']])) skipMap[e.uid] = true;

      const fresh = row.slice();
      fresh[colIdx['Date']] = dateStr;
      fresh[colIdx['Day']] = dayStr;
      fresh[colIdx['Start']] = startStr;
      fresh[colIdx['End']] = endStr;
      fresh[colIdx['Player']] = e.player;
      fresh[colIdx['Team']] = e.team;
      fresh[colIdx['Type']] = e.type;
      fresh[colIdx['Event']] = e.title;
      fresh[colIdx['Location']] = e.location;
      fresh[colIdx['Notes']] = notes;
      fresh[colIdx['Link']] = e.link;

      if (rowChanged_(row, fresh)) updates.push({ row: r + 1, values: fresh });
    } else {
      if (e.start < cutoff) return;  // don't backfill past events

      const blank = new Array(lastCol).fill('');
      blank[colIdx['UID']] = e.uid;
      blank[colIdx['Date']] = dateStr;
      blank[colIdx['Day']] = dayStr;
      blank[colIdx['Start']] = startStr;
      blank[colIdx['End']] = endStr;
      blank[colIdx['Player']] = e.player;
      blank[colIdx['Team']] = e.team;
      blank[colIdx['Type']] = e.type;
      blank[colIdx['Event']] = e.title;
      blank[colIdx['Location']] = e.location;
      blank[colIdx['Notes']] = notes;
      blank[colIdx['Link']] = e.link;
      blank[colIdx['Going?']] = '';
      additions.push(blank);
    }
  });

  // Flag rows whose event disappeared from the feed entirely.
  for (let r = 1; r < data.length; r++) {
    const uid = String(data[r][colIdx['UID']]).trim();
    if (!uid || feedUids[uid]) continue;

    const existingNote = String(data[r][colIdx['Notes']]);
    if (existingNote.indexOf('REMOVED FROM TEAMSNAP') === -1) {
      const fresh = data[r].slice();
      fresh[colIdx['Notes']] = 'REMOVED FROM TEAMSNAP' +
        (existingNote ? ' | ' + existingNote : '');
      updates.push({ row: r + 1, values: fresh });
    }
  }

  updates.forEach(function (u) {
    sheet.getRange(u.row, 1, 1, lastCol).setValues([u.values]);
  });

  if (additions.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, lastCol)
         .setValues(additions);
  }

  // Keep the sheet in date order as new events are appended.
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol)
         .sort([{ column: colIdx['Date'] + 1, ascending: true },
                { column: colIdx['Player'] + 1, ascending: true }]);
  }

  log_('Sheet: ' + additions.length + ' added, ' + updates.length + ' updated, ' +
       Object.keys(skipMap).length + ' marked as not going.');

  return skipMap;
}

/** Appends any COLS header that isn't already present. */
function ensureColumns_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, width).getValues()[0]
                      .map(function (h) { return String(h).trim(); });

  const missing = COLS.filter(function (c) { return header.indexOf(c) === -1; });
  if (!missing.length) return;

  sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
  log_('Added missing sheet columns: ' + missing.join(', '));
}

/**
 * Google Sheets auto-detects text like "6:00 PM" or "2026-08-06" and
 * silently converts it to a real Date/Time value. A bare time then gets
 * anchored to Dec 30 1899, which is why the web app once showed that date.
 * Locking these columns to Plain Text stops Sheets from re-interpreting
 * what we write. Cheap to call every sync; only touches formatting.
 */
function lockTextColumns_(sheet, colIdx) {
  ['Date', 'Day', 'Start', 'End'].forEach(function (name) {
    const col = colIdx[name] + 1;
    if (col > 0) {
      sheet.getRange(1, col, sheet.getMaxRows(), 1).setNumberFormat('@');
    }
  });
}

/** Treats "no", "n", "skip", "false", and "x" as a skip. */
function isSkipped_(value) {
  const v = String(value).trim().toLowerCase();
  return v === 'no' || v === 'n' || v === 'skip' || v === 'false' || v === 'x';
}

function rowChanged_(oldRow, newRow) {
  for (let i = 0; i < newRow.length; i++) {
    if (String(oldRow[i]) !== String(newRow[i])) return true;
  }
  return false;
}

// ============================================================================
// CALENDAR SYNC
// ============================================================================

function syncCalendar_(feedEvents, skipMap) {
  const cal = getOrCreateCalendar_();

  const from = new Date();
  from.setDate(from.getDate() - CONFIG.LOOKBACK_DAYS);
  const to = new Date();
  to.setDate(to.getDate() + CONFIG.LOOKAHEAD_DAYS);

  // Map what's already on the calendar, keyed by the tag we set.
  const existing = {};
  cal.getEvents(from, to).forEach(function (ev) {
    const uid = ev.getTag(TAG_KEY);
    if (!uid) return;
    if (existing[uid]) {
      ev.deleteEvent();   // defensive: clean up any duplicate
    } else {
      existing[uid] = ev;
    }
  });

  const stats = { created: 0, updated: 0, removed: 0, skipped: 0 };
  const shouldExist = {};

  feedEvents.forEach(function (e) {
    if (e.start < from || e.start > to) return;

    const skip = skipMap[e.uid] === true;
    const drop = skip || e.cancelled;

    if (drop) {
      if (skip) stats.skipped++;
      if (existing[e.uid]) {
        existing[e.uid].deleteEvent();
        stats.removed++;
        log_('Removed: ' + e.title + ' on ' + e.start.toDateString() +
             (skip ? ' (you marked No)' : ' (cancelled in TeamSnap)'));
      }
      return;
    }

    shouldExist[e.uid] = true;
    const desc = buildDescription_(e);

    if (existing[e.uid]) {
      const ev = existing[e.uid];
      let changed = false;

      if (ev.getTitle() !== e.title) { ev.setTitle(e.title); changed = true; }
      if (ev.getLocation() !== e.location) { ev.setLocation(e.location); changed = true; }
      if (ev.getDescription() !== desc) { ev.setDescription(desc); changed = true; }
      if (applyColor_(ev, e.player)) changed = true;

      if (!e.allDay) {
        if (ev.getStartTime().getTime() !== e.start.getTime() ||
            ev.getEndTime().getTime() !== e.end.getTime()) {
          ev.setTime(e.start, e.end);
          changed = true;
        }
      }

      if (changed) {
        stats.updated++;
        log_('Updated: ' + e.title + ' -> ' + e.start.toLocaleString());
        Utilities.sleep(150);
      }
    } else {
      let ev;
      if (e.allDay) {
        ev = cal.createAllDayEvent(e.title, e.start,
              { location: e.location, description: desc });
      } else {
        ev = cal.createEvent(e.title, e.start, e.end,
              { location: e.location, description: desc });
      }
      ev.setTag(TAG_KEY, e.uid);

      applyColor_(ev, e.player);

      stats.created++;
      log_('Created: ' + e.title + ' on ' + e.start.toDateString());
      Utilities.sleep(150);
    }
  });

  // Anything tagged on the calendar that the feed no longer contains: remove it.
  Object.keys(existing).forEach(function (uid) {
    if (shouldExist[uid] || skipMap[uid]) return;
    const stillInFeed = feedEvents.some(function (e) { return e.uid === uid; });
    if (!stillInFeed) {
      existing[uid].deleteEvent();
      stats.removed++;
      log_('Removed (no longer in feed): ' + uid);
    }
  });

  return stats;
}

function buildDescription_(e) {
  const parts = [];
  if (e.player) parts.push('Player: ' + e.player);
  if (e.team) parts.push('Team: ' + e.team);
  if (e.squad && playerHasMultipleTeams_(e.player)) parts.push('Squad: ' + e.squad);
  if (e.arrival) parts.push('Arrive: ' + e.arrival);
  if (e.uniform) parts.push('Uniform: ' + titleCase_(e.uniform));
  if (e.feedNotes) parts.push('Notes: ' + e.feedNotes);
  if (e.link) {
    parts.push('');
    parts.push('Open in TeamSnap: ' + e.link);
  }

  // --- CARPOOL HOOK ---------------------------------------------------
  // Appends the block between its own markers, or nothing at all when no
  // leg has been assigned. Because the regular sync compares the whole
  // description before writing, a carpool change made in the web app
  // propagates here on the next pass without any extra plumbing.
  try {
    if (typeof carpoolBlock_ === 'function') {
      const cp = carpoolBlock_(e.uid);
      if (cp) { parts.push(''); parts.push(cp); }
    }
  } catch (err) {
    log_('Carpool block failed for ' + e.uid + ': ' + err);
  }
  // --------------------------------------------------------------------

  parts.push('');
  parts.push('Not going? Update it here: ' + CONFIG.WEB_APP_URL);
  return parts.join('\n');
}

/** "WHITE" -> "White". Leaves multi-word values like "WHITE / BLACK" readable too. */
function titleCase_(s) {
  return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

function getOrCreateCalendar_() {
  const found = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
  if (found && found.length > 0) return found[0];

  const cal = CalendarApp.createCalendar(CONFIG.CALENDAR_NAME);
  cal.setDescription('Auto-synced from TeamSnap ONE. Managed by TeamSnapSync script.');
  log_('Created new calendar: ' + CONFIG.CALENDAR_NAME);
  return cal;
}

// ============================================================================
// UTILITIES
// ============================================================================

function log_(msg) {
  if (CONFIG.VERBOSE) Logger.log(msg);
}

/**
 * Diagnostic helper. Shows how each feed event will be renamed, without
 * touching your calendar or sheet. Run this first if titles look wrong.
 */
function previewFeed() {
  const events = fetchAndParseFeed_();
  const tz = CONFIG.TIME_ZONE;
  Logger.log('=== ' + events.length + ' events in feed ===');
  events
    .sort(function (a, b) { return a.start - b.start; })
    .slice(0, 40)
    .forEach(function (e) {
      Logger.log(
        Utilities.formatDate(e.start, tz, 'yyyy-MM-dd EEE h:mm a') + '  ' +
        e.title +
        (e.notes ? '   [' + e.notes + ']' : '') +
        (e.cancelled ? '  *** CANCELLED ***' : '')
      );
    });
}

/**
 * Nuclear option: deletes every event this script created, then re-syncs
 * from scratch. Rarely needed - the regular syncNow() already catches and
 * fixes title, time, location, and description changes in place. Only
 * reach for this after something structural changes (a player renamed,
 * the title format itself changed) where every event genuinely needs to
 * be rebuilt rather than patched.
 *
 * Your "Going?" column and every carpool assignment are preserved, since
 * both live in the spreadsheet rather than on the calendar.
 *
 * A short pause between deletes avoids Google's own rate limit on rapid
 * calendar writes ("creating or deleting too many events in a short
 * time"), which a full-speed loop over 100+ events can trip.
 */
function resetCalendar() {
  const cal = getOrCreateCalendar_();
  const from = new Date();
  from.setDate(from.getDate() - CONFIG.LOOKBACK_DAYS);
  const to = new Date();
  to.setDate(to.getDate() + CONFIG.LOOKAHEAD_DAYS);

  let n = 0;
  cal.getEvents(from, to).forEach(function (ev) {
    if (ev.getTag(TAG_KEY)) {
      ev.deleteEvent();
      n++;
      Utilities.sleep(300);
    }
  });
  log_('Deleted ' + n + ' synced events. Re-syncing...');
  syncNow();
}
