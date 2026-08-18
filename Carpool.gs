/**
 * Carpool coordination — server side.
 *
 * Add this as a THIRD script file in the same Apps Script project as
 * TeamSnapSync.gs and WebApp.gs. It reuses CONFIG, COLS, TAG_KEY, log_(),
 * and getOrCreateCalendar_() from TeamSnapSync.gs, so all three must live
 * in the same project.
 *
 * Nothing here touches the schedule sync. Four new tabs are created in the
 * existing spreadsheet; the "Kids Soccer Schedules" tab is never written to
 * by this file.
 *
 * FIRST RUN:
 *   1. Paste this file in.
 *   2. Run setupCarpoolTabs() once. It creates the four tabs and seeds
 *      Teams from CONFIG.PLAYERS.
 *   3. Open the spreadsheet, fill in Families by hand (or wait for the UI).
 *   4. Run previewCarpoolNote() to see what a calendar note will look like.
 *
 * IDs used throughout:
 *   DriverID   'scott' / 'steph'        -> one of us
 *              'F3:1' / 'F3:2'          -> parent 1 or 2 of family F3
 *   RiderID    'k:Maddie'               -> one of our kids
 *              'F3'                     -> the player in family F3
 *   Direction  'out'  = ride to the event
 *              'back' = ride home from the event
 */

// ============================================================================
// CONFIG
// ============================================================================

const CARPOOL = {
  // Us. Add or remove names freely; the DriverID is the lowercased name.
  US: ['Scott', 'Steph'],

  // Our kids, in the order they should appear in pickers.
  KIDS: ['Maddie', 'Emily', 'Cooper'],

  // Tab names.
  TAB_TEAMS:    'Teams',
  TAB_FAMILIES: 'Families',
  TAB_DEFAULTS: 'Defaults',
  TAB_CARPOOL:  'Carpool',

  // How far ahead applyDefaults / the UI look.
  WINDOW_DAYS: 21
};

const TABS = {
  Teams:    ['TeamID', 'Team', 'Kid', 'Display'],
  Families: ['FamilyID', 'TeamID', 'Player', 'P1Name', 'P1Phone',
             'P2Name', 'P2Phone', 'Address', 'Notes'],
  Defaults: ['RuleID', 'TeamID', 'Type', 'Weekday', 'Direction', 'DriverID'],
  Carpool:  ['UID', 'Direction', 'DriverID', 'RiderIDs', 'Updated']
};

// Markers in the calendar description. Everything between them is ours to
// rewrite; everything outside is left exactly as TeamSnapSync built it.
const CP_START = '--- CARPOOL ---';
const CP_END   = '--- END CARPOOL ---';

// ============================================================================
// ONE-TIME SETUP
// ============================================================================

/**
 * Creates the four tabs if they're missing and seeds Teams from
 * CONFIG.PLAYERS. Safe to re-run: existing rows are never overwritten,
 * and only genuinely new teams get appended.
 */
function setupCarpoolTabs() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  Object.keys(TABS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      log_('Created tab: ' + name);
    }
    const cols = TABS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    }
    // Phone columns must stay text or Sheets mangles them.
    ['P1Phone', 'P2Phone'].forEach(function (c) {
      const i = cols.indexOf(c);
      if (i !== -1) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });

  seedTeams_();
  log_('Carpool tabs ready.');
}

/** Adds a Teams row for any CONFIG.PLAYERS entry not already listed. */
function seedTeams_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CARPOOL.TAB_TEAMS);
  const existing = readTab_(CARPOOL.TAB_TEAMS);
  const have = {};
  existing.forEach(function (r) { have[String(r.Team).trim()] = true; });

  const rows = [];
  let n = existing.length;
  Object.keys(CONFIG.PLAYERS).forEach(function (team) {
    if (have[team]) return;
    n++;
    const kid = CONFIG.PLAYERS[team];
    rows.push(['T' + n, team, kid, suggestDisplay_(team, kid)]);
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, TABS.Teams.length).setValues(rows);
    log_('Seeded ' + rows.length + ' team(s).');
  }
}

/** 'Panthers U12G Black' + 'Emily' -> "Emily's U12 Black" */
function suggestDisplay_(team, kid) {
  const m = team.match(/U(\d+)([BG])\s*(Black|Gold)?/);
  if (!m) return kid + "'s team";
  return kid + "'s U" + m[1] + (m[3] ? ' ' + m[3] : '');
}

// ============================================================================
// TAB IO
// ============================================================================

/** Reads a tab into an array of plain objects keyed by header name. */
function readTab_(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];

  const cols = TABS[name];
  const data = sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), cols.length))
                 .getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });

  const out = [];
  for (let r = 1; r < data.length; r++) {
    const rec = { _row: r + 1 };
    let blank = true;
    header.forEach(function (h, i) {
      if (!h) return;
      const v = data[r][i];
      rec[h] = (v instanceof Date) ? v : String(v == null ? '' : v).trim();
      if (rec[h] !== '') blank = false;
    });
    if (!blank) out.push(rec);
  }
  return out;
}

/** Appends or updates one row, matched on the tab's first column. */
function upsertRow_(name, keyValue, values) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(name);
  const cols = TABS[name];
  if (!sh) throw new Error('Missing tab: ' + name + '. Run setupCarpoolTabs().');

  const rows = readTab_(name);
  const keyCol = cols[0];
  let target = null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][keyCol] === keyValue) { target = rows[i]._row; break; }
  }

  const line = cols.map(function (c) { return values.hasOwnProperty(c) ? values[c] : ''; });
  if (target) sh.getRange(target, 1, 1, cols.length).setValues([line]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length).setValues([line]);
}

/** Deletes rows whose first column matches. Returns how many went. */
function deleteRows_(name, keyValue) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(name);
  const rows = readTab_(name).filter(function (r) { return r[TABS[name][0]] === keyValue; });
  rows.sort(function (a, b) { return b._row - a._row; });  // bottom up
  rows.forEach(function (r) { sh.deleteRow(r._row); });
  return rows.length;
}

// ============================================================================
// CARPOOL STORAGE
// ============================================================================

/** Composite key for a leg. */
function legKey_(uid, dir) { return uid + '|' + dir; }

/**
 * Every carpool assignment, keyed 'uid|out' / 'uid|back'.
 *
 * A row that exists with a blank DriverID means "deliberately cleared" and
 * defaults will NOT refill it. No row at all means "never touched" and
 * defaults are free to fill it. That distinction is the whole reason
 * clearing a leg still writes a row.
 */
function readCarpoolMap_() {
  const map = {};
  readCarpoolRows_().forEach(function (r) {
    map[legKey_(r.UID, r.Direction)] = {
      driver: r.DriverID || '',
      riders: r.RiderIDs ? r.RiderIDs.split(',').map(function (s) { return s.trim(); })
                                     .filter(String) : []
    };
  });
  return map;
}

/** Raw Carpool rows. Direction is normalised to out/back. */
function readCarpoolRows_() {
  return readTab_(CARPOOL.TAB_CARPOOL).map(function (r) {
    r.Direction = String(r.Direction).toLowerCase() === 'back' ? 'back' : 'out';
    return r;
  });
}

/**
 * Writes one leg, then refreshes that event's calendar note immediately so
 * the change shows up without waiting for the 15-minute sync.
 */
function setCarpoolLeg(uid, dir, driverId, riderIds) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('The schedule is syncing right now. Try again in a moment.');
  }

  try {
    dir = (String(dir).toLowerCase() === 'back') ? 'back' : 'out';
    const riders = (riderIds || []).filter(String);

    // The Carpool tab has a composite key, so upsertRow_ (single-column
    // match) won't do. Find the exact uid+direction row by hand.
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh = ss.getSheetByName(CARPOOL.TAB_CARPOOL);
    if (!sh) throw new Error('Missing Carpool tab. Run setupCarpoolTabs().');

    const rows = readCarpoolRows_();
    let target = null;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].UID === uid && rows[i].Direction === dir) { target = rows[i]._row; break; }
    }

    const line = [uid, dir, driverId || '', riders.join(','), new Date()];
    if (target) sh.getRange(target, 1, 1, TABS.Carpool.length).setValues([line]);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, TABS.Carpool.length).setValues([line]);

    SpreadsheetApp.flush();

    let calendarUpdated = true;
    try {
      updateEventNote_(uid);
    } catch (err) {
      calendarUpdated = false;   // next sync will catch up
      log_('Immediate note update failed for ' + uid + ': ' + err);
    }

    return { uid: uid, dir: dir, calendarUpdated: calendarUpdated };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// DRIVERS AND RIDERS
// ============================================================================

/** Every selectable driver, us first, then each family's parents. */
function buildDrivers_(families) {
  const out = CARPOOL.US.map(function (n) {
    return { id: n.toLowerCase(), name: n, us: true, teamId: '' };
  });
  families.forEach(function (f) {
    if (f.P1Name) out.push({ id: f.FamilyID + ':1', name: f.P1Name,
                             phone: f.P1Phone, us: false, teamId: f.TeamID, famId: f.FamilyID });
    if (f.P2Name) out.push({ id: f.FamilyID + ':2', name: f.P2Name,
                             phone: f.P2Phone, us: false, teamId: f.TeamID, famId: f.FamilyID });
  });
  return out;
}

function driverName_(id, drivers) {
  for (let i = 0; i < drivers.length; i++) if (drivers[i].id === id) return drivers[i].name;
  return '';
}

function isUsDriver_(id) {
  return CARPOOL.US.some(function (n) { return n.toLowerCase() === String(id).toLowerCase(); });
}

/** 'k:Maddie' -> 'Maddie'.  'F3' -> that family's player name. */
function riderName_(id, families) {
  if (String(id).indexOf('k:') === 0) return String(id).substring(2);
  for (let i = 0; i < families.length; i++) {
    if (families[i].FamilyID === id) return families[i].Player;
  }
  return id;
}

function familyById_(id, families) {
  for (let i = 0; i < families.length; i++) if (families[i].FamilyID === id) return families[i];
  return null;
}

// ============================================================================
// THE CALENDAR NOTE
// ============================================================================

/**
 * The carpool block for one event, markers included. Returns '' when
 * neither leg has been touched, so events without a carpool keep exactly
 * the description they have today.
 */
function carpoolBlock_(uid) {
  const map = readCarpoolMap_();
  const out = map[legKey_(uid, 'out')];
  const back = map[legKey_(uid, 'back')];
  if (!out && !back) return '';
  if ((!out || !out.driver) && (!back || !back.driver)) return '';

  const families = readTab_(CARPOOL.TAB_FAMILIES);
  const drivers = buildDrivers_(families);

  const lines = [CP_START];
  lines.push(legLine_('TO:   ', out, families, drivers));
  lines.push(legLine_('HOME: ', back, families, drivers));
  lines.push('');
  lines.push('Change it: ' + CONFIG.WEB_APP_URL);
  lines.push(CP_END);
  return lines.join('\n');
}

function legLine_(label, leg, families, drivers) {
  if (!leg || !leg.driver) return label + '** no driver assigned **';

  const name = driverName_(leg.driver, drivers) || leg.driver;
  const mine = leg.riders.filter(function (r) { return r.indexOf('k:') === 0; })
                         .map(function (r) { return r.substring(2); });
  const guests = leg.riders.filter(function (r) { return r.indexOf('k:') !== 0; })
                           .map(function (r) { return familyById_(r, families); })
                           .filter(Boolean);

  let s = label + name;
  s += '  (' + (mine.length ? mine.join(', ') : 'none of ours') +
       (guests.length ? ' +' + guests.length : '') + ')';

  guests.forEach(function (f) {
    s += '\n        ' + f.Player + ' — ' + f.Address + (f.Notes ? ' (' + f.Notes + ')' : '');
    s += '\n        ' + f.P1Name + ' ' + f.P1Phone +
         (f.P2Name ? '  /  ' + f.P2Name + ' ' + f.P2Phone : '');
  });
  return s;
}

/**
 * Rewrites just the carpool block on one calendar event, leaving every
 * other line alone. Used for the immediate update after an assignment;
 * the regular sync goes through buildDescription_ instead.
 */
function updateEventNote_(uid) {
  const row = findScheduleRow_(uid);
  if (!row) return false;

  const cal = getOrCreateCalendar_();
  const from = new Date(row.date.getTime() - 36 * 60 * 60 * 1000);
  const to   = new Date(row.date.getTime() + 36 * 60 * 60 * 1000);

  let ev = null;
  cal.getEvents(from, to).forEach(function (c) {
    if (c.getTag(TAG_KEY) === uid) ev = c;
  });
  if (!ev) return false;   // not going, or not synced yet

  const current = ev.getDescription() || '';
  const block = carpoolBlock_(uid);
  const next = spliceBlock_(current, block);

  if (next !== current) {
    ev.setDescription(next);
    log_('Carpool note updated: ' + ev.getTitle());
    return true;
  }
  return false;
}

/** Replaces the marked region, appends it if absent, removes it if block is ''. */
function spliceBlock_(text, block) {
  const s = text.indexOf(CP_START);
  const e = text.indexOf(CP_END);

  if (s !== -1 && e !== -1 && e > s) {
    const before = text.substring(0, s).replace(/\s+$/, '');
    const after = text.substring(e + CP_END.length).replace(/^\s+/, '');
    if (!block) return (before + (after ? '\n\n' + after : '')).trim();
    return (before + '\n\n' + block + (after ? '\n\n' + after : '')).trim();
  }
  if (!block) return text;
  return (text.replace(/\s+$/, '') + '\n\n' + block).trim();
}

/** Looks up one schedule row by UID. Returns { date, team, type, day, title }. */
function findScheduleRow_(uid) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() < 2) return null;

  const lastCol = Math.max(sh.getLastColumn(), COLS.length);
  const data = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  COLS.forEach(function (c) { idx[c] = header.indexOf(c); });

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idx['UID']]).trim() !== uid) continue;
    const dv = data[r][idx['Date']];
    const d = (dv instanceof Date)
      ? new Date(dv.getFullYear(), dv.getMonth(), dv.getDate())
      : parseYmd_(String(dv));
    if (!d) return null;
    return {
      uid: uid,
      date: d,
      day: String(data[r][idx['Day']] || ''),
      team: String(data[r][idx['Team']] || ''),
      type: String(data[r][idx['Type']] || ''),
      title: String(data[r][idx['Event']] || '')
    };
  }
  return null;
}

// ============================================================================
// DEFAULTS
// ============================================================================

/**
 * Fills every untouched leg on upcoming events from the Defaults tab.
 *
 * "Untouched" means no Carpool row exists for that uid+direction. A row
 * with a blank driver was cleared on purpose and stays cleared. Nothing
 * already assigned is ever overwritten.
 *
 * Called at the end of syncNow(), and by the Setup screen's button.
 */
function applyDefaultsToUpcoming() {
  const rules = readTab_(CARPOOL.TAB_DEFAULTS);
  if (!rules.length) return { filled: 0 };

  const teams = readTab_(CARPOOL.TAB_TEAMS);
  const teamIdByName = {};
  teams.forEach(function (t) { teamIdByName[t.Team] = t.TeamID; });

  const map = readCarpoolMap_();
  const events = readUpcomingSchedule_();

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CARPOOL.TAB_CARPOOL);
  const additions = [];
  const touched = {};

  events.forEach(function (e) {
    const teamId = teamIdByName[e.team];
    if (!teamId) return;

    ['out', 'back'].forEach(function (dir) {
      if (map.hasOwnProperty(legKey_(e.uid, dir))) return;   // touched already

      const rule = matchRule_(rules, teamId, e.type, e.day, dir);
      if (!rule || !rule.DriverID) return;

      additions.push([e.uid, dir, rule.DriverID, 'k:' + e.player, new Date()]);
      touched[e.uid] = true;
    });
  });

  if (additions.length) {
    sh.getRange(sh.getLastRow() + 1, 1, additions.length, TABS.Carpool.length)
      .setValues(additions);
    SpreadsheetApp.flush();
    Object.keys(touched).forEach(function (uid) {
      try { updateEventNote_(uid); } catch (err) { log_('Note update failed: ' + uid); }
    });
  }

  log_('Defaults filled ' + additions.length + ' leg(s).');
  return { filled: additions.length };
}

/** Blank Weekday matches any day. Type must match the sheet's Type value. */
function matchRule_(rules, teamId, type, day, dir) {
  let best = null;
  rules.forEach(function (r) {
    if (r.TeamID !== teamId) return;
    if (String(r.Direction).toLowerCase() !== dir) return;
    if (r.Type && r.Type !== type) return;
    if (r.Weekday && r.Weekday !== day) return;
    // A rule naming a specific weekday beats a catch-all.
    if (!best || (r.Weekday && !best.Weekday)) best = r;
  });
  return best;
}

/** Upcoming schedule rows, skipping anything removed from TeamSnap. */
function readUpcomingSchedule_(days) {
  const windowDays = days || CARPOOL.WINDOW_DAYS;
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() < 2) return [];

  const tz = CONFIG.TIME_ZONE;
  const lastCol = Math.max(sh.getLastColumn(), COLS.length);
  const data = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });
  const idx = {};
  COLS.forEach(function (c) { idx[c] = header.indexOf(c); });

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const horizon = new Date(today.getTime());
  horizon.setDate(horizon.getDate() + windowDays);

  const out = [];
  for (let r = 1; r < data.length; r++) {
    const uid = String(data[r][idx['UID']]).trim();
    if (!uid) continue;

    const dv = data[r][idx['Date']];
    const d = (dv instanceof Date)
      ? new Date(dv.getFullYear(), dv.getMonth(), dv.getDate())
      : parseYmd_(String(dv));
    if (!d || d < today || d > horizon) continue;

    const notes = String(data[r][idx['Notes']] || '');
    if (notes.indexOf('REMOVED FROM TEAMSNAP') !== -1) continue;
    if (isSkipped_(data[r][idx['Going?']])) continue;   // not going, no ride needed

    out.push({
      uid: uid,
      dateKey: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      dateLabel: Utilities.formatDate(d, tz, 'EEE d MMM'),
      day: String(data[r][idx['Day']] || ''),
      start: cellToTimeString_(data[r][idx['Start']], tz),
      end: cellToTimeString_(data[r][idx['End']], tz),
      player: String(data[r][idx['Player']] || ''),
      team: String(data[r][idx['Team']] || ''),
      type: String(data[r][idx['Type']] || ''),
      title: String(data[r][idx['Event']] || ''),
      location: String(data[r][idx['Location']] || ''),
      cancelled: notes.indexOf('CANCELLED') !== -1
    });
  }
  return out;
}

// ============================================================================
// WEB APP ENDPOINTS
// ============================================================================

/** Everything the Carpool and Setup screens need, in one round trip. */
function getCarpoolData(days) {
  const families = readTab_(CARPOOL.TAB_FAMILIES);
  const teams = readTab_(CARPOOL.TAB_TEAMS);
  const map = readCarpoolMap_();
  const events = readUpcomingSchedule_(days);

  const teamIdByName = {};
  teams.forEach(function (t) { teamIdByName[t.Team] = t.TeamID; });

  events.forEach(function (e) {
    e.teamId = teamIdByName[e.team] || '';
    e.out = map[legKey_(e.uid, 'out')] || { driver: '', riders: [] };
    e.back = map[legKey_(e.uid, 'back')] || { driver: '', riders: [] };
  });

  return {
    events: events,
    teams: teams,
    families: families,
    defaults: readTab_(CARPOOL.TAB_DEFAULTS),
    drivers: buildDrivers_(families),
    kids: CARPOOL.KIDS
  };
}

/** Creates or updates one family. Pass FamilyID '' to create. */
function saveFamily(rec) {
  if (!rec || !rec.TeamID) throw new Error('A family needs a team.');
  if (!rec.Player) throw new Error('A family needs a player name.');

  let id = rec.FamilyID;
  if (!id) {
    const existing = readTab_(CARPOOL.TAB_FAMILIES);
    let n = existing.length + 1;
    const used = {};
    existing.forEach(function (f) { used[f.FamilyID] = true; });
    while (used['F' + n]) n++;
    id = 'F' + n;
  }

  upsertRow_(CARPOOL.TAB_FAMILIES, id, {
    FamilyID: id, TeamID: rec.TeamID, Player: rec.Player,
    P1Name: rec.P1Name || '', P1Phone: rec.P1Phone || '',
    P2Name: rec.P2Name || '', P2Phone: rec.P2Phone || '',
    Address: rec.Address || '', Notes: rec.Notes || ''
  });
  return id;
}

/**
 * Removes a family and scrubs it out of every leg it appears in, so no
 * assignment is left pointing at someone who isn't on the roster.
 */
function deleteFamily(familyId) {
  const rows = readCarpoolRows_();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CARPOOL.TAB_CARPOOL);
  const touched = {};

  rows.forEach(function (r) {
    let driver = r.DriverID;
    let riders = r.RiderIDs ? r.RiderIDs.split(',').map(function (s) { return s.trim(); }) : [];
    let changed = false;

    if (driver && driver.indexOf(familyId + ':') === 0) { driver = ''; changed = true; }
    const kept = riders.filter(function (x) { return x !== familyId; });
    if (kept.length !== riders.length) { riders = kept; changed = true; }

    if (changed) {
      sh.getRange(r._row, 1, 1, TABS.Carpool.length)
        .setValues([[r.UID, r.Direction, driver, riders.join(','), new Date()]]);
      touched[r.UID] = true;
    }
  });

  // Any default rule naming one of their parents goes back to "ask each time".
  readTab_(CARPOOL.TAB_DEFAULTS).forEach(function (d) {
    if (d.DriverID && d.DriverID.indexOf(familyId + ':') === 0) {
      upsertRow_(CARPOOL.TAB_DEFAULTS, d.RuleID, {
        RuleID: d.RuleID, TeamID: d.TeamID, Type: d.Type,
        Weekday: d.Weekday, Direction: d.Direction, DriverID: ''
      });
    }
  });

  const n = deleteRows_(CARPOOL.TAB_FAMILIES, familyId);
  SpreadsheetApp.flush();
  Object.keys(touched).forEach(function (uid) {
    try { updateEventNote_(uid); } catch (err) { log_('Note update failed: ' + uid); }
  });
  return n;
}

/** Creates or updates one default rule. Pass RuleID '' to create. */
function saveDefaultRule(rec) {
  if (!rec || !rec.TeamID) throw new Error('A rule needs a team.');
  const dir = String(rec.Direction).toLowerCase() === 'back' ? 'back' : 'out';

  let id = rec.RuleID;
  if (!id) {
    const existing = readTab_(CARPOOL.TAB_DEFAULTS);
    let n = existing.length + 1;
    const used = {};
    existing.forEach(function (d) { used[d.RuleID] = true; });
    while (used['R' + n]) n++;
    id = 'R' + n;
  }

  upsertRow_(CARPOOL.TAB_DEFAULTS, id, {
    RuleID: id, TeamID: rec.TeamID, Type: rec.Type || '',
    Weekday: rec.Weekday || '', Direction: dir, DriverID: rec.DriverID || ''
  });
  return id;
}

function deleteDefaultRule(ruleId) {
  return deleteRows_(CARPOOL.TAB_DEFAULTS, ruleId);
}

/** Plain-text summary for pasting into a group chat. */
function carpoolTextForDate(dateKey) {
  const data = getCarpoolData();
  const day = data.events.filter(function (e) { return e.dateKey === dateKey; });
  if (!day.length) return 'Nothing scheduled that day.';

  const lines = ['Panthers carpool — ' + day[0].dateLabel];
  day.forEach(function (e) {
    const t = teamDisplay_(e.teamId, data.teams) || e.player;
    lines.push('');
    lines.push(t + ' · ' + e.start + ' · ' + (e.location || '').split(',')[0]);
    lines.push('  there: ' + legText_(e.out, data));
    lines.push('  home: ' + legText_(e.back, data));
  });
  return lines.join('\n');
}

function legText_(leg, data) {
  if (!leg || !leg.driver) return 'NEEDS A DRIVER';
  let s = driverName_(leg.driver, data.drivers) || leg.driver;
  if (leg.riders.length) {
    s += ' w/ ' + leg.riders.map(function (r) {
      return riderName_(r, data.families);
    }).join(', ');
  }
  return s;
}

function teamDisplay_(teamId, teams) {
  for (let i = 0; i < teams.length; i++) {
    if (teams[i].TeamID === teamId) return teams[i].Display || teams[i].Team;
  }
  return '';
}

// ============================================================================
// DIAGNOSTICS
// ============================================================================

/** Prints the carpool block for the next event that has one. */
function previewCarpoolNote() {
  const events = readUpcomingSchedule_();
  if (!events.length) { Logger.log('No upcoming events.'); return; }

  let shown = 0;
  events.forEach(function (e) {
    const block = carpoolBlock_(e.uid);
    if (!block || shown >= 3) return;
    shown++;
    Logger.log('=== ' + e.title + ' — ' + e.dateLabel + ' ' + e.start + ' ===\n' + block + '\n');
  });
  if (!shown) Logger.log('No events have a carpool assignment yet. ' +
                         'Add a Defaults rule and run applyDefaultsToUpcoming().');
}

/** Sanity check: tabs exist, teams line up with CONFIG.PLAYERS, IDs resolve. */
function carpoolSelfTest() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const problems = [];

  Object.keys(TABS).forEach(function (t) {
    if (!ss.getSheetByName(t)) problems.push('Missing tab: ' + t);
  });
  if (problems.length) { Logger.log(problems.join('\n')); return; }

  const teams = readTab_(CARPOOL.TAB_TEAMS);
  const families = readTab_(CARPOOL.TAB_FAMILIES);
  const rules = readTab_(CARPOOL.TAB_DEFAULTS);
  const drivers = buildDrivers_(families);
  const teamIds = {};
  teams.forEach(function (t) { teamIds[t.TeamID] = true; });

  Object.keys(CONFIG.PLAYERS).forEach(function (name) {
    if (!teams.some(function (t) { return t.Team === name; })) {
      problems.push('CONFIG.PLAYERS has "' + name + '" but Teams does not. Run setupCarpoolTabs().');
    }
  });
  families.forEach(function (f) {
    if (!teamIds[f.TeamID]) problems.push('Family ' + f.FamilyID + ' points at unknown team ' + f.TeamID);
    if (!f.P1Name && !f.P2Name) problems.push('Family ' + f.FamilyID + ' has no parents, so nobody can drive.');
  });
  rules.forEach(function (r) {
    if (!teamIds[r.TeamID]) problems.push('Rule ' + r.RuleID + ' points at unknown team ' + r.TeamID);
    if (r.DriverID && !drivers.some(function (d) { return d.id === r.DriverID; })) {
      problems.push('Rule ' + r.RuleID + ' names driver ' + r.DriverID + ' who no longer exists.');
    }
  });
  readCarpoolRows_().forEach(function (c) {
    if (c.DriverID && !drivers.some(function (d) { return d.id === c.DriverID; })) {
      problems.push('Leg ' + c.UID + '|' + c.Direction + ' names driver ' + c.DriverID + ' who no longer exists.');
    }
  });

  Logger.log(problems.length
    ? 'Problems found:\n' + problems.join('\n')
    : 'All good. ' + teams.length + ' teams, ' + families.length + ' families, ' +
      rules.length + ' rules, ' + drivers.length + ' drivers.');
}
