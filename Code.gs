const APP_TITLE = 'JMAC 28-Day Training Log';
const SESSION_TTL_SECONDS = 21600;

const SHEETS = {
  PROGRAM: 'Program',
  WARMUPS: 'Warmups',
  MEALS: 'Meals',
  DAILY: 'Daily Targets',
  SAUNA: 'Sauna',
  NOTES: 'Upgrade Notes',
  LOG: 'Workout Log',
  SESSIONS: 'Session Summary',
  SETTINGS: 'Settings',
  PROFILES: 'Profiles',
  CHECKINS: 'Check-ins'
};

const HEADERS = {
  PROGRAM: ['Day','Block','Slot','Exercise','Notes','Week 1','Week 2','Week 3','Week 4','Suggested Load','Order','Input Type','Rest Seconds','Video URL'],
  WARMUPS: ['Day','Drill','Purpose','Dose','Order'],
  MEALS: ['Day','Meal','Time','Food','Order'],
  DAILY: ['Day','Item','Target','Order'],
  SAUNA: ['Day','Item','Detail','Order'],
  NOTES: ['Day','Note'],
  LOG: ['Timestamp','Athlete','Date','Day','Week','Block','Slot','Exercise','Set','Target','Metric Type','Reps','Weight','Distance','Duration Seconds','Rounds','Breaths','RPE','Completed','Notes','Session ID'],
  SESSIONS: ['Timestamp','Athlete','Date','Day','Week','Session ID','Duration Seconds','Completed Sets','Total Sets','Training Volume','Session Notes'],
  SETTINGS: ['Key','Value'],
  PROFILES: ['Athlete','PIN Hash','Active','Role'],
  CHECKINS: ['Timestamp','Athlete','Date','Bodyweight','Waist','Sleep Hours','Steps','Calories','Protein','Water','Energy','Soreness','Notes']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .setFaviconUrl('https://drive.google.com/uc?export=download&id=11Ijp1h2BH2SIKQcVlzdYPySN9CXUSOYO&filename=icon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, SHEETS.PROGRAM, HEADERS.PROGRAM);
  ensureSheetWithHeaders_(ss, SHEETS.WARMUPS, HEADERS.WARMUPS);
  ensureSheetWithHeaders_(ss, SHEETS.MEALS, HEADERS.MEALS);
  ensureSheetWithHeaders_(ss, SHEETS.DAILY, HEADERS.DAILY);
  ensureSheetWithHeaders_(ss, SHEETS.SAUNA, HEADERS.SAUNA);
  ensureSheetWithHeaders_(ss, SHEETS.NOTES, HEADERS.NOTES);
  ensureSheetWithHeaders_(ss, SHEETS.LOG, HEADERS.LOG);
  ensureSheetWithHeaders_(ss, SHEETS.SESSIONS, HEADERS.SESSIONS);
  ensureSheetWithHeaders_(ss, SHEETS.SETTINGS, HEADERS.SETTINGS);
  ensureSheetWithHeaders_(ss, SHEETS.PROFILES, HEADERS.PROFILES);
  ensureSheetWithHeaders_(ss, SHEETS.CHECKINS, HEADERS.CHECKINS);

  seedSettings_();
  seedDefaultProfile_();
  if (ss.getSheetByName(SHEETS.PROGRAM).getLastRow() < 2) syncAllProgramData_();
  return getLoginBootstrap();
}

function getLoginBootstrap() {
  setupStructureOnly_();
  return {
    title: readSetting_('App Title') || APP_TITLE,
    profiles: getActiveProfileNames_()
  };
}

function login(athlete, pin) {
  setupStructureOnly_();
  const profile = findProfile_(String(athlete || '').trim());
  if (!profile || profile.Active.toUpperCase() !== 'TRUE') throw new Error('Profile is not active.');
  if (profile['PIN Hash'] !== hashPin_(String(pin || ''))) throw new Error('Incorrect PIN.');

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session:' + token, JSON.stringify({athlete: profile.Athlete, role: profile.Role || 'Athlete'}), SESSION_TTL_SECONDS);
  return {token: token, bootstrap: getBootstrap(token)};
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session:' + token);
  return true;
}

function getBootstrap(token) {
  const session = requireSession_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const program = enrichProgram_(readObjects_(ss.getSheetByName(SHEETS.PROGRAM)));
  return {
    title: readSetting_('App Title') || APP_TITLE,
    athlete: session.athlete,
    role: session.role,
    currentWeek: Number(readSetting_('Current Week') || 1),
    defaultDay: readSetting_('Default Day') || 'Monday',
    spreadsheetUrl: ss.getUrl(),
    program: program,
    warmups: readObjects_(ss.getSheetByName(SHEETS.WARMUPS)),
    meals: readObjects_(ss.getSheetByName(SHEETS.MEALS)),
    dailyTargets: readObjects_(ss.getSheetByName(SHEETS.DAILY)),
    sauna: readObjects_(ss.getSheetByName(SHEETS.SAUNA)),
    upgradeNotes: readObjects_(ss.getSheetByName(SHEETS.NOTES)),
    previous: getPreviousPerformance_(session.athlete),
    dashboard: getDashboard_(session.athlete),
    history: getRecentHistory(token, 40),
    profiles: session.role.toLowerCase() === 'coach' ? listProfiles(token) : []
  };
}

function saveWorkout(token, payload) {
  const session = requireSession_(token);
  if (!payload || !Array.isArray(payload.entries)) throw new Error('No workout entries received.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  const sessionSheet = ss.getSheetByName(SHEETS.SESSIONS);
  const now = new Date();
  const sessionId = payload.sessionId || Utilities.getUuid();
  const date = payload.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const rows = payload.entries.filter(e => e.exercise).map(e => [
    now,
    session.athlete,
    date,
    payload.day || '',
    Number(payload.week || 1),
    e.block || '',
    e.slot || '',
    e.exercise,
    Number(e.set || 1),
    e.target || '',
    e.metricType || 'REPS_WEIGHT',
    numberOrBlank_(e.reps),
    numberOrBlank_(e.weight),
    numberOrBlank_(e.distance),
    numberOrBlank_(e.durationSeconds),
    numberOrBlank_(e.rounds),
    numberOrBlank_(e.breaths),
    numberOrBlank_(e.rpe),
    Boolean(e.completed),
    e.notes || '',
    sessionId
  ]);

  if (!rows.length) throw new Error('Nothing to save.');
  const personalRecords = detectSessionPRs_(session.athlete, payload.entries);
  logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  const completedSets = rows.filter(r => r[18] === true).length;
  const totalSets = rows.length;
  const volume = rows.reduce((sum, r) => sum + ((Number(r[12]) || 0) * (Number(r[11]) || 0)), 0);
  const durationSeconds = Number(payload.durationSeconds || 0);
  sessionSheet.appendRow([
    now, session.athlete, date, payload.day || '', Number(payload.week || 1), sessionId,
    durationSeconds, completedSets, totalSets, volume, payload.sessionNotes || ''
  ]);

  const summary = {
    ok: true,
    sessionId: sessionId,
    rowsSaved: rows.length,
    completedSets: completedSets,
    totalSets: totalSets,
    volume: Math.round(volume),
    durationSeconds: durationSeconds,
    personalRecords: personalRecords
  };
  return summary;
}

function getRecentHistory(token, limit) {
  const session = requireSession_(token);
  return getLogRows_(session.athlete, Number(limit || 40));
}

function getExerciseHistory(token, exercise, limit) {
  const session = requireSession_(token);
  const all = getLogRows_(session.athlete, 5000).filter(r => String(r.Exercise).toUpperCase() === String(exercise).toUpperCase());
  const rows = all.slice(0, Number(limit || 30));
  let maxWeight = 0, maxReps = 0, estimated1RM = 0, bestVolume = 0;
  all.forEach(r => {
    const weight = Number(r.Weight || 0), reps = Number(r.Reps || 0);
    maxWeight = Math.max(maxWeight, weight);
    maxReps = Math.max(maxReps, reps);
    bestVolume = Math.max(bestVolume, weight * reps);
    if (weight > 0 && reps > 0 && reps <= 15) estimated1RM = Math.max(estimated1RM, weight * (1 + reps / 30));
  });
  return {exercise: exercise, rows: rows, pr: {maxWeight: maxWeight, maxReps: maxReps, estimated1RM: Math.round(estimated1RM), bestVolume: Math.round(bestVolume)}};
}


function saveCheckIn(token, payload) {
  const session = requireSession_(token);
  payload = payload || {};

  const date = payload.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKINS);

  sheet.appendRow([
    new Date(),
    session.athlete,
    date,
    numberOrBlank_(payload.bodyweight),
    numberOrBlank_(payload.waist),
    numberOrBlank_(payload.sleepHours),
    numberOrBlank_(payload.steps),
    numberOrBlank_(payload.calories),
    numberOrBlank_(payload.protein),
    numberOrBlank_(payload.water),
    numberOrBlank_(payload.energy),
    numberOrBlank_(payload.soreness),
    payload.notes || ''
  ]);

  return getDashboard_(session.athlete);
}

function getDashboard(token) {
  const session = requireSession_(token);
  return getDashboard_(session.athlete);
}

function getDashboard_(athlete) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = startOfDay_(new Date());
  const sevenDaysAgo = new Date(now); sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 6);
  const twentyEightDaysAgo = new Date(now); twentyEightDaysAgo.setDate(twentyEightDaysAgo.getDate() - 27);
  const weekStart = startOfWeek_(now);

  const sessions = readObjectsRaw_(ss.getSheetByName(SHEETS.SESSIONS))
    .filter(r => String(r.Athlete) === String(athlete))
    .map(r => Object.assign({}, r, {_date: parseDateValue_(r.Date || r.Timestamp)}))
    .filter(r => r._date)
    .sort((a, b) => b._date - a._date);

  const checkins = readObjectsRaw_(ss.getSheetByName(SHEETS.CHECKINS))
    .filter(r => String(r.Athlete) === String(athlete))
    .map(r => Object.assign({}, r, {_date: parseDateValue_(r.Date || r.Timestamp)}))
    .filter(r => r._date)
    .sort((a, b) => b._date - a._date);

  const last7Sessions = sessions.filter(r => r._date >= sevenDaysAgo);
  const last28Sessions = sessions.filter(r => r._date >= twentyEightDaysAgo);
  const thisWeekSessions = sessions.filter(r => r._date >= weekStart);
  const last7Checkins = checkins.filter(r => r._date >= sevenDaysAgo);

  const completedSets7 = sum_(last7Sessions, 'Completed Sets');
  const totalSets7 = sum_(last7Sessions, 'Total Sets');
  const volume7 = sum_(last7Sessions, 'Training Volume');
  const duration7 = sum_(last7Sessions, 'Duration Seconds');

  const weighted = checkins
    .filter(r => Number(r.Bodyweight || 0) > 0)
    .slice()
    .sort((a, b) => a._date - b._date);

  const latestWeight = weighted.length ? Number(weighted[weighted.length - 1].Bodyweight) : 0;
  const startingWeight = weighted.length ? Number(weighted[0].Bodyweight) : 0;
  const latestCheckin = checkins.length ? checkins[0] : null;

  const bodyweightSeries = weighted.slice(-14).map(r => ({
    date: Utilities.formatDate(r._date, Session.getScriptTimeZone(), 'M/d'),
    weight: Number(r.Bodyweight)
  }));

  const days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  const weekStatus = days.map(day => ({
    day: day,
    complete: thisWeekSessions.some(r => String(r.Day).toLowerCase() === day.toLowerCase())
  }));

  const recentCheckIns = checkins.slice(0, 7).map(r => ({
    date: Utilities.formatDate(r._date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    bodyweight: r.Bodyweight || '',
    waist: r.Waist || '',
    sleepHours: r['Sleep Hours'] || '',
    steps: r.Steps || '',
    calories: r.Calories || '',
    protein: r.Protein || '',
    energy: r.Energy || '',
    soreness: r.Soreness || '',
    notes: r.Notes || ''
  }));

  return {
    sessionsThisWeek: thisWeekSessions.length,
    workoutsLast7: last7Sessions.length,
    workoutsLast28: last28Sessions.length,
    completedSetsLast7: completedSets7,
    totalSetsLast7: totalSets7,
    completionRateLast7: totalSets7 ? Math.round(completedSets7 / totalSets7 * 100) : 0,
    volumeLast7: Math.round(volume7),
    averageDurationLast7: last7Sessions.length ? Math.round(duration7 / last7Sessions.length) : 0,
    latestWeight: latestWeight || '',
    startingWeight: startingWeight || '',
    weightChange: latestWeight && startingWeight ? Math.round((latestWeight - startingWeight) * 10) / 10 : '',
    latestWaist: latestCheckin ? latestCheckin.Waist || '' : '',
    averageSleep7: roundOne_(average_(last7Checkins, 'Sleep Hours')),
    averageSteps7: Math.round(average_(last7Checkins, 'Steps') || 0),
    averageProtein7: Math.round(average_(last7Checkins, 'Protein') || 0),
    averageEnergy7: roundOne_(average_(last7Checkins, 'Energy')),
    averageSoreness7: roundOne_(average_(last7Checkins, 'Soreness')),
    lastCheckInDate: latestCheckin ? Utilities.formatDate(latestCheckin._date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    bodyweightSeries: bodyweightSeries,
    weekStatus: weekStatus,
    recentCheckIns: recentCheckIns,
    recentSessions: sessions.slice(0, 8).map(r => ({
      date: Utilities.formatDate(r._date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      day: r.Day || '',
      completedSets: Number(r['Completed Sets'] || 0),
      totalSets: Number(r['Total Sets'] || 0),
      volume: Number(r['Training Volume'] || 0),
      durationSeconds: Number(r['Duration Seconds'] || 0)
    }))
  };
}

function readObjectsRaw_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.filter(row => row.some(v => v !== '' && v !== null)).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function parseDateValue_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return startOfDay_(value);

  const text = String(value).trim();
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));

  const parsed = new Date(text);
  return isNaN(parsed) ? null : startOfDay_(parsed);
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function startOfWeek_(date) {
  const d = startOfDay_(date);
  const day = d.getDay();
  const offset = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + offset);
  return d;
}

function sum_(rows, key) {
  return rows.reduce((sum, row) => sum + (Number(row[key]) || 0), 0);
}

function average_(rows, key) {
  const values = rows.map(row => Number(row[key])).filter(v => isFinite(v) && v > 0);
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0;
}

function roundOne_(value) {
  return value ? Math.round(value * 10) / 10 : 0;
}


function setCurrentWeek(token, week) {
  requireSession_(token);
  return writeSetting_('Current Week', Math.max(1, Math.min(4, Number(week || 1))));
}

function setDefaultDay(token, day) {
  requireSession_(token);
  return writeSetting_('Default Day', String(day || 'Monday'));
}

function listProfiles(token) {
  const session = requireCoach_(token);
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES)).map(p => ({Athlete: p.Athlete, Active: p.Active, Role: p.Role}));
}

function addOrUpdateProfile(token, athlete, pin, role) {
  requireCoach_(token);
  const name = String(athlete || '').trim();
  if (!name) throw new Error('Athlete name is required.');
  if (String(pin || '').length < 4) throw new Error('PIN must be at least 4 digits.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES);
  const rows = readObjects_(sheet);
  const idx = rows.findIndex(r => r.Athlete.toLowerCase() === name.toLowerCase());
  const values = [name, hashPin_(String(pin)), true, role || 'Athlete'];
  if (idx >= 0) sheet.getRange(idx + 2, 1, 1, 4).setValues([values]);
  else sheet.appendRow(values);
  return listProfiles(token);
}

function setProfileActive(token, athlete, active) {
  requireCoach_(token);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES);
  const rows = readObjects_(sheet);
  const idx = rows.findIndex(r => r.Athlete.toLowerCase() === String(athlete).toLowerCase());
  if (idx < 0) throw new Error('Profile not found.');
  sheet.getRange(idx + 2, 3).setValue(Boolean(active));
  return listProfiles(token);
}

function forceRefreshProgram(token) {
  requireCoach_(token);
  syncAllProgramData_();
  return getBootstrap(token);
}

function setupStructureOnly_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, SHEETS.PROGRAM, HEADERS.PROGRAM);
  ensureSheetWithHeaders_(ss, SHEETS.WARMUPS, HEADERS.WARMUPS);
  ensureSheetWithHeaders_(ss, SHEETS.MEALS, HEADERS.MEALS);
  ensureSheetWithHeaders_(ss, SHEETS.DAILY, HEADERS.DAILY);
  ensureSheetWithHeaders_(ss, SHEETS.SAUNA, HEADERS.SAUNA);
  ensureSheetWithHeaders_(ss, SHEETS.NOTES, HEADERS.NOTES);
  ensureSheetWithHeaders_(ss, SHEETS.LOG, HEADERS.LOG);
  ensureSheetWithHeaders_(ss, SHEETS.SESSIONS, HEADERS.SESSIONS);
  ensureSheetWithHeaders_(ss, SHEETS.SETTINGS, HEADERS.SETTINGS);
  ensureSheetWithHeaders_(ss, SHEETS.PROFILES, HEADERS.PROFILES);
  ensureSheetWithHeaders_(ss, SHEETS.CHECKINS, HEADERS.CHECKINS);
  seedSettings_();
  seedDefaultProfile_();
}

function seedSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const defaults = [['App Title', APP_TITLE], ['Current Week', '1'], ['Default Day', 'Monday']];
  defaults.forEach(([k, v]) => { if (!readSetting_(k)) sheet.appendRow([k, v]); });
}

function seedDefaultProfile_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES);
  if (sheet.getLastRow() < 2) sheet.appendRow(['Jeremy', hashPin_('2468'), true, 'Coach']);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const oldValues = sheet.getDataRange().getValues();
    const oldHeaders = oldValues[0].map(String);
    const exact = headers.length === oldHeaders.length && headers.every((h, i) => oldHeaders[i] === h);

    if (!exact) {
      const oldIndex = {};
      oldHeaders.forEach((h, i) => oldIndex[h] = i);
      const migrated = oldValues.slice(1).filter(row => row.some(v => v !== '' && v !== null)).map(row =>
        headers.map(h => {
          if (Object.prototype.hasOwnProperty.call(oldIndex, h)) return row[oldIndex[h]];
          if (h === 'Athlete' && name === SHEETS.LOG) return 'Jeremy';
          return '';
        })
      );
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (migrated.length) sheet.getRange(2, 1, migrated.length, headers.length).setValues(migrated);
    }
  }

  sheet.getRange(1, 1, 1, headers.length).setBackground('#1f2937').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function syncAllProgramData_() {
  replaceSheetData_(SHEETS.PROGRAM, HEADERS.PROGRAM, PROGRAM_DATA.map((x, i) => [
    x.day, x.block, x.slot, x.exercise, x.notes, x.week1, x.week2, x.week3, x.week4,
    x.suggestedLoad, i + 1, x.inputType || inferInputType_(x.exercise, x.week1),
    x.restSeconds || inferRestSeconds_(x.exercise, x.block), x.videoUrl || ''
  ]));
  replaceSheetData_(SHEETS.WARMUPS, HEADERS.WARMUPS, WARMUP_DATA.map((x, i) => [x.day, x.drill, x.purpose, x.dose, i + 1]));
  replaceSheetData_(SHEETS.MEALS, HEADERS.MEALS, MEAL_DATA.map((x, i) => [x.day, x.meal, x.time, x.food, i + 1]));
  replaceSheetData_(SHEETS.DAILY, HEADERS.DAILY, DAILY_DATA.map((x, i) => [x.day, x.item, x.target, i + 1]));
  replaceSheetData_(SHEETS.SAUNA, HEADERS.SAUNA, SAUNA_DATA.map((x, i) => [x.day, x.item, x.detail, i + 1]));
  replaceSheetData_(SHEETS.NOTES, HEADERS.NOTES, UPGRADE_NOTES.map(x => [x.day, x.note]));
}

function replaceSheetData_(name, headers, rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1f2937').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function enrichProgram_(rows) {
  return rows.map(r => {
    r['Input Type'] = r['Input Type'] || inferInputType_(r.Exercise, r['Week 1']);
    r['Rest Seconds'] = r['Rest Seconds'] || String(inferRestSeconds_(r.Exercise, r.Block));
    r['Video URL'] = r['Video URL'] || '';
    return r;
  });
}

function inferInputType_(exercise, target) {
  const text = (String(exercise) + ' ' + String(target)).toUpperCase();
  if (/BREATH|VACUUM/.test(text)) return 'BREATHS_TIME';
  if (/ROUND/.test(text)) return 'ROUNDS_TIME';
  if (/WALK|BIKE|AIRBIKE|AIRDYNE|SAUNA|MOBILITY|RECOVERY|TIME|SPRINTS/.test(text) && /MIN|SEC|\/|EASY/.test(text)) return 'TIME';
  if (/CARRY|SLED|YD|YARD|MARCH|DRAG/.test(text)) return 'DISTANCE_LOAD';
  return 'REPS_WEIGHT';
}

function inferRestSeconds_(exercise, block) {
  const text = (String(exercise) + ' ' + String(block)).toUpperCase();
  if (/FRONT SQUAT|PUSH PRESS|JUMP|THRUST|RFESS|CHIN|PULL-UP/.test(text)) return 120;
  if (/FINISHER|BIKE|WALK|RECOVERY|MOBILITY/.test(text)) return 0;
  return 75;
}

function getPreviousPerformance_(athlete) {
  const rows = getLogRows_(athlete, 5000);
  const previous = {};
  rows.forEach(r => {
    const key = String(r.Exercise).toUpperCase();
    if (!previous[key]) previous[key] = {date: r.Date, sets: []};
    if (previous[key].date === r.Date && previous[key].sets.length < 12) {
      previous[key].sets.push({set: r.Set, reps: r.Reps, weight: r.Weight, distance: r.Distance, durationSeconds: r['Duration Seconds'], rounds: r.Rounds, breaths: r.Breaths, rpe: r.RPE});
    }
  });
  Object.keys(previous).forEach(key => previous[key].sets.sort((a, b) => Number(a.set || 0) - Number(b.set || 0)));
  return previous;
}

function getLogRows_(athlete, limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  const athleteIdx = headers.indexOf('Athlete');
  return values.reverse().filter(row => athleteIdx < 0 || row[athleteIdx] === athlete).slice(0, limit).map(row => {
    const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
  });
}

function detectSessionPRs_(athlete, entries) {
  const prs = [];
  const history = getLogRows_(athlete, 5000);
  const grouped = {};
  history.forEach(r => {
    const key = String(r.Exercise).toUpperCase();
    if (!grouped[key]) grouped[key] = {maxWeight: 0, maxReps: 0, e1rm: 0};
    const w = Number(r.Weight || 0), reps = Number(r.Reps || 0);
    grouped[key].maxWeight = Math.max(grouped[key].maxWeight, w);
    grouped[key].maxReps = Math.max(grouped[key].maxReps, reps);
    if (w > 0 && reps > 0 && reps <= 15) grouped[key].e1rm = Math.max(grouped[key].e1rm, w * (1 + reps / 30));
  });
  entries.forEach(e => {
    const key = String(e.exercise).toUpperCase();
    const old = grouped[key] || {maxWeight: 0, maxReps: 0, e1rm: 0};
    const w = Number(e.weight || 0), reps = Number(e.reps || 0), e1rm = w > 0 && reps > 0 && reps <= 15 ? w * (1 + reps / 30) : 0;
    if (w > old.maxWeight) prs.push(e.exercise + ': weight PR ' + w);
    else if (e1rm > old.e1rm + 0.5) prs.push(e.exercise + ': estimated 1RM PR ' + Math.round(e1rm));
    else if (reps > old.maxReps && w >= old.maxWeight * 0.9) prs.push(e.exercise + ': rep PR ' + reps);
  });
  return [...new Set(prs)];
}

function getActiveProfileNames_() {
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES))
    .filter(p => String(p.Active).toUpperCase() === 'TRUE')
    .map(p => p.Athlete);
}

function findProfile_(athlete) {
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROFILES))
    .find(p => p.Athlete.toLowerCase() === String(athlete).toLowerCase());
}

function hashPin_(pin) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin), Utilities.Charset.UTF_8));
}

function requireSession_(token) {
  const raw = CacheService.getScriptCache().get('session:' + token);
  if (!raw) throw new Error('Your session expired. Log in again.');
  CacheService.getScriptCache().put('session:' + token, raw, SESSION_TTL_SECONDS);
  return JSON.parse(raw);
}

function requireCoach_(token) {
  const session = requireSession_(token);
  if (String(session.role).toLowerCase() !== 'coach') throw new Error('Coach access required.');
  return session;
}

function readObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  return values.filter(row => row.some(v => v !== '')).map(row => {
    const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
  });
}

function readSetting_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const row = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues().find(r => r[0] === key);
  return row ? row[1] : '';
}

function writeSetting_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return value; }
  }
  sheet.appendRow([key, value]);
  return value;
}

function numberOrBlank_(value) {
  return value === '' || value === null || typeof value === 'undefined' ? '' : Number(value);
}

const PROGRAM_DATA = [
  {
    "day": "Monday",
    "block": "TRISET A",
    "slot": "A1",
    "exercise": "FRONT SQUAT",
    "notes": "Heavy but clean",
    "week1": "4x6",
    "week2": "5x5",
    "week3": "5x4",
    "week4": "3x5",
    "suggestedLoad": "185-225",
    "order": 1
  },
  {
    "day": "Monday",
    "block": "TRISET A",
    "slot": "A2",
    "exercise": "MED BALL CHEST PASS",
    "notes": "Explosive",
    "week1": "4x3",
    "week2": "5x3",
    "week3": "5x3",
    "week4": "3x3",
    "suggestedLoad": "10-20 lb",
    "order": 2
  },
  {
    "day": "Monday",
    "block": "TRISET A",
    "slot": "A3",
    "exercise": "HOLLOW BODY HOLD",
    "notes": "Anti-extension core",
    "week1": "3x20s",
    "week2": "3x25s",
    "week3": "4x20s",
    "week4": "2x20s",
    "suggestedLoad": "BW",
    "order": 3
  },
  {
    "day": "Monday",
    "block": "TRISET B",
    "slot": "B1",
    "exercise": "DB BENCH PRESS",
    "notes": "1-2 reps in reserve",
    "week1": "4x10",
    "week2": "4x8",
    "week3": "5x8",
    "week4": "3x10",
    "suggestedLoad": "70-90s",
    "order": 4
  },
  {
    "day": "Monday",
    "block": "TRISET B",
    "slot": "B2",
    "exercise": "CHEST SUPPORTED ROW",
    "notes": "Full squeeze",
    "week1": "4x12",
    "week2": "4x10",
    "week3": "5x10",
    "week4": "3x12",
    "suggestedLoad": "70-100s",
    "order": 5
  },
  {
    "day": "Monday",
    "block": "TRISET B",
    "slot": "B3",
    "exercise": "DUAL CABLE LIFT / CHOP",
    "notes": "Athletic trunk",
    "week1": "3x12e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x12e",
    "suggestedLoad": "30-50",
    "order": 6
  },
  {
    "day": "Monday",
    "block": "TRISET C",
    "slot": "C1",
    "exercise": "WALKING LUNGE",
    "notes": "Smooth continuous reps",
    "week1": "3x10e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x20 steps",
    "suggestedLoad": "40-60s",
    "order": 7
  },
  {
    "day": "Monday",
    "block": "TRISET C",
    "slot": "C2",
    "exercise": "WEIGHTED CHIN-UP",
    "notes": "Lat width + upper-body density",
    "week1": "4x6",
    "week2": "4x8",
    "week3": "5x6",
    "week4": "3x8",
    "suggestedLoad": "BW / +25",
    "order": 8
  },
  {
    "day": "Monday",
    "block": "TRISET C",
    "slot": "C3",
    "exercise": "CABLE LATERAL RAISE",
    "notes": "Shoulder cap pump",
    "week1": "3x15",
    "week2": "3x20",
    "week3": "4x15",
    "week4": "2x25",
    "suggestedLoad": "10-20",
    "order": 9
  },
  {
    "day": "Monday",
    "block": "FINISHER",
    "slot": "F",
    "exercise": "BIKE SPRINTS",
    "notes": "Stop before sloppy",
    "week1": "10/50 x8",
    "week2": "12/48 x10",
    "week3": "15/45 x10",
    "week4": "8 min easy",
    "suggestedLoad": "Max effort",
    "order": 10
  },
  {
    "day": "Tuesday",
    "block": "TRISET A",
    "slot": "A1",
    "exercise": "DB JUMP SQUAT",
    "notes": "Fast and explosive",
    "week1": "4x5",
    "week2": "5x4",
    "week3": "5x3",
    "week4": "3x5",
    "suggestedLoad": "20-35s",
    "order": 1
  },
  {
    "day": "Tuesday",
    "block": "TRISET A",
    "slot": "A2",
    "exercise": "BOX JUMP",
    "notes": "Stick landing",
    "week1": "4x3",
    "week2": "5x3",
    "week3": "5x2",
    "week4": "3x3",
    "suggestedLoad": "BW",
    "order": 2
  },
  {
    "day": "Tuesday",
    "block": "TRISET A",
    "slot": "A3",
    "exercise": "SPRINTER SIT-UP",
    "notes": "Dynamic trunk / hip flexor",
    "week1": "3x10e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x10e",
    "suggestedLoad": "BW",
    "order": 3
  },
  {
    "day": "Tuesday",
    "block": "TRISET B",
    "slot": "B1",
    "exercise": "RFESS",
    "notes": "Single-leg strength",
    "week1": "3x8e",
    "week2": "4x8e",
    "week3": "4x10e",
    "week4": "2x12e",
    "suggestedLoad": "40-60s",
    "order": 4
  },
  {
    "day": "Tuesday",
    "block": "TRISET B",
    "slot": "B2",
    "exercise": "PULL-UP",
    "notes": "Full range",
    "week1": "4x8",
    "week2": "4x10",
    "week3": "5x8",
    "week4": "3x10",
    "suggestedLoad": "BW / +25",
    "order": 5
  },
  {
    "day": "Tuesday",
    "block": "TRISET B",
    "slot": "B3",
    "exercise": "FARMER CARRY",
    "notes": "Tall posture",
    "week1": "3x40 yd",
    "week2": "4x40 yd",
    "week3": "5x30 yd",
    "week4": "2x40 yd",
    "suggestedLoad": "70-100s",
    "order": 6
  },
  {
    "day": "Tuesday",
    "block": "TRISET C",
    "slot": "C1",
    "exercise": "PUSH PRESS",
    "notes": "Explosive drive",
    "week1": "4x5",
    "week2": "5x4",
    "week3": "5x3",
    "week4": "3x5",
    "suggestedLoad": "135-185",
    "order": 7
  },
  {
    "day": "Tuesday",
    "block": "TRISET C",
    "slot": "C2",
    "exercise": "GLUTE HAM RAISE",
    "notes": "Direct hamstring strength",
    "week1": "3x8",
    "week2": "4x8",
    "week3": "4x10",
    "week4": "2x12",
    "suggestedLoad": "BW / +25",
    "order": 8
  },
  {
    "day": "Tuesday",
    "block": "TRISET C",
    "slot": "C3",
    "exercise": "SLED PUSH",
    "notes": "Moderate continuous effort",
    "week1": "10 min",
    "week2": "12 min",
    "week3": "15 min",
    "week4": "8 min",
    "suggestedLoad": "Moderate-heavy",
    "order": 9
  },
  {
    "day": "Wednesday",
    "block": "RECOVERY",
    "slot": "A1",
    "exercise": "INCLINE WALK",
    "notes": "Nasal breathing if possible",
    "week1": "45 min",
    "week2": "50 min",
    "week3": "60 min",
    "week4": "40 min",
    "suggestedLoad": "3.0-3.5 mph",
    "order": 1
  },
  {
    "day": "Wednesday",
    "block": "RECOVERY",
    "slot": "A2",
    "exercise": "RECOVERY MOBILITY",
    "notes": "Move continuously",
    "week1": "2 rounds",
    "week2": "3 rounds",
    "week3": "3 rounds",
    "week4": "2 rounds",
    "suggestedLoad": "BW",
    "order": 2
  },
  {
    "day": "Wednesday",
    "block": "RECOVERY",
    "slot": "A3",
    "exercise": "CROCODILE BREATHING",
    "notes": "Relaxation",
    "week1": "2 min",
    "week2": "3 min",
    "week3": "3 min",
    "week4": "2 min",
    "suggestedLoad": "BW",
    "order": 3
  },
  {
    "day": "Wednesday",
    "block": "CORE",
    "slot": "B1",
    "exercise": "HANGING KNEE RAISE",
    "notes": "Lower abs",
    "week1": "3x12",
    "week2": "3x15",
    "week3": "4x12",
    "week4": "2x15",
    "suggestedLoad": "BW",
    "order": 4
  },
  {
    "day": "Wednesday",
    "block": "CORE",
    "slot": "B2",
    "exercise": "PALLOF PRESS",
    "notes": "Anti-rotation",
    "week1": "3x10e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x12e",
    "suggestedLoad": "30-50",
    "order": 5
  },
  {
    "day": "Wednesday",
    "block": "CORE",
    "slot": "B3",
    "exercise": "VACUUM BREATHING",
    "notes": "Waist control",
    "week1": "3x5 breaths",
    "week2": "4x5 breaths",
    "week3": "4x6 breaths",
    "week4": "3x5 breaths",
    "suggestedLoad": "BW",
    "order": 6
  },
  {
    "day": "Thursday",
    "block": "TRISET A",
    "slot": "A1",
    "exercise": "HEEL-ELEVATED DB SQUAT",
    "notes": "Quad focus",
    "week1": "4x12",
    "week2": "4x15",
    "week3": "5x12",
    "week4": "3x15",
    "suggestedLoad": "50-70s",
    "order": 1
  },
  {
    "day": "Thursday",
    "block": "TRISET A",
    "slot": "A2",
    "exercise": "DB INCLINE BENCH",
    "notes": "Upper chest",
    "week1": "4x10",
    "week2": "4x12",
    "week3": "5x10",
    "week4": "3x15",
    "suggestedLoad": "60-80s",
    "order": 2
  },
  {
    "day": "Thursday",
    "block": "TRISET A",
    "slot": "A3",
    "exercise": "REAR DELT CABLE FLY",
    "notes": "Rear delt cap / posture",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "10-20",
    "order": 3
  },
  {
    "day": "Thursday",
    "block": "TRISET B",
    "slot": "B1",
    "exercise": "SEATED CABLE ROW",
    "notes": "Controlled squeeze",
    "week1": "4x12",
    "week2": "4x15",
    "week3": "5x12",
    "week4": "3x15",
    "suggestedLoad": "120-180",
    "order": 4
  },
  {
    "day": "Thursday",
    "block": "TRISET B",
    "slot": "B2",
    "exercise": "SINGLE LEG RDL",
    "notes": "Control",
    "week1": "3x10e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x12e",
    "suggestedLoad": "40-60s",
    "order": 5
  },
  {
    "day": "Thursday",
    "block": "TRISET B",
    "slot": "B3",
    "exercise": "LAT PULLDOWN",
    "notes": "Extra vertical pull / lat width",
    "week1": "4x10",
    "week2": "4x12",
    "week3": "5x10",
    "week4": "3x12",
    "suggestedLoad": "120-180",
    "order": 6
  },
  {
    "day": "Thursday",
    "block": "TRISET C",
    "slot": "C1",
    "exercise": "LEG CURL",
    "notes": "Hamstring pump",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x12",
    "week4": "2x20",
    "suggestedLoad": "90-140",
    "order": 7
  },
  {
    "day": "Thursday",
    "block": "TRISET C",
    "slot": "C2",
    "exercise": "STRAIGHT ARM PULLDOWN",
    "notes": "Lat tie-in + trunk stiffness",
    "week1": "3x12",
    "week2": "3x15",
    "week3": "4x12",
    "week4": "2x15",
    "suggestedLoad": "50-80",
    "order": 8
  },
  {
    "day": "Thursday",
    "block": "TRISET C",
    "slot": "C3",
    "exercise": "ASSAULT BIKE",
    "notes": "Hard but controlled",
    "week1": "15/45 x8",
    "week2": "15/45 x10",
    "week3": "20/40 x10",
    "week4": "10 min easy",
    "suggestedLoad": "Max effort",
    "order": 9
  },
  {
    "day": "Friday",
    "block": "TRISET A",
    "slot": "A1",
    "exercise": "HIP THRUST",
    "notes": "Glute focus",
    "week1": "4x10",
    "week2": "4x12",
    "week3": "5x10",
    "week4": "3x15",
    "suggestedLoad": "225-315",
    "order": 1
  },
  {
    "day": "Friday",
    "block": "TRISET A",
    "slot": "A2",
    "exercise": "LOW INCLINE DB PRESS",
    "notes": "Upper chest fullness",
    "week1": "4x12",
    "week2": "4x10",
    "week3": "5x8",
    "week4": "3x15",
    "suggestedLoad": "50-70s",
    "order": 2
  },
  {
    "day": "Friday",
    "block": "TRISET A",
    "slot": "A3",
    "exercise": "AB WHEEL",
    "notes": "Brace hard",
    "week1": "3x10",
    "week2": "4x10",
    "week3": "4x12",
    "week4": "2x10",
    "suggestedLoad": "BW",
    "order": 3
  },
  {
    "day": "Friday",
    "block": "TRISET B",
    "slot": "B1",
    "exercise": "1-ARM DB ROW",
    "notes": "Upper back",
    "week1": "4x12e",
    "week2": "4x15e",
    "week3": "5x12e",
    "week4": "3x15e",
    "suggestedLoad": "80-110",
    "order": 4
  },
  {
    "day": "Friday",
    "block": "TRISET B",
    "slot": "B2",
    "exercise": "SLIDING LEG CURL",
    "notes": "Hamstring hypertrophy",
    "week1": "3x12",
    "week2": "4x12",
    "week3": "4x15",
    "week4": "2x20",
    "suggestedLoad": "BW / sliders",
    "order": 5
  },
  {
    "day": "Friday",
    "block": "TRISET B",
    "slot": "B3",
    "exercise": "DUAL CABLE LATERAL RAISE",
    "notes": "Shoulder detail",
    "week1": "3x20",
    "week2": "3x25",
    "week3": "4x20",
    "week4": "2x25",
    "suggestedLoad": "10-20",
    "order": 6
  },
  {
    "day": "Friday",
    "block": "TRISET C",
    "slot": "C1",
    "exercise": "CABLE CURL",
    "notes": "Constant tension",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "40-70",
    "order": 7
  },
  {
    "day": "Friday",
    "block": "TRISET C",
    "slot": "C2",
    "exercise": "SUITCASE CARRY",
    "notes": "Core stability",
    "week1": "3x30 yd",
    "week2": "4x40 yd",
    "week3": "5x30 yd",
    "week4": "2x40 yd",
    "suggestedLoad": "70-100s",
    "order": 8
  },
  {
    "day": "Friday",
    "block": "TRISET C",
    "slot": "C3",
    "exercise": "AIRDYNE RECOVERY",
    "notes": "Low-impact recovery conditioning",
    "week1": "15 min",
    "week2": "20 min",
    "week3": "25 min",
    "week4": "15 min",
    "suggestedLoad": "Easy",
    "order": 9
  },
  {
    "day": "Friday",
    "block": "TRISET D",
    "slot": "D1",
    "exercise": "WEIGHTED DIPS",
    "notes": "Chest / triceps / serratus",
    "week1": "3x10",
    "week2": "4x10",
    "week3": "4x12",
    "week4": "2xAMRAP",
    "suggestedLoad": "BW / +25",
    "order": 10
  },
  {
    "day": "Friday",
    "block": "TRISET D",
    "slot": "D2",
    "exercise": "FACE PULL",
    "notes": "Rear delt + posture",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "40-70",
    "order": 11
  },
  {
    "day": "Friday",
    "block": "TRISET D",
    "slot": "D3",
    "exercise": "LEAN-AWAY LATERAL RAISE",
    "notes": "Shoulder cap fullness",
    "week1": "3x15",
    "week2": "3x20",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "10-20",
    "order": 12
  },
  {
    "day": "Friday",
    "block": "TRISET D",
    "slot": "D4",
    "exercise": "OVERHEAD ROPE EXTENSION",
    "notes": "Long head triceps",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "40-70",
    "order": 13
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT A",
    "slot": "A1",
    "exercise": "SLED PUSH",
    "notes": "Moderate pace",
    "week1": "20 yd",
    "week2": "25 yd",
    "week3": "30 yd",
    "week4": "20 yd",
    "suggestedLoad": "Moderate",
    "order": 1
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT A",
    "slot": "A2",
    "exercise": "BIKE",
    "notes": "Moderate-hard",
    "week1": "30 sec",
    "week2": "40 sec",
    "week3": "45 sec",
    "week4": "30 sec",
    "suggestedLoad": "Moderate",
    "order": 2
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT A",
    "slot": "A3",
    "exercise": "ZERCHER CARRY",
    "notes": "Anterior core + posture",
    "week1": "40 yd",
    "week2": "50 yd",
    "week3": "60 yd",
    "week4": "40 yd",
    "suggestedLoad": "Moderate",
    "order": 3
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT B",
    "slot": "B1",
    "exercise": "FACE PULL",
    "notes": "Rear delt + posture",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "40-70",
    "order": 4
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT B",
    "slot": "B2",
    "exercise": "PALLOF PRESS",
    "notes": "Anti-rotation trunk",
    "week1": "3x10e",
    "week2": "3x12e",
    "week3": "4x10e",
    "week4": "2x12e",
    "suggestedLoad": "30-50",
    "order": 5
  },
  {
    "day": "Saturday",
    "block": "CIRCUIT B",
    "slot": "B3",
    "exercise": "WALKING LUNGE",
    "notes": "Continuous reps",
    "week1": "10e",
    "week2": "12e",
    "week3": "15e",
    "week4": "10e",
    "suggestedLoad": "40-60s",
    "order": 6
  },
  {
    "day": "Saturday",
    "block": "CORE",
    "slot": "C1",
    "exercise": "SUITCASE CARRY",
    "notes": "Obliques + posture",
    "week1": "3x30 yd",
    "week2": "4x40 yd",
    "week3": "5x30 yd",
    "week4": "2x40 yd",
    "suggestedLoad": "70-100s",
    "order": 7
  },
  {
    "day": "Saturday",
    "block": "CORE",
    "slot": "C2",
    "exercise": "AB WHEEL",
    "notes": "Anti-extension",
    "week1": "3x10",
    "week2": "4x10",
    "week3": "4x12",
    "week4": "2x10",
    "suggestedLoad": "BW",
    "order": 8
  },
  {
    "day": "Saturday",
    "block": "CORE",
    "slot": "C3",
    "exercise": "VACUUM BREATHING",
    "notes": "Waist control",
    "week1": "3x5 breaths",
    "week2": "4x5 breaths",
    "week3": "4x6 breaths",
    "week4": "3x5 breaths",
    "suggestedLoad": "BW",
    "order": 9
  },
  {
    "day": "Sunday",
    "block": "RECOVERY",
    "slot": "A1",
    "exercise": "LONG WALK",
    "notes": "Easy pace",
    "week1": "45 min",
    "week2": "60 min",
    "week3": "60 min",
    "week4": "40 min",
    "suggestedLoad": "Easy",
    "order": 1
  },
  {
    "day": "Sunday",
    "block": "RECOVERY",
    "slot": "A2",
    "exercise": "RECOVERY MOBILITY",
    "notes": "Low intensity",
    "week1": "10 min",
    "week2": "15 min",
    "week3": "15 min",
    "week4": "10 min",
    "suggestedLoad": "BW",
    "order": 2
  },
  {
    "day": "Sunday",
    "block": "RECOVERY",
    "slot": "A3",
    "exercise": "BREATHING",
    "notes": "Crocodile / box breathing",
    "week1": "5 min",
    "week2": "5 min",
    "week3": "8 min",
    "week4": "5 min",
    "suggestedLoad": "BW",
    "order": 3
  },
  {
    "day": "Sunday",
    "block": "CORE",
    "slot": "B1",
    "exercise": "CABLE CRUNCH",
    "notes": "Upper abs",
    "week1": "3x15",
    "week2": "4x15",
    "week3": "4x20",
    "week4": "2x20",
    "suggestedLoad": "50-80",
    "order": 4
  },
  {
    "day": "Sunday",
    "block": "CORE",
    "slot": "B2",
    "exercise": "SIDE PLANK",
    "notes": "Obliques",
    "week1": "3x30s",
    "week2": "3x40s",
    "week3": "4x40s",
    "week4": "2x30s",
    "suggestedLoad": "BW",
    "order": 5
  },
  {
    "day": "Sunday",
    "block": "CORE",
    "slot": "B3",
    "exercise": "VACUUM BREATHING",
    "notes": "Waist control",
    "week1": "3x5 breaths",
    "week2": "4x5 breaths",
    "week3": "4x6 breaths",
    "week4": "3x5 breaths",
    "suggestedLoad": "BW",
    "order": 6
  }
];
const WARMUP_DATA = [
  {
    "day": "Monday",
    "drill": "A-SKIP",
    "purpose": "Sprint mechanics + stiffness",
    "dose": "2x20 yd",
    "order": 1
  },
  {
    "day": "Monday",
    "drill": "ANKLE ROCKER",
    "purpose": "Ankle mobility",
    "dose": "10/side",
    "order": 2
  },
  {
    "day": "Monday",
    "drill": "90/90 HIP SWITCH",
    "purpose": "Hip rotation",
    "dose": "6/side",
    "order": 3
  },
  {
    "day": "Monday",
    "drill": "GOBLET SQUAT ISO",
    "purpose": "Front squat pattern",
    "dose": "2x20 sec",
    "order": 4
  },
  {
    "day": "Monday",
    "drill": "BOX JUMP",
    "purpose": "Potentiation",
    "dose": "3x3",
    "order": 5
  },
  {
    "day": "Tuesday",
    "drill": "POGO HOPS",
    "purpose": "Elastic stiffness",
    "dose": "2x20",
    "order": 1
  },
  {
    "day": "Tuesday",
    "drill": "WALL DRILL SWITCH",
    "purpose": "Acceleration mechanics",
    "dose": "2x10",
    "order": 2
  },
  {
    "day": "Tuesday",
    "drill": "HIP AIRPLANE",
    "purpose": "Hip stability",
    "dose": "5/side",
    "order": 3
  },
  {
    "day": "Tuesday",
    "drill": "MED BALL SLAM",
    "purpose": "CNS activation",
    "dose": "3x3",
    "order": 4
  },
  {
    "day": "Tuesday",
    "drill": "BROAD JUMP",
    "purpose": "Explosive prep",
    "dose": "3x2",
    "order": 5
  },
  {
    "day": "Wednesday",
    "drill": "INCLINE WALK",
    "purpose": "Raise temperature",
    "dose": "5 min",
    "order": 1
  },
  {
    "day": "Wednesday",
    "drill": "90/90 BREATHING",
    "purpose": "Downshift nervous system",
    "dose": "2 min",
    "order": 2
  },
  {
    "day": "Wednesday",
    "drill": "WORLD'S GREATEST STRETCH",
    "purpose": "Hip + T-spine mobility",
    "dose": "5/side",
    "order": 3
  },
  {
    "day": "Wednesday",
    "drill": "CAT-CAMEL",
    "purpose": "Spinal motion",
    "dose": "8",
    "order": 4
  },
  {
    "day": "Wednesday",
    "drill": "ANKLE ROCKER",
    "purpose": "Restore ankle mobility",
    "dose": "10/side",
    "order": 5
  },
  {
    "day": "Thursday",
    "drill": "BIKE",
    "purpose": "Raise temperature",
    "dose": "5 min",
    "order": 1
  },
  {
    "day": "Thursday",
    "drill": "COSSACK SQUAT",
    "purpose": "Adductors + hips",
    "dose": "6/side",
    "order": 2
  },
  {
    "day": "Thursday",
    "drill": "BAND PULL-APART",
    "purpose": "Upper back activation",
    "dose": "15",
    "order": 3
  },
  {
    "day": "Thursday",
    "drill": "HEEL-ELEVATED GOBLET SQUAT",
    "purpose": "Quad prep",
    "dose": "2x8",
    "order": 4
  },
  {
    "day": "Thursday",
    "drill": "MED BALL CHEST PASS",
    "purpose": "Potentiation",
    "dose": "3x3",
    "order": 5
  },
  {
    "day": "Friday",
    "drill": "INCLINE WALK",
    "purpose": "Raise temperature",
    "dose": "5 min",
    "order": 1
  },
  {
    "day": "Friday",
    "drill": "GLUTE BRIDGE",
    "purpose": "Glute activation",
    "dose": "2x12",
    "order": 2
  },
  {
    "day": "Friday",
    "drill": "SCAP PUSH-UP",
    "purpose": "Scap control",
    "dose": "2x10",
    "order": 3
  },
  {
    "day": "Friday",
    "drill": "DEADBUG",
    "purpose": "Core brace",
    "dose": "2x8e",
    "order": 4
  },
  {
    "day": "Friday",
    "drill": "MED BALL ROTATIONAL THROW",
    "purpose": "Athletic activation",
    "dose": "3x3e",
    "order": 5
  },
  {
    "day": "Saturday",
    "drill": "BIKE",
    "purpose": "Raise temperature",
    "dose": "5 min",
    "order": 1
  },
  {
    "day": "Saturday",
    "drill": "WORLD'S GREATEST STRETCH",
    "purpose": "Mobility",
    "dose": "5/side",
    "order": 2
  },
  {
    "day": "Saturday",
    "drill": "POGO HOPS",
    "purpose": "Elasticity",
    "dose": "2x20",
    "order": 3
  },
  {
    "day": "Saturday",
    "drill": "SLED MARCH",
    "purpose": "Prep circuit",
    "dose": "2x20 yd",
    "order": 4
  },
  {
    "day": "Saturday",
    "drill": "MED BALL SLAM",
    "purpose": "Activation",
    "dose": "3x3",
    "order": 5
  },
  {
    "day": "Sunday",
    "drill": "WALK",
    "purpose": "Blood flow",
    "dose": "5-10 min",
    "order": 1
  },
  {
    "day": "Sunday",
    "drill": "90/90 BREATHING",
    "purpose": "Relaxation",
    "dose": "2 min",
    "order": 2
  },
  {
    "day": "Sunday",
    "drill": "CAT-CAMEL",
    "purpose": "Spinal mobility",
    "dose": "8",
    "order": 3
  },
  {
    "day": "Sunday",
    "drill": "CHILD'S POSE BREATHING",
    "purpose": "Recovery",
    "dose": "2 min",
    "order": 4
  },
  {
    "day": "Sunday",
    "drill": "ANKLE ROCKER",
    "purpose": "Restore movement",
    "dose": "10/side",
    "order": 5
  }
];
const MEAL_DATA = [
  {
    "day": "Monday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Eggs + oats + berries",
    "order": 1
  },
  {
    "day": "Monday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Chicken + jasmine rice + broccoli",
    "order": 2
  },
  {
    "day": "Monday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Greek yogurt + walnuts + chia",
    "order": 3
  },
  {
    "day": "Monday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Lean beef + potatoes + greens",
    "order": 4
  },
  {
    "day": "Monday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Protein shake + magnesium",
    "order": 5
  },
  {
    "day": "Tuesday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Protein smoothie + banana",
    "order": 1
  },
  {
    "day": "Tuesday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Chicken rice bowl + vegetables",
    "order": 2
  },
  {
    "day": "Tuesday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Eggs + avocado + fruit",
    "order": 3
  },
  {
    "day": "Tuesday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Salmon + asparagus + potatoes",
    "order": 4
  },
  {
    "day": "Tuesday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Greek yogurt or protein shake",
    "order": 5
  },
  {
    "day": "Wednesday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Eggs + spinach + berries",
    "order": 1
  },
  {
    "day": "Wednesday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Ground turkey salad + olive oil",
    "order": 2
  },
  {
    "day": "Wednesday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Protein shake + almonds",
    "order": 3
  },
  {
    "day": "Wednesday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Chicken + vegetables + small potato if needed",
    "order": 4
  },
  {
    "day": "Wednesday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Tea + magnesium",
    "order": 5
  },
  {
    "day": "Thursday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Oats + whey + fruit",
    "order": 1
  },
  {
    "day": "Thursday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Chicken + rice + broccoli",
    "order": 2
  },
  {
    "day": "Thursday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Greek yogurt + chia + berries",
    "order": 3
  },
  {
    "day": "Thursday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Lean beef + sweet potato + greens",
    "order": 4
  },
  {
    "day": "Thursday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Casein / protein shake",
    "order": 5
  },
  {
    "day": "Friday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Eggs + turkey bacon + fruit",
    "order": 1
  },
  {
    "day": "Friday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Tuna/chicken + rice cakes + vegetables",
    "order": 2
  },
  {
    "day": "Friday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Protein smoothie",
    "order": 3
  },
  {
    "day": "Friday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Steak + vegetables + potatoes if needed",
    "order": 4
  },
  {
    "day": "Friday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Greek yogurt",
    "order": 5
  },
  {
    "day": "Saturday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Protein pancakes or eggs + fruit",
    "order": 1
  },
  {
    "day": "Saturday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Chicken salad + rice if training",
    "order": 2
  },
  {
    "day": "Saturday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Protein shake + almonds",
    "order": 3
  },
  {
    "day": "Saturday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Salmon + potatoes + vegetables",
    "order": 4
  },
  {
    "day": "Saturday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Tea + magnesium",
    "order": 5
  },
  {
    "day": "Sunday",
    "meal": "Meal 1",
    "time": "7:00 AM",
    "food": "Eggs + fruit",
    "order": 1
  },
  {
    "day": "Sunday",
    "meal": "Meal 2",
    "time": "12:00 PM",
    "food": "Lean burger bowl + salad",
    "order": 2
  },
  {
    "day": "Sunday",
    "meal": "Meal 3",
    "time": "3:30 PM",
    "food": "Greek yogurt + berries",
    "order": 3
  },
  {
    "day": "Sunday",
    "meal": "Meal 4",
    "time": "6:30 PM",
    "food": "Chicken + vegetables + potatoes",
    "order": 4
  },
  {
    "day": "Sunday",
    "meal": "Before Bed",
    "time": "9:30 PM",
    "food": "Protein shake + magnesium",
    "order": 5
  }
];
const DAILY_DATA = [
  {
    "day": "Monday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Monday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Monday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Monday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Monday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Tuesday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Tuesday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Tuesday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Tuesday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Tuesday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Wednesday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Wednesday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Wednesday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Wednesday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Wednesday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Thursday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Thursday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Thursday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Thursday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Thursday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Friday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Friday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Friday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Friday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Friday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Saturday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Saturday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Saturday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Saturday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Saturday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  },
  {
    "day": "Sunday",
    "item": "Water",
    "target": "1-1.5 gallons",
    "order": 1
  },
  {
    "day": "Sunday",
    "item": "Protein",
    "target": "200-220g",
    "order": 2
  },
  {
    "day": "Sunday",
    "item": "Steps",
    "target": "10k-14k",
    "order": 3
  },
  {
    "day": "Sunday",
    "item": "Alcohol",
    "target": "0",
    "order": 4
  },
  {
    "day": "Sunday",
    "item": "Sleep",
    "target": "7.5-9 hrs",
    "order": 5
  }
];
const SAUNA_DATA = [
  {
    "day": "Monday",
    "item": "Best Time",
    "detail": "Post-workout",
    "order": 1
  },
  {
    "day": "Monday",
    "item": "Duration",
    "detail": "15-20 min",
    "order": 2
  },
  {
    "day": "Monday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Monday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Monday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Tuesday",
    "item": "Best Time",
    "detail": "Evening",
    "order": 1
  },
  {
    "day": "Tuesday",
    "item": "Duration",
    "detail": "15 min",
    "order": 2
  },
  {
    "day": "Tuesday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Tuesday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Tuesday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Wednesday",
    "item": "Best Time",
    "detail": "After incline walk",
    "order": 1
  },
  {
    "day": "Wednesday",
    "item": "Duration",
    "detail": "20-25 min",
    "order": 2
  },
  {
    "day": "Wednesday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Wednesday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Wednesday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Thursday",
    "item": "Best Time",
    "detail": "Post-workout",
    "order": 1
  },
  {
    "day": "Thursday",
    "item": "Duration",
    "detail": "15-20 min",
    "order": 2
  },
  {
    "day": "Thursday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Thursday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Thursday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Friday",
    "item": "Best Time",
    "detail": "Evening recovery",
    "order": 1
  },
  {
    "day": "Friday",
    "item": "Duration",
    "detail": "20 min",
    "order": 2
  },
  {
    "day": "Friday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Friday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Friday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Saturday",
    "item": "Best Time",
    "detail": "After lean-out circuit",
    "order": 1
  },
  {
    "day": "Saturday",
    "item": "Duration",
    "detail": "15-20 min",
    "order": 2
  },
  {
    "day": "Saturday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Saturday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Saturday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  },
  {
    "day": "Sunday",
    "item": "Best Time",
    "detail": "Optional recovery",
    "order": 1
  },
  {
    "day": "Sunday",
    "item": "Duration",
    "detail": "15 min",
    "order": 2
  },
  {
    "day": "Sunday",
    "item": "Temperature",
    "detail": "170-200\u00b0F if tolerated",
    "order": 3
  },
  {
    "day": "Sunday",
    "item": "Hydration",
    "detail": "Water + electrolytes after",
    "order": 4
  },
  {
    "day": "Sunday",
    "item": "Goal",
    "detail": "Recovery / circulation / relaxation",
    "order": 5
  }
];
const UPGRADE_NOTES = [
  {
    "day": "Monday",
    "note": "Replace any chest-supported row with 1-arm cable row if duplicated elsewhere. Keep front squat. Add 1 set of weighted chin-ups if recovery good."
  },
  {
    "day": "Tuesday",
    "note": "Replace any repeated core movement with Copenhagen plank. Prioritize DB jump squat or cable high pull over redundant jumps."
  },
  {
    "day": "Wednesday",
    "note": "Use Recovery Engine: 35-45 min Zone 2 + sled drags + mobility + trunk only."
  },
  {
    "day": "Thursday",
    "note": "Increase upper-back volume: add straight-arm pulldown or rear delt fly if missing. Keep incline press."
  },
  {
    "day": "Friday",
    "note": "Keep incline + weighted dips. Add overhead rope extension if triceps volume low. Replace duplicate chop/core with face pull or curl."
  },
  {
    "day": "Saturday",
    "note": "Replace lifting with Recovery Engine (Zone 2, sled drags, mobility, carries, sauna)."
  }
];
