// Paste this entire file into Extensions > Apps Script in your Google Sheet
// Then deploy as a web app (see README for steps)

function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'getMeds')          result = getMeds();
    else if (action === 'getTodayLog') result = getTodayLog();
    else if (action === 'getConfig')   result = getConfig();
    else if (action === 'logDose')     result = logDose(e.parameter);
    else if (action === 'logRomSession')  result = logRomSession(e.parameter);
    else if (action === 'getTodayRom')    result = getTodayRom();
    else if (action === 'getAppointments') result = getAppointments();
    else if (action === 'getPtExercises')  result = getPtExercises();
    else if (action === 'logPtSets')      result = logPtSets(e.parameter);
    else if (action === 'getTodayPt')     result = getTodayPt();
    else if (action === 'getLastDose')    result = getLastDose(e.parameter);
    else if (action === 'logPain')        result = logPain(e.parameter);
    else if (action === 'getTodayPain')      result = getTodayPain();
    else if (action === 'getMilestones')     result = getMilestones();
    else if (action === 'getWeekStats')      result = getWeekStats();
    else if (action === 'logColdTherapy')    result = logColdTherapy(e.parameter);
    else if (action === 'getColdTherapyLog') result = getColdTherapyLog();
    else result = { error: 'Unknown action: ' + action };
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMeds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Medications');
  const data = sheet.getDataRange().getValues();
  const meds = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const active = row[6];
    if (active === true || String(active).toUpperCase() === 'TRUE') {
      meds.push({
        id:             String(row[0]),
        name:           String(row[1]),
        dose:           String(row[2]),
        apap_mg:        Number(row[3]) || 0,
        interval_hours: Number(row[4]) || 6,
        color:          String(row[5])
      });
    }
  }

  return meds;
}

function getTodayLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const log = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate === todayStr) {
      log.push({
        timestamp: row[0],
        med_id:    String(row[1]),
        med_name:  String(row[2]),
        dose:      String(row[3]),
        apap_mg:   Number(row[4]) || 0
      });
    }
  }

  return log;
}

function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (const row of data) {
    if (row[0]) config[String(row[0])] = row[1];
  }
  return config;
}

function logDose(params) {
  const medId = params.medId;
  if (!medId) return { error: 'medId required' };

  const medsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Medications');
  const medsData = medsSheet.getDataRange().getValues();

  let med = null;
  for (let i = 1; i < medsData.length; i++) {
    if (String(medsData[i][0]) === medId) {
      med = {
        id:      String(medsData[i][0]),
        name:    String(medsData[i][1]),
        dose:    String(medsData[i][2]),
        apap_mg: Number(medsData[i][3]) || 0
      };
      break;
    }
  }

  if (!med) return { error: 'Medication not found: ' + medId };

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  logSheet.appendRow([timestamp, med.id, med.name, med.dose, med.apap_mg]);

  return { success: true, timestamp, med: med.name };
}

function logRomSession(params) {
  const startTime       = params.startTime;
  const endTime         = params.endTime;
  const durationMinutes = Number(params.durationMinutes);

  if (!startTime || !endTime || !durationMinutes) return { error: 'startTime, endTime, and durationMinutes required' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROM Log');
  const tz    = Session.getScriptTimeZone();
  const date  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  sheet.appendRow([date, startTime, endTime, durationMinutes]);

  return { success: true, date, durationMinutes };
}

function getTodayRom() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROM Log');
  const data  = sheet.getDataRange().getValues();

  if (data.length <= 1) return { sessions: [], totalMinutes: 0 };

  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sessions = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate === todayStr) {
      sessions.push({
        date:            rowDate,
        startTime:       String(row[1]),
        endTime:         String(row[2]),
        durationMinutes: Number(row[3]) || 0
      });
    }
  }

  const totalMinutes = sessions.reduce((s, r) => s + r.durationMinutes, 0);
  return { sessions, totalMinutes };
}

function getAppointments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  const data  = sheet.getDataRange().getValues();

  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const appts    = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const dateStr = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (dateStr >= todayStr) {
      appts.push({
        date:             dateStr,
        time:             row[1] instanceof Date ? Utilities.formatDate(row[1], tz, 'h:mm a') : String(row[1] || ''),
        title:            String(row[2] || ''),
        location_name:    String(row[3] || ''),
        location_address: String(row[4] || ''),
        notes:            String(row[5] || '')
      });
    }
  }

  appts.sort((a, b) => a.date.localeCompare(b.date));
  return appts;
}

function getPtExercises() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Exercises');
  const data  = sheet.getDataRange().getValues();
  const exercises = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const active = row[5];
    if (active === true || String(active).toUpperCase() === 'TRUE') {
      exercises.push({
        id:             String(row[0]),
        name:           String(row[1]),
        reps_per_set:   String(row[2]),
        sets:           Number(row[3]) || 3,
        image_filename: String(row[4] || ''),
      });
    }
  }

  return exercises;
}

function logPtSets(params) {
  const exerciseId   = params.exerciseId;
  const exerciseName = params.exerciseName;
  const setsAdded    = Number(params.setsAdded) || 1;

  if (!exerciseId) return { error: 'exerciseId required' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Log');
  const tz    = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([timestamp, exerciseId, exerciseName, setsAdded]);

  return { success: true, timestamp, exerciseId, setsAdded };
}

function getLastDose(params) {
  const medId = params.medId;
  if (!medId) return { timestamp: null };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();

  let bestMs        = 0;
  let bestTimestamp = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || String(row[1]) !== medId) continue;
    const ms = new Date(row[0]).getTime();
    if (ms > bestMs) {
      bestMs        = ms;
      bestTimestamp = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd HH:mm:ss');
    }
  }

  return { timestamp: bestTimestamp };
}

function logPain(params) {
  const score = Number(params.score);
  if (!score || score < 1 || score > 10) return { error: 'score must be 1-10' };
  const notes = params.notes ? String(params.notes) : '';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pain Log');
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([timestamp, score, notes]);
  return { success: true, timestamp, score };
}

function getTodayPain() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pain Log');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { entries: [], average: null };
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate === todayStr) {
      entries.push({ timestamp: String(row[0]), score: Number(row[1]) || 0, notes: String(row[2] || '') });
    }
  }
  const average = entries.length
    ? Math.round((entries.reduce((s, e) => s + e.score, 0) / entries.length) * 10) / 10
    : null;
  return { entries, average };
}

function logColdTherapy(params) {
  const event       = params.event;
  const durationSec = Number(params.duration_sec);

  if (!event || !durationSec) return { error: 'event and duration_sec required' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const tz    = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([timestamp, event, durationSec]);

  return { success: true, timestamp, event, duration_sec: durationSec };
}

function getColdTherapyLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const data  = sheet.getDataRange().getValues();

  if (data.length <= 1) return { sessions: [] };

  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sessions = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate === todayStr) {
      const event = String(row[1]).trim().toLowerCase();
      sessions.push({
        timestamp:    String(row[0]),
        event:        event === 'stopped' ? 'stopped' : 'complete',
        duration_sec: Number(row[2]) || 0
      });
    }
  }

  return { sessions };
}

function getWeekStats() {
  const tz    = Session.getScriptTimeZone();
  const today = new Date();
  const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  const days = [], dayLabels = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    days.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
    dayLabels.push(DAY_NAMES[d.getDay()]);
  }

  function initMap() { const m = {}; days.forEach(d => m[d] = 0); return m; }

  // ROM
  const romSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROM Log');
  const romData  = romSheet.getDataRange().getValues();
  const romByDay = initMap();
  for (let i = 1; i < romData.length; i++) {
    if (!romData[i][0]) continue;
    const ds = Utilities.formatDate(new Date(romData[i][0]), tz, 'yyyy-MM-dd');
    if (ds in romByDay) romByDay[ds] += Number(romData[i][3]) || 0;
  }

  // Pain
  const painSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pain Log');
  const painData  = painSheet.getDataRange().getValues();
  const painSum   = initMap(), painCount = initMap();
  for (let i = 1; i < painData.length; i++) {
    if (!painData[i][0]) continue;
    const ds = Utilities.formatDate(new Date(painData[i][0]), tz, 'yyyy-MM-dd');
    if (ds in painSum) { painSum[ds] += Number(painData[i][1]) || 0; painCount[ds]++; }
  }

  // PT log
  const ptLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Log');
  const ptLogData  = ptLogSheet.getDataRange().getValues();
  const ptByDay    = initMap();
  for (let i = 1; i < ptLogData.length; i++) {
    if (!ptLogData[i][0]) continue;
    const ds = Utilities.formatDate(new Date(ptLogData[i][0]), tz, 'yyyy-MM-dd');
    if (ds in ptByDay) ptByDay[ds] += Number(ptLogData[i][3]) || 0;
  }

  // PT goal (total sets per day across active exercises)
  const ptExSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Exercises');
  const ptExData  = ptExSheet.getDataRange().getValues();
  let ptGoal = 0;
  for (let i = 1; i < ptExData.length; i++) {
    const active = ptExData[i][5];
    if (active === true || String(active).toUpperCase() === 'TRUE') ptGoal += Number(ptExData[i][3]) || 0;
  }

  // Cold therapy
  const coldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const coldData  = coldSheet.getDataRange().getValues();
  const coldByDay = initMap();
  for (let i = 1; i < coldData.length; i++) {
    if (!coldData[i][0]) continue;
    const ds    = Utilities.formatDate(new Date(coldData[i][0]), tz, 'yyyy-MM-dd');
    const event = String(coldData[i][1]).trim().toLowerCase();
    if (ds in coldByDay && event !== 'stopped') coldByDay[ds]++;
  }

  return {
    days,
    dayLabels,
    rom:    days.map(d => romByDay[d]),
    pain:   days.map(d => painCount[d] > 0 ? Math.round(painSum[d] / painCount[d] * 10) / 10 : null),
    pt:     days.map(d => ptByDay[d]),
    ptGoal,
    cold:   days.map(d => coldByDay[d])
  };
}

function getMilestones() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Milestones');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();
  const milestones = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    milestones.push({
      id:    String(row[0]),
      name:  String(row[1]),
      date:  row[2] ? Utilities.formatDate(new Date(row[2]), tz, 'yyyy-MM-dd') : null,
      notes: String(row[3] || '')
    });
  }

  milestones.sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  return milestones;
}

function getTodayPt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Log');
  const data  = sheet.getDataRange().getValues();

  if (data.length <= 1) return {};

  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const totals   = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = Utilities.formatDate(new Date(row[0]), tz, 'yyyy-MM-dd');
    if (rowDate !== todayStr) continue;
    const id = String(row[1]);
    totals[id] = (totals[id] || 0) + (Number(row[3]) || 0);
  }

  return totals;
}
