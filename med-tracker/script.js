// Paste this entire file into Extensions > Apps Script in your Google Sheet
// Then deploy as a web app (see README for steps)

function parseCount_(str, def) {
  const m = String(str || '').match(/(\d+)/);
  return m ? Number(m[1]) : def;
}

// Parses a date cell (Date object or "M/d/yyyy" string) → "yyyy-MM-dd"
function cellDateStr_(dateCell, tz) {
  if (dateCell instanceof Date) return Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
  const str = String(dateCell).trim();
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
  const d = new Date(str);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

// Combines a date cell and a time cell (both may be Date objects or strings)
// into a formatted 'yyyy-MM-dd HH:mm:ss' string.
// Parses text time strings directly via regex — new Date('2000-01-01 2:30 PM')
// fails in GAS V8 and returns midnight, so we can't rely on it.
function combineDateTime_(dateCell, timeCell, tz) {
  const datePart = cellDateStr_(dateCell, tz);

  let timePart;
  if (timeCell instanceof Date) {
    timePart = Utilities.formatDate(timeCell, tz, 'HH:mm:ss');
  } else {
    const str = String(timeCell).trim();
    const m   = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i);
    if (m) {
      let h = parseInt(m[1], 10);
      if ((m[4] || '').toLowerCase() === 'pm' && h < 12) h += 12;
      if ((m[4] || '').toLowerCase() === 'am' && h === 12) h = 0;
      timePart = `${String(h).padStart(2,'0')}:${m[2]}:${m[3] || '00'}`;
    } else {
      timePart = '00:00:00';
    }
  }

  return `${datePart} ${timePart}`;
}

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
    else if (action === 'getLastDoses')   result = getLastDoses();
    else if (action === 'logPain')        result = logPain(e.parameter);
    else if (action === 'getTodayPain')      result = getTodayPain();
    else if (action === 'getMilestones')     result = getMilestones();
    else if (action === 'getWeekStats')      result = getWeekStats();
    else if (action === 'startColdTherapy')  result = startColdTherapy();
    else if (action === 'stopColdTherapy')   result = stopColdTherapy(e.parameter);
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
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const meds = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    // Auto-activate if start_date = today and currently inactive
    let active = row[8];
    const isActive = active === true || String(active).toUpperCase() === 'TRUE';
    if (!isActive && row[6]) {
      const startStr = Utilities.formatDate(new Date(row[6]), tz, 'yyyy-MM-dd');
      if (startStr === todayStr) {
        sheet.getRange(i + 1, 9).setValue(true);
        active = true;
      }
    }

    if (active === true || String(active).toUpperCase() === 'TRUE') {
      const asNeeded = row[9];
      meds.push({
        id:             String(row[0]),
        name:           String(row[1]),
        dose:           String(row[2]),
        num_pills:      Number(row[3]) || 0,
        interval_hours: Number(row[4]) || 6,
        color:          String(row[5]),
        as_needed:      asNeeded === true || String(asNeeded).toUpperCase() === 'TRUE',
        notes:          String(row[10] || '')
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
    const ts      = String(row[6] || '').trim();
    const rowDate = ts ? ts.slice(0, 10) : cellDateStr_(row[0], tz);
    if (rowDate !== todayStr) continue;
    log.push({
      timestamp: ts || combineDateTime_(row[0], row[1], tz),
      med_id:    String(row[2]),
      med_name:  String(row[3]),
      dose:      String(row[4]),
      num_pills: Number(row[5]) || 0
    });
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
        num_pills: Number(medsData[i][3]) || 0
      };
      break;
    }
  }

  if (!med) return { error: 'Medication not found: ' + medId };

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  const tz = Session.getScriptTimeZone();
  const now = params.timestamp ? new Date(params.timestamp) : new Date();
  const dateVal = Utilities.formatDate(now, tz, 'M/d/yyyy');
  const timeVal = Utilities.formatDate(now, tz, 'h:mm a');
  const timestamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  logSheet.appendRow([dateVal, timeVal, med.id, med.name, med.dose, med.num_pills, timestamp]);

  return { success: true, timestamp, med: med.name };
}

function logRomSession(params) {
  const startTime       = params.startTime;
  const endTime         = params.endTime;
  const durationMinutes = Number(params.durationMinutes);

  if (!startTime || !endTime || !durationMinutes) return { error: 'startTime, endTime, and durationMinutes required' };

  const degrees = params.degrees !== undefined && params.degrees !== '' ? Number(params.degrees) : '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROM Log');
  const tz    = Session.getScriptTimeZone();
  const date  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  sheet.appendRow([date, startTime, endTime, durationMinutes, degrees]);

  return { success: true, date, durationMinutes, degrees: degrees !== '' ? degrees : null };
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
    const rowDate = cellDateStr_(row[0], tz);
    if (rowDate === todayStr) {
      sessions.push({
        date:            rowDate,
        startTime:       String(row[1]),
        endTime:         String(row[2]),
        durationMinutes: Number(row[3]) || 0,
        degrees:         row[4] !== '' && row[4] !== undefined ? Number(row[4]) : null
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
    const dateStr = cellDateStr_(row[0], tz);
    if (dateStr >= todayStr) {
      // row[1] is a time cell — Sheets returns these as Date objects anchored to Dec 30 1899
      let timeStr = '';
      if (row[1] instanceof Date) {
        const yr = row[1].getFullYear();
        if (yr === 1899 || yr === 1900) {
          timeStr = Utilities.formatDate(row[1], tz, 'h:mm a');
        }
      } else {
        timeStr = String(row[1] || '');
      }
      appts.push({
        date:             dateStr,
        time:             timeStr,
        title:            String(row[2] || ''),
        with:             String(row[3] || ''),
        location_name:    String(row[4] || ''),
        location_address: String(row[5] || ''),
        notes:            String(row[6] || '')
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
    const active = row[7];
    if (active === true || String(active).toUpperCase() === 'TRUE') {
      const setsPerSession = parseCount_(String(row[4] || '1'), 1);
      const performPerDay  = parseCount_(String(row[5] || '1'), 1);
      exercises.push({
        id:               String(row[0]),
        name:             String(row[1]),
        repeat:           String(row[2] || ''),
        hold:             String(row[3] || ''),
        sets_per_session: setsPerSession,
        perform_per_day:  performPerDay,
        daily_sets:       setsPerSession * performPerDay,
        image_filename:   String(row[6] || ''),
        instructions:     String(row[8] || ''),
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

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (!row[0] || String(row[2]) !== medId) continue;
    const ts = String(row[6] || '').trim();
    return { timestamp: ts || combineDateTime_(row[0], row[1], tz) };
  }

  return { timestamp: null };
}

function getLastDoses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  const data  = sheet.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();

  const best = {};

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if (!row[0]) continue;
    const medId = String(row[2]);
    if (!best[medId]) {
      const ts = String(row[6] || '').trim();
      best[medId] = ts || combineDateTime_(row[0], row[1], tz);
    }
  }

  return best;
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
    const rowDate = cellDateStr_(row[0], tz);
    if (rowDate === todayStr) {
      entries.push({ timestamp: String(row[0]), score: Number(row[1]) || 0, notes: String(row[2] || '') });
    }
  }
  const average = entries.length
    ? Math.round((entries.reduce((s, e) => s + e.score, 0) / entries.length) * 10) / 10
    : null;
  return { entries, average };
}

function startColdTherapy() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const tz    = Session.getScriptTimeZone();
  const now   = new Date();

  // Columns: Start Date | Start Time | End Date | End Time | Duration (hours) | Status
  sheet.appendRow([
    Utilities.formatDate(now, tz, 'M/d/yyyy'),
    Utilities.formatDate(now, tz, 'h:mm a'),
    '', '', '',
    'In Progress'
  ]);

  return { success: true };
}

function stopColdTherapy(params) {
  const durationSec = Number(params.duration_sec);
  if (!durationSec) return { error: 'duration_sec required' };

  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const tz       = Session.getScriptTimeZone();
  const now      = new Date();
  const data     = sheet.getDataRange().getValues();
  const endDate  = Utilities.formatDate(now, tz, 'M/d/yyyy');
  const endTime  = Utilities.formatDate(now, tz, 'h:mm a');
  const durHours = Math.round((durationSec / 3600) * 10000) / 10000;

  // Find the last In Progress row and fill it in
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][5]).trim().toLowerCase() === 'in progress') {
      sheet.getRange(i + 1, 3).setValue(endDate);
      sheet.getRange(i + 1, 4).setValue(endTime);
      sheet.getRange(i + 1, 5).setValue(durHours);
      sheet.getRange(i + 1, 6).setValue('Completed');
      return { success: true, duration_hours: durHours };
    }
  }

  return { error: 'No in-progress session found' };
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

    const rowDate = cellDateStr_(row[0], tz);
    if (rowDate !== todayStr) continue;

    const startTs    = combineDateTime_(row[0], row[1], tz);
    const status     = String(row[5] || '').trim();
    const inProgress = status.toLowerCase() === 'in progress';

    const session = {
      timestamp:      startTs,
      start_iso:      new Date(startTs).toISOString(),
      duration_hours: Number(row[4]) || 0,
      status:         status,
      event:          inProgress ? 'in_progress' : 'complete'
    };

    if (!inProgress && row[2] && row[3]) {
      session.end_timestamp = combineDateTime_(row[2], row[3], tz);
    }

    sessions.push(session);
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
    const ds = cellDateStr_(romData[i][0], tz);
    if (ds in romByDay) romByDay[ds] += Number(romData[i][3]) || 0;
  }

  // Pain
  const painSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pain Log');
  const painData  = painSheet.getDataRange().getValues();
  const painSum   = initMap(), painCount = initMap();
  for (let i = 1; i < painData.length; i++) {
    if (!painData[i][0]) continue;
    const ds = cellDateStr_(painData[i][0], tz);
    if (ds in painSum) { painSum[ds] += Number(painData[i][1]) || 0; painCount[ds]++; }
  }

  // PT log
  const ptLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Log');
  const ptLogData  = ptLogSheet.getDataRange().getValues();
  const ptByDay    = initMap();
  for (let i = 1; i < ptLogData.length; i++) {
    if (!ptLogData[i][0]) continue;
    const ds = cellDateStr_(ptLogData[i][0], tz);
    if (ds in ptByDay) ptByDay[ds] += Number(ptLogData[i][3]) || 0;
  }

  // PT goal (total sets per day across active exercises)
  const ptExSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PT Exercises');
  const ptExData  = ptExSheet.getDataRange().getValues();
  let ptGoal = 0;
  for (let i = 1; i < ptExData.length; i++) {
    const active = ptExData[i][7];
    if (active === true || String(active).toUpperCase() === 'TRUE') {
      ptGoal += parseCount_(String(ptExData[i][4] || '1'), 1) * parseCount_(String(ptExData[i][5] || '1'), 1);
    }
  }

  // Cold therapy
  const coldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Cold Therapy Log');
  const coldData  = coldSheet.getDataRange().getValues();
  const coldByDay = initMap();
  for (let i = 1; i < coldData.length; i++) {
    if (!coldData[i][0]) continue;
    const ds     = cellDateStr_(coldData[i][0], tz);
    const status = String(coldData[i][5]).trim().toLowerCase();
    if (ds in coldByDay && status === 'completed') coldByDay[ds]++;
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
    const rowDate = cellDateStr_(row[0], tz);
    if (rowDate !== todayStr) continue;
    const id = String(row[1]);
    totals[id] = (totals[id] || 0) + (Number(row[3]) || 0);
  }

  return totals;
}
