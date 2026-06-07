/**
 * @fileoverview Core business logic and backend routing for CEMS
 */

// --- CONSTANTS ---
const BURN_TIERS = {
  11: { name: 'High Burn',   multiplier:  2.0 },
   6: { name: 'Medium Burn', multiplier:  1.0 },
   5: { name: 'Low Burn',    multiplier:  0.5 },
   1: { name: 'Desk',        multiplier:  0.25 }, // Desk baseline tier
   2: { name: 'Recovery',    multiplier: -0.5 },
   0: { name: 'Neutral',     multiplier:  0.0 }
};

const COLOUR_TO_TIER_ID = {
  "11": 11,
  "6": 6,
  "5": 5,
  "8": 1, 
  "2": 2,
  "": 0
};

// Global settings cache for the execution run
let SETTINGS_CACHE = null;

// --- WEB APP ROUTING ---
/**
 * Web App entry point rendering the CEMS UI using HTML templates.
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet() {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('CEMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Helper function to inject external HTML files into the main template.
 * @param {string} filename - Name of the HTML file to include (without extension).
 * @returns {string} Raw HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- UTILITIES & LOGGING ---
/**
 * Reads the Settings sheet and returns the value for a given parameter via an in-memory cache.
 */
function getSetting(parameterName) {
  if (!SETTINGS_CACHE) {
    const sheet = getSheetByName('Settings');
    const data = getSheetData(sheet); // Uses Utils cache
    const idxParam = getColumnIndex(sheet, 'A (Parameter)');
    const idxVal = getColumnIndex(sheet, 'B (Value)');

    SETTINGS_CACHE = {};
    data.forEach(row => {
      if (row[idxParam]) {
        SETTINGS_CACHE[row[idxParam]] = row[idxVal];
      }
    });
  }
  
  if (SETTINGS_CACHE[parameterName] === undefined) {
    throw new Error(`Setting parameter '${parameterName}' not found in Settings sheet.`);
  }
  
  return String(SETTINGS_CACHE[parameterName]);
}

function getSettingAsNumber(parameterName) {
  const val = getSetting(parameterName);
  const num = parseFloat(val);
  if (isNaN(num)) throw new Error(`Setting '${parameterName}' requires a numeric value, got '${val}'.`);
  return num;
}

function getSettingAsBool(parameterName) {
  const val = getSetting(parameterName);
  return val.trim().toUpperCase() === "TRUE";
}

function logEvent(functionName, status, message) {
  try {
    const sheet = getSheetByName('SystemLog');
    const timestamp = new Date().toISOString();
    appendRow(sheet, [timestamp, functionName, status, message]);
  } catch (error) {
    // Fail silently
  }
}

function calculatePoints(durationMinutes, multiplier) {
  const points = (durationMinutes / 30) * multiplier;
  return Math.round(points * 100) / 100;
}


// --- DAY MANAGEMENT ---
function getDayType(dateString) {
  const overridesSheet = getSheetByName('DayOverrides');
  const overrideRow = getRowsForDate(overridesSheet, dateString);
  
  if (overrideRow.length > 0) {
    const dayTypeIndex = getColumnIndex(overridesSheet, 'DayType');
    return String(overrideRow[0][dayTypeIndex]).toLowerCase();
  }
  
  const parts = dateString.split('-');
  const dateObj = new Date(parts[0], parts[1] - 1, parts[2]);
  const day = dateObj.getDay(); 
  
  if (day >= 1 && day <= 4) return 'work';
  return 'autonomous';
}

function setDayOverride(dateString, dayType) {
  const validTypes = ['work', 'autonomous', 'rest'];
  const typeLower = dayType.toLowerCase();
  if (!validTypes.includes(typeLower)) throw new Error(`Invalid dayType '${dayType}'.`);

  const sheet = getSheetByName('DayOverrides');
  const rowNum = findRowByColumnValue(sheet, 'Date', dateString);
  const typeColIndex = getColumnIndex(sheet, 'DayType') + 1;

  if (rowNum !== -1) {
    sheet.getRange(rowNum, typeColIndex).setValue(typeLower);
    if(typeof EXECUTION_CACHE !== 'undefined') EXECUTION_CACHE.data['DayOverrides'] = null;
  } else {
    appendRow(sheet, [dateString, typeLower]);
  }
}

// --- BASELINE EVENT GENERATION ---
function isBaselineDisabled(dateString) {
  const sheet = getSheetByName('DailyLog');
  const data = getSheetData(sheet);
  const idxDate = getColumnIndex(sheet, 'Date');
  const idxID = getColumnIndex(sheet, 'CalendarEventID');
  const idxName = getColumnIndex(sheet, 'EventName');

  return data.some(row => {
    const rowDate = (row[idxDate] instanceof Date) ? getDateString(row[idxDate]) : String(row[idxDate]);
    return rowDate === dateString && row[idxID] === 'SYSTEM_BASELINE' && String(row[idxName]).includes('(Off)');
  });
}

function ensureBaselineEvent(dateString) {
  if (getDayType(dateString) !== 'work' || isBaselineDisabled(dateString)) return;
  
  const sheet = getSheetByName('DailyLog');
  const data = getSheetData(sheet);
  const idxDate = getColumnIndex(sheet, 'Date');
  const idxID = getColumnIndex(sheet, 'CalendarEventID');
  
  let exists = false;
  for (let i = 0; i < data.length; i++) {
    const d = (data[i][idxDate] instanceof Date) ? getDateString(data[i][idxDate]) : String(data[i][idxDate]);
    if (d === dateString && data[i][idxID] === 'SYSTEM_BASELINE') {
      exists = true;
      break;
    }
  }
  
  if (!exists) {
    const pts = getSettingAsNumber('DeskBaseline_Points');
    appendRow(sheet, [dateString, 'Desk Baseline', '09:00', '17:00', 480, 'Desk', pts, 0, false, false, true, 'SYSTEM_BASELINE']);
    if(typeof EXECUTION_CACHE !== 'undefined') EXECUTION_CACHE.data['DailyLog'] = null;
  }
}


// --- CORE DATA FETCHING ---
function getDailyTotals(dateString) {
  const sheet = getSheetByName('DailyLog');
  const rows = getRowsForDate(sheet, dateString);
  const dayType = getDayType(dateString);
  
  const idxStartTime = getColumnIndex(sheet, 'StartTime');
  const idxEndTime = getColumnIndex(sheet, 'EndTime');
  const idxBurnTier = getColumnIndex(sheet, 'BurnTier');
  const idxPoints = getColumnIndex(sheet, 'Points');
  
  let eventPoints = 0;
  let rawRecoveryCredit = 0;
  let penalties = 0;
  let highBurnEvents = [];

  rows.forEach(row => {
    const tier = row[idxBurnTier];
    const pts = parseFloat(row[idxPoints]) || 0;
    const startTimeStr = (row[idxStartTime] instanceof Date) ? formatTime(row[idxStartTime]) : String(row[idxStartTime]);
    const endTimeStr = (row[idxEndTime] instanceof Date) ? formatTime(row[idxEndTime]) : String(row[idxEndTime]);

    if (tier === 'Recovery') rawRecoveryCredit += pts; 
    else eventPoints += pts;

    if (tier === 'High Burn') highBurnEvents.push({ startStr: startTimeStr, endStr: endTimeStr });
  });

  const cap = getSettingAsNumber('RecoveryCreditCap');
  const recoveryCredit = (rawRecoveryCredit < cap) ? cap : rawRecoveryCredit;

  if (getSettingAsBool('ConsecutivePenalty_Enabled')) {
    const penaltyValue = getSettingAsNumber('ConsecutivePenalty_Value');
    const gapThreshold = getSettingAsNumber('ConsecutiveGap_Minutes');
    const refDate = new Date();
    highBurnEvents.sort((a, b) => parseTimeString(a.startStr, refDate).getTime() - parseTimeString(b.startStr, refDate).getTime());

    for (let i = 1; i < highBurnEvents.length; i++) {
      const prevEnd = parseTimeString(highBurnEvents[i - 1].endStr, refDate);
      const currStart = parseTimeString(highBurnEvents[i].startStr, refDate);
      
      if (minutesBetween(prevEnd, currStart) < gapThreshold) {
        penalties += penaltyValue;
      }
    }
  }

  let totalPoints = eventPoints + recoveryCredit + penalties;
  totalPoints = Math.round(totalPoints * 100) / 100;

  const safeThreshold = getSettingAsNumber('Safe_Threshold');
  const warningThreshold = getSettingAsNumber('Warning_Threshold');
  const allowance = dayType === 'work' ? getSettingAsNumber('DailyAllowance_WorkDay') : getSettingAsNumber('DailyAllowance_RestDay');

  let status = 'warning';
  if (totalPoints <= safeThreshold) status = 'safe';
  else if (totalPoints >= warningThreshold) status = 'critical';

  return {
    totalPoints: totalPoints,
    eventPoints: Math.round(eventPoints * 100) / 100,
    baselinePoints: 0, // Kept at 0 to avoid breaking frontend elements relying on this property
    recoveryCredit: recoveryCredit,
    penalties: Math.round(penalties * 100) / 100,
    allowance: allowance,
    dayType: dayType,
    status: status
  };
}

function getAppData() {
  try {
    const todayStr = getDateString(new Date());
    // Inject baseline if needed before calculations
    ensureBaselineEvent(todayStr); 
    
    const totals = getDailyTotals(todayStr);
    const events = getEventsForDate(todayStr);
    
    const alerts = [];
    
    // --- OVERLAP DETECTION ---
    // Exclude the desk baseline from throwing false positive overlap warnings
    const schedulableEvents = events.filter(e => e.CalendarEventID !== 'SYSTEM_BASELINE');
    let overlapFound = false;
    let maxEnd = "00:00";

    for (let i = 0; i < schedulableEvents.length; i++) {
      const start = schedulableEvents[i].StartTime || "00:00";
      const end = schedulableEvents[i].EndTime || "00:00";

      if (start < maxEnd) {
        overlapFound = true;
        break;
      }
      if (end > maxEnd) {
        maxEnd = end;
      }
    }

    if (overlapFound) {
      alerts.push({ type: 'overlap', message: '⚠ Warning: Overlapping events detected.' });
    }
    // -------------------------

    const weekDays = [];
    const today = new Date();
    const dayOfWeek = today.getDay(); 
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    
    for (let i = 0; i < 7; i++) {
      const d = new Date(today);
      d.setDate(today.getDate() + mondayOffset + i);
      const ds = getDateString(d);
      
      // Inject baseline for each day of week
      ensureBaselineEvent(ds);
      
      const dt = getDailyTotals(ds);
      weekDays.push({ 
        date: ds, dayType: dt.dayType, totalPoints: dt.totalPoints, status: dt.status, allowance: dt.allowance 
      });
    }
    
    const settingKeys = [
      'DailyAllowance_WorkDay','DailyAllowance_RestDay','DeskBaseline_Points',
      'RecoveryCreditCap','ConsecutivePenalty_Enabled','ConsecutivePenalty_Value',
      'ConsecutiveGap_Minutes','Safe_Threshold','Warning_Threshold'
    ];

    const settings = {};
    settingKeys.forEach(k => {
      try { settings[k] = getSetting(k); } catch (e) { settings[k] = null; }
    });

    return {
      success: true,
      today: { date: todayStr, dayType: totals.dayType, totals, events, alerts },
      week:  { days: weekDays },
      settings
    };
  } catch(e) {
    logEvent('getAppData', 'ERROR', e.toString());
    return { success: false, message: e.toString() };
  }
}

function getAnalyticsData(period) {
  try {
    const daysToSubtract = period === '7days' ? 7 : (period === '30days' ? 30 : 90);
    const refDate = new Date();
    const endDateStr = getDateString(refDate);
    const startDate = new Date();
    startDate.setDate(refDate.getDate() - daysToSubtract + 1);
    const startDateStr = getDateString(startDate);

    const baselinePoints = getSettingAsNumber('DeskBaseline_Points');
    const warningThreshold = getSettingAsNumber('Warning_Threshold');
    const recoveryCap = getSettingAsNumber('RecoveryCreditCap');

    const archiveSheet = getSheetByName('Archive');
    const dailyLogSheet = getSheetByName('DailyLog');
    
    const archiveData = getSheetData(archiveSheet);
    const dailyLogData = getSheetData(dailyLogSheet); 
    const combinedData = [...archiveData, ...dailyLogData];

    const idxDate = getColumnIndex(dailyLogSheet, 'Date');
    const idxBurnTier = getColumnIndex(dailyLogSheet, 'BurnTier');
    const idxDuration = getColumnIndex(dailyLogSheet, 'Duration_Min');
    const idxPoints = getColumnIndex(dailyLogSheet, 'Points');
    const idxName = getColumnIndex(dailyLogSheet, 'EventName');

    const groupedDays = {};
    combinedData.forEach(row => {
      const dateVal = row[idxDate];
      const rowDateStr = (dateVal instanceof Date) ? getDateString(dateVal) : String(dateVal);
      
      if (rowDateStr >= startDateStr && rowDateStr <= endDateStr) {
        if (!groupedDays[rowDateStr]) groupedDays[rowDateStr] = { points: 0, rawRecovery: 0, dayType: getDayType(rowDateStr), hasBaselineEvent: false };
        
        const tier = row[idxBurnTier];
        const pts = parseFloat(row[idxPoints]) || 0;
        const name = String(row[idxName]);
        
        if (tier === 'Recovery') groupedDays[rowDateStr].rawRecovery += pts;
        else groupedDays[rowDateStr].points += pts;
        
        if (name === 'Desk Baseline') groupedDays[rowDateStr].hasBaselineEvent = true;
      }
    });

    const result = {
      success: true, period: period, avgAll: 0, avgWorkDays: 0, avgAutonomousDays: 0,
      maxDay: { date: '-', points: 0 }, minDay: { date: '-', points: 999 },
      daysOverWarning: 0, totalHighBurnMins: 0, totalRecoveryMins: 0, consecutiveOverloadStreaks: 0,
      tierDistribution: { highBurnMins: 0, mediumBurnMins: 0, lowBurnMins: 0, deskMins: 0, recoveryMins: 0, neutralMins: 0 },
      dailyPoints: []
    };

    combinedData.forEach(row => {
      const dateVal = row[idxDate];
      const rowDateStr = (dateVal instanceof Date) ? getDateString(dateVal) : String(dateVal);
      if (rowDateStr >= startDateStr && rowDateStr <= endDateStr) {
        const tier = row[idxBurnTier];
        const mins = parseFloat(row[idxDuration]) || 0;
        
        if (tier === 'High Burn') { result.totalHighBurnMins += mins; result.tierDistribution.highBurnMins += mins; }
        else if (tier === 'Medium Burn') result.tierDistribution.mediumBurnMins += mins;
        else if (tier === 'Low Burn') result.tierDistribution.lowBurnMins += mins;
        else if (tier === 'Desk') result.tierDistribution.deskMins += mins; 
        else if (tier === 'Recovery') { result.totalRecoveryMins += mins; result.tierDistribution.recoveryMins += mins; }
        else result.tierDistribution.neutralMins += mins;
      }
    });

    let sumAll = 0, sumWork = 0, sumAuto = 0, countWork = 0, countAuto = 0, currentStreak = 0;
    const sortedDates = Object.keys(groupedDays).sort();
    
    sortedDates.forEach(dateStr => {
      const dayData = groupedDays[dateStr];
      
      // Backward compatibility for pre-event-baseline historical data
      const baseline = (dayData.dayType === 'work' && !dayData.hasBaselineEvent && !isBaselineDisabled(dateStr)) ? baselinePoints : 0;
      const clampedRecovery = (dayData.rawRecovery < recoveryCap) ? recoveryCap : dayData.rawRecovery;
      
      let dailyTotal = dayData.points + baseline + clampedRecovery;
      dailyTotal = Math.round(dailyTotal * 100) / 100;
  
      let status = 'safe';
      if (dailyTotal >= warningThreshold) status = 'critical';

      if (dailyTotal >= warningThreshold) {
        currentStreak++;
        if (currentStreak === 2) result.consecutiveOverloadStreaks++; 
      } else {
        currentStreak = 0;
      }

      result.dailyPoints.push({ date: dateStr, points: dailyTotal, status: status });
      sumAll += dailyTotal;
      
      if (dayData.dayType === 'work') { sumWork += dailyTotal; countWork++; }
      else { sumAuto += dailyTotal; countAuto++; }

      if (dailyTotal >= warningThreshold) result.daysOverWarning++;
      if (dailyTotal > result.maxDay.points) { result.maxDay.points = dailyTotal; result.maxDay.date = dateStr; }
      if (dailyTotal < result.minDay.points) { result.minDay.points = dailyTotal; result.minDay.date = dateStr; }
    });

    const countAll = sortedDates.length;
    if (countAll > 0) {
      result.avgAll = Number((sumAll / countAll).toFixed(1));
      result.avgWorkDays = countWork ? Number((sumWork / countWork).toFixed(1)) : 0;
      result.avgAutonomousDays = countAuto ? Number((sumAuto / countAuto).toFixed(1)) : 0;
    } else {
      result.minDay.points = 0; 
    }

    return result;

  } catch (error) {
    logEvent('getAnalyticsData', 'ERROR', error.toString());
    return { success: false, message: error.toString() };
  }
}
