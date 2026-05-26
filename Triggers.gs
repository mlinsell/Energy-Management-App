/**
 * @fileoverview Automated triggers and scheduled background tasks for CEMS.
 */

/**
 * Flushes old triggers and installs the core automation schedule.
 */
function installTriggers() {
  removeTriggers();

  // 15-minute sync
  ScriptApp.newTrigger('syncCalendar')
    .timeBased()
    .everyMinutes(15)
    .create();

  // Midnight archive 
  ScriptApp.newTrigger('midnightReset')
    .timeBased()
    .atHour(0)
    .nearMinute(5)
    .everyDays(1)
    .create();

  // Weekly analytics
  ScriptApp.newTrigger('weeklyAnalytics')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  logEvent('installTriggers', 'SUCCESS', 'All scheduled triggers installed successfully.');
}

/**
 * Removes all triggers associated with this Apps Script project.
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

/**
 * Safely moves yesterday's data from DailyLog to Archive, ensuring no data loss.
 */
function midnightReset() {
  try {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayStr = getDateString(yesterday);

    const logSheet = getSheetByName('DailyLog');
    const rows = getRowsForDate(logSheet, yesterdayStr);

    if (rows.length === 0) {
      logEvent('midnightReset', 'SUCCESS', `No data to archive for ${yesterdayStr}`);
      return;
    }

    const archiveSheet = getSheetByName('Archive');
    const oldCount = archiveSheet.getLastRow();
    const timestamp = new Date().toISOString();
    
    // Map rows to append the DateArchived column
    const rowsToArchive = rows.map(row => [...row, timestamp]);
    
    // Bulk append to Archive sheet
    archiveSheet.getRange(oldCount + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    SpreadsheetApp.flush();
    
    const newCount = archiveSheet.getLastRow();
    
    if (newCount === oldCount + rowsToArchive.length) {
      // Verification passed: Delete rows backwards to prevent index shifting
      for (let i = logSheet.getLastRow(); i >= 2; i--) {
        const cellDate = logSheet.getRange(i, 1).getValue();
        const cellDateStr = (cellDate instanceof Date) ? getDateString(cellDate) : String(cellDate);
        if (cellDateStr === yesterdayStr) {
          logSheet.deleteRow(i);
        }
      }
      logEvent('midnightReset', 'SUCCESS', `Archived ${rows.length} rows for ${yesterdayStr}`);
    } else {
      throw new Error('Archive verification failed (Row count mismatch) — DailyLog NOT cleared');
    }

  } catch (error) {
    logEvent('midnightReset', 'ERROR', error.toString());
    const email = Session.getActiveUser().getEmail();
    if (email) {
      GmailApp.sendEmail(
        email, 
        'CEMS Archive Failed', 
        'The midnight archive failed verification. DailyLog has not been cleared. Check SystemLog.'
      );
    }
  }
}

/**
 * Calculates aggregated statistics for the previous week (Mon-Sun) and saves to WeeklyAnalytics.
 */
function weeklyAnalytics() {
  try {
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0 = Sun
    const daysSinceMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    
    // Calculate last Monday
    const lastMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysSinceMonday - 7);
    const targetDates = [];
    
    for (let i = 0; i < 7; i++) {
      const d = new Date(lastMonday);
      d.setDate(lastMonday.getDate() + i);
      targetDates.push(getDateString(d));
    }

    const archiveSheet = getSheetByName('Archive');
    const allData = archiveSheet.getDataRange().getValues().slice(1); // skip header
    const idxDate = getColumnIndex(archiveSheet, 'Date');
    const idxBurnTier = getColumnIndex(archiveSheet, 'BurnTier');
    const idxDuration = getColumnIndex(archiveSheet, 'Duration_Min');
    const idxPoints = getColumnIndex(archiveSheet, 'Points');
    const idxStart = getColumnIndex(archiveSheet, 'StartTime');
    const idxEnd = getColumnIndex(archiveSheet, 'EndTime');

    // Group rows by target dates
    const groupedDays = {};
    allData.forEach(row => {
      const dateVal = row[idxDate];
      const rowDateStr = (dateVal instanceof Date) ? getDateString(dateVal) : String(dateVal);
      
      if (targetDates.includes(rowDateStr)) {
        if (!groupedDays[rowDateStr]) groupedDays[rowDateStr] = [];
        groupedDays[rowDateStr].push(row);
      }
    });

    const dailyScores = [];
    let overWarningCount = 0;
    let totalHighBurnMins = 0;
    let totalRecoveryMins = 0;

    const warningThreshold = getSettingAsNumber('Warning_Threshold');
    const recoveryCap = getSettingAsNumber('RecoveryCreditCap');
    const deskBaseline = getSettingAsNumber('DeskBaseline_Points');
    const penaltyEnabled = getSettingAsBool('ConsecutivePenalty_Enabled');
    const penaltyVal = getSettingAsNumber('ConsecutivePenalty_Value');
    const penaltyGap = getSettingAsNumber('ConsecutiveGap_Minutes');

    targetDates.forEach(dateStr => {
      const rows = groupedDays[dateStr] || [];
      if (rows.length === 0) return; // Skip days with no data entirely

      const dayType = getDayType(dateStr);
      let eventPts = 0;
      let rawRecovery = 0;
      let highBurnEvents = [];

      rows.forEach(row => {
        const tier = row[idxBurnTier];
        const mins = parseFloat(row[idxDuration]) || 0;
        const pts = parseFloat(row[idxPoints]) || 0;
        
        if (tier === 'Recovery') {
          rawRecovery += pts;
          totalRecoveryMins += mins;
        } else {
          eventPts += pts;
        }
        if (tier === 'High Burn') {
          totalHighBurnMins += mins;
          
          let stStr = (row[idxStart] instanceof Date) ? formatTime(row[idxStart]) : String(row[idxStart]);
          let endStr = (row[idxEnd] instanceof Date) ? formatTime(row[idxEnd]) : String(row[idxEnd]);
          highBurnEvents.push({ startStr: stStr, endStr: endStr });
        }
      });

      const baseline = (dayType === 'work') ? deskBaseline : 0;
      const clampedRec = (rawRecovery < recoveryCap) ? recoveryCap : rawRecovery;

      let penaltyPts = 0;
      if (penaltyEnabled) {
        const refD = new Date();
        highBurnEvents.sort((a, b) => parseTimeString(a.startStr, refD).getTime() - parseTimeString(b.startStr, refD).getTime());
        
        for (let i = 1; i < highBurnEvents.length; i++) {
          const prevE = parseTimeString(highBurnEvents[i - 1].endStr, refD);
          const currS = parseTimeString(highBurnEvents[i].startStr, refD);
          if (minutesBetween(prevE, currS) < penaltyGap) {
            penaltyPts += penaltyVal;
          }
        }
      }

      const dailyTotal = eventPts + baseline + clampedRec + penaltyPts;
      dailyScores.push({ total: dailyTotal, dayType: dayType });

      if (dailyTotal >= warningThreshold) {
        overWarningCount++;
      }
    });

    // Compute Averages safely
    let sumAll = 0, sumWork = 0, sumAuto = 0;
    let countWork = 0, countAuto = 0;
    let maxPts = 0, minPts = 0;

    if (dailyScores.length > 0) {
      maxPts = Math.max(...dailyScores.map(d => d.total));
      minPts = Math.min(...dailyScores.map(d => d.total));

      dailyScores.forEach(d => {
        sumAll += d.total;
        if (d.dayType === 'work') {
          sumWork += d.total;
          countWork++;
        } else {
          sumAuto += d.total;
          countAuto++;
        }
      });
    }

    const avgAll = dailyScores.length ? (sumAll / dailyScores.length).toFixed(2) : 0;
    const avgWork = countWork ? (sumWork / countWork).toFixed(2) : 0;
    const avgAuto = countAuto ? (sumAuto / countAuto).toFixed(2) : 0;

    const rowToAppend = [
      targetDates[0], // WeekStartDate
      avgAll,
      avgWork,
      avgAuto,
      maxPts,
      minPts,
      overWarningCount,
      totalHighBurnMins,
      totalRecoveryMins
    ];

    const weeklySheet = getSheetByName('WeeklyAnalytics');
    appendRow(weeklySheet, rowToAppend);

    logEvent('weeklyAnalytics', 'SUCCESS', `Processed analytics for week of ${targetDates[0]}`);

  } catch (error) {
    logEvent('weeklyAnalytics', 'ERROR', error.toString());
  }
}
