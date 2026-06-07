/**
 * @fileoverview Calendar synchronization and manual event entry functions.
 */

/**
 * Synchronizes events from the specified Google Calendar into the DailyLog sheet.
 */
function syncCalendar(daysBack = 0, daysAhead = 2) {
  try {
    if (typeof daysBack !== 'number') daysBack = 0;
    if (typeof daysAhead !== 'number') daysAhead = 2;

    const calendarIdStr = getSetting('CalendarID');

    if (!calendarIdStr || calendarIdStr.trim() === '' || calendarIdStr.includes('(leave blank')) {
      throw new Error('CalendarID is blank or invalid in Settings sheet.');
    }

    const calendar = CalendarApp.getCalendarById(calendarIdStr.trim());
    if (!calendar) {
      throw new Error(`Calendar with ID '${calendarIdStr}' not found. Verify permissions.`);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const startDate = new Date(today);
    startDate.setDate(startDate.getDate() - daysBack);
    
    const endDate = new Date(today);
    endDate.setDate(endDate.getDate() + daysAhead);
    endDate.setHours(23, 59, 59, 999);

    const events = calendar.getEvents(startDate, endDate);
    const sheet = getSheetByName('DailyLog');
    const manualOverrideColIndex = getColumnIndex(sheet, 'ManualOverride') + 1; 

    events.forEach(event => {
      const title = event.getTitle();
      const eventStartDate = event.getStartTime();
      const eventEndDate = event.getEndTime();
      const colorId = event.getColor();
      const eventId = event.getId();

      let tierId = COLOUR_TO_TIER_ID[colorId];
      if (colorId === "" || colorId === undefined || tierId === undefined) {
        tierId = 0; 
      }

      const tier = BURN_TIERS[tierId];
      const durationMin = minutesBetween(eventStartDate, eventEndDate);
      const points = calculatePoints(durationMin, tier.multiplier);
      
      const dateString = getDateString(eventStartDate);
      const startTimeStr = formatTime(eventStartDate);
      const endTimeStr = formatTime(eventEndDate);

      let isCompleted = false;
      const existingRowNum = findRowByColumnValue(sheet, 'CalendarEventID', eventId);

      // Extract existing completed state if the row already exists so sync doesn't un-tick it
      if (existingRowNum !== -1) {
        try {
          const completedColIdx = getColumnIndex(sheet, 'Completed') + 1;
          isCompleted = sheet.getRange(existingRowNum, completedColIdx).getValue() === true;
        } catch(e) { } // Fails safely if user hasn't renamed the column yet
      }

      const rowArray = [
        dateString,
        title,
        startTimeStr,
        endTimeStr,
        durationMin,
        tier.name,
        points,
        0, // ConsecutivePenalty
        isCompleted, // Replaces RecoveryWindowFlag
        false, // LateLoadFlag
        false, // ManualOverride
        eventId
      ];

      if (existingRowNum !== -1) {
        const isManual = sheet.getRange(existingRowNum, manualOverrideColIndex).getValue();
        if (String(isManual).toUpperCase() !== 'TRUE') {
          sheet.getRange(existingRowNum, 1, 1, rowArray.length).setValues([rowArray]);
        }
      } else {
        appendRow(sheet, rowArray);
      }
    });

    logEvent('syncCalendar', 'SUCCESS', `Synced ${events.length} events from past ${daysBack} to next ${daysAhead} days`);
    return { success: true, eventCount: events.length };

  } catch (error) {
    logEvent('syncCalendar', 'ERROR', error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Validates and appends a manual event entry to the DailyLog sheet.
 */
function addManualEntry(entryObject) {
  try {
    const { eventName, dateString, startTime, endTime, burnTierName } = entryObject;

    let multiplier = null;
    for (const key in BURN_TIERS) {
      if (BURN_TIERS[key].name === burnTierName) {
        multiplier = BURN_TIERS[key].multiplier;
        break;
      }
    }

    if (multiplier === null) {
      throw new Error(`Burn tier '${burnTierName}' not found.`);
    }

    const refDate = new Date();
    const startObj = parseTimeString(startTime, refDate);
    const endObj = parseTimeString(endTime, refDate);
    const durationMin = minutesBetween(startObj, endObj);
    const points = calculatePoints(durationMin, multiplier);

    const sheet = getSheetByName('DailyLog');
    const rowArray = [
      dateString,
      eventName,
      startTime,
      endTime,
      durationMin,
      burnTierName,
      points,
      0, // ConsecutivePenalty
      false, // Completed
      false, // LateLoadFlag
      true,  // ManualOverride
      ''     // CalendarEventID
    ];

    appendRow(sheet, rowArray);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Reads and formats all DailyLog events for a specific date into an array of objects.
 */
function getEventsForDate(dateString) {
  const sheet = getSheetByName('DailyLog');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = getRowsForDate(sheet, dateString);

  const events = rows.map(row => {
    const eventObj = {};
    headers.forEach((header, index) => {
      let val = row[index];
      
      if (val instanceof Date) {
        if (header.includes('Time')) {
          val = formatTime(val);
        } else if (header === 'Date') {
           val = getDateString(val);
        }
      }
      eventObj[header] = val;
    });
    return eventObj;
  });

  const activeEvents = events.filter(ev => !(ev.CalendarEventID === 'SYSTEM_BASELINE' && String(ev.EventName).includes('(Off)')));

  activeEvents.sort((a, b) => {
    const timeA = a.StartTime || "00:00";
    const timeB = b.StartTime || "00:00";
    return timeA.localeCompare(timeB);
  });

  return activeEvents;
}

/**
 * Toggles the completed status of an event in the DailyLog.
 */
function toggleEventCompletion(eventId, eventName, startTime, dateString, isCompleted) {
  try {
    const sheet = getSheetByName('DailyLog');
    const data = getSheetData(sheet);
    const idxDate = getColumnIndex(sheet, 'Date');
    const idxID = getColumnIndex(sheet, 'CalendarEventID');
    const idxName = getColumnIndex(sheet, 'EventName');
    const idxStart = getColumnIndex(sheet, 'StartTime');
    const idxCompleted = getColumnIndex(sheet, 'Completed') + 1; 

    let targetRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      const rowDate = (data[i][idxDate] instanceof Date) ? getDateString(data[i][idxDate]) : String(data[i][idxDate]);
      const rowStart = (data[i][idxStart] instanceof Date) ? formatTime(data[i][idxStart]) : String(data[i][idxStart]);

      if (rowDate === dateString) {
        // Match by Event ID if it exists, otherwise fall back to matching Name/Time (for Manual events)
        if (eventId && data[i][idxID] === eventId) {
          targetRowIndex = i + 2; break;
        } else if (!eventId && data[i][idxName] === eventName && rowStart === startTime) {
          targetRowIndex = i + 2; break;
        }
      }
    }

    if (targetRowIndex !== -1) {
      sheet.getRange(targetRowIndex, idxCompleted).setValue(isCompleted);
      if(typeof EXECUTION_CACHE !== 'undefined') EXECUTION_CACHE.data['DailyLog'] = null;
      return { success: true };
    }
    return { success: false, message: "Event not found in log." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Creates a new calendar event, assigns the correct GAS colour, and syncs to DailyLog.
 */
function createCalendarEvent(title, dateString, startTime, endTime, burnTierName) {
  try {
    const calendarIdStr = getSetting('CalendarID');
    const calendar = CalendarApp.getCalendarById(calendarIdStr.trim());
    if (!calendar) throw new Error('Calendar not found.');

    const dateParts = dateString.split('-');
    const refDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
    const startDate = parseTimeString(startTime, refDate);
    const endDate = parseTimeString(endTime, refDate);
    const event = calendar.createEvent(title, startDate, endDate);

    let targetTierId = 0;
    for (const [tierId, tierData] of Object.entries(BURN_TIERS)) {
      if (tierData.name === burnTierName) {
        targetTierId = parseInt(tierId);
        break;
      }
    }

    let targetColorId = "";
    for (const [colId, tierId] of Object.entries(COLOUR_TO_TIER_ID)) {
      if (tierId === targetTierId && colId !== "") {
        targetColorId = colId;
        break;
      }
    }

    if (targetColorId) {
      event.setColor(targetColorId);
    }

    syncCalendar(7, 2);
    return { success: true, eventId: event.getId() };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Updates an existing calendar event's title, times, and colour, then syncs.
 */
function updateCalendarEvent(eventId, title, dateString, startTime, endTime, burnTierName, oldDateString) {
  try {
    const sheet = getSheetByName('DailyLog');
    
    if (eventId === 'SYSTEM_BASELINE') {
      const data = getSheetData(sheet);
      const idxDate = getColumnIndex(sheet, 'Date');
      const idxID = getColumnIndex(sheet, 'CalendarEventID');
      
      const searchDate = oldDateString || dateString;
      let targetRowIndex = -1;
      for (let i = 0; i < data.length; i++) {
        const rowDate = (data[i][idxDate] instanceof Date) ? getDateString(data[i][idxDate]) : String(data[i][idxDate]);
        if (rowDate === searchDate && data[i][idxID] === 'SYSTEM_BASELINE') {
          targetRowIndex = i + 2;
          break;
        }
      }
      
      if (targetRowIndex !== -1) {
        let multiplier = 0;
        for (const key in BURN_TIERS) {
          if (BURN_TIERS[key].name === burnTierName) {
            multiplier = BURN_TIERS[key].multiplier;
            break;
          }
        }
        
        const refDate = new Date();
        const durationMin = minutesBetween(parseTimeString(startTime, refDate), parseTimeString(endTime, refDate));
        const points = calculatePoints(durationMin, multiplier);
        
        const dateCol = getColumnIndex(sheet, 'Date') + 1;
        const nameCol = getColumnIndex(sheet, 'EventName') + 1;
        const startCol = getColumnIndex(sheet, 'StartTime') + 1;
        const endCol = getColumnIndex(sheet, 'EndTime') + 1;
        const durCol = getColumnIndex(sheet, 'Duration_Min') + 1;
        const tierCol = getColumnIndex(sheet, 'BurnTier') + 1;
        const ptsCol = getColumnIndex(sheet, 'Points') + 1;
        
        sheet.getRange(targetRowIndex, dateCol).setValue(dateString);
        sheet.getRange(targetRowIndex, nameCol).setValue(title);
        sheet.getRange(targetRowIndex, startCol).setValue(startTime);
        sheet.getRange(targetRowIndex, endCol).setValue(endTime);
        sheet.getRange(targetRowIndex, durCol).setValue(durationMin);
        sheet.getRange(targetRowIndex, tierCol).setValue(burnTierName);
        sheet.getRange(targetRowIndex, ptsCol).setValue(points);
        
        if(typeof EXECUTION_CACHE !== 'undefined') EXECUTION_CACHE.data['DailyLog'] = null;
      }
      return { success: true };
    }

    const calendarIdStr = getSetting('CalendarID');
    const calendar = CalendarApp.getCalendarById(calendarIdStr.trim());
    const event = calendar.getEventById(eventId);
    if (!event) throw new Error('Event not found.');

    const dateParts = dateString.split('-');
    const refDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
    
    event.setTitle(title);
    event.setTime(parseTimeString(startTime, refDate), parseTimeString(endTime, refDate));

    let targetTierId = 0;
    for (const [tierId, tierData] of Object.entries(BURN_TIERS)) {
      if (tierData.name === burnTierName) {
        targetTierId = parseInt(tierId);
        break;
      }
    }

    let targetColorId = "";
    for (const [colId, tierId] of Object.entries(COLOUR_TO_TIER_ID)) {
      if (tierId === targetTierId && colId !== "") {
        targetColorId = colId;
        break;
      }
    }

    event.setColor(targetColorId);
    syncCalendar(7, 2);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Deletes a calendar event and its corresponding row in the DailyLog.
 */
function deleteCalendarEvent(eventId, dateString) {
  try {
    const sheet = getSheetByName('DailyLog');
    
    if (eventId === 'SYSTEM_BASELINE') {
      if (!dateString) throw new Error('Date string is required to delete a baseline.');
      
      const data = getSheetData(sheet);
      const idxDate = getColumnIndex(sheet, 'Date');
      const idxID = getColumnIndex(sheet, 'CalendarEventID');
      
      let targetRowIndex = -1;
      for (let i = 0; i < data.length; i++) {
        const rowDate = (data[i][idxDate] instanceof Date) ? getDateString(data[i][idxDate]) : String(data[i][idxDate]);
        if (rowDate === dateString && data[i][idxID] === 'SYSTEM_BASELINE') {
          targetRowIndex = i + 2; 
          break;
        }
      }
      
      if (targetRowIndex !== -1) {
        const nameCol = getColumnIndex(sheet, 'EventName') + 1;
        const pointsCol = getColumnIndex(sheet, 'Points') + 1;
        const tierCol = getColumnIndex(sheet, 'BurnTier') + 1;
        
        sheet.getRange(targetRowIndex, nameCol).setValue('Desk Baseline (Off)');
        sheet.getRange(targetRowIndex, pointsCol).setValue(0);
        sheet.getRange(targetRowIndex, tierCol).setValue('Neutral');
        
        if(typeof EXECUTION_CACHE !== 'undefined') EXECUTION_CACHE.data['DailyLog'] = null;
      }
      
      return { success: true };
    }

    const calendarIdStr = getSetting('CalendarID');
    const calendar = CalendarApp.getCalendarById(calendarIdStr.trim());
    const event = calendar.getEventById(eventId);
    
    if (event) {
      event.deleteEvent();
    }

    const rowNum = findRowByColumnValue(sheet, 'CalendarEventID', eventId);
    if (rowNum !== -1) {
      sheet.deleteRow(rowNum);
    }

    syncCalendar(7, 2);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Batch updates the Settings sheet based on a key-value object map.
 */
function saveSettings(settingsObject) {
  try {
    const sheet = getSheetByName('Settings');
    const valueColIndex = getColumnIndex(sheet, 'B (Value)') + 1; 

    for (const [key, value] of Object.entries(settingsObject)) {
      const rowNum = findRowByColumnValue(sheet, 'A (Parameter)', key);
      if (rowNum !== -1) {
        sheet.getRange(rowNum, valueColIndex).setValue(value);
      }
    }
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
