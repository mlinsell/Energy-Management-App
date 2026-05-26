/**
 * @fileoverview Calendar synchronization and manual event entry functions.
 */

/**
 * Synchronizes events from the specified Google Calendar into the DailyLog sheet.
 * @param {number} [daysAhead=2] - Number of days ahead (including today) to fetch events for.
 * @returns {Object} Success status and count of events synced, or error message.
 */
function syncCalendar(daysAhead = 2) {
  try {
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
    
    const endDate = new Date(today);
    endDate.setDate(endDate.getDate() + daysAhead);
    endDate.setHours(23, 59, 59, 999);

    const events = calendar.getEvents(today, endDate);
    const sheet = getSheetByName('DailyLog');
    const manualOverrideColIndex = getColumnIndex(sheet, 'ManualOverride') + 1; // 1-based for getRange()

    events.forEach(event => {
      const title = event.getTitle();
      const startDate = event.getStartTime();
      const eventEndDate = event.getEndTime();
      const colorId = event.getColor();
      const eventId = event.getId();

      let unclassifiedFlag = false;
      let tierId = COLOUR_TO_TIER_ID[colorId];
      
      if (colorId === "" || colorId === undefined || tierId === undefined) {
        unclassifiedFlag = true;
        tierId = 0; // Default to Neutral
      }

      const tier = BURN_TIERS[tierId];
      const durationMin = minutesBetween(startDate, eventEndDate);
      const points = calculatePoints(durationMin, tier.multiplier);
      
      const dateString = getDateString(startDate);
      const startTimeStr = formatTime(startDate);
      const endTimeStr = formatTime(eventEndDate);
      const flags = applyValidationFlags(startTimeStr, endTimeStr, dateString);

      const rowArray = [
        dateString,
        title,
        startTimeStr,
        endTimeStr,
        durationMin,
        tier.name,
        points,
        0, // ConsecutivePenalty
        flags.recoveryWindowFlag,
        flags.lateLoadFlag,
        false, // ManualOverride
        eventId
      ];

      const existingRowNum = findRowByColumnValue(sheet, 'CalendarEventID', eventId);

      if (existingRowNum !== -1) {
        const isManual = sheet.getRange(existingRowNum, manualOverrideColIndex).getValue();
        if (String(isManual).toUpperCase() !== 'TRUE') {
          sheet.getRange(existingRowNum, 1, 1, rowArray.length).setValues([rowArray]);
        }
      } else {
        appendRow(sheet, rowArray);
      }
    });

    logEvent('syncCalendar', 'SUCCESS', `Synced ${events.length} events for ${daysAhead + 1} days`);
    return { success: true, eventCount: events.length };

  } catch (error) {
    logEvent('syncCalendar', 'ERROR', error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Validates and appends a manual event entry to the DailyLog sheet.
 * @param {Object} entryObject - The manual entry payload.
 * @param {string} entryObject.eventName - Title of the event.
 * @param {string} entryObject.dateString - YYYY-MM-DD date string.
 * @param {string} entryObject.startTime - HH:mm start time string.
 * @param {string} entryObject.endTime - HH:mm end time string.
 * @param {string} entryObject.burnTierName - String matching a name property in BURN_TIERS.
 * @returns {Object} Success status or error message payload.
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
    const flags = applyValidationFlags(startTime, endTime, dateString);

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
      flags.recoveryWindowFlag,
      flags.lateLoadFlag,
      true, // ManualOverride
      ''    // CalendarEventID
    ];

    appendRow(sheet, rowArray);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Reads and formats all DailyLog events for a specific date into an array of objects.
 * @param {string} dateString - YYYY-MM-DD string to query.
 * @returns {Array<Object>} List of formatted event objects with column headers as keys.
 */
function getEventsForDate(dateString) {
  const sheet = getSheetByName('DailyLog');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = getRowsForDate(sheet, dateString);

  return rows.map(row => {
    const eventObj = {};
    headers.forEach((header, index) => {
      let val = row[index];
      
      // Cleanup GAS Date parsing quirks back to string format
      if (val instanceof Date) {
        if (header.includes('Time')) {
          val = formatTime(val);
        } else if (header === 'Date') {
          val = getDateString(val);
        }
      }
      eventObj[header] = val;
    });

    const isManual = String(eventObj['ManualOverride']).toUpperCase() === 'TRUE';
    eventObj.unclassifiedFlag = (eventObj['BurnTier'] === 'Neutral' && !isManual);

    return eventObj;
  });
}


/**
 * Creates a new calendar event, assigns the correct GAS colour, and syncs to DailyLog.
 * @param {string} title - The title of the event.
 * @param {string} dateString - Target date (YYYY-MM-DD).
 * @param {string} startTime - Event start time (HH:mm).
 * @param {string} endTime - Event end time (HH:mm).
 * @param {string} burnTierName - Name of the burn tier to assign.
 * @returns {Object} Success status and eventId, or error message.
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

    // Find Tier ID by matching name in BURN_TIERS
    let targetTierId = 0;
    for (const [tierId, tierData] of Object.entries(BURN_TIERS)) {
      if (tierData.name === burnTierName) {
        targetTierId = parseInt(tierId);
        break;
      }
    }

    // Find GAS String Color ID by mapping backwards through COLOUR_TO_TIER_ID
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

    // Immediately trigger a sync (daysAhead = 0 as it's targeted for today/specific day)
    syncCalendar(0);

    return { success: true, eventId: event.getId() };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Updates an existing calendar event's title, times, and colour, then syncs.
 * @param {string} eventId - Calendar Event ID.
 * @param {string} title - Updated title.
 * @param {string} dateString - Target date (YYYY-MM-DD).
 * @param {string} startTime - Updated start time (HH:mm).
 * @param {string} endTime - Updated end time (HH:mm).
 * @param {string} burnTierName - Updated burn tier name.
 * @returns {Object} Success status or error message.
 */
function updateCalendarEvent(eventId, title, dateString, startTime, endTime, burnTierName) {
  try {
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

    // GAS clears color if set to empty string, fallback handled safely by API
    event.setColor(targetColorId); 
    
    syncCalendar(0);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Deletes a calendar event and its corresponding row in the DailyLog.
 * @param {string} eventId - Calendar Event ID to delete.
 * @returns {Object} Success status or error message.
 */
function deleteCalendarEvent(eventId) {
  try {
    const calendarIdStr = getSetting('CalendarID');
    const calendar = CalendarApp.getCalendarById(calendarIdStr.trim());
    const event = calendar.getEventById(eventId);
    
    if (event) {
      event.deleteEvent();
    }

    const sheet = getSheetByName('DailyLog');
    const rowNum = findRowByColumnValue(sheet, 'CalendarEventID', eventId);
    
    if (rowNum !== -1) {
      sheet.deleteRow(rowNum);
    }

    syncCalendar(0);
    return { success: true };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Batch updates the Settings sheet based on a key-value object map.
 * @param {Object} settingsObject - Dictionary of parameter names and their new values.
 * @returns {Object} Success status or error message.
 */
function saveSettings(settingsObject) {
  try {
    const sheet = getSheetByName('Settings');
    const valueColIndex = getColumnIndex(sheet, 'B (Value)') + 1; // Apps Script uses 1-based index

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
