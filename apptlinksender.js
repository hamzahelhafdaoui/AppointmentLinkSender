/**
 * Gold Standard Medical Group - Telehealth SMS Sender via ClickSend
 * 
 
 */

// ============ CONFIGURATION ============
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE', // Replace with your actual spreadsheet ID
  
  CLICKSEND_USERNAME: '',
  CLICKSEND_API_KEY: '',   // Updated ClickSend API key
  CLICKSEND_API_URL: 'https://rest.clicksend.com/v3/sms/send',
  
  // Time window in minutes for matching appointment times (±5 minutes)
  TIME_WINDOW_MINUTES: 5,
  
  // Tabs to stop processing at (not inclusive) - these are therapist tabs, not provider tabs
  STOP_AT_TABS: [
    
  ],
  
  // Error logging sheet name (will be created if doesn't exist)
  ERROR_LOG_SHEET: 'Messaging Errors'
};

/**
 * Main function to send telehealth SMS messages to ALL providers
 */
function sendAllProvidersSMS() {
  try {
    Logger.log('=== Starting Telehealth SMS Send Process ===');
    Logger.log('Current time: ' + new Date());
    
    // Find today's appointment sheet
    const sheet = findTodaysAppointmentSheet();
    if (!sheet) {
      Logger.log('No appointment sheet found for today');
      return;
    }
    
    Logger.log('Found sheet: ' + sheet.getName());
    
    // Get current time for comparison
    const currentTime = new Date();
    
    const tabs = sheet.getSheets();
    
    // Determine which stop tabs to use
    let stopTabs = CONFIG.STOP_AT_TABS;
    let stopTabExists = false;
    
    for (let tab of tabs) {
      if (stopTabs.includes(tab.getName())) {
        stopTabExists = true;
        break;
      }
    }
    
    if (!stopTabExists) {
      Logger.log('None of the primary stop tabs found, proceeding with all tabs.');
    } else {
      Logger.log('Using stop tabs: ' + stopTabs.join(', '));
    }
    
    // Process each provider tab until we reach the stop tabs
    let providersProcessed = 0;
    let sentMessages = 0;
    
    for (let tab of tabs) {
      const tabName = tab.getName();
      
      // Stop before processing the designated stop tabs
      if (stopTabs.some(stopTab => tabName.toLowerCase().includes(stopTab.toLowerCase()))) {
        Logger.log(`[v0] Reached stop tab: ${tabName}. Ending processing.`);
        break;
      }
      
      Logger.log('Processing tab: ' + tabName);
      
      // Process this provider tab
      const result = processProviderTab(tab, currentTime);
      providersProcessed++;
      sentMessages += result.sentCount;
    }
    
    Logger.log('=== Process Complete ===');
    Logger.log('Providers processed: ' + providersProcessed);
    Logger.log('SMS messages sent: ' + sentMessages);
    
  } catch (error) {
    Logger.log('CRITICAL ERROR: ' + error.toString());
    logError('System', 'N/A', 'N/A', 'N/A', error.toString());
  }
}

/**
 * Find the spreadsheet FILE with today's date in Google Drive
 */
function findTodaysAppointmentSheet() {
  const today = new Date();
  const dateFormats = [
    Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM-dd-yyyy'), // 11-12-2025
    Utilities.formatDate(today, Session.getScriptTimeZone(), 'M-d-yyyy')    // 11-12-2025 (no leading zeros)
  ];
  
  Logger.log('Looking for sheets with date: ' + dateFormats.join(' or '));
  
  // First, try to find native Google Sheets files
  for (let dateFormat of dateFormats) {
    // Try exact match with "appointment"
    let searchQuery = 'mimeType = "application/vnd.google-apps.spreadsheet" and title contains "' + dateFormat + ' appointment"';
    Logger.log('[v0] Searching for Google Sheets with query: ' + searchQuery);
    let files = DriveApp.searchFiles(searchQuery);
    
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('Found sheet: ' + file.getName());
      return SpreadsheetApp.openById(file.getId());
    }
  }
  
  // Note: We'll create a converted copy the first time, then use that going forward
  for (let dateFormat of dateFormats) {
    let searchQuery = 'title contains "' + dateFormat + ' appointment"';
    Logger.log('[v0] Searching for any file with query: ' + searchQuery);
    let files = DriveApp.searchFiles(searchQuery);
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // Check if it's an Excel file
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        Logger.log('[v0] Found XLSX file: ' + fileName);
        Logger.log('[v0] ERROR: Cannot process XLSX files directly. Please convert to Google Sheets:');
        Logger.log('[v0] 1. Open the file: ' + fileName);
        Logger.log('[v0] 2. Go to File > Save as Google Sheets');
        Logger.log('[v0] 3. Run this script again');
        return null;
      }
      
      // If it's already a Google Sheets file, use it
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        Logger.log('Found sheet: ' + fileName);
        return SpreadsheetApp.openById(file.getId());
      }
    }
  }
  
  Logger.log('[v0] No matching files found.');
  return null;
}

/**
 * Process a single provider tab
 */
function processProviderTab(tab, currentTime) {
  const result = { sentCount: 0 };
  
  try {
    // Get all data from the tab
    const data = tab.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('Tab has no data rows');
      return result;
    }
    
    const headers = data[0].map(function(h) { return h.toString().trim(); });
    
    const colIndices = {
      otColumn: headers.indexOf('O/T'),
      timeColumn: headers.indexOf('Time'),
      phoneColumn: headers.indexOf('Phone #'),
      linkSentColumn: headers.indexOf('Link Sent'),
      apptWithColumn: headers.indexOf('Appt With'),
      followUpColumn: headers.indexOf('Follow Up'),
      dobColumn: headers.indexOf('DOB')
    };
    
    if (colIndices.timeColumn === -1) {
      colIndices.timeColumn = 0;
    }
    
    // Validate required columns exist
    if (colIndices.otColumn === -1 || colIndices.phoneColumn === -1 || colIndices.linkSentColumn === -1) {
      Logger.log('Missing required columns in tab: ' + tab.getName());
      return result;
    }
    
    Logger.log('[v0] First 5 data rows raw times:');
    for (let i = 1; i < Math.min(6, data.length); i++) {
      const row = data[i];
      const otValue = row[colIndices.otColumn] ? row[colIndices.otColumn].toString().trim() : '';
      const timeValue = row[colIndices.timeColumn];
      const timeType = typeof timeValue;
      const isDate = timeValue instanceof Date;
      Logger.log('[v0] Row ' + (i+1) + ': Time=' + timeValue + ', Type=' + timeType + ', IsDate=' + isDate + ', O/T=' + otValue);
    }
    
    // Get all telehealth rows
    const telehealthRows = [];
    for (let i = 1; i < data.length; i++) {
      const row = { rowIndex: i + 1, data: data[i] };
      const otValue = row.data[colIndices.otColumn];
      
      // Check if O/T column starts with T (covers T, TF, TE, etc.)
      if (otValue && otValue.toString().trim().toUpperCase().startsWith('T')) {
        telehealthRows.push(row);
      }
    }

    if (telehealthRows.length === 0) {
      Logger.log('No telehealth appointments found');
      return result;
    }
    
    Logger.log('Telehealth rows found: ' + telehealthRows.length);
    
    let earliestApptTime = null;
    let validAppointmentsFound = 0;
    
    for (let row of telehealthRows) {
      const timeValue = row.data[colIndices.timeColumn];
      Logger.log('[v0] DEBUG Raw time value: "' + timeValue + '" (type: ' + typeof timeValue + ')');
      
      const parsedTime = parseAppointmentTime(timeValue);
      
      if (parsedTime) {
        validAppointmentsFound++;
        Logger.log('[v0] Parsed appointment time: ' + timeValue + ' -> ' + Utilities.formatDate(parsedTime, Session.getScriptTimeZone(), 'h:mm a'));
        
        if (earliestApptTime === null || parsedTime < earliestApptTime) {
          earliestApptTime = parsedTime;
        }
      } else {
        Logger.log('[v0] DEBUG Failed to parse time: "' + timeValue + '"');
      }
    }
    
    Logger.log('[v0] Valid appointments with times found: ' + validAppointmentsFound);
    
    if (earliestApptTime && validAppointmentsFound > 0) {
      // Provider shift starts 5 minutes before their earliest appointment
      const shiftStartTime = new Date(earliestApptTime.getTime() - (CONFIG.TIME_WINDOW_MINUTES * 60 * 1000));
      
      Logger.log('[v0] Earliest appointment: ' + Utilities.formatDate(earliestApptTime, Session.getScriptTimeZone(), 'h:mm a'));
      Logger.log('[v0] Shift starts at: ' + Utilities.formatDate(shiftStartTime, Session.getScriptTimeZone(), 'h:mm a'));
      Logger.log('[v0] Current time: ' + Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'h:mm a'));
      
      if (currentTime < shiftStartTime) {
        Logger.log('[v0] Provider shift has NOT started yet (current < shift start) - SKIPPING tab');
        return result;
      }
      
      Logger.log('[v0] Provider shift HAS started - PROCEEDING with SMS');
    } else {
      Logger.log('[v0] Could not parse appointment times - PROCEEDING with SMS anyway (better to send than not send)');
    }
    
    // Extract provider first name
    const providerName = extractProviderFirstName(tab.getName(), 
      colIndices.apptWithColumn !== -1 ? data[1][colIndices.apptWithColumn] : null);
    
    for (let row of telehealthRows) {
      const phoneNumber = row.data[colIndices.phoneColumn];
      const patientName = row.data[0] || 'Patient';
      
      if (!phoneNumber || phoneNumber.toString().trim() === '') {
        Logger.log('Skipping row ' + row.rowIndex + ' - no phone number');
        continue;
      }
      
      if (colIndices.dobColumn !== -1) {
        const dobValue = row.data[colIndices.dobColumn];
        if (dobValue && dobValue.toString().trim() === '*') {
          Logger.log('Skipping row ' + row.rowIndex + ' - incomplete intake (DOB is *)');
          continue;
        }
      }
      
      const followUpValue = colIndices.followUpColumn !== -1 ? row.data[colIndices.followUpColumn] : null;
      if (patientHasBeenSeen(followUpValue)) {
        Logger.log('Skipping row ' + row.rowIndex + ' - patient already seen: ' + followUpValue);
        continue;
      }
      
      // Send SMS
      const smsResult = sendClickSendSMS(phoneNumber, providerName);
      
      if (smsResult.success) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'h:mm a');
        const existingValue = row.data[colIndices.linkSentColumn];
        let newValue;
        
        if (existingValue && existingValue.toString().trim() !== '') {
          // Append to existing time (e.g., "8:20" becomes "8:20 9:16 AM")
          newValue = existingValue.toString() + ' ' + timestamp;
        } else {
          newValue = timestamp;
        }
        
        tab.getRange(row.rowIndex, colIndices.linkSentColumn + 1).setValue(newValue);
        
        result.sentCount++;
        Logger.log('SMS sent successfully to ' + phoneNumber + ' - Timestamp: ' + timestamp);
      } else {
        logError(
          tab.getName(),
          patientName,
          row.data[colIndices.timeColumn],
          phoneNumber,
          smsResult.error
        );
        Logger.log('Failed to send SMS to ' + phoneNumber + ': ' + smsResult.error);
      }
    }
    
  } catch (error) {
    Logger.log('Error processing tab ' + tab.getName() + ': ' + error.toString());
    logError(tab.getName(), 'N/A', 'N/A', 'N/A', error.toString());
  }
  
  return result;
}

/**
 * Parse appointment time string to Date object
 */
function parseAppointmentTime(timeValue) {
  try {
    if (!timeValue) return null;
    
    const referenceDate = new Date();
    
    if (timeValue instanceof Date || (timeValue && typeof timeValue.getHours === 'function')) {
      // Google Sheets stores 9:00 AM as 12:00, 10:00 AM as 13:00, etc.
      let hours = timeValue.getHours() - 3;
      if (hours < 0) hours += 24; // Handle wrap-around for times before 3 AM
      const minutes = timeValue.getMinutes();
      return new Date(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate(), hours, minutes, 0, 0);
    }
    
    // Fallback for string values
    const timeStr = String(timeValue).trim();
    if (!timeStr) return null;
    
    // Simple regex for "9:00", "9:00 AM", "14:30" etc
    const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})(?:\s*(AM|PM|am|pm))?/);
    if (!timeMatch) return null;
    
    let hours = parseInt(timeMatch[1]);
    const minutes = parseInt(timeMatch[2]);
    const meridiem = timeMatch[3] ? timeMatch[3].toUpperCase() : null;
    
    // Convert to 24-hour
    if (meridiem === 'PM' && hours !== 12) hours += 12;
    else if (meridiem === 'AM' && hours === 12) hours = 0;
    else if (!meridiem && hours >= 1 && hours <= 7) hours += 12; // 1-7 without AM/PM = PM
    
    return new Date(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate(), hours, minutes, 0, 0);
  } catch (e) {
    return null;
  }
}

/**
 * Extract provider first name from tab name or Appt With column
 */
function extractProviderFirstName(tabName, apptWithValue) {
  // Try tab name first
  let name = tabName;
  
  // If Appt With column has value, use that
  if (apptWithValue && apptWithValue.toString().trim() !== '') {
    name = apptWithValue.toString();
  }
  
  // Clean up common titles and prefixes including /PC
  name = name.replace(/^(Dr\.|Doctor|DR|Mr\.|Mrs\.|Ms\.|Miss)\s*/gi, '');
  name = name.replace(/\/PC/gi, '');
  
  // Extract first name (first word)
  const nameParts = name.trim().split(/\s+/);
  const firstName = nameParts[0];
  
  // Remove any special characters and convert to lowercase
  return firstName.replace(/[^a-zA-Z]/g, '').toLowerCase();
}

/**
 * Send SMS via ClickSend API
 */
function sendClickSendSMS(phoneNumber, providerFirstName) {
  try {
    // Format phone number (remove any non-numeric characters except +)
    const formattedPhone = phoneNumber.toString().replace(/[^\d+]/g, '');
    
    let doxyLink;
    if (providerFirstName.toLowerCase() === 'vivian') {
      doxyLink = 'goldstandard.doxy.me/viviangsmg1';
    } else {
      doxyLink = 'goldstandard.doxy.me/' + providerFirstName + 'gsmg';
    }
    
    const message = 'Gold Standard Medical Group invites you to a secure video call: ' + doxyLink + '.';
    
    // Prepare API request
    const payload = {
      messages: [
        {
          source: 'google-apps-script',
          from: 'GSMG',
          body: message,
          to: formattedPhone
        }
      ]
    };
    
    // Create authorization header
    const authHeader = 'Basic ' + Utilities.base64Encode(
      CONFIG.CLICKSEND_USERNAME + ':' + CONFIG.CLICKSEND_API_KEY
    );
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': authHeader
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    // Send request
    const response = UrlFetchApp.fetch(CONFIG.CLICKSEND_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result.response_code === 'SUCCESS') {
        return { success: true };
      } else {
        return { success: false, error: 'ClickSend API error: ' + result.response_msg };
      }
    } else {
      return { success: false, error: 'HTTP ' + responseCode + ': ' + responseBody };
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Log errors to the Messaging Errors sheet
 */
function logError(providerTab, patientName, appointmentTime, phoneNumber, errorMessage) {
  try {
    // Get or create error log sheet
    let errorSheet = getOrCreateErrorLogSheet();
    
    // Add error row
    const timestamp = new Date();
    errorSheet.appendRow([
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss'),
      providerTab,
      patientName,
      appointmentTime,
      phoneNumber,
      errorMessage
    ]);
    
  } catch (error) {
    Logger.log('Failed to log error: ' + error.toString());
  }
}

/**
 * Get or create the error log sheet
 */
function getOrCreateErrorLogSheet() {
  // Try to find the sheet in the script's associated spreadsheet
  // If running standalone, create a new spreadsheet
  let ss;
  
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // If no active spreadsheet, search for existing error log or create new one
    const files = DriveApp.searchFiles('title = "' + CONFIG.ERROR_LOG_SHEET + '" and mimeType = "application/vnd.google-apps.spreadsheet"');
    
    if (files.hasNext()) {
      ss = SpreadsheetApp.openById(files.next().getId());
    } else {
      ss = SpreadsheetApp.create(CONFIG.ERROR_LOG_SHEET);
    }
  }
  
  let sheet = ss.getSheetByName(CONFIG.ERROR_LOG_SHEET);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ERROR_LOG_SHEET);
    // Add headers
    sheet.appendRow(['Timestamp', 'Provider Tab', 'Patient Name', 'Appointment Time', 'Phone Number', 'Error Message']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * Test function to verify configuration
 */
function testConfiguration() {
  Logger.log('Testing configuration...');
  Logger.log('ClickSend Username: ' + CONFIG.CLICKSEND_USERNAME);
  Logger.log('ClickSend API Key: ' + (CONFIG.CLICKSEND_API_KEY ? 'Set' : 'NOT SET'));
  Logger.log('Current time: ' + new Date());
  Logger.log('Time zone: ' + Session.getScriptTimeZone());
  
  // Try to find today's sheet
  const sheet = findTodaysAppointmentSheet();
  if (sheet) {
    Logger.log('Found appointment sheet: ' + sheet.getName());
    Logger.log('Number of tabs: ' + sheet.getSheets().length);
  } else {
    Logger.log('No appointment sheet found for today');
  }
}

/**
 * Check if patient has already been seen based on Follow Up column value
 * @param {string} followUpValue - The value from the Follow Up column
 * @returns {boolean} - True if patient has been seen, false otherwise
 */
function patientHasBeenSeen(followUpValue) {
  if (!followUpValue || followUpValue.toString().trim() === '') {
    return false;
  }
  
  const value = followUpValue.toString().trim().toLowerCase();
  
  const notSeenIndicators = ['vm', 'lvm', 'confirmed', 'conf'];
  for (let i = 0; i < notSeenIndicators.length; i++) {
    if (value === notSeenIndicators[i] || value.includes(notSeenIndicators[i])) {
      return false;
    }
  }
  
  if (value.includes('cancel') || value.includes('resch') || value.includes('doxy')) {
    return true;
  }
  
  if (value.includes('no f/u') || value.includes('no fu') || value.includes('nof/u') || value.includes('nofu')) {
    return true;
  }
  
  const datePattern = /\d{1,2}[\/\-]\d{1,2}([\/\-]\d{2,4})?/;
  if (datePattern.test(value)) {
    return true;
  }
  
  if (value === 'rs' || value.match(/\brs\b/)) {
    return true;
  }
  
  const weekPatterns = [
    /^\d+\s*w$/i,
    /^\d+\s*wks?$/i,
    /^\d+\s*weeks?$/i,
    /\d+\s*weeks?\s+with/i,
    /\d+\s*wks?\s+with/i,
    /\d+\s*w\s+with/i,
    /^w$/i
  ];
  
  for (let i = 0; i < weekPatterns.length; i++) {
    if (weekPatterns[i].test(value)) {
      return true;
    }
  }
  
  const monthPatterns = [
    /^\d+\s*m$/i,
    /^\d+\s*mo$/i,
    /^\d+\s*months?$/i,
    /\d+\s*months?\s+with/i,
    /\d+\s*mo\s+with/i
  ];
  
  for (let i = 0; i < monthPatterns.length; i++) {
    if (monthPatterns[i].test(value)) {
      return true;
    }
  }
  
  if (value.includes('seen')) {
    const excludePatterns = [
      /recently\s+seen/i,
      /seen\s+yesterday/i,
      /seen\s+last/i,
      /just\s+seen/i,
      /been\s+seen/i,
      /was\s+seen/i,
      /already\s+seen/i
    ];
    
    let shouldExclude = false;
    for (let i = 0; i < excludePatterns.length; i++) {
      if (excludePatterns[i].test(value)) {
        shouldExclude = true;
        break;
      }
    }
    
    if (shouldExclude) {
      return false;
    }
    
    if (value === 'seen' || value.match(/,\s*seen/) || value.match(/\bseen\b/)) {
      return true;
    }
  }
  
  return false;
}




