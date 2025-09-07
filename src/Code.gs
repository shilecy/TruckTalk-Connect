/**
 * @fileoverview Main Google Apps Script file for the TruckTalk Connect add-on.
 * Contains core logic for sidebar UI, data analysis, and API communication.
 */

// Global constant for the required header fields based on the provided data schema.
const REQUIRED_FIELDS = [
  'loadId',
  'fromAddress',
  'fromAppointmentDateTimeUTC',
  'toAddress',
  'toAppointmentDateTimeUTC',
  'status',
  'driverName',
  'unitNumber',
  'broker'
];

// Header synonyms for robust column matching.
const HEADER_MAPPINGS = {
  loadId: ['loadId', 'load id', 'ref', 'vrid', 'reference', 'ref #'],
  fromAddress: ['fromAddress', 'from', 'pu', 'pickup', 'origin', 'pickup address', 'pickup location'],
  fromAppointmentDateTimeUTC: ['fromAppointmentDateTimeUTC', 'pu time', 'pickup appt', 'pickup date/time'],
  toAddress: ['toAddress', 'to', 'drop', 'delivery', 'destination', 'delivery address', 'delivery location'],
  toAppointmentDateTimeUTC: ['toAppointmentDateTimeUTC', 'del time', 'delivery appt', 'delivery date/time'],
  status: ['status', 'load status', 'stage', 'load status'],
  driverName: ['driverName', 'driver', 'driver name', 'driver/carrier'],
  driverPhone: ['driverPhone', 'phone', 'driver phone', 'contact'],
  unitNumber: ['unitNumber', 'unit', 'truck', 'truck #', 'tractor', 'unit number'],
  broker: ['broker', 'customer', 'shipper']
};

const PROXY_ENDPOINT = "https://truck-talk-connect.vercel.app/openai-proxy";

/**
 * Creates the menu in Google Sheets to open the sidebar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('TruckTalk Connect')
      .addItem('Open Chat', 'showSidebar')
      .addToUi();
}

/**
 * Displays the HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui')
      .setTitle('TruckTalk Connect')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Handles all incoming messages and commands from the UI.
 * This function acts as the central router for the bot's logic.
 * @param {object} payload The data sent from the UI, containing command and message.
 * @return {object|string} The response to be sent back to the UI.
 */
function handleChatMessage(payload) {
  const userMessage = payload.message;
  const command = payload.command;
  
  if (command === 'analyze_sheet') {
    return processSheetAnalysisCommand();
  }
  
  // Handle general chat or other commands
  return processGeneralChat(userMessage);
}

/**
 * Executes the core analysis of the active sheet.
 * @return {object} A result object containing issues or the generated load data.
 */
function processSheetAnalysisCommand() {
  const analysisResult = analyzeSheet();
  
  if (analysisResult.issues.length > 0) {
    // Correctly returns the full analysis object
    return analysisResult;
  }
  
  // If no issues, we can return the success message with the loads data
  return analysisResult;
}

/**
 * Processes a general chat message from the user.
 * @param {string} message The user's message.
 * @return {string} The bot's chat response.
 */
function processGeneralChat(message) {
  const welcomeMessages = [
    "Hello! How can I assist you today?",
    "Hi there! What can I do for you?",
    "Hello! How can I assist you today with TruckTalk Connect?"
  ];
  const greeting = welcomeMessages[Math.floor(Math.random() * welcomeMessages.length)];
  return greeting;
}

/**
 * Analyzes the active sheet to validate data and find issues.
 * @return {object} An object with `ok` status and an `issues` array.
 */
function analyzeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { ok: false, issues: [{ severity: 'error', message: 'No data found in the sheet.', suggestion: 'Please ensure you have at least one header row and one data row.', action: null }] };
  }
  const headers = data.shift();
  
  const issues = [];
  const loads = [];
  
  const headerMap = mapHeaders(headers);
  
  const missingHeaders = REQUIRED_FIELDS.filter(field => !headerMap.map.hasOwnProperty(field));
  if (missingHeaders.length > 0) {
    issues.push({
      code: 'MISSING_COLUMN',
      severity: 'error',
      message: 'Missing or misspelled required headers.',
      suggestion: `Please ensure all required headers are present. Missing: ${missingHeaders.join(', ')}`,
      action: null,
      rows: []
    });
    return { ok: false, issues: issues, loads: [], mapping: headerMap.map, meta: { analyzedRows: data.length, analyzedAt: new Date().toISOString() } };
  }

  // Use a map to group similar issues
  const issueMap = new Map();

  function addIssue(newIssue) {
    const key = `${newIssue.code}-${newIssue.column}`;
    if (issueMap.has(key)) {
      const existingIssue = issueMap.get(key);
      if (newIssue.rows && newIssue.rows.length > 0) {
        existingIssue.rows.push(...newIssue.rows);
      }
    } else {
      issueMap.set(key, newIssue);
    }
  }

  // Find inconsistent status values
  const statusColumnIndex = headerMap.map.status;
  const statusValues = [...new Set(data.map(row => row[statusColumnIndex] && row[statusColumnIndex].toString().trim()).filter(Boolean))];
  if (statusValues.length > 1) {
    addIssue({
      code: 'INCONSISTENT_STATUS',
      severity: 'warn',
      message: 'Inconsistent status vocabulary found.',
      suggestion: `Please normalize your status values. Found: ${statusValues.join(', ')}.`,
      action: null,
      rows: []
    });
  }

  data.forEach((row, rowIndex) => {
    const load = {};
    let rowHasError = false;
    
    REQUIRED_FIELDS.forEach(field => {
      const colIndex = headerMap.map[field];
      const cellValue = row[colIndex];
      
      if (!cellValue || cellValue.toString().trim() === '') {
        addIssue({
          code: 'EMPTY_REQUIRED_CELL',
          severity: 'error',
          message: `Missing value for required field '${field}'.`,
          suggestion: `Enter a value in the column '${field}'.`,
          rows: [rowIndex + 2],
          column: field,
          action: {
            command: 'selectCell',
            column: colIndex,
            row: rowIndex + 2
          }
        });
        rowHasError = true;
      } else {
        load[field] = cellValue;
      }
    });

    // Check for optional driverPhone field
    const driverPhoneMapping = headerMap.map.driverPhone;
    if (driverPhoneMapping !== undefined) {
      load.driverPhone = row[driverPhoneMapping] || null;
    }
    
    if (!rowHasError) {
      loads.push(load);
    }
  });

  // Check for duplicate loadIds
  const loadIdMap = new Map();
  loads.forEach((load, index) => {
    const loadId = load.loadId;
    if (loadIdMap.has(loadId)) {
      addIssue({
        code: 'DUPLICATE_ID',
        severity: 'error',
        message: `Duplicate loadId found: '${loadId}'.`,
        suggestion: `Ensure each load has a unique ID.`,
        rows: [loadIdMap.get(loadId) + 2, index + 2],
        column: 'loadId',
        action: null
      });
    } else {
      loadIdMap.set(loadId, index);
    }
  });

  // Check for datetime format
  ['fromAppointmentDateTimeUTC', 'toAppointmentDateTimeUTC'].forEach(field => {
    const colIndex = headerMap.map[field];
    if (colIndex !== undefined) {
      data.forEach((row, rowIndex) => {
        const cellValue = row[colIndex];
        if (cellValue) {
          try {
            const date = new Date(cellValue);
            if (isNaN(date.getTime())) {
              addIssue({
                code: 'BAD_DATE_FORMAT',
                severity: 'error',
                message: `Invalid date/time format for field '${field}'.`,
                suggestion: `Please use a valid ISO 8601 format.`,
                rows: [rowIndex + 2],
                column: field,
                action: {
                  command: 'selectCell',
                  column: colIndex,
                  row: rowIndex + 2
                }
              });
            } else {
              // This is a warning for non-ISO 8601 but valid dates
              // Simplified check, a full check would be more complex
              const isISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z?$/.test(cellValue);
              if (!isISO) {
                 addIssue({
                   code: 'NON_ISO_OUTPUT',
                   severity: 'warn',
                   message: `Date/time format for '${field}' is not ISO 8601.`,
                   suggestion: `Please normalize to ISO 8601 UTC.`,
                   rows: [rowIndex + 2],
                   column: field,
                   action: {
                     command: 'selectCell',
                     column: colIndex,
                     row: rowIndex + 2
                   }
                 });
              }
            }
          } catch (e) {
            addIssue({
              code: 'BAD_DATE_FORMAT',
              severity: 'error',
              message: `Invalid date/time format for field '${field}'.`,
              suggestion: `Please use a valid ISO 8601 format.`,
              rows: [rowIndex + 2],
              column: field,
              action: {
                command: 'selectCell',
                column: colIndex,
                row: rowIndex + 2
              }
            });
          }
        }
      });
    }
  });

  const finalIssues = Array.from(issueMap.values());

  if (finalIssues.length > 0) {
    return { ok: false, issues: finalIssues, loads: [], mapping: headerMap.map, meta: { analyzedRows: data.length, analyzedAt: new Date().toISOString() } };
  } else {
    return { ok: true, issues: [], loads: loads, mapping: headerMap.map, meta: { analyzedRows: data.length, analyzedAt: new Date().toISOString() } };
  }
}

/**
 * Maps headers in the spreadsheet to the required fields using synonyms.
 * @param {string[]} headers The header row from the spreadsheet.
 * @return {object} An object containing the mapping and a status flag.
 */
function mapHeaders(headers) {
  const map = {};
  const lowerCaseHeaders = headers.map(h => h.toLowerCase());
  
  for (const field in HEADER_MAPPINGS) {
    const synonyms = HEADER_MAPPINGS[field];
    for (let i = 0; i < synonyms.length; i++) {
      const index = lowerCaseHeaders.indexOf(synonyms[i].toLowerCase());
      if (index !== -1) {
        map[field] = index;
        break; // Found a match, move to the next field
      }
    }
  }

  const allRequiredFound = REQUIRED_FIELDS.every(field => map.hasOwnProperty(field));
  
  return { map: map, allFound: allRequiredFound };
}

/**
 * Handles a user-requested fix by executing a predefined action.
 * @param {object} action The action object from the UI containing a command and parameters.
 * @return {boolean} True if a fix was applied, false otherwise.
 */
function handleSuggestionClick(action) {
  if (action && action.command === 'selectCell') {
    selectSheetCell(action.column, action.row);
    return true; // Return true because a fix was applied
  }
  return false; // Return false because no fix was applied
}

/**
 * Selects a specific cell in the active sheet.
 * @param {number} colIndex The column index (0-based).
 * @param {number} rowNum The row number (1-based).
 */
function selectSheetCell(colIndex, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getRange(rowNum, colIndex + 1);
  sheet.setActiveRange(cell);
}