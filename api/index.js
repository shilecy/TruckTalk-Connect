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

// Define which issue codes can be fixed automatically by the AI.
const FIXABLE_ISSUES = new Set([
  'MISSING_REQUIRED_FIELD',
  'INVALID_DATE_FORMAT'
]);

const PROXY_ENDPOINT = "https://us-central1-your-gcp-project-id.cloudfunctions.net/openai-proxy"; // REPLACE WITH YOUR PROXY URL

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
    return sendDataForAnalysis(userMessage, 'gpt-3.5-turbo');
  } else if (command === 'apply_fix') {
    return applyFix(payload.action);
  } else {
    return processGeneralChat(userMessage);
  }
}

/**
 * Prepares and sends sheet data to the server-side AI for analysis.
 * @param {string} userMessage The message from the user.
 * @param {string} model The AI model to use for the analysis.
 * @return {object} The analysis result from the AI proxy.
 */
function sendDataForAnalysis(userMessage, model) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { ok: false, issues: [{ severity: 'error', message: 'No data found in the sheet.', suggestion: 'Please ensure you have at least one header row and one data row.' }] };
  }
  
  const headers = data[0];
  const sampleData = data.slice(1, 11); // Send first 10 rows for efficiency
  
  const payload = {
    headers: headers,
    sampleData: sampleData,
    userMessage: userMessage,
    requiredFields: REQUIRED_FIELDS,
    model: model,
    headerMappings: HEADER_MAPPINGS
  };
  
  try {
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(PROXY_ENDPOINT, options);
    const result = JSON.parse(response.getContentText());
    
    // Group issues by code and severity for better UI presentation.
    return groupIssues(result);
    
  } catch (e) {
    return { ok: false, issues: [{ severity: 'error', message: `Server error: ${e.message}`, suggestion: 'Please check your proxy server logs for details.' }] };
  }
}

/**
 * Groups issues with the same code and message into a single entry,
 * consolidating affected rows.
 * @param {object} analysisResult The raw analysis result from the AI.
 * @return {object} The grouped analysis result.
 */
function groupIssues(analysisResult) {
  if (!analysisResult.issues) {
    return analysisResult;
  }
  const groupedIssues = {};
  analysisResult.issues.forEach(issue => {
    const key = `${issue.code}-${issue.column}`;
    if (!groupedIssues[key]) {
      groupedIssues[key] = {
        code: issue.code,
        severity: issue.severity,
        message: issue.message,
        suggestion: issue.suggestion,
        column: issue.column,
        rows: []
      };
      if (FIXABLE_ISSUES.has(issue.code)) {
        groupedIssues[key].action = {
          command: `fix_${issue.code.toLowerCase()}`,
          column: issue.column,
          rows: []
        };
      }
    }
    groupedIssues[key].rows.push(...(issue.rows || []));
    if (groupedIssues[key].action) {
      groupedIssues[key].action.rows.push(...(issue.rows || []));
    }
  });

  analysisResult.issues = Object.values(groupedIssues);
  return analysisResult;
}

/**
 * Applies a fix to the spreadsheet based on the action provided by the UI.
 * @param {object} action The action object containing the command and data.
 * @return {boolean} True if the fix was applied, false otherwise.
 */
function applyFix(action) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(action.column);

  if (colIndex === -1) {
    return false;
  }

  try {
    switch (action.command) {
      case 'fix_invalid_date_format':
        action.rows.forEach(row => {
          const cell = sheet.getRange(row, colIndex + 1);
          const value = cell.getValue();
          if (typeof value === 'string') {
            const date = new Date(value);
            if (!isNaN(date.getTime())) {
              cell.setValue(date);
            }
          }
        });
        return true;
      case 'fix_missing_required_field':
        action.rows.forEach(row => {
          const cell = sheet.getRange(row, colIndex + 1);
          if (!cell.getValue()) {
            cell.setValue('N/A');
          }
        });
        return true;
      default:
        return false;
    }
  } catch (e) {
    console.error(`Failed to apply fix: ${e.message}`);
    return false;
  }
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
 * Selects a specific cell in the active sheet.
 * @param {string} columnName - The header name of the column to select.
 * @param {number} rowNum - The row number (1-based index).
 */
function selectSheetCell(columnName, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(columnName);

  if (colIndex !== -1) {
    const colNum = colIndex + 1; 
    const range = sheet.getRange(rowNum, colNum);
    sheet.setActiveRange(range);
  } else {
    Logger.log('Column not found: ' + columnName);
  }
}
