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
 * @param {object} payload The data sent from the UI, containing command and message.
 * @return {object|string} The response to be sent back to the UI.
 */
function handleChatMessage(payload) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const lastState = JSON.parse(userProperties.getProperty('chatState') || '{}');

    switch (payload.command) {
      case 'analyze_sheet':
        userProperties.deleteProperty('chatState');
        return sendDataForAnalysis(payload.message, lastState.mapping);
      case 'apply_fix':
        userProperties.deleteProperty('chatState');
        return applyFix(payload.action);
      case 'apply_mapping':
        const newMapping = payload.mapping;
        lastState.mapping = { ...lastState.mapping, ...newMapping };
        userProperties.setProperty('chatState', JSON.stringify(lastState));
        return { message: "Mapping applied. Re-running analysis...", success: true };
      case 'general_chat':
        return processGeneralChatWithAI(payload.message);
      default:
        return processGeneralChat(payload.message);
    }
  } catch (e) {
    Logger.log(e);
    return `Error: ${e.message}`;
  }
}

/**
 * Sends a general chat message to the AI proxy for a natural language response.
 * @param {string} message The user's message.
 * @return {object} The AI's conversational response.
 */
function processGeneralChatWithAI(message) {
  const payload = {
    chatMessage: message,
    context: "You are a helpful assistant for the TruckTalk Connect Google Sheets add-on. You help users manage and validate trucking load data. Your responses should be conversational, professional, and directly related to the user's intent within the context of the spreadsheet."
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  try {
    const response = UrlFetchApp.fetch(PROXY_ENDPOINT, options);
    const result = JSON.parse(response.getContentText());
    return { ok: true, message: result.chatResponse };
  } catch (e) {
    return { ok: false, message: "Sorry, I'm having trouble connecting to my brain. Please try again later." };
  }
}

/**
 * Prepares and sends sheet data to the server-side AI for analysis.
 * @param {string} userMessage The original message from the user.
 * @param {object} existingMapping The mapping from the previous conversation state.
 * @return {object} The analysis result from the AI proxy.
 */
function sendDataForAnalysis(userMessage, existingMapping = {}) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values.shift();
  const rows = values.slice(0, 200); // Limit rows for performance and payload size

  const payload = {
    headers: headers,
    rows: rows,
    knownSynonyms: HEADER_MAPPINGS,
    requiredFields: REQUIRED_FIELDS,
    userMapping: existingMapping
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(PROXY_ENDPOINT, options);
    const responseText = response.getContentText();

    // Check for an empty or non-JSON response body even if the status is 200
    if (!responseText || response.getResponseCode() !== 200) {
      return { 
        ok: false, 
        issues: [{ 
          code: "PROXY_ERROR", 
          severity: "error", 
          message: `Proxy server returned an error: HTTP ${response.getResponseCode()}.`, 
          suggestion: "Please check the proxy server's logs for more details." 
        }] 
      };
    }

    // Try to parse the JSON response. This is the key fix.
    let result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      return { 
        ok: false, 
        issues: [{ 
          code: "INVALID_RESPONSE", 
          severity: "error", 
          message: `Received an invalid response from the proxy server.`, 
          suggestion: "The response was not valid JSON. Check the proxy server's logs." 
        }] 
      };
    }

    if (result.error) {
      return { 
        ok: false, 
        issues: [{ 
          code: "PROXY_ERROR", 
          severity: "error", 
          message: `Proxy Error: ${result.error}`, 
          suggestion: "Check the Vercel logs for more details." 
        }] 
      };
    }

    const groupedResult = groupIssues(result);

    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('lastAnalysisResult', JSON.stringify(groupedResult));
    
    return groupedResult;

  } catch (e) {
    return { 
      ok: false, 
      issues: [{ 
        code: "NETWORK_ERROR", 
        severity: "error", 
        message: `Could not connect to analysis server: ${e.message}`, 
        suggestion: "Check your internet connection or try again later." 
      }] 
    };
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
    const key = `${issue.code}-${issue.column || 'no_column'}`;
    if (!groupedIssues[key]) {
      groupedIssues[key] = {
        code: issue.code,
        severity: issue.severity,
        message: issue.message,
        suggestion: issue.suggestion,
        column: issue.column,
        rows: []
      };
      if (FIXABLE_ISSUES.has(issue.code) || issue.code === 'MISSING_COLUMN') {
        groupedIssues[key].action = {
          command: `fix_${issue.code.toLowerCase()}`,
          column: issue.column,
          rows: []
        };
      } else {
        groupedIssues[key].action = null;
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
 * Applies a fix to the spreadsheet based on the action provided by the AI.
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
 * Jumps to the specified cell in the active sheet.
 * @param {string} columnName The header name of the column.
 * @param {number} rowNum The 1-based row number.
 */
function jumpToCell(columnName, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(columnName);

  if (colIndex > -1) {
    const range = sheet.getRange(rowNum, colIndex + 1);
    sheet.setActiveRange(range);
  } else {
    throw new Error(`Column not found: ${columnName}`);
  }
}
