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
    // The command is now correctly handled by calling the function
    // that sends data to the proxy, as outlined in the brief.
    return sendDataForAnalysis(userMessage, 'gpt-3.5-turbo');
  } else {
    // This is where you would implement logic to handle user commands like
    // "Use DEL Time for delivery appt."
    // For now, it will just return a generic response.
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
    // CORRECTED: Added the model to the payload.
    model: model
  };
  
  try {
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(PROXY_ENDPOINT, options);
    const result = JSON.parse(response.getContentText());
    
    // The AI-generated analysis result is returned directly
    return result;
    
  } catch (e) {
    return { ok: false, issues: [{ severity: 'error', message: `Server error: ${e.message}`, suggestion: 'Please check your proxy server logs for details.' }] };
  }
}

/**
 * Handles a suggestion click from the UI to jump to a cell.
 * @param {object} action The action object from the AI response.
 */
function handleSuggestionClick(action) {
  if (action && action.command === 'selectCell') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(action.column);
    
    if (colIndex !== -1) {
      selectSheetCell(colIndex, action.row);
      return true;
    } else {
      SpreadsheetApp.getUi().alert(`Could not find column '${action.column}'.`);
      return false;
    }
  }
  return false;
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
 * @param {number} colIndex The column index (0-based).
 * @param {number} rowNum The row number (1-based).
 */
function selectSheetCell(colIndex, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getRange(rowNum, colIndex + 1);
  sheet.setActiveRange(cell);
}