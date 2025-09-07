/**
 * @fileoverview Main Google Apps Script file for the TruckTalk Connect add-on.
 * Contains core logic for sidebar UI, data analysis, and API communication.
 */

// Global constant for the required header fields.
// This is our data model, used throughout the analysis.
const REQUIRED_FIELDS = [
  'loadNumber',
  'originName',
  'originAddress',
  'destinationName',
  'destinationAddress',
  'pickupDate',
  'pickupTime',
  'deliveryDate',
  'deliveryTime',
  'trailerType',
  'product',
  'quantity',
  'unit',
  'weight',
  'rate',
  'rateType',
  'contact',
  'reference'
];

// URL for your Vercel proxy that handles all AI API calls.
const PROXY_ENDPOINT = 'https://truck-talk-connect.vercel.app/openai-proxy';

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
  const trimmedMessage = userMessage.toLowerCase().trim();
  
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
  const headers = data.shift();
  
  const issues = [];
  const loads = [];
  
  const headerMap = mapHeaders(headers);
  
  if (!headerMap.allFound) {
    issues.push({
      severity: 'error',
      message: 'Missing or misspelled required headers.',
      suggestion: `Please ensure all headers are present: ${REQUIRED_FIELDS.join(', ')}`,
      action: null
    });
    return { ok: false, issues: issues };
  }
  
  data.forEach((row, rowIndex) => {
    const load = {};
    let rowHasError = false;
    
    REQUIRED_FIELDS.forEach(field => {
      const colIndex = headerMap.map[field];
      const cellValue = row[colIndex];
      
      if (!cellValue || cellValue.toString().trim() === '') {
        issues.push({
          severity: 'error',
          message: `Missing value for required field '${field}'.`,
          suggestion: `Enter a value in cell ${sheet.getRange(rowIndex + 2, colIndex + 1).getA1Notation()}.`,
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
    
    if (!rowHasError) {
      loads.push(load);
    }
  });
  
  if (issues.length > 0) {
    return { ok: false, issues: issues };
  } else {
    return { ok: true, issues: [], loads: loads };
  }
}

/**
 * Maps headers in the spreadsheet to the required fields.
 * @param {string[]} headers The header row from the spreadsheet.
 * @return {object} An object containing the mapping and a status flag.
 */
function mapHeaders(headers) {
  const map = {};
  const lowerCaseHeaders = headers.map(h => h.toLowerCase());
  
  REQUIRED_FIELDS.forEach(field => {
    const index = lowerCaseHeaders.indexOf(field.toLowerCase());
    if (index !== -1) {
      map[field] = index;
    }
  });
  
  const allFound = REQUIRED_FIELDS.every(field => map.hasOwnProperty(field));
  
  return { map: map, allFound: allFound };
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