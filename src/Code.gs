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
  'originAppointmentDateTimeUTC',
  'destinationAppointmentDateTimeUTC',
  'status',
  'driverName',
  'unitNumber',
  'broker'
];

// Header synonyms for the AI to reference.
const HEADER_SYNONYMS = {
  loadNumber: ["Load #", "Load ID", "Ref", "VRID", "Reference", "Ref #"],
  originName: ["Origin", "Pickup", "PU"],
  originAddress: ["Pickup Address", "From", "PU Address"],
  destinationName: ["Destination", "Drop", "DEL"],
  destinationAddress: ["Delivery Address", "To", "DEL Address"],
  originAppointmentDateTimeUTC: ["PU Time", "Pickup Appt", "Pickup Date/Time"],
  destinationAppointmentDateTimeUTC: ["DEL Time", "Delivery Appt", "Delivery Date/Time"],
  status: ["Status", "Load Status", "Stage"],
  driverName: ["Driver", "Driver Name"],
  unitNumber: ["Unit", "Truck", "Truck #", "Tractor", "Unit Number"],
  broker: ["Broker", "Customer", "Shipper"]
};

// URL for your Vercel proxy that handles all AI API calls.
const PROXY_ENDPOINT = 'https://truck-talk-connect.vercel.app/openai-proxy';

// --- UI Functions ---

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TruckTalk Connect")
    .addItem("Open Chat", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("TruckTalk Connect")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function selectSheetCell(colName, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headerRow.indexOf(colName);
  
  if (colIndex !== -1) {
    sheet.getRange(rowNum, colIndex + 1).activate();
  } else {
    throw new Error(`Column with header "${colName}" not found.`);
  }
}

// --- Main Chat & Analysis Functions ---

/**
 * Handles all chat messages from the sidebar UI by using AI to determine intent.
 * @param {Object} payload The message object from the UI.
 * @return {Object|string} Analysis result object, a structured message, or a general chat message from the AI.
 */
function handleChatMessage(payload) {
  const userMessage = payload.message || '';
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const prompt = {
    userMessage: userMessage,
    headers: headers,
    sampleData: rows.slice(0, 15), // Send a small sample for the AI to analyze.
    knownSynonyms: HEADER_SYNONYMS,
    requiredFields: REQUIRED_FIELDS,
    systemPrompt: `You are an AI assistant for a Google Sheets add-on named "TruckTalk Connect". Your primary goal is to help users validate their logistics data.
      - **If the user asks to analyze the sheet**: Perform a full analysis. Map headers to the provided required fields and synonyms. Identify and report issues like missing columns, duplicate IDs, invalid dates, or inconsistent statuses. Normalize all dates to ISO 8601 UTC format. Return a JSON object strictly following the 'AnalysisResult' contract.
      - **If the user is asking for general help**: Provide a helpful, conversational response.
      - **If the header mapping is ambiguous**: Ask the user for clarification in a conversational manner.
      - **Important**: Your response MUST be valid JSON if an analysis is performed; otherwise, it should be plain text. Do not include any extra prose outside of the JSON.`
  };

  try {
    const response = UrlFetchApp.fetch(PROXY_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(prompt)
    });

    const textResponse = response.getContentText();

    try {
      const jsonResponse = JSON.parse(textResponse);
      // The new logic to check for the structured response from the proxy
      if (jsonResponse.type === 'suggestion_list') {
        return jsonResponse;
      }
      
      if (jsonResponse.issues !== undefined && jsonResponse.mapping !== undefined) {
        // It's a structured AnalysisResult.
        return jsonResponse;
      }
    } catch (e) {
      // It's a general conversational response, not JSON.
      return textResponse;
    }
    
    return textResponse;
    
  } catch (e) {
    console.error('API call failed:', e);
    return 'I am unable to connect to the AI at the moment. Please try again later.';
  }
}

/**
 * Handles clicks on the "fix" buttons in the chat, jumping to the relevant cell.
 * @param {Object} action The action object from the UI.
 */
function handleSuggestionClick(action) {
  if (action.type === 'jump_to_cell') {
    selectSheetCell(action.column, action.row);
  }
}