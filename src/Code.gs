/**
 * @fileoverview Main Google Apps Script file for the TruckTalk Connect add-on.
 * Contains core logic for sidebar UI, data analysis, and API communication.
 */

// Global constant for the required header fields.
// This is our data model, used throughout the analysis.
const REQUIRED_FIELDS = [
  'loadId', 'fromAddress', 'fromAppointmentDateTimeUTC',
  'toAddress', 'toAppointmentDateTimeUTC', 'status',
  'driverName', 'unitNumber', 'broker'
];

// Header synonyms for the AI to reference.
const HEADER_SYNONYMS = {
  loadId: ['Load ID', 'Ref', 'VRID', 'Reference', 'Ref #'],
  fromAddress: ['From', 'PU', 'Pickup', 'Origin', 'Pickup Address'],
  fromAppointmentDateTimeUTC: ['PU Time', 'Pickup Appt', 'Pickup Date/Time'],
  toAddress: ['To', 'Drop', 'Delivery', 'Destination', 'Delivery Address'],
  toAppointmentDateTimeUTC: ['DEL Time', 'Delivery Appt', 'Delivery Date/Time'],
  status: ['Status', 'Load Status', 'Stage'],
  driverName: ['Driver', 'Driver Name'],
  driverPhone: ['Phone', 'Driver Phone', 'Contact'],
  unitNumber: ['Unit', 'Truck', 'Truck #', 'Tractor', 'Unit Number'],
  broker: ['Broker', 'Customer', 'Shipper']
};

// URL for your Vercel proxy that handles all AI API calls.
const PROXY_ENDPOINT = 'https://truck-talk-connect.vercel.app/openapi-proxy';

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
 * This function is the single entry point for all user messages.
 * @param {Object} payload The message object from the UI.
 * @return {AnalysisResult|string} Analysis result object or a general chat message from the AI.
 */
function handleChatMessage(payload) {
  const userMessage = payload.message || '';
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // The prompt object is sent to the Vercel proxy.
  // The proxy uses this to determine intent (analysis or general chat)
  // and perform the necessary logic.
  const prompt = {
    userMessage: userMessage,
    headers: headers,
    sampleData: rows.slice(0, 15), // Send a small sample for the AI to analyze.
    knownSynonyms: HEADER_SYNONYMS,
    requiredFields: REQUIRED_FIELDS,
    // The system prompt guides the AI's behavior.
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

    // Check if the AI's response is a valid JSON for analysis results.
    try {
      const jsonResponse = JSON.parse(textResponse);
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

// The following functions are no longer directly called by handleChatMessage,
// but can be kept for future use or to serve as a reference for your backend logic.
function mapHeaders(headers) {
  const mapping = {};
  const issues = [];
  const requiredFields = ["loadId", "fromAddress", "fromAppointmentDateTimeUTC", "toAddress", "toAppointmentDateTimeUTC", "status", "driverName", "unitNumber", "broker"];

  for (const field of requiredFields) {
    const synonyms = HEADER_SYNONYMS[field] || [];
    const foundHeader = headers.find(h => synonyms.map(s => s.toLowerCase()).includes(h.toLowerCase().trim()));
    if (foundHeader) {
      mapping[field] = foundHeader;
    } else {
      issues.push({
        code: "MISSING_COLUMN",
        severity: "error",
        message: `Missing required column for field: ${field}`,
        suggestion: `Add a column with a header like: ${synonyms.join(", ")}`
      });
    }
  }
  return { mapping, issues };
}

function validateSheet(sheet, existingMapping = null) {
  // This function would contain your detailed validation logic.
  // In the new, AI-driven model, this logic would likely be moved to your Vercel proxy.
  // The code is kept here for reference.
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const { mapping, issues: headerIssues } = existingMapping ? { mapping: existingMapping, issues: [] } : mapHeaders(headerRow);
  const issues = [...headerIssues];

  if (issues.some(i => i.severity === 'error')) {
    return { ok: false, issues, mapping, meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() } };
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const loads = [];
  const loadIds = new Set();
  const statusValues = new Set();
  const allowedStatuses = ["Pending", "In Progress", "Completed", "Cancelled"];

  data.forEach((row, idx) => {
    const rowNumber = idx + 2;
    const load = {};

    for (const field in mapping) {
      const colIdx = headerRow.indexOf(mapping[field]);
      load[field] = (colIdx !== -1) ? (row[colIdx] || null) : null;
    }

    if (!load.loadId) {
      issues.push({ code: "MISSING_ID", severity: "error", message: `Missing loadId in row ${rowNumber}`, suggestion: "Add a unique loadId", rows: [rowNumber] });
    } else if (loadIds.has(load.loadId)) {
      issues.push({ code: "DUPLICATE_ID", severity: "error", message: `Duplicate loadId '${load.loadId}' in row ${rowNumber}`, suggestion: "Ensure all loadId values are unique", rows: [rowNumber], column: mapping.loadId });
    } else {
      loadIds.add(load.loadId);
    }
    
    ["fromAppointmentDateTimeUTC", "toAppointmentDateTimeUTC"].forEach(dateField => {
      const dateVal = load[dateField];
      if (dateVal) {
        const date = new Date(dateVal);
        if (isNaN(date.getTime())) {
          issues.push({ code: "INVALID_DATE", severity: "error", message: `Invalid date '${dateVal}' in row ${rowNumber}`, suggestion: "Correct the date format (ISO 8601 recommended)", rows: [rowNumber], column: mapping[dateField] });
        } else {
          load[dateField] = date.toISOString();
        }
      }
    });

    if (load.status) {
      statusValues.add(load.status);
      if (!allowedStatuses.includes(load.status)) {
        issues.push({ code: "INVALID_STATUS", severity: "warn", message: `Unrecognized status '${load.status}' in row ${rowNumber}`, suggestion: `Consider normalizing to: ${allowedStatuses.join(", ")}`, rows: [rowNumber], column: mapping.status });
      }
    }

    loads.push(load);
  });
  
  return {
    ok: issues.filter(i => i.severity === "error").length === 0,
    issues,
    mapping,
    meta: { analyzedRows: data.length, analyzedAt: new Date().toISOString() },
    loads: issues.filter(i => i.severity === "error").length === 0 ? loads : undefined
  };
}