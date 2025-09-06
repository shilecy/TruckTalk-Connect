// This is the complete and final code for your Code.gs file.
// It combines all functions into one file to eliminate conflicts.

// --- UI Functions ---

// Runs automatically when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TruckTalk Connect")
    .addItem("Open Chat", "showSidebar")
    .addToUi();
}

// Opens the sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("TruckTalk Connect")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Main Chat & Analysis Functions ---

/**
 * Handles a chat message by checking for sheet analysis commands or sending to the Vercel proxy.
 * @param {string} userMessage The user's message from the UI.
 * @return {string} The AI's response or an analysis result.
 */
function handleChatMessage(userMessage) {
  const trimmedMessage = userMessage.toLowerCase().trim();

  // Check if the user is asking for sheet analysis
  if (trimmedMessage.includes("analyze") || trimmedMessage.includes("summary")) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const result = validateSheet(sheet);
    return result; // The logic for analysis is in the separate functions below
  }
  
  // If not, send the message to the Vercel proxy
  const proxyUrl = 'https://truck-talk-connect.vercel.app/openai-proxy';
  
  const payload = {
    prompt: userMessage,
    model: 'gpt-3.5-turbo'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(proxyUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (jsonResponse.choices && jsonResponse.choices.length > 0) {
      return jsonResponse.choices[0].message.content;
    } else {
      return 'No response from the proxy. Check your request and proxy URL.';
    }
  } catch (e) {
    return 'Error communicating with the proxy server: ' + e.message;
  }
}

// --- Analysis & Helper Functions (moved from the original api.gs file) ---

/**
 * Maps sheet headers to data model fields based on synonyms.
 * @param {string[]} headers
 * @returns {{mapping: Object<string, string>, issues: Issue[]}}
 */
function mapHeaders(headers) {
  const synonyms = {
    loadId: ["Load ID", "Ref", "VRID", "Reference", "Ref #"],
    fromAddress: ["From", "PU", "Pickup", "Origin", "Pickup Address"],
    fromAppointmentDateTimeUTC: ["PU Time", "Pickup Appt", "Pickup Date/Time"],
    toAddress: ["To", "Drop", "Delivery", "Destination", "Delivery Address"],
    toAppointmentDateTimeUTC: ["DEL Time", "Delivery Appt", "Delivery Date/Time"],
    status: ["Status", "Load Status", "Stage"],
    driverName: ["Driver", "Driver Name"],
    driverPhone: ["Phone", "Driver Phone", "Contact"],
    unitNumber: ["Unit", "Truck", "Truck #", "Tractor", "Unit Number"],
    broker: ["Broker", "Customer", "Shipper"]
  };
  
  const requiredFields = ["loadId", "fromAddress", "fromAppointmentDateTimeUTC", "toAddress", "toAppointmentDateTimeUTC", "status", "driverName", "unitNumber", "broker"];
  const mapping = {};
  const issues = [];
  const headerSet = new Set(headers.map(h => h.toLowerCase().trim()));

  for (const field in synonyms) {
    const foundHeader = headers.find(h => synonyms[field].map(s => s.toLowerCase()).includes(h.toLowerCase().trim()));
    if (foundHeader) {
      mapping[field] = foundHeader;
    }
  }

  // Check for missing required columns
  requiredFields.forEach(field => {
    if (!mapping[field]) {
      issues.push({
        code: "MISSING_COLUMN",
        severity: "error",
        message: `Missing required column for field: ${field}`,
        column: field,
        suggestion: `Add a column with a header like: ${synonyms[field].join(", ")}`
      });
    }
  });

  return { mapping, issues };
}

/**
 * Validates the sheet data and returns an AnalysisResult.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {AnalysisResult}
 */
function validateSheet(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const { mapping, issues: headerIssues } = mapHeaders(headerRow);
  const issues = [...headerIssues];

  if (issues.some(i => i.severity === 'error')) {
    return {
      ok: false,
      issues,
      mapping,
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    };
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

    // Validation checks
    if (!load.loadId) {
      issues.push({ code: "MISSING_ID", severity: "error", message: `Missing loadId in row ${rowNumber}`, suggestion: "Add a unique loadId", rows: [rowNumber] });
    } else if (loadIds.has(load.loadId)) {
      issues.push({ code: "DUPLICATE_ID", severity: "error", message: `Duplicate loadId '${load.loadId}' in row ${rowNumber}`, suggestion: "Ensure all loadId values are unique", rows: [rowNumber], column: mapping.loadId });
    } else {
      loadIds.add(load.loadId);
    }
    
    ["fromAppointmentDateTimeUTC", "toAppointmentDateTimeUTC"].forEach(dateField => {
      const dateVal = load[dateField];
      if (!dateVal) {
        issues.push({ code: "MISSING_DATE", severity: "warn", message: `Missing date for ${dateField} in row ${rowNumber}`, suggestion: "Fill in the missing date", rows: [rowNumber], column: mapping[dateField] });
      } else {
        const date = new Date(dateVal);
        if (isNaN(date.getTime())) {
          issues.push({ code: "INVALID_DATE", severity: "error", message: `Invalid date '${dateVal}' in row ${rowNumber}`, suggestion: "Correct the date format (ISO 8601 recommended)", rows: [rowNumber], column: mapping[dateField] });
        } else {
          load[dateField] = date.toISOString(); // Normalize to ISO 8601
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
    meta: {
      analyzedRows: data.length,
      analyzedAt: new Date().toISOString()
    },
    loads: issues.filter(i => i.severity === "error").length === 0 ? loads : undefined
  };
}