/**
 * @fileoverview Server-side logic for the TruckTalk Connect add-on.
 * This file handles sheet analysis, validation, and chat interactions.
 */

// Define the core data models using JSDoc for clarity and type-checking.

/**
 * @typedef {Object} Load
 * @property {string} loadId
 * @property {string} fromAddress
 * @property {string} fromAppointmentDateTimeUTC - ISO 8601
 * @property {string} toAddress
 * @property {string} toAppointmentDateTimeUTC - ISO 8601
 * @property {string} status
 * @property {string} driverName
 * @property {string} [driverPhone] - optional
 * @property {string} unitNumber - vehicle/truck id
 * @property {string} broker
 */

/**
 * @typedef {Object} Issue
 * @property {string} code
 * @property {'error'|'warn'} severity
 * @property {string} message
 * @property {number[]} [rows] - affected rows (1-based)
 * @property {string} [column] - header name
 * @property {string} [suggestion]
 */

/**
 * @typedef {Object} AnalysisResult
 * @property {boolean} ok
 * @property {Issue[]} issues
 * @property {Load[]} [loads]
 * @property {Object<string, string>} mapping - header→field mapping
 * @property {Object} meta
 * @property {number} meta.analyzedRows
 * @property {string} meta.analyzedAt
 */

// --- Menu and UI Functions ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Open Chat", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("TruckTalk Connect")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Main Analysis Functions ---

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

/**
 * Handle chat messages from the sidebar
 * @param {string} message
 * @returns {AnalysisResult | string} response
 */
function handleChatMessage(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const trimmedMessage = message.toLowerCase().trim();

  if (trimmedMessage.includes("analyze")) {
    return validateSheet(sheet);
  } else if (trimmedMessage.includes("summary")) {
    const result = validateSheet(sheet);
    return `Analyzed ${result.meta.analyzedRows} rows. ${result.issues.length} issue(s) detected.`;
  } else {
    return "I can help you analyze the sheet. Try typing 'analyze' or 'summary'.";
  }
}

/**
 * Finds a cell by column name and row number and activates it.
 * @param {string} colName The header name of the column.
 * @param {number} rowNum The 1-based row number.
 * @returns {string} Success message.
 */
function selectSheetCell(colName, rowNum) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(colName);
    
    if (colIndex !== -1) {
      const cell = sheet.getRange(rowNum, colIndex + 1);
      sheet.setActiveRange(cell);
      return "Cell selected successfully.";
    } else {
      throw new Error(`Column '${colName}' not found.`);
    }
  } catch (e) {
    console.error("Error selecting cell:", e.message);
    throw new Error("Could not select the cell.");
  }
}