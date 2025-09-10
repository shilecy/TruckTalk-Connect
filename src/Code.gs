/**
 * Opens the sidebar when add-on is launched
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("TruckTalk Connect")
    .addItem("Open Sidebar", "showSidebar")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Testing")
      .addItem("Run All Tests", "runTests")
      .addItem("Insert Sample Data", "insertSampleData"))
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui.html")
    .setTitle("TruckTalk Connect");
  SpreadsheetApp.getUi().showSidebar(html);
}

// Required fields (driverPhone optional)
const REQUIRED_FIELDS = [
  "loadId",
  "fromAddress",
  "fromAppointmentDateTimeUTC",
  "toAddress",
  "toAppointmentDateTimeUTC",
  "status",
  "driverName",
  "unitNumber",
  "broker"
];

// Known header synonyms
const HEADER_MAPPINGS = {
  loadId: ["Load ID", "Ref", "VRID", "Reference", "Ref #"],
  fromAddress: ["From", "PU", "Pickup", "Origin", "Pickup Address"],
  fromAppointmentDateTimeUTC: ["PU Time", "Pickup Appt", "Pickup Date/Time"],
  toAddress: ["To", "Drop", "Delivery", "Destination", "Delivery Address"],
  toAppointmentDateTimeUTC: ["DEL Time", "Delivery Appt", "Delivery Date/Time"],
  status: ["Status", "Load Status", "Stage"],
  driverName: ["Driver", "Driver Name", "Carrier"],
  driverPhone: ["Phone", "Driver Phone", "Contact"],
  unitNumber: ["Unit", "Truck", "Truck #", "Tractor", "Unit Number"],
  broker: ["Broker", "Customer", "Shipper"]
};

/**
 * Central entry point for UI messages
 */
function handleChatMessage(payload) {
  let result;
  let chatMessage = "";

  try {
    // Dispatch commands based on payload
    if (payload.command === "analyze_sheet") {
      result = analyzeActiveSheet({ returnLoads: true });
      chatMessage = generateAIResponse(payload.chatRequest.prompt, result);
      return { chatMessage: chatMessage, result: result };
    }

    if (payload.command === "apply_fix") {
      const issue = payload.issue;
      result = applyFix(issue);
      chatMessage = generateAIResponse(payload.chatRequest.prompt, result, {
        issueMessage: issue.message,
        suggestion: issue.suggestion
      });
      return { chatMessage: chatMessage, result: result };
    }

    if (payload.command === "apply_mapping") {
      result = analyzeActiveSheet({ headerOverrides: payload.mapping, returnLoads: true });
      chatMessage = generateAIResponse(payload.chatRequest.prompt, result);
      return { chatMessage: chatMessage, result: result };
    }

    if (payload.command === "general_chat") {
      chatMessage = generateAIResponse(payload.chatRequest.prompt);
      return { chatMessage: chatMessage };
    }
    
    // Fallback for unknown commands
    return {
      chatMessage: "I'm not sure how to handle that request. Please try asking me something else.",
      result: null
    };

  } catch (err) {
    // Global error handler
    return {
      chatMessage: `An unexpected error occurred: ${err.message}`,
      result: {
        ok: false,
        issues: [{
          code: "SERVER_ERROR",
          severity: "error",
          message: err.message
        }],
        mapping: {},
        meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
      }
    };
  }
}

/**
 * Generates an AI-driven chat response.
 *
 * @param {string} userPrompt The prompt to send to the AI.
 * @param {object} [result] The analysis result from a sheet operation.
 * @param {object} [context] Additional context for the AI.
 * @returns {string} The AI-generated chat message.
 */
function generateAIResponse(userPrompt, result, context = {}) {
  const systemPrompt = `
    You are TruckTalk Connect, a helpful and friendly AI assistant for Google Sheets.
    Your main goal is to make a user's life easier by providing clear, concise, and friendly updates.
    - Keep responses short and conversational.
    - If provided, use the analysis result JSON to inform your answer.
    - If there are no issues, give a positive and encouraging response.
    - When there are issues, summarize them briefly and refer the user to the "Results" tab for details.
    - When a fix is applied, provide a quick confirmation of what was fixed.
    - Never mention the API or the AI model you are using.
    - Always use a friendly, conversational tone.
    `;

  let fullPrompt = userPrompt;
  if (result) {
    fullPrompt += `\n\nSheet analysis result:\n${JSON.stringify(result, null, 2)}`;
  }
  if (context.issueMessage) {
    fullPrompt += `\n\nIssue fixed: "${context.issueMessage}". Suggested fix: "${context.suggestion}"`;
  }

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + getOpenAIKey() },
    payload: JSON.stringify({
      model: "gpt-4.1-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: fullPrompt }
      ],
      temperature: 0.7
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const data = JSON.parse(response.getContentText());
    return data.choices[0].message.content;
  } catch (err) {
    console.error("Error contacting AI:", err.message);
    return "⚠️ I'm having trouble connecting right now. Please try again in a moment.";
  }
}

function applyFix(issue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  // Add helper function to combine date and time into ISO 8601
  function combineDateTime(dateValue, timeValue) {
    if (!dateValue || !timeValue) return null;
    let date = new Date(dateValue);
    let time = new Date(timeValue);
    // If time has the 1899-12-30 date, extract just the time part
    if (time.getFullYear() === 1899 && time.getMonth() === 11 && time.getDate() === 30) {
      date.setHours(time.getHours());
      date.setMinutes(time.getMinutes());
      date.setSeconds(time.getSeconds());
    }
    // Convert to ISO string and ensure UTC
    return date.toISOString();
  }

 // prepare AI prompt
  const systemPrompt = `
    You are TruckTalk Connect, AI assistant for fixing logistics data in Google Sheets.

    Rules:
    - Fix ONLY the described issue.
    - Never invent values. If a value is invalid but repairable, normalize it.
    - Only use provided HEADER_MAPPINGS for header→field mapping.
    - For datetime fixes:
      * Combine separate date and time columns into ISO 8601 UTC
      * If time has '1899-12-30', extract only the time part
      * Handle common formats: MM/DD/YY, DD-MM-YYYY, etc.
      * Convert all times to UTC (assume ET if no timezone)
      * Example outputs:
        - Date only: 2025-09-08T00:00:00Z
        - With time: 2025-09-08T14:30:00Z
      * If both date and time are missing, leave blank and flag only.
      * After each successfully fixed issue, must always return json output.
    - Always output STRICT JSON in this format:
 {
  "fixes": [{ 
    "row": number,
    "column": string,
    "newValue": string,
    "sourceColumns": string[],  // for combined date/time fixes
    "sourceValues": string[]    // original values used
  }],
  loads?: any[],
  mapping: Record<string,string>,
  meta: { analyzedRows: number, analyzedAt: string }
 }`;

  // Enhanced prompt with datetime context
  const userPrompt = `
  Issue to fix: ${issue.message}
  Headers: ${JSON.stringify(headers)}
  Sample rows: ${JSON.stringify(rows.slice(0,5))}
  
  Context for datetime fixes:
  - Current year: ${new Date().getFullYear()}
  - Default timezone: ET (UTC-4)
  - Date columns: ${headers.filter(h => h.toLowerCase().includes('date')).join(', ')}
  - Time columns: ${headers.filter(h => h.toLowerCase().includes('time')).join(', ')}
  
  Special handling:
  1. If you see '1899-12-30' in time fields, extract only the time part
  2. For separate date/time columns:
     - Find matching pairs (e.g., 'PU date' + 'PU time')
     - Combine them into ISO 8601 UTC
  3. Return detailed transformation explanation
  `;

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + getOpenAIKey() },
    payload: JSON.stringify({
      model: "gpt-4.1-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt }
      ],
      temperature: 0
    })
  });

  const data = JSON.parse(response.getContentText());
  const parsed = JSON.parse(data.choices[0].message.content);

  // Apply fixes with special handling for datetime issues
  parsed.fixes.forEach(fix => {
    if (fix.sourceColumns && fix.sourceColumns.length > 1) {
      // This is a combined date/time fix
      const targetColIndex = headers.indexOf(fix.column) + 1;
      if (targetColIndex > 0) {
        // If column doesn't exist, create it
        if (targetColIndex > sheet.getLastColumn()) {
          sheet.insertColumnAfter(sheet.getLastColumn());
          sheet.getRange(1, targetColIndex).setValue(fix.column);
        }
        
        // Apply the combined datetime value
        sheet.getRange(fix.row + 2, targetColIndex).setValue(fix.newValue);
        
        // Optionally hide or mark original columns as processed
        fix.sourceColumns.forEach(sourceCol => {
          const sourceColIndex = headers.indexOf(sourceCol) + 1;
          if (sourceColIndex > 0) {
            const cell = sheet.getRange(fix.row + 2, sourceColIndex);
            cell.setBackground('#e8f0fe');  // Light blue to indicate processed
          }
        });
      }
    } else {
      // Regular single-column fix
      const colIndex = headers.indexOf(fix.column) + 1;
      if (colIndex > 0) {
        sheet.getRange(fix.row + 1, colIndex).setValue(fix.newValue);
      }
    }
  });

  // Re-run analysis
  const newResult = analyzeActiveSheet({ returnLoads: true });
  
  // Attach JSON for fixed issue(s)
 try {
   newResult.fixedJson = parsed.fixes.map(fix => {
     const rowIndex = fix.row; // because your AI prompt now outputs real sheet row numbers
     const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
     const obj = {};
     headers.forEach((h, i) => {
       obj[h] = rowValues[i];
     });
     return obj;
   });
 } catch (e) {
   newResult.fixedJson = [];
   console.error("Error building fixedJson:", e.message);
 }

  // Enhance result with fix details
  newResult.aiSummary = parsed.summary;
  newResult.transformations = parsed.transformations;
  newResult.fixedLoadJson = true;  // Indicate JSON should be shown
  
  return newResult;
}


function detectIntent(userMessage) {
  const msg = userMessage.toLowerCase().trim();

  // Flexible detection
  if (/\banal(yse|yze)\b/.test(msg) || msg.includes("check") || msg.includes("review") || msg.includes("scan")) {
    return "analyze_sheet";
  }

  return "general_chat";
}


/**
 * Read sheet data and run analysis
 */
function analyzeActiveSheet(opts) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return {
      ok: false,
      issues: [{
        code: "NO_DATA",
        severity: "error",
        message: "Sheet has no data rows."
      }],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    };
  }

  const headers = data[0];
  const rows = data.slice(1);

  // ✅ Pass headerOverrides if provided and get the result from OpenAI
  const result = callOpenAI(headers, rows, opts?.headerOverrides || {});

  // Group duplicate issues + always attach suggestion
  result.issues = groupIssues(result.issues || []).map(issue => ({
    ...issue,
    suggestion: issue.suggestion || "Please review and update the sheet manually."
  }));

    // Only include the 'loads' data if the returnLoads flag is true
  if (opts?.returnLoads) {
    // Re-call OpenAI to get the loads if they weren't returned before
    if (!result.loads) {
        const loadsResult = callOpenAI(headers, rows, opts?.headerOverrides || {}, true);
        result.loads = loadsResult.loads;
    }
  } else {
    // Delete the loads key to prevent it from being sent back
    delete result.loads;
  }

  return result;
}


/**
 * Groups duplicate issues by code+column
 */
function groupIssues(issues) {
  const grouped = {};
  issues.forEach(issue => {
    const key = issue.code + "|" + (issue.column || "");
    if (!grouped[key]) {
      grouped[key] = { ...issue, rows: issue.rows ? [...issue.rows] : [] };
    } else {
      grouped[key].rows = [
        ...new Set([...(grouped[key].rows || []), ...(issue.rows || [])])
      ];
    }
  });
  return Object.values(grouped);
}

/**
 * Calls OpenAI API with sheet snapshot
 */
function callOpenAI(rawHeaders, sampleData, headerOverrides, returnLoads = true) {
  const systemPrompt = `
You are TruckTalk Connect, an AI assistant working inside Google Sheets.

Responsibilities:
1. Interpret sheet headers and propose header→field mapping.
2. Detect missing or ambiguous mappings and propose solutions (ask user for confirmation).
3. Normalize bad formats (dates → ISO 8601 UTC) but never invent missing values.
4. Flag unknown or missing values as issues.
5. Summarize issues in plain language with suggested fixes.

Validation rules (must strictly follow):
- Required columns missing → ERROR
- Duplicate loadId → ERROR
- Invalid datetime → ERROR
- Empty required cell → ERROR
- Non-ISO datetime → WARN
- Inconsistent status vocabulary → WARN

Row Indexing Rule (must strictly follow):
- Sheet row 1 = headers
- When analyzing rows, the first data row is sheet row 2.
- ALWAYS report issue row numbers as the actual Google Sheet row numbers, not zero-based indexes.
- Example: if the 1st row of data has an issue, output row = 2.

Rules:
- Never invent data. Unknowns must stay blank and flagged as issues (ERROR).
- Dates must be ISO 8601 UTC.
- Detect missing required columns, duplicate loadId, invalid datetimes, empty required cells, inconsistent statuses.
- Normalize header synonyms.
Return JSON strictly as:
{
  ok: boolean,
  issues: Array<{
    code: string,
    severity: "error"|"warn",
    message: string,
    rows?: number[],
    column?: string,
    suggestion?: string,
    suggestionTarget?: string // New field for mapping suggestions
  }>,
  loads?: any[],
  mapping: Record<string,string>,
  meta: { analyzedRows: number, analyzedAt: string }
}
  `;

  const payload = {
    headers: rawHeaders,
    rows: sampleData,
    requiredFields: REQUIRED_FIELDS,
    knownSynonyms: HEADER_MAPPINGS
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + getOpenAIKey()
    },
    payload: JSON.stringify({
      model: "gpt-4.1-mini", 
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: JSON.stringify(payload) }
      ],
      temperature: 0
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const data = JSON.parse(response.getContentText());

    if (!data.choices || !data.choices[0].message) {
      throw new Error("No response from OpenAI");
    }

    let parsed;
    try {
      parsed = JSON.parse(data.choices[0].message.content);
    } catch (e) {
      throw new Error("AI did not return valid JSON: " + e.message);
    }

    // Attach meta
    parsed.meta = {
      analyzedRows: sampleData.length,
      analyzedAt: new Date().toISOString()
    };

    // If returnLoads is false, remove the loads property
    if (!returnLoads) {
      delete parsed.loads;
    }

    // group duplicate issues (same code+column)
    parsed.issues = groupIssues(parsed.issues || []).map(issue => ({
      ...issue,
      suggestion: issue.suggestion || "Please review and update the sheet manually."
    }));

    return parsed;

  } catch (err) {
    return {
      ok: false,
      issues: [{
        code: "SERVER_ERROR",
        severity: "error",
        message: err.message
      }],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    };
  }
}

/**
 * Select a specific cell in the sheet
 */
function selectSheetCell(columnName, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getDataRange().getValues()[0].map(h => (h || '').toString().trim().toLowerCase());
  const target = (columnName || '').toLowerCase();

  let colIndex = headers.indexOf(target);

  if (colIndex === -1) {
    // No exact match: select the whole row
    const range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
    SpreadsheetApp.setActiveRange(range);
    return;
  }

  colIndex = colIndex + 1; // convert 0-based to 1-based
  const range = sheet.getRange(rowNum, colIndex);
  SpreadsheetApp.setActiveRange(range);
}

/**
 * Retrieve OpenAI API key from Script Properties
 */
function getOpenAIKey() {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty("OPENAI_API_KEY");
  if (!key) throw new Error("Missing OpenAI API key. Set OPENAI_API_KEY in Script Properties.");
  return key;
}

// ================== UNIT TESTS ==================

/**
 * Runs all unit tests and logs the results.
 * This function can be run directly from the Apps Script editor.
 */
function runTests() {
  const tests = [
    testDateTimeNormalization,
    testAddressValidation,
    testLoadIdValidation,
    testRequiredFieldValidation,
    testEmptyCellDetection
  ];
  let passed = 0;
  let failed = 0;
  Logger.log('--- Starting Unit Tests ---');
  tests.forEach(test => {
    try {
      test();
      passed++;
      Logger.log(`✅ ${test.name} passed`);
    } catch (e) {
      failed++;
      Logger.log(`❌ ${test.name} failed: ${e.message}`);
    }
  });
  Logger.log(`\nTest Summary: ${passed} passed, ${failed} failed`);
}

/**
 * Test suite for the normalizeDateTime function.
 */
function testDateTimeNormalization() {
  const cases = [
    { input: "9/10/2025 2:30 PM ET", expected: "2025-09-10T18:30:00.000Z", name: "MM/DD/YYYY with time" },
    { input: "2025-09-10 14:00:00", expected: "2025-09-10T18:00:00.000Z", name: "YYYY-MM-DD with 24-hour time" },
    { input: "10-Sep-2025 09:15", expected: "2025-09-10T13:15:00.000Z", name: "DD-MMM-YYYY with time" },
    { input: "Invalid Date", expected: null, name: "Invalid string" },
    { input: null, expected: null, name: "Null input" },
    { input: "", expected: null, name: "Empty string" }
  ];
  cases.forEach(({input, expected, name}, i) => {
    const result = normalizeDateTime(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} (${name}) failed: expected ${expected}, got ${result}`);
    }
  });
}

/**
 * Test suite for the validateAddress function.
 */
function testAddressValidation() {
  const cases = [
    { input: "123 Main St, Atlanta, GA 30303", expected: true, name: "Valid full address" },
    { input: "Houston", expected: false, name: "City only" },
    { input: "", expected: false, name: "Empty address" }
  ];
  cases.forEach(({input, expected, name}, i) => {
    const result = validateAddress(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} (${name}) failed: expected ${expected}, got ${result}`);
    }
  });
}

/**
 * Test suite for the validateLoadId function.
 */
function testLoadIdValidation() {
  const existingIds = ["TL123456", "TL789012"];
  const cases = [
    { input: "TL345678", existingIds: existingIds, expected: true, name: "Unique ID" },
    { input: "TL123456", existingIds: existingIds, expected: false, name: "Duplicate ID" },
    { input: "", existingIds: existingIds, expected: false, name: "Empty ID" }
  ];
  cases.forEach(({input, existingIds, expected, name}, i) => {
    const result = validateLoadId(input, existingIds);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} (${name}) failed: expected ${expected}, got ${result}`);
    }
  });
}

/**
 * Test suite for the validateRequiredFields function.
 */
function testRequiredFieldValidation() {
  const requiredFields = REQUIRED_FIELDS;
  const cases = [
    { input: { loadId: "TL123", fromAddress: "123 Main St", fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z", toAddress: "456 Oak Ave", toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z", status: "assigned", driverName: "John Smith", unitNumber: "4721", broker: "BigShipper Inc" }, expected: true, name: "All fields present" },
    { input: { loadId: "TL123", fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z", toAddress: "456 Oak Ave", toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z", status: "assigned", driverName: "John Smith", unitNumber: "4721", broker: "BigShipper Inc" }, expected: false, name: "Missing 'fromAddress'" },
    { input: { loadId: "", fromAddress: "123 Main St", fromAppointmentDateTimeUTC: "2025-09-10T14:00:00Z", toAddress: "456 Oak Ave", toAppointmentDateTimeUTC: "2025-09-11T16:00:00Z", status: "assigned", driverName: "John Smith", unitNumber: "4721", broker: "BigShipper Inc" }, expected: false, name: "Empty 'loadId'" }
  ];
  cases.forEach(({input, expected, name}, i) => {
    const result = validateRequiredFields(input, requiredFields);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} (${name}) failed: expected ${expected}, got ${result}`);
    }
  });
}

/**
 * Test suite for the isEmptyCell function.
 */
function testEmptyCellDetection() {
  const cases = [
    { input: "", expected: true, name: "Empty string" },
    { input: "    ", expected: true, name: "Whitespace string" },
    { input: null, expected: true, name: "Null value" },
    { input: undefined, expected: true, name: "Undefined value" },
    { input: "Not empty", expected: false, name: "Non-empty string" }
  ];
  cases.forEach(({input, expected, name}, i) => {
    const result = isEmptyCell(input);
    if (result !== expected) {
      throw new Error(`Case ${i + 1} (${name}) failed: expected ${expected}, got ${result}`);
    }
  });
}

// ================== HELPER FUNCTIONS ==================

function normalizeDateTime(input) {
  if (!input) return null;
  try {
    const date = new Date(input);
    if (isNaN(date.getTime())) return null;
    if (!input.includes('Z') && !input.includes('+') && !input.match(/[A-Z]{2,3}$/)) {
      date.setHours(date.getHours() + 4); // ET -> UTC
    }
    return date.toISOString();
  } catch {
    return null;
  }
}

function validateAddress(address) {
  if (!address) return false;
  return address.length >= 10 && address.includes(',');
}

function validateLoadId(loadId, existingIds) {
  if (!loadId) return false;
  return !existingIds.includes(loadId);
}

function validateRequiredFields(data, requiredFields) {
  return requiredFields.every(field => {
    const value = data[field];
    return value !== undefined && value !== null && value.toString().trim() !== '';
  });
}

function isEmptyCell(value) {
  if (value === null || value === undefined) return true;
  return value.toString().trim() === '';
}

function insertSampleData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sampleData = generateSampleData();
  sheet.clear();
  const allHeaders = new Set();
  [...sampleData.happyRows, ...sampleData.brokenRows].forEach(row => {
    Object.keys(row).forEach(header => allHeaders.add(header));
  });
  const headers = Array.from(allHeaders);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const happyRowsValues = sampleData.happyRows.map(row => headers.map(header => row[header] || ''));
  if (happyRowsValues.length > 0) {
    sheet.getRange(2, 1, happyRowsValues.length, headers.length).setValues(happyRowsValues);
  }
  const brokenRowsValues = sampleData.brokenRows.map(row => headers.map(header => row[header] || ''));
  if (brokenRowsValues.length > 0) {
    sheet.getRange(2 + happyRowsValues.length, 1, brokenRowsValues.length, headers.length).setValues(brokenRowsValues);
  }
  sheet.autoResizeColumns(1, headers.length);
}
