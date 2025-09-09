/**
 * Opens the sidebar when add-on is launched
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("TruckTalk Connect")
    .addItem("Open Sidebar", "showSidebar")
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
  fromAddress: ["Pickup location", "From", "PU", "Pickup", "Origin", "Pickup Address"],
  fromAppointmentDateTimeUTC: ["PU Time", "Pickup Appt", "Pickup Date/Time"],
  toAddress: ["Delivery location", "To", "Drop", "Delivery", "Destination", "Delivery Address"],
  toAppointmentDateTimeUTC: ["DEL Time", "Delivery Appt", "Delivery Date/Time"],
  status: ["Status", "Load Status", "Stage"],
  driverName: ["Driver", "Driver Name", "Carrier", "Driver/Carrier"],
  driverPhone: ["Phone", "Driver Phone", "Contact"],
  unitNumber: ["Unit", "Truck", "Truck #", "Tractor", "Unit Number"],
  broker: ["Broker", "Customer", "Shipper"]
};

/**
 * Central entry point for UI messages
 */
function handleChatMessage(payload) { 
  try {
    if (payload.command === "analyze_sheet") {
      return analyzeActiveSheet({ returnLoads: false });
    }

    if (payload.command === "apply_fix") {
      return applyFix(payload.issue);
    }

    if (payload.command === "apply_mapping") {
      // user confirmed mapping -> re-run analysis with overrides
      return analyzeActiveSheet({ headerOverrides: payload.mapping, returnLoads: true });
    }

    // default → let AI chat handle it
    return runChatAI(payload.message);

  } catch (err) {
    return {
      ok: false,
      issues: [{
        code: "ERROR",
        severity: "error",
        message: err.message
      }],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    };
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
    - For datetime fixes:
      * Combine separate date and time columns into ISO 8601 UTC
      * If time has '1899-12-30', extract only the time part
      * Handle common formats: MM/DD/YY, DD-MM-YYYY, etc.
      * Convert all times to UTC (assume ET if no timezone)
      * Example outputs:
        - Date only: 2025-09-08T00:00:00Z
        - With time: 2025-09-08T14:30:00Z
      * If both date and time are missing, leave blank and flag only.
    - Always output STRICT JSON in this format:
 {
  "fixes": [{ 
    "row": number,
    "column": string,
    "newValue": string,
    "sourceColumns": string[],  // for combined date/time fixes
    "sourceValues": string[]    // original values used
  }],
  "summary": string,
  "transformations": [{         // explain what was done
    "type": "datetime",
    "from": string,
    "to": string,
    "logic": string
  }]
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
        sheet.getRange(fix.row + 2, colIndex).setValue(fix.newValue);
      }
    }
  });

  // Re-run analysis
  const newResult = analyzeActiveSheet({ returnLoads: true });
  
  // Enhance result with fix details
  newResult.aiSummary = parsed.summary;
  newResult.transformations = parsed.transformations;
  newResult.fixedLoadJson = true;  // Indicate JSON should be shown
  
  // If this was a datetime fix, add the transformation details
  if (parsed.transformations && parsed.transformations.some(t => t.type === 'datetime')) {
    newResult.changes = parsed.transformations.map(t => 
      `📅 ${t.from} → ${t.to}\n${t.logic}`
    ).join('\n\n');
  }
  
  return newResult;
}


function runChatAI(userMessage) {
  const systemPrompt = `
You are TruckTalk Connect, an AI assistant inside Google Sheets.
- If the user asks to 'analyze' or 'analyse' or 'scan' or 'review' or 'check' this tab, you may proceed to analyze."
- You can explain analysis results, suggest fixes, or just chat casually.
- Never fabricate data; if unknown, say so.
- Keep responses short, friendly, and helpful.
`;

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + getOpenAIKey() },
    payload: JSON.stringify({
      model: "gpt-4.1-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMessage }
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
    return "⚠️ Error contacting AI: " + err.message;
  }
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

    // ✅ NEW: Only include the 'loads' data if the returnLoads flag is true
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

    // ✅ NEW: If returnLoads is false, remove the loads property
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



