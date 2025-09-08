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
  try {
    if (payload.command === "analyze_sheet") {
      return analyzeActiveSheet(payload.opts || {});
    }

    if (payload.command === "apply_fix") {
      return applyFix(payload.issue);
    }

    if (payload.command === "apply_mapping") {
      // user confirmed mapping -> re-run analysis with overrides
      return analyzeActiveSheet({ headerOverrides: payload.mapping });
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
  const rows = sheet.getRange(2, 1, Math.min(20, sheet.getLastRow()-1), headers.length).getValues(); // sample

  const systemPrompt = `
You are an AI fixing Google Sheet logistics loads data.
Given headers and rows, fix ONLY the issue described.
Never fabricate values. Leave unknowns blank.
Output JSON:
{
  fixes: [ { row: number, column: string, newValue: string } ],
  summary: string
}`;

  const userPrompt = `
Issue to fix: ${issue.message}
Headers: ${JSON.stringify(headers)}
Sample rows: ${JSON.stringify(rows.slice(0,5))}
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

  // Apply fixes
  parsed.fixes.forEach(fix => {
    const colIndex = headers.indexOf(fix.column) + 1;
    if (colIndex > 0) {
      sheet.getRange(fix.row + 1, colIndex).setValue(fix.newValue);
    }
  });

  switch (issue.code) {
    case "MISSING_COLUMN":
      // Example: insert missing column
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(issue.column);
      break;

    case "EMPTY_CELL":
      // Example: flag empty cells (real fix could ask AI to fill with best guess)
      issue.rows.forEach(row => {
        sheet.getRange(row + 1, findColumnIndex(sheet, issue.column)).setBackground("#fff3cd"); // highlight yellow
      });
      break;

    case "BAD_DATE_FORMAT":
      // Example: try to parse & normalize
      issue.rows.forEach(row => {
        const cell = sheet.getRange(row + 1, findColumnIndex(sheet, issue.column));
        const val = cell.getValue();
        const parsed = tryNormalizeDate(val);
        if (parsed) cell.setValue(parsed);
      });
      break;
  }

  // Re-run analysis
  const newResult = analyzeActiveSheet();
  newResult.aiSummary = parsed.summary;  // so UI can show "Great. 24 loads valid..."
  return newResult;
}

function runChatAI(userMessage) {
  const systemPrompt = `
You are TruckTalk Connect, an AI assistant inside Google Sheets.
- If the user asks to analyze, politely suggest "Type 'analyze' to analyze this tab."
- If they say things like "analyse this tab" or "check sheet", you may also directly suggest analysis.
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

  // ✅ Pass headerOverrides if provided
  const result = callOpenAI(headers, rows, opts?.headerOverrides || {});

  // Group duplicate issues + always attach suggestion
  result.issues = groupIssues(result.issues || []).map(issue => ({
    ...issue,
    suggestion: issue.suggestion || "Please review and update the sheet manually."
  }));

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
function callOpenAI(rawHeaders, sampleData) {
  const systemPrompt = `
You are TruckTalk Connect, an AI assistant working inside Google Sheets.

Responsibilities:
1. Interpret sheet headers and propose header→field mapping.
2. Detect missing or ambiguous mappings and propose solutions (ask user for confirmation).
3. Normalize bad formats (dates → ISO 8601 UTC) but never invent missing values.
4. Flag unknown or missing values as issues.
5. Summarize issues in plain language with suggested fixes.

Rules:
- Never invent data. Unknowns must stay blank and flagged as issues.
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
    suggestion?: string
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
      model: "gpt-4.1-mini", // or gpt-3.5-turbo for cheaper
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
