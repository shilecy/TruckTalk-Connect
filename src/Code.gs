/**
 * TruckTalk Connect - Code.gs
 * Consolidated, updated Apps Script backend implementing:
 * - Loads schema & AnalysisResult contract
 * - Header → field mapping with synonyms & user-confirmed mapping persistence
 * - Validation + normalization (dates normalized to ISO only when TZ present; missing TZ => error)
 * - Issues with severity 'error'|'warn'
 * - Duplicate loadId detection
 * - Status vocabulary check (warn)
 * - Preview push (non-destructive): writes a "preview-ttc-<timestamp>" sheet
 * - Rate limiting: 10 analyses / min per user (soft)
 * - Non-destructive behavior by default
 *
 * IMPORTANT: mapping internal representation is field -> header for analysis convenience.
 * The returned AnalysisResult.mapping follows the brief: header -> field.
 */

/* -----------------------------
   Config / Data model
   ----------------------------- */
const LOAD_FIELDS = [
  "loadId",
  "fromAddress",
  "fromAppointmentDateTimeUTC",
  "toAddress",
  "toAppointmentDateTimeUTC",
  "status",
  "driverName",
  "driverPhone",     // optional
  "unitNumber",
  "broker"
];

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

// header synonyms (case-insensitive). Add more synonyms here when needed.
const HEADER_SYNONYMS = {
  loadId: ["Load ID", "Ref", "VRID", "Reference", "Ref #", "Load", "load no", "load number"],
  fromAddress: ["From", "PU", "Pickup", "Origin", "Pickup Address", "Pick Up", "Pick-Up"],
  fromAppointmentDateTimeUTC: ["PU Time", "Pickup Appt", "Pickup Date/Time", "PU Date", "Pickup Date"],
  toAddress: ["To", "Drop", "Delivery", "Destination", "Delivery Address", "Drop Address"],
  toAppointmentDateTimeUTC: ["DEL Time", "Delivery Appt", "Delivery Date/Time", "DEL Date"],
  status: ["Status", "Load Status", "Stage"],
  driverName: ["Driver", "Driver Name", "DriverName"],
  driverPhone: ["Phone", "Driver Phone", "Contact", "Contact Phone"],
  unitNumber: ["Unit", "Truck", "Truck #", "Tractor", "Unit Number", "Vehicle"],
  broker: ["Broker", "Customer", "Shipper", "Consignor"]
};

// Allowed (canonical) statuses for vocabulary check — adjust as needed
const ALLOWED_STATUSES = ["Pending", "In Progress", "Completed", "Cancelled", "On Hold", "Assigned"];

/* -----------------------------
   Utilities
   ----------------------------- */

function nowIso() {
  return new Date().toISOString();
}

function userKeyPrefix() {
  const email = (Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || "anon";
  return email.replace(/[@.]/g, "_");
}

/* -----------------------------
   Rate limiting (soft)
   ----------------------------- */
function enforceRateLimit_() {
  const userProps = PropertiesService.getUserProperties();
  const key = "rate_" + userKeyPrefix();
  const raw = userProps.getProperty(key);
  const now = Date.now();
  const bucket = raw ? JSON.parse(raw) : [];
  // keep timestamps within last 60s
  const pruned = bucket.filter(ts => now - ts < 60 * 1000);
  if (pruned.length >= 10) {
    throw new Error("Rate limit exceeded: please wait a few seconds before analyzing again.");
  }
  pruned.push(now);
  userProps.setProperty(key, JSON.stringify(pruned));
}

/* -----------------------------
   Mapping persistence
   - We store header->field (as required by Output contract)
   - Internally for analysis we convert to field->header for quick lookup
   ----------------------------- */
function saveUserMapping(headerToField) {
  // headerToField: { "Pickup Time" : "fromAppointmentDateTimeUTC", ... }
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("ttc_mapping_" + userKeyPrefix(), JSON.stringify(headerToField));
  return { ok: true };
}

function getUserMapping() {
  const userProps = PropertiesService.getUserProperties();
  const raw = userProps.getProperty("ttc_mapping_" + userKeyPrefix());
  return raw ? JSON.parse(raw) : null;
}

/* -----------------------------
   Sidebar + UI Helpers
   ----------------------------- */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TruckTalk Connect").addItem("Open Chat", "showSidebar").addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui").setTitle("TruckTalk Connect").setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function selectSheetCell(colName, rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headerRow.findIndex(h => String(h).trim() === String(colName).trim());
  if (colIndex !== -1) {
    sheet.setActiveSelection(sheet.getRange(rowNum, colIndex + 1));
    return { ok: true };
  } else {
    throw new Error(`Column with header "${colName}" not found.`);
  }
}

/* -----------------------------
   Header mapping heuristics
   - returns: { fieldToHeader: {...}, headerToField: {...}, issues: [...] }
   - uses saved mapping to prefer user mapping
   ----------------------------- */
function proposeHeaderMapping_(headers) {
  const saved = getUserMapping(); // header->field
  const headerLower = headers.map(h => String(h || "").toLowerCase().trim());

  // Start with mapping according to synonyms
  const fieldToHeader = {};
  const headerToField = {};

  // 1) Apply saved mappings first (if present)
  if (saved) {
    for (const h in saved) {
      const f = saved[h];
      // ensure header exists
      const idx = headerLower.indexOf(String(h).toLowerCase().trim());
      if (idx !== -1 && LOAD_FIELDS.indexOf(f) !== -1) {
        fieldToHeader[f] = headers[idx];
        headerToField[headers[idx]] = f;
      }
    }
  }

  // 2) Fill remaining with synonym matches
  LOAD_FIELDS.forEach(field => {
    if (fieldToHeader[field]) return;
    const synonyms = HEADER_SYNONYMS[field] || [];
    // prefer exact header name match
    const exactIdx = headers.findIndex(h => String(h).trim().toLowerCase() === (field.toLowerCase()));
    if (exactIdx !== -1) {
      fieldToHeader[field] = headers[exactIdx];
      headerToField[headers[exactIdx]] = field;
      return;
    }
    // try synonyms
    for (const s of synonyms) {
      const idx = headerLower.indexOf(String(s).toLowerCase().trim());
      if (idx !== -1) {
        fieldToHeader[field] = headers[idx];
        headerToField[headers[idx]] = field;
        return;
      }
    }
    // otherwise leave unmapped for now
  });

  // 3) Detect ambiguous headers: if a header maps to multiple fields (very rare because we assigned once),
  // or fields unmapped -> ask user.
  const issues = [];
  const unmappedFields = LOAD_FIELDS.filter(f => !fieldToHeader[f] && REQUIRED_FIELDS.includes(f));
  if (unmappedFields.length) {
    issues.push({
      code: "MISSING_COLUMN",
      severity: "error",
      message: `Missing required columns: ${unmappedFields.join(", ")}`,
      suggestion: `Map existing headers or add columns for: ${unmappedFields.join(", ")}`
    });
  }

  return { fieldToHeader, headerToField, issues };
}

/* -----------------------------
   Date parsing & normalization rules
   - Per brief:
     * If date is non-parsable OR timezone is missing -> INVALID DATETIME (error).
     * If date includes timezone or offset we normalize to ISO 8601 UTC.
     * If original value is not ISO but parseable with TZ, add a 'warn' NON_ISO_DATE (suggest normalized).
   ----------------------------- */
function parseAndNormalizeDate(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === "") {
    return { ok: false, reason: "empty" };
  }
  const s = String(raw).trim();
  // Quick detect timezone tokens or offsets:
  const tzRegex = /([zZ]\b|[+\-]\d{2}(:?\d{2})?|\b(UTC|GMT|MST|MDT|PST|PDT|CST|EDT|CEST|CET|MYT|SGT|IST|WIB|WITA|WIT)\b)/i;
  const hasTZ = tzRegex.test(s);

  // Try Date.parse
  const parsed = Date.parse(s);
  if (isNaN(parsed)) {
    // Try some common formats with simple replacements (e.g., slash -> dash)
    const try2 = Date.parse(s.replace(/\//g, "-"));
    if (isNaN(try2)) return { ok: false, reason: "unparsable" };
    // parsed but might lack TZ
    if (!hasTZ) return { ok: false, reason: "missing_timezone" };
    const iso = new Date(try2).toISOString();
    return { ok: true, iso, original: s, isoWarning: !/^\d{4}-\d{2}-\d{2}T/.test(s) ? true : false };
  } else {
    if (!hasTZ) {
      // per brief: timezone missing => treat as error (do not fabricate)
      return { ok: false, reason: "missing_timezone" };
    }
    const iso = new Date(parsed).toISOString();
    return { ok: true, iso, original: s, isoWarning: !/^\d{4}-\d{2}-\d{2}T/.test(s) ? true : false };
  }
}

/* -----------------------------
   Main validation logic
   - validateSheet: reads the sheet, maps headers, validates rows, builds AnalysisResult
   ----------------------------- */
function validateSheet(sheet, overrideFieldToHeader) {
  // enforce rate limit
  enforceRateLimit_();

  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headerRow = headerRange.getValues()[0].map(h => (h === null || h === undefined) ? "" : String(h));
  const headers = headerRow.slice(); // array of header strings
  const dataRange = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow() - 1), sheet.getLastColumn());
  const rawRows = dataRange.getValues(); // array of arrays

  // Build mapping proposals (field->header) using synonyms and saved mapping, but allow override
  const { fieldToHeader, headerToField, issues: headerIssues } = proposeHeaderMapping_(headers);
  // If caller provided override mapping (field->header or header->field) convert accordingly
  let effectiveFieldToHeader = Object.assign({}, fieldToHeader);
  if (overrideFieldToHeader) {
    // overrideFieldToHeader may be header->field or field->header.
    // detect by checking keys: if keys are field names we accept directly; otherwise invert.
    const keys = Object.keys(overrideFieldToHeader);
    const looksLikeFieldKeys = keys.every(k => LOAD_FIELDS.indexOf(k) !== -1);
    if (looksLikeFieldKeys) {
      effectiveFieldToHeader = Object.assign({}, effectiveFieldToHeader, overrideFieldToHeader);
    } else {
      // keys appear to be headers -> field; convert to field->header
      for (const h in overrideFieldToHeader) {
        const f = overrideFieldToHeader[h];
        if (LOAD_FIELDS.indexOf(f) !== -1) effectiveFieldToHeader[f] = h;
      }
    }
  }

  const issues = headerIssues.slice();

  // If any header issues are errors we will still attempt row analysis but final ok will reflect errors
  const loads = [];
  const seenLoadIds = new Map(); // map of id->firstRowNumber
  const uniqueStatusValues = new Set();

  rawRows.forEach((rowArray, idx) => {
    const rowNum = idx + 2; // actual sheet row number
    const load = {};

    // Collect per-row problems
    const rowProblems = [];

    // Extract values for each LOAD_FIELDS
    LOAD_FIELDS.forEach(field => {
      const header = effectiveFieldToHeader[field];
      if (header) {
        const colIdx = headers.findIndex(h => String(h).trim() === String(header).trim());
        const rawVal = (colIdx !== -1) ? rowArray[colIdx] : null;
        // Do not fabricate: if header exists but cell empty -> set null and error for required fields
        load[field] = (rawVal === "" || rawVal === null || rawVal === undefined) ? null : rawVal;
      } else {
        load[field] = null;
      }
    });

    // Required field empties
    REQUIRED_FIELDS.forEach(req => {
      if (!load[req]) {
        issues.push({
          code: "EMPTY_REQUIRED",
          severity: "error",
          message: `Row ${rowNum}: Empty required field ${req}`,
          rows: [rowNum],
          column: effectiveFieldToHeader[req] || null,
          suggestion: `Fill '${req}' or map a header to this field`
        });
      }
    });

    // loadId duplicate check
    if (load.loadId) {
      const idStr = String(load.loadId).trim();
      if (seenLoadIds.has(idStr)) {
        issues.push({
          code: "DUPLICATE_ID",
          severity: "error",
          message: `Row ${rowNum}: Duplicate loadId '${idStr}' (also in row ${seenLoadIds.get(idStr)})`,
          rows: [rowNum, seenLoadIds.get(idStr)],
          column: effectiveFieldToHeader.loadId,
          suggestion: "Ensure loadId is unique"
        });
      } else {
        seenLoadIds.set(idStr, rowNum);
      }
    }

    // Date validations & normalization rules for pickup/delivery appointment fields
    ["fromAppointmentDateTimeUTC", "toAppointmentDateTimeUTC"].forEach(dateField => {
      const rawVal = load[dateField];
      if (!rawVal) {
        // We already created EMPTY_REQUIRED above for required fields
        return;
      }
      const parsed = parseAndNormalizeDate(rawVal);
      if (!parsed.ok) {
        // per brief, missing timezone or unparsable => error
        const code = parsed.reason === "missing_timezone" ? "BAD_DATE_MISSING_TZ" : "BAD_DATE_FORMAT";
        issues.push({
          code,
          severity: "error",
          message: `Row ${rowNum}: Invalid datetime for '${dateField}': '${rawVal}' (${parsed.reason})`,
          rows: [rowNum],
          column: effectiveFieldToHeader[dateField],
          suggestion: "Include timezone or provide ISO8601 timestamps with timezone (e.g. 2025-08-29T14:00:00+08:00)"
        });
        // Do not set normalized field (never fabricate)
        load[dateField] = null;
      } else {
        // parsed.ok true -> normalized iso present
        load[dateField] = parsed.iso;
        // If original wasn't ISO, raise a warn that we normalized
        if (parsed.isoWarning) {
          issues.push({
            code: "NON_ISO_DATE",
            severity: "warn",
            message: `Row ${rowNum}: Non-ISO datetime normalized for '${dateField}'`,
            rows: [rowNum],
            column: effectiveFieldToHeader[dateField],
            suggestion: `Normalized to ${parsed.iso}. Consider updating source to ISO 8601 with timezone.`
          });
        }
      }
    });

    // Status vocabulary check (warn)
    if (load.status) {
      const s = String(load.status).trim();
      uniqueStatusValues.add(s);
      if (!ALLOWED_STATUSES.includes(s)) {
        issues.push({
          code: "INVALID_STATUS",
          severity: "warn",
          message: `Row ${rowNum}: Unrecognized status '${s}'`,
          rows: [rowNum],
          column: effectiveFieldToHeader.status,
          suggestion: `Consider normalizing to one of: ${ALLOWED_STATUSES.join(", ")}`
        });
      }
    }

    // driverPhone: optional, but if present do a basic normalization (keep but don't validate heavily)
    if (load.driverPhone) {
      // basic cleaning, do not fabricate country codes
      const cleaned = String(load.driverPhone).replace(/[^\d+]/g, "");
      load.driverPhone = cleaned || load.driverPhone;
    }

    // unitNumber & broker & driverName are strings; keep as-is (but ensure not fabricated)
    // store load object (values might include nulls where we couldn't normalize)
    loads.push(load);
  });

  // If there were any header-level or row-level 'error' severity items, ok=false
  const ok = issues.filter(i => i.severity === "error").length === 0;

  // Build mapping for output contract: header -> field mapping used
  const headerToFieldOut = {};
  for (const f of Object.keys(effectiveFieldToHeader)) {
    const h = effectiveFieldToHeader[f];
    if (h) headerToFieldOut[h] = f;
  }

  const result = {
    ok,
    issues,
    loads: ok ? loads : undefined,
    mapping: headerToFieldOut,
    meta: {
      analyzedRows: loads.length,
      analyzedAt: nowIso()
    }
  };

  // Save lastAnalysisResult in user properties (for follow-ups)
  const props = PropertiesService.getUserProperties();
  props.setProperty("lastAnalysisResult_" + userKeyPrefix(), JSON.stringify(result));

  return result;
}

/* -----------------------------
   API exposed to client
   - handleChatMessage(payload): main router used by UI
   - saveUserMapping(headerToField) already above
   - previewPushToSheet: writes preview sheet (non-destructive)
   ----------------------------- */

/**
 * Main router for UI chat messages and commands.
 * payload may be { command: 'analyze_sheet' } or { message: '...' } or mapping commands
 */
function handleChatMessage(payload) {
  try {
    // Determine command
    const command = payload && payload.command ? payload.command : null;
    const message = payload && payload.message ? String(payload.message).trim() : "";

    // If UI sends explicit analyze command
    if (command === "analyze_sheet" || /(^|\s)analyz(e|e current tab|e sheet|sis)/i.test(message) || /summary/i.test(message)) {
      return processSheetAnalysisCommand();
    }

    // If UI sends mapping action as command
    if (command === "save_mapping" && payload.mapping) {
      saveUserMapping(payload.mapping);
      // Re-run analysis with saved mapping
      return processSheetAnalysisCommand();
    }

    // If message is a header mapping natural-language command: "Use 'DEL Time' for toAppointmentDateTimeUTC"
    const mapMatch = message.match(/use\s+['"]?(.+?)['"]?\s+for\s+['"]?(.+?)['"]?/i);
    if (mapMatch) {
      const sheetHeader = mapMatch[1];
      const fieldName = mapMatch[2];
      return processHeaderMappingCommand(sheetHeader, fieldName);
    }

    // Fallback: route to proxy chat (existing implementation)
    return processGeneralChat(message);
  } catch (e) {
    return { ok: false, issues: [{ code: "INTERNAL_ERROR", severity: "error", message: e.message, suggestion: "See logs" }], meta: { analyzedRows: 0, analyzedAt: nowIso() } };
  }
}

function processSheetAnalysisCommand() {
  // Grab active sheet and run validation
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const result = validateSheet(sheet);
  return result;
}

function processHeaderMappingCommand(sheetHeader, fieldName) {
  // Accept both header->field or "Use 'Header' for field" mapping.
  // Validate fieldName is a known load field
  if (LOAD_FIELDS.indexOf(fieldName) === -1) {
    return { ok: false, issues: [{ code: "UNKNOWN_FIELD", severity: "error", message: `Unknown field '${fieldName}'`, suggestion: `Use one of: ${LOAD_FIELDS.join(", ")}` }], meta: { analyzedRows: 0, analyzedAt: nowIso() } };
  }

  // Save mapping (merge with existing)
  const existing = getUserMapping() || {};
  existing[sheetHeader] = fieldName;
  saveUserMapping(existing);

  // Re-run analysis with new saved mapping
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  return validateSheet(sheet);
}

/**
 * Calls OpenAI proxy for free-form chat messages.
 * Keep as-is from your previous implementation (uses '/openai-proxy').
 */
function processGeneralChat(userMessage) {
  const proxyUrl = 'https://truck-talk-connect.vercel.app/openai-proxy'; // keep or change to your proxy
  if (!userMessage || String(userMessage).trim() === "") return "Please enter a message.";

  const payload = { prompt: userMessage, model: 'gpt-3.5-turbo' };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(proxyUrl, options);
    if (response.getResponseCode() >= 400) {
      return `Proxy error: ${response.getContentText()}`;
    }
    const jsonResponse = JSON.parse(response.getContentText());
    if (jsonResponse && jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message) {
      return jsonResponse.choices[0].message.content;
    } else {
      return JSON.stringify(jsonResponse);
    }
  } catch (e) {
    return 'Error communicating with the proxy server: ' + e.message;
  }
}

/**
 * Non-destructive preview push:
 * - writes a new sheet named: 'preview-ttc-<timestamp>'
 * - columns follow the canonical LOAD_FIELDS order
 * - returns {ok:true, sheetName} on success
 */
function previewPushToSheet(loads) {
  try {
    if (!loads || !Array.isArray(loads) || loads.length === 0) {
      return { ok: false, message: "No loads provided for preview push." };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kuala_Lumpur", "yyyyMMdd_HHmmss");
    const previewName = `preview-ttc-${ts}`;
    // create a new sheet
    const previewSheet = ss.insertSheet(previewName);
    // header row -> canonical fields
    previewSheet.getRange(1, 1, 1, LOAD_FIELDS.length).setValues([LOAD_FIELDS]);
    // rows
    const values = loads.map(l => LOAD_FIELDS.map(f => (l && Object.prototype.hasOwnProperty.call(l, f)) ? (l[f] === null ? "" : l[f]) : ""));
    previewSheet.getRange(2, 1, values.length, LOAD_FIELDS.length).setValues(values);
    return { ok: true, sheetName: previewName };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

/* For debugging / quick test */
function test_ping() {
  return { ok: true, now: new Date().toISOString() };
}
