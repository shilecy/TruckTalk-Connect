const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

const PROXY_ENDPOINT = '/openai-proxy';
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

// The core AI analysis endpoint
app.post(PROXY_ENDPOINT, async (req, res) => {
  if (!OPENAI_API_KEY) {
    return res.status(500).json({ error: 'OpenAI API key not set on server.' });
  }

  const { headers, rows, knownSynonyms, requiredFields } = req.body;

  // The system prompt defines the AI's role and the strict output contract.
  const systemPrompt = `You are a specialized AI assistant for logistics data. Your task is to analyze a 2D table representing a list of loads and return a single, structured JSON object.
    
    Your output MUST strictly adhere to this JSON schema:
    {
      "ok": boolean,
      "issues": Array<{
        "code": string,              // e.g., MISSING_COLUMN, BAD_DATE_FORMAT, DUPLICATE_ID
        "severity": 'error'|'warn',
        "message": string,           // user-friendly
        "rows"?: number[],           // affected rows (1-based)
        "column"?: string,           // header name
        "suggestion"?: string       // how to fix
      }>,
      "loads"?: Array<{
        "loadId": string;
        "fromAddress": string;
        "fromAppointmentDateTimeUTC": string;
        "toAddress": string;
        "toAppointmentDateTimeUTC": string;
        "status": string;
        "driverName": string;
        "driverPhone"?: string;
        "unitNumber": string;
        "broker": string;
      }>,
      "mapping": Record<string,string>, // header→field mapping used
      "meta": { "analyzedRows": number; "analyzedAt": string; }
    }
    
    You will perform the following tasks:
    1.  **Interpret Headers:** Map the user's headers to the standard Load fields using the provided synonyms. If a field has no matching header, report a MISSING_COLUMN error.
    2.  **Validate Data:**
        * Required Columns: 'error' if any of the required fields are missing from the mapping.
        * Duplicate 'loadId': 'error' for any duplicate 'loadId' values.
        * Empty Cells: 'error' for any empty cells in required fields.
        * Invalid Datetime: 'error' for non-parsable or timezone-missing dates.
        * Inconsistent Status: 'warn' and list all unique values for user normalization.
    3.  **Normalize Values:** Convert all date-time strings to ISO 8601 UTC format. For example, '08/29 2pm MST' becomes '2025-08-29T20:00:00Z'. State assumptions in the message if the timezone is not provided.
    4.  **Issue Summarization:** For each issue, provide a plain-language 'message' and a 'suggestion' for fixing it.
    5.  **Fabrication:** Never invent data. If a cell value is missing, leave the corresponding JSON field blank and flag it with an issue.
    6.  **Return JSON:** Only return the final JSON object. Do not include any other text or explanation outside the JSON.
    
    The user's input is a JSON object with the following keys:
    - 'headers': The header row of the sheet.
    - 'rows': A small sample of the data for analysis (first 200 rows to save cost).
    - 'knownSynonyms': The standard synonyms to use for header mapping.
    - 'requiredFields': The list of required fields.
    
    Your analysis should be based solely on this provided data.`;

  // The user prompt provides the actual data for the AI to analyze.
  const userPrompt = `
    Headers: ${JSON.stringify(headers)}
    Rows: ${JSON.stringify(rows)}
    Known Synonyms: ${JSON.stringify(knownSynonyms)}
    Required Fields: ${JSON.stringify(requiredFields)}
    `;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-4o", // Or another suitable model
        response_format: { type: "json_object" },
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ],
      }),
    });

    if (!response.ok) {
      throw new Error(`OpenAI API request failed with status ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    const messageContent = data?.choices?.[0]?.message?.content;

    // Check if the message content exists before trying to parse it
    if (!messageContent) {
      console.error('OpenAI response content is missing:', JSON.stringify(data));
      return res.status(500).json({ error: 'OpenAI response content is missing.' });
    }

    let resultJson;
    try {
      resultJson = JSON.parse(messageContent);
    } catch (parseError) {
      console.error('Failed to parse OpenAI JSON response:', messageContent);
      return res.status(500).json({ error: 'Invalid JSON response from OpenAI API.' });
    }

    // Add the meta field before sending the final result back
    const finalResult = {
      ...resultJson,
      meta: {
        analyzedRows: rows.length,
        analyzedAt: new Date().toISOString(),
      }
    };

    res.json(finalResult);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// For Vercel, we export the app
module.exports = app;
