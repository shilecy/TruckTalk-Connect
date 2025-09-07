const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');

const app = express();
app.use(express.json());
app.use(cors());

// A root route to confirm the server is running.
app.get('/', (req, res) => {
  res.status(200).send('TruckTalk Connect AI Proxy is running.');
});

// The single, unified endpoint that handles both chat and analysis.
app.post('/openai-proxy', async (req, res) => {
  const { userMessage, headers, sampleData } = req.body;
  const openaiApiKey = process.env.OPENAI_API_KEY;

  if (!openaiApiKey) {
    return res.status(500).json({ error: 'OpenAI API key not set on server.' });
  }

  // Define prompts for different intents
  const analysisPrompt = `
    You are an expert data analyst AI for "TruckTalk Connect". Your task is to analyze spreadsheet data for truck loads, validate it, and convert it into a structured JSON format.

    **Your Goal:** Return a single JSON object that strictly matches the 'AnalysisResult' 
    schema.
    **Core Rules:**
    1. **Header Mapping:** Map the user's headers to the canonical 'Load' schema fields. Use the provided 'HEADER_SYNONYMS'.
    2. **Data Validation & Normalization:**
      * **Dates:** All date fields MUST be converted to ISO 8601 UTC format.
      * **Required Fields:** All fields in the Load schema are required.
      * **Uniqueness:** 'loadId' must be unique.
    3. **NEVER FABRICATE DATA:** If a value is unknown or invalid, its corresponding JSON field must be \`null\`, and you must generate an issue.
    4. **Issue Column:** The 'column' property for each issue MUST be the original header name from the user's sheet.
    5. **Output:** Your entire response MUST be a single, valid JSON object that strictly adheres to the 'AnalysisResult' schema.
    \`\`\`json
    {
      "HEADER_SYNONYMS": {
        "loadId": ["Load ID", "Ref", "VRID", "Reference", "Ref #"],
        "fromAddress": ["From", "PU", "Pickup", "Origin", "Pickup Address"],
        "fromAppointmentDateTimeUTC": ["PU Time", "Pickup Appt", "Pickup Date/Time"],
        "toAddress": ["To", "Drop", "Delivery", "Destination", "Delivery Address"],
        "toAppointmentDateTimeUTC": ["DEL Time", "Delivery Appt", "Delivery Date/Time"],
        "status": ["Status", "Load Status", "Stage"],
        "driverName": ["Driver", "Driver Name"],
        "unitNumber": ["Unit", "Truck", "Truck #", "Tractor", "Unit Number"],
        "broker": ["Broker", "Customer", "Shipper"]
      },
      "AnalysisResultSchema": {
        "ok": "boolean",
        "issues": [{"code": "string", "severity": "'error'|'warn'", "message": "string", "rows": "number[]", "column": "string", "suggestion": "string"}],
        "loads": "Load[] | undefined",
        "mapping": "{ originalHeader: canonicalField, ... }",
        "meta": {"analyzedRows": "number", "analyzedAt": "string (ISO 8601)"}
      }
    }
    \`\`\`
    `;

  const generalChatPrompt = `You are a helpful AI assistant for "TruckTalk Connect".
  You are designed to answer questions and provide general information about the add-on. Be concise and friendly.`;

  try {
    let openaiResponse;
    let finalPayload;

    // Check for the presence of headers and sample data to determine intent.
    if (headers && sampleData) {
      finalPayload = {
        model: 'gpt-4o',
        messages: [{
          role: 'system',
          content: analysisPrompt
        }, {
          role: 'user',
          content: `Analyze the headers and data below and return a single JSON object conforming to the schema.
            Headers: ${JSON.stringify(headers)}
            Sample Data: ${JSON.stringify(sampleData)}`
        }],
        response_format: { type: 'json_object' }
      };
    } else {
      finalPayload = {
        model: 'gpt-3.5-turbo',
        messages: [{
          role: 'system',
          content: generalChatPrompt
        }, {
          role: 'user',
          content: userMessage
        }]
      };
    }
    
    openaiResponse = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${openaiApiKey}`
      },
      body: JSON.stringify(finalPayload)
    });

    if (!openaiResponse.ok) {
      throw new Error(`OpenAI API error: ${openaiResponse.statusText}`);
    }

    const data = await openaiResponse.json();
    const botResponse = data.choices[0].message.content;

    try {
      const jsonResult = JSON.parse(botResponse);
      return res.status(200).json(jsonResult);
    } catch (e) {
      return res.status(200).send(botResponse);
    }

  } catch (error) {
    console.error("Proxy Error:", error);
    res.status(500).json({ error: 'Failed to communicate with the OpenAI API.' });
  }
});

module.exports = app;