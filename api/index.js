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

// New GET handler to provide a helpful message for misdirected requests.
app.get('/openai-proxy', (req, res) => {
  res.status(405).json({
    error: 'Method Not Allowed',
    message: 'This endpoint only accepts POST requests for data analysis or chat. Please send a POST request with the required body.'
  });
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

    **Your Goal:** Return a single JSON object that strictly matches the 'AnalysisResult' schema.
    **Core Rules:**
    1. **Header Mapping:** Map the user's headers to the canonical 'Load' schema fields. Use the provided 'HEADER_SYNONYMS'.
    2. **Data Validation & Normalization:**
      * **Dates:** All date fields MUST be converted to ISO 8601 UTC format.
      * **Required Fields:** All fields in the Load schema are required.
      * **Uniqueness:** 'loadNumber' must be unique.
    3. **NEVER FABRICATE DATA:** If a value is unknown or invalid, its corresponding JSON field must be \`null\`, and you must generate an issue.
    4. **Issue Details:** For each issue, you MUST include 'severity', 'message', 'suggestion', 'rows', 'column', and an optional 'action' object if the issue is fixable by jumping to a specific cell.
    5. **Action Object:** If an issue requires a user to look at a specific cell, the 'action' object MUST be structured as: \`{ "command": "selectCell", "column": "<Original Header Name>", "row": <Row Number> }\`
    6. **Output:** Your entire response MUST be a single, valid JSON object that strictly adheres to the 'AnalysisResult' schema.
    
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
        "issues": [{"severity": "'error'|'warn'", "message": "string", "suggestion": "string", "rows": "number[]", "column": "string (Original Header Name)", "action": {"command": "string", "column": "string", "row": "number"}}],
        "loads": "Load[] | undefined",
        "mapping": "{ originalHeader: canonicalField, ... }"
      }
    }
    \`\`\`
    `;

  const generalChatPrompt = `You are a helpful AI assistant for "TruckTalk Connect". You are designed to answer questions and provide general information about the add-on. Be concise and friendly.`;

  try {
    // First, determine the user's intent with a light-weight model
    const intentResponse = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${openaiApiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-3.5-turbo',
        messages: [{
          role: 'system',
          content: 'You are an intent detection AI. Analyze the user\'s message. If they are asking for data analysis or validation, respond with "analyze". For any other message, respond with "chat". Respond with only one word.'
        }, {
          role: 'user',
          content: userMessage
        }]
      })
    });
    
    if (!intentResponse.ok) {
        throw new Error('Failed to get intent from OpenAI.');
    }

    const intentData = await intentResponse.json();
    const intent = intentData.choices[0].message.content.trim().toLowerCase();
    
    let openaiResponse;
    let finalPayload;

    if (intent === 'analyze') {
      finalPayload = {
        model: 'gpt-4o',
        messages: [{
          role: 'system',
          content: analysisPrompt
        }, {
          role: 'user',
          content: `Analyze the headers and data below. If there are no errors, provide the parsed JSON. If there are errors, provide the issues array.\n\nHeaders: ${JSON.stringify(headers)}\nSample Data: ${JSON.stringify(sampleData)}`
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

    // Now, call OpenAI with the correct payload based on intent
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
        return res.status(200).json({
          ok: true,
          issues: [],
          loads: null,
          message: botResponse
        });
    }
    
  } catch (error) {
    console.error("Proxy Error:", error);
    res.status(500).json({ error: 'Failed to communicate with the OpenAI API.' });
  }
});

module.exports = app;