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

    **Your Goal:** Return a single JSON object that strictly matches the 'AnalysisResult' schema.
    **Core Rules:**
    1. **Header Mapping:** Map the user's headers to the canonical 'Load' schema fields. Use the provided 'HEADER_SYNONYMS'.
    2. **Data Validation & Normalization:**
      * **Dates:** All date fields MUST be converted to ISO 8601 UTC format.
      * **Required Fields:** All fields in the Load schema are required.
      * **Uniqueness:** 'loadNumber' must be unique.
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
        "issues": [{"severity": "'error'|'warn'", "message": "string", "suggestion": "string", "rows": "number[]", "column": "string (Original Header Name)"}],
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
    
    // Log the raw response status and text for debugging
    console.log('OpenAI Intent API Response Status:', intentResponse.status);
    const intentText = await intentResponse.text();
    console.log('OpenAI Intent API Response Body:', intentText);

    if (!intentResponse.ok) {
        throw new Error('Failed to get intent from OpenAI.');
    }

    const intentData = JSON.parse(intentText);
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
        // It's a structured analysis result.
        
        // This is the new logic for handling a structured response.
        // It checks if there are issues and returns a structured response for the UI to handle.
        if (jsonResult.ok === false && jsonResult.issues && jsonResult.issues.length > 0) {
            return res.status(200).json({
                type: 'suggestion_list',
                message: 'Analysis complete. Found issues:',
                issues: jsonResult.issues.map(issue => ({
                    message: `[${issue.severity.toUpperCase()}] ${issue.message}`,
                    suggestion: issue.suggestion,
                    action: {
                        type: 'jump_to_cell',
                        column: issue.column,
                        row: issue.rows[0] // Assuming one row per issue for this example
                    }
                }))
            });
        }
        
        return res.status(200).json(jsonResult);
    } catch (e) {
        // It's a general conversational response.
        return res.status(200).send(botResponse);
    }
    
  } catch (error) {
    console.error("Proxy Error:", error);
    res.status(500).json({ error: 'Failed to communicate with the OpenAI API.' });
  }
});

module.exports = app;