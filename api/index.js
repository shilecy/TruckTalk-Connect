import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

// Proxy endpoint for Sheets Add-on
app.post("/api/index.js", async (req, res) => {
  const { headers, sampleData, requiredFields, headerMappings } = req.body;

  const systemPrompt = `
  You convert a 2D table of logistics loads into typed JSON.
  Never invent data. Unknowns stay blank and flagged as issues.
  Dates must be ISO 8601 UTC.
  Return JSON strictly as:
  {
    ok: boolean,
    issues: Array<{
      code: string,
      severity: "error" | "warn",
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

  const userPrompt = {
    headers,
    rows: sampleData,
    knownSynonyms: headerMappings,
    requiredFields,
  };

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${process.env.OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-4.1-mini", // or gpt-3.5-turbo if you want cheaper
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: JSON.stringify(userPrompt) }
        ],
        temperature: 0,
      }),
    });

    const data = await response.json();

    let parsed;
    try {
      parsed = JSON.parse(data.choices[0].message.content);
    } catch (err) {
      throw new Error("AI did not return valid JSON.");
    }

    res.json(parsed);
  } catch (err) {
    console.error("Proxy error:", err);
    res.status(500).json({
      ok: false,
      issues: [{ code: "SERVER_ERROR", severity: "error", message: err.message }],
      mapping: {},
      meta: { analyzedRows: 0, analyzedAt: new Date().toISOString() }
    });
  }
});

export default app;
