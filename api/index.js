const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');

const app = express();

// A simple in-memory cache to store generated results and avoid re-generation.
const cache = new Map();

// Vercel deployment requires a specific port, often provided by process.env.PORT.
const port = process.env.PORT || 3000;

// Configure CORS to allow requests from any origin.
// This is critical for web apps to communicate with the Vercel backend.
const corsOptions = {
    origin: '*', // Allow all origins for simplicity in a public API
    methods: ['POST', 'GET', 'OPTIONS'], // Explicitly allow POST, GET, and OPTIONS methods
    allowedHeaders: ['Content-Type', 'Authorization'], // Allow these headers
    credentials: true,
};

// Use the configured CORS middleware.
app.use(cors(corsOptions));
// Use express.json() to parse incoming JSON payloads.
app.use(express.json());

// Main route for processing OpenAI requests.
app.post('/openai-proxy', async (req, res) => {
    // Acknowledge the OPTIONS preflight request immediately
    // to avoid potential CORS errors on the client side.
    if (req.method === 'OPTIONS') {
        res.status(200).send();
        return;
    }

    try {
        const { prompt, model, useCache } = req.body;

        // Check if a response for this prompt and model is already in the cache.
        const cacheKey = `${model}:${prompt}`;
        if (useCache && cache.has(cacheKey)) {
            console.log("Serving from cache.");
            return res.status(200).json({
                text: cache.get(cacheKey),
                fromCache: true
            });
        }

        const apiKey = process.env.OPENAI_API_KEY;
        if (!apiKey) {
            console.error("OPENAI_API_KEY environment variable is not set.");
            return res.status(500).json({
                error: "Server configuration error: OpenAI API key is missing."
            });
        }

        const openAIResponse = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`,
            },
            body: JSON.stringify({
                model: model,
                messages: [{ role: 'user', content: prompt }],
            }),
        });

        if (!openAIResponse.ok) {
            const errorText = await openAIResponse.text();
            console.error(`OpenAI API request failed with status ${openAIResponse.status}: ${errorText}`);
            return res.status(openAIResponse.status).json({
                error: `OpenAI API error: ${errorText}`
            });
        }

        const data = await openAIResponse.json();
        const generatedText = data.choices[0].message.content.trim();

        // Store the new result in the cache.
        if (useCache) {
            cache.set(cacheKey, generatedText);
            console.log("Response cached.");
        }

        res.status(200).json({
            text: generatedText,
            fromCache: false
        });

    } catch (error) {
        console.error("Server error during OpenAI proxy request:", error);
        res.status(500).json({
            error: `Internal server error: ${error.message}`
        });
    }
});

// Catch-all route for requests to other paths or with unsupported methods.
app.all('*', (req, res) => {
    // This provides a helpful error message for methods other than POST, GET, and OPTIONS.
    res.status(405).json({
        error: 'Method Not Allowed',
        message: `This endpoint only accepts POST requests for data analysis or chat. The requested path is ${req.path}.`
    });
});

// Start the server. Vercel automatically handles this with a serverless function.
app.listen(port, () => {
    console.log(`Server listening at http://localhost:${port}`);
});

module.exports = app;