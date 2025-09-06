const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(cors());

// A basic root route to confirm the server is running.
app.get('/', (req, res) => {
  res.status(200).send('OpenAI Proxy server is running.');
});

// The main endpoint for the chatbot.
app.post('/openai-proxy', async (req, res) => {
  const { prompt, model } = req.body;
  const openaiApiKey = process.env.OPENAI_API_KEY;

  if (!openaiApiKey) {
    return res.status(500).json({ error: 'OpenAI API key not set.' });
  }

  try {
    const openaiResponse = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${openaiApiKey}`
      },
      body: JSON.stringify({
        model: model || 'gpt-3.5-turbo',
        messages: [{ role: 'user', content: prompt }]
      })
    });

    const data = await openaiResponse.json();
    res.status(openaiResponse.status).json(data);
  } catch (error) {
    res.status(500).json({ error: 'Failed to communicate with the OpenAI API.' });
  }
});

// This is the part that is not used by Vercel's serverless platform.
// The vercel.json file handles the server startup.
app.listen(PORT, () => {
  console.log(`Proxy server listening on port ${PORT}`);
});