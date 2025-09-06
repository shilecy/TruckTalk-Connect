// Runs automatically when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TruckTalk Connect")
    .addItem("Open Chat", "showSidebar")
    .addToUi();
}

// Opens the sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("TruckTalk Connect")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Handles a chat message by sending it to the OpenAI API and returning the response.
 * @param {string} userMessage The user's message from the UI.
 * @return {string} The AI's response.
 */
function handleChatMessage(userMessage) {
  // Retrieve the API key from the script properties.
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const endpoint = 'https://api.openai.com/v1/chat/completions';
  
  if (!apiKey) {
    return 'Error: OpenAI API key not found. Please add it to your project properties.';
  }

  const payload = {
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'system',
        content: 'You are a helpful assistant for analyzing and summarizing data from a Google Sheet.'
      },
      {
        role: 'user',
        content: userMessage
      }
    ],
    temperature: 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (jsonResponse.choices && jsonResponse.choices.length > 0) {
      return jsonResponse.choices[0].message.content;
    } else {
      return 'No response from the API. Check your request and API key.';
    }
  } catch (e) {
    return 'Error communicating with the OpenAI API: ' + e.message;
  }
}