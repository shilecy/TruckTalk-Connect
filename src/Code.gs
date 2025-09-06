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
 * Handles a chat message by checking for sheet analysis commands or sending to the Vercel proxy.
 * @param {string} userMessage The user's message from the UI.
 * @return {string} The AI's response or an analysis result.
 */
function handleChatMessage(userMessage) {
  const trimmedMessage = userMessage.toLowerCase().trim();

  // Check if the user is asking for sheet analysis
  if (trimmedMessage.includes("analyze") || trimmedMessage.includes("summary")) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const result = validateSheet(sheet); // Assuming validateSheet is still in api.gs
    return result;
  }
  
  // If not, send the message to the Vercel proxy
  const proxyUrl = 'https://truck-talk-connect.vercel.app/openai-proxy';
  
  const payload = {
    prompt: userMessage,
    model: 'gpt-3.5-turbo'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(proxyUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (jsonResponse.choices && jsonResponse.choices.length > 0) {
      return jsonResponse.choices[0].message.content;
    } else {
      return 'No response from the proxy. Check your request and proxy URL.';
    }
  } catch (e) {
    return 'Error communicating with the proxy server: ' + e.message;
  }
}