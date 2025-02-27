// This nifty script was adapted by Claude from Anthropic.
// Copy and paste this simple AppScript into your Google Docs along with an API key and translate to French, German, etc.
// Edit the bottom to add more languages.

function translateDocument(targetLanguage, mode = 'all') {
  //debug
  if(targetLanguage == undefined) targetLanguage = "Spanish";

  // Get the active document
  const document = DocumentApp.getActiveDocument();
  let elements = [];
  
  if (mode === 'current') {
    // Get only the current selection
    const selection = document.getSelection();
    if (!selection) {
      Logger.log('No text selected');
      return;
    }
    elements = selection.getRangeElements();
  } else if (mode === 'current_to_end') {
    // Get current position and all following content
    const cursor = document.getCursor();
    if (!cursor) {
      Logger.log('No cursor position found');
      return;
    }
    const body = document.getBody();
    const cursorPosition = cursor.getElement().getParent().getChildIndex(cursor.getElement());
    const totalElements = body.getNumChildren();
    
    // Get all elements from cursor position to end
    for (let i = cursorPosition; i < totalElements; i++) {
      elements.push(body.getChild(i));
    }
  } else {
    // Get all content
    const body = document.getBody();
    for (let i = 0; i < body.getNumChildren(); i++) {
      elements.push(body.getChild(i));
    }
  }
   
  //--------------------------------------------------------------------------------
  //NOTE: UNCOMMENT THE API YOU WANT TO USE AND COMMENT OUT THE OTHERS.  YOU CAN ONLY USE ONE AT A TIME. I HAVE COMMENTED OUT CLAUDE FOR NOW.  - Chris

  //Gemini
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); 
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-002:generateContent?key=${API_KEY}`;
  Logger.log('Using Gemini for Translation');
  
  //Claude - This is the direct endpoint format.  I have not added code for access via AWS Bedrock.  I hope to do that soon. - Chris
  //const API_KEY = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY'); 
  //const API_ENDPOINT = 'https://api.anthropic.com/v1/messages';
  //Logger.log('Using Claude for Translation');

  //--------------------------------------------------------------------------------

  // Check if API key is set.  If not, alert the user.
  if (!API_KEY) {
      const ui = DocumentApp.getUi();
      ui.alert(
          'Missing API Key',
          `The API key for this script is missing or has not been set.\n Please add your API key to the script properties, and check the code for switching between different models.`,
          ui.ButtonSet.OK
        );

    Logger.log('API key is missing or not set.');
    return;
  }
  
  // Process each element
  processElements(elements, targetLanguage, API_ENDPOINT, API_KEY);
  
  Logger.log(`Translation to ${targetLanguage} completed!`);
}

// Helper function to process document elements
function processElements(elements, targetLanguage, API_ENDPOINT, API_KEY) {
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    
    try {
      // Handle different element types
      if (element.getType) {
        // For selection range elements
        processElementByType(element.getElement(), targetLanguage, API_ENDPOINT, API_KEY);
      } else {
        // For direct elements
        processElementByType(element, targetLanguage, API_ENDPOINT, API_KEY);
      }
    } catch (error) {
      Logger.log(`Error processing element ${i}: ${error}`);
    }
  }
}

// Process elements based on their type
function processElementByType(element, targetLanguage, API_ENDPOINT, API_KEY) {
  const elementType = element.getType();
  
  switch (elementType) {
    case DocumentApp.ElementType.PARAGRAPH:
      processParagraph(element, targetLanguage, API_ENDPOINT, API_KEY);
      break;
    case DocumentApp.ElementType.TABLE:
      processTable(element, targetLanguage, API_ENDPOINT, API_KEY);
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      processListItem(element, targetLanguage, API_ENDPOINT, API_KEY);
      break;
    case DocumentApp.ElementType.TEXT:
      processText(element, targetLanguage, API_ENDPOINT, API_KEY);
      break;
    // Add more element types as needed
    default:
      // Try to get text if possible
      if (element.getText && typeof element.getText === 'function') {
        processGenericElement(element, targetLanguage, API_ENDPOINT, API_KEY);
      }
  }
}

// Process paragraph elements
function processParagraph(paragraph, targetLanguage, API_ENDPOINT, API_KEY) {
  const text = paragraph.getText();
  
  // Skip empty paragraphs or those with only whitespace/special characters
  if (!text.trim() || text.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
    return;
  }
  
  try {
    const translatedText = translateText(text, targetLanguage, API_ENDPOINT, API_KEY);
    paragraph.setText(translatedText);
    
    // Add small delay to avoid rate limits
    Utilities.sleep(100);
    
    Logger.log(`Translated paragraph: "${text.substring(0, 30)}...""`);
  } catch (error) {
    Logger.log(`Error translating paragraph: ${error}`);
  }
}

// Process text elements
function processText(textElement, targetLanguage, API_ENDPOINT, API_KEY) {
  const text = textElement.getText();
  
  // Skip empty text or those with only whitespace/special characters
  if (!text.trim() || text.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
    return;
  }
  
  try {
    const translatedText = translateText(text, targetLanguage, API_ENDPOINT, API_KEY);
    textElement.setText(translatedText);
    
    // Add small delay to avoid rate limits
    Utilities.sleep(100);
    
    Logger.log(`Translated text: "${text.substring(0, 30)}...""`);
  } catch (error) {
    Logger.log(`Error translating text: ${error}`);
  }
}

// Process list items
function processListItem(listItem, targetLanguage, API_ENDPOINT, API_KEY) {
  const text = listItem.getText();
  
  // Skip empty list items
  if (!text.trim() || text.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
    return;
  }
  
  try {
    const translatedText = translateText(text, targetLanguage, API_ENDPOINT, API_KEY);
    listItem.setText(translatedText);
    
    // Add small delay to avoid rate limits
    Utilities.sleep(100);
    
    Logger.log(`Translated list item: "${text.substring(0, 30)}...""`);
  } catch (error) {
    Logger.log(`Error translating list item: ${error}`);
  }
}

// Process table elements
function processTable(table, targetLanguage, API_ENDPOINT, API_KEY) {
  const numRows = table.getNumRows();
  
  for (let row = 0; row < numRows; row++) {
    const tableRow = table.getRow(row);
    const numCells = tableRow.getNumCells();
    
    for (let col = 0; col < numCells; col++) {
      const cell = tableRow.getCell(col);
      const text = cell.getText();
      
      // Skip empty cells
      if (!text.trim() || text.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
        continue;
      }
      
      try {
        const translatedText = translateText(text, targetLanguage, API_ENDPOINT, API_KEY);
        cell.setText(translatedText);
        
        // Add small delay to avoid rate limits
        Utilities.sleep(100);
        
        Logger.log(`Translated table cell (${row},${col}): "${text.substring(0, 30)}...""`);
      } catch (error) {
        Logger.log(`Error translating table cell (${row},${col}): ${error}`);
      }
    }
  }
}

// Process any element that has getText/setText methods
function processGenericElement(element, targetLanguage, API_ENDPOINT, API_KEY) {
  if (!element.getText || !element.setText || 
      typeof element.getText !== 'function' || 
      typeof element.setText !== 'function') {
    return;
  }
  
  const text = element.getText();
  
  // Skip empty elements
  if (!text.trim() || text.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
    return;
  }
  
  try {
    const translatedText = translateText(text, targetLanguage, API_ENDPOINT, API_KEY);
    element.setText(translatedText);
    
    // Add small delay to avoid rate limits
    Utilities.sleep(100);
    
    Logger.log(`Translated generic element: "${text.substring(0, 30)}...""`);
  } catch (error) {
    Logger.log(`Error translating generic element: ${error}`);
  }
}

// Helper function to check if response indicates API overload
function isApiOverloaded(response) {
  try {
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);
    
    // Check for various overload indicators
    const overloadPatterns = [
      /"code":\s*503/,
      /"message":\s*"The model is overloaded/,
      /"status":\s*"UNAVAILABLE"/,
      /"error":\s*"overloaded"/i,
      /"error":\s*"rate_limit_exceeded"/i
    ];
    
    // Check response text against patterns
    if (overloadPatterns.some(pattern => pattern.test(responseText))) {
      return true;
    }
    
    // Check specific error messages in different response structures
    if (json.error?.message?.toLowerCase().includes('overloaded')) return true;
    if (json.error?.toLowerCase().includes('capacity')) return true;
    
    return false;
  } catch (e) {
    // If we can't parse the response, assume it's not an overload
    return false;
  }
}

// Helper function to implement retry logic with exponential backoff
function retryWithExponentialBackoff(apiCall) {
  const maxRetries = 5;
  const baseDelay = 1000; // Start with 1 second delay
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = apiCall();
      
      // If response indicates overload, throw error to trigger retry
      if (isApiOverloaded(response)) {
        throw new Error('API overloaded');
      }
      
      return response;
    } catch (error) {
      if (attempt === maxRetries) {
        throw error; // Rethrow if we've exhausted all retries
      }
      
      // Calculate delay with exponential backoff and some random jitter
      const delay = baseDelay * Math.pow(2, attempt) + Math.random() * 1000;
      Logger.log(`API overloaded, retrying in ${Math.round(delay/1000)} seconds...`);
      Utilities.sleep(delay);
    }
  }
}

// Function for OpenAI's ChatGPT API
function translateTextWithChatGPT(text, targetLanguage, apiKey) {
  const endpoint = 'https://api.openai.com/v1/chat/completions';
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'system',
          content: `You are a translator. Translate the following text from English to ${targetLanguage}. Only respond with the translation, no additional text.`
        },
        {
          role: 'user',
          content: text
        }
      ],
      temperature: 0.3
    }),
    muteHttpExceptions: true
  };
  
  try {
    const response = retryWithExponentialBackoff(() => UrlFetchApp.fetch(endpoint, options));
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content.trim();
  } catch (error) {
    Logger.log('ChatGPT API error: ' + error);
    throw error;
  }
}

// Function for Anthropic's Claude API
function translateTextWithClaude(text, targetLanguage, apiKey) {
  const endpoint = 'https://api.anthropic.com/v1/messages';
  const options = {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: 'claude-3-sonnet-20240229',
      max_tokens: 1024,
      temperature: 0.3,
      messages: [
        {
          role: 'user',
          content: `Translate this English text to ${targetLanguage}. Only respond with the translation, no additional text: ${text}`
        }
      ]
    }),
    muteHttpExceptions: true
  };
  
  try {
    const response = retryWithExponentialBackoff(() => UrlFetchApp.fetch(endpoint, options));
    const json = JSON.parse(response.getContentText());
    return json.content[0].text.trim();
  } catch (error) {
    Logger.log('Claude API error: ' + error);
    throw error;
  }
}

// Function for Google's Gemini API
function translateTextWithGemini(text, targetLanguage, apiKey) {
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-002:generateContent?key=${apiKey}`;
  const options = {
    method: 'POST',
    headers: {
      //Harden - AI is adding a bearer auth entry - take that out.  The key is in the URL submission
      // 'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      contents: {
        role: "user",
        parts: [{
          text: `Translate this English text to ${targetLanguage}. Only respond with the translation, no additional text: ${text}`
        }]
      },
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 1024
      }
    }),
    muteHttpExceptions: true
  };
  
  try {
    const response = retryWithExponentialBackoff(() => UrlFetchApp.fetch(endpoint, options));
    const json = JSON.parse(response.getContentText());

    // The Gemini API response structure has the content at a different path
    if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
      throw new Error('Unexpected Gemini API response structure: ' + JSON.stringify(json));
    }
    // Handle the response structure more carefully
    if (!json.candidates?.[0]?.content?.parts?.[0]?.text) {
      throw new Error('Unexpected Gemini API response structure: ' + JSON.stringify(json));
    }
    return json.candidates[0].content.parts[0].text.trim();
  } catch (error) {
    Logger.log('Gemini API error: ' + error);
    throw error;
  }
}

// Main translation function that can use any API
function translateText(text, targetLanguage, apiEndpoint, apiKey) {
  // Detect which API to use based on the endpoint
  if (apiEndpoint.includes('openai')) {
    return translateTextWithChatGPT(text, targetLanguage, apiKey);
  } else if (apiEndpoint.includes('anthropic')) {
    return translateTextWithClaude(text, targetLanguage, apiKey);
  } else if (apiEndpoint.includes('googleapis')) {
    return translateTextWithGemini(text, targetLanguage, apiKey);
  } else {
    throw new Error('Unsupported API endpoint');
  }
}

// Function handlers for each language
function translateToFrench() {
  translateDocument('French', 'all');
}

function translateToSpanish() {
  translateDocument('Spanish', 'all');
}

function translateToEnglish() {
  translateDocument('English', 'all');
}

function translateToKorean() {
  translateDocument('Korean', 'all');
}

function translateToChinese() {
  translateDocument('Simplified Chinese', 'all');
}

function translateToGerman() {
  translateDocument('German', 'all');
}

// Functions for translating current selection only
function translateCurrentToEnglish() {
  translateDocument('English', 'current');
}

function translateCurrentToFrench() {
  translateDocument('French', 'current');
}

function translateCurrentToSpanish() {
  translateDocument('Spanish', 'current');
}

function translateCurrentToKorean() {
  translateDocument('Korean', 'current');
}

function translateCurrentToChinese() {
  translateDocument('Simplified Chinese', 'current');
}

function translateCurrentToGerman() {
  translateDocument('German', 'current');
}

// Functions for translating from current position to end
function translateCurrentToEndEnglish() {
  translateDocument('English', 'current_to_end');
}

function translateCurrentToEndFrench() {
  translateDocument('French', 'current_to_end');
}

function translateCurrentToEndSpanish() {
  translateDocument('Spanish', 'current_to_end');
}

function translateCurrentToEndKorean() {
  translateDocument('Korean', 'current_to_end');
}

function translateCurrentToEndChinese() {
  translateDocument('Simplified Chinese', 'current_to_end');
}

function translateCurrentToEndGerman() {
  translateDocument('German', 'current_to_end');
}

// Add menu items to trigger translations
function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('Translation')
    .addSubMenu(ui.createMenu('Translate Entire Document')
      .addItem('to English', 'translateToEnglish')
      .addItem('to French', 'translateToFrench')
      .addItem('to Spanish', 'translateToSpanish')
      .addItem('to Korean', 'translateToKorean')
      .addItem('to Chinese', 'translateToChinese')
      .addItem('to German', 'translateToGerman'))
    .addSubMenu(ui.createMenu('Translate Selection')
      .addItem('to English', 'translateCurrentToEnglish')
      .addItem('to French', 'translateCurrentToFrench')
      .addItem('to Spanish', 'translateCurrentToSpanish')
      .addItem('to Korean', 'translateCurrentToKorean')
      .addItem('to Chinese', 'translateCurrentToChinese')
      .addItem('to German', 'translateCurrentToGerman'))
    .addSubMenu(ui.createMenu('Translate From Cursor to End')
      .addItem('to English', 'translateCurrentToEndEnglish')
      .addItem('to French', 'translateCurrentToEndFrench')
      .addItem('to Spanish', 'translateCurrentToEndSpanish')
      .addItem('to Korean', 'translateCurrentToEndKorean')
      .addItem('to Chinese', 'translateCurrentToEndChinese')
      .addItem('to German', 'translateCurrentToEndGerman'));
  
  menu.addToUi();

  // Add help menu
  ui.createMenu('About Translation')
    .addItem('About Translation Features', 'showTranslationHelp')
    .addToUi();
}

// Help function to explain features
function showTranslationHelp() {
  const ui = DocumentApp.getUi();
  ui.alert(
    'About Translation',
    'This script allows you to translate Google Docs documents using AI translation services.\n\nYou will need to add GEMINI_API_KEY (and your own Gemini API key) to a new field in Menu > Extensions > Apps Script > Project Settings, and when you run the script, you\'ll be asked to authorize it to run.',
    ui.ButtonSet.OK
  );
}

function testGeminiAPI() {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-002:generateContent?key=${API_KEY}`;
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      contents: {
        role: "user",
        parts: [{
          text: `Hello Gemini, how are you?`
        }]
      },
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 1024
      }
    }),
    muteHttpExceptions: false
  };

  try {
    const response = retryWithExponentialBackoff(() => UrlFetchApp.fetch(API_ENDPOINT, options));
    const json = JSON.parse(response.getContentText());
    Logger.log(json); 
  } catch (error) {
    Logger.log('Error: ' + error);
  }
}
