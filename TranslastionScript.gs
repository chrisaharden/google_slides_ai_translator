  function translatePresentation(targetLanguage) {
  // Get the active presentation
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  // Replace these with your API endpoint and key
  const API_ENDPOINT = 'YOUR_API_ENDPOINT';
  const API_KEY = 'YOUR_API_KEY';
  
  // Iterate through each slide
  slides.forEach((slide, slideIndex) => {
    // Get all shape elements on the slide that might contain text
    const shapes = slide.getShapes();
    
    shapes.forEach((shape, shapeIndex) => {
      // Check if the shape has text
      if (shape.getText()) {
        const textRange = shape.getText();
        const originalText = textRange.asString();
        
        // Skip empty text
        if (!originalText.trim()) return;
        
        try {
          // Make API request to translate text
          const translatedText = translateText(originalText, targetLanguage, API_ENDPOINT, API_KEY);
          
          // Update the text in the shape
          textRange.setText(translatedText);
          
          // Add small delay to avoid rate limits
          Utilities.sleep(100);
          
          // Log progress
          Logger.log(`Translated text to ${targetLanguage} on slide ${slideIndex + 1}, shape ${shapeIndex + 1}`);
        } catch (error) {
          Logger.log(`Error translating text on slide ${slideIndex + 1}, shape ${shapeIndex + 1}: ${error}`);
        }
      }
    });
    
    // Also check for any tables on the slide
    const tables = slide.getTables();
    tables.forEach((table, tableIndex) => {
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      
      // Iterate through each cell
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          const cell = table.getCell(row, col);
          const originalText = cell.getText().asString();
          
          // Skip empty cells
          if (!originalText.trim()) continue;
          
          try {
            // Translate cell content
            const translatedText = translateText(originalText, targetLanguage, API_ENDPOINT, API_KEY);
            cell.getText().setText(translatedText);
            
            // Add small delay to avoid rate limits
            Utilities.sleep(100);
            
            Logger.log(`Translated table ${tableIndex + 1} cell (${row},${col}) to ${targetLanguage} on slide ${slideIndex + 1}`);
          } catch (error) {
            Logger.log(`Error translating table cell on slide ${slideIndex + 1}: ${error}`);
          }
        }
      }
    });
  });
  
  Logger.log(`Translation to ${targetLanguage} completed!`);
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
    const response = UrlFetchApp.fetch(endpoint, options);
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
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());
    return json.content[0].text.trim();
  } catch (error) {
    Logger.log('Claude API error: ' + error);
    throw error;
  }
}

// Main translation function that can use either API
function translateText(text, targetLanguage, apiEndpoint, apiKey) {
  // Detect which API to use based on the endpoint
  if (apiEndpoint.includes('openai')) {
    return translateTextWithChatGPT(text, targetLanguage, apiKey);
  } else if (apiEndpoint.includes('anthropic')) {
    return translateTextWithClaude(text, targetLanguage, apiKey);
  } else {
    throw new Error('Unsupported API endpoint');
  }
}

// Function handlers for each language
function translateToFrench() {
  translatePresentation('French');
}

function translateToSpanish() {
  translatePresentation('Spanish');
}

function translateToEnglish() {
  translatePresentation('English');
}

function translateToKorean() {
  translatePresentation('Korean');
}

function translateToChinese() {
  translatePresentation('Simplified Chinese');
}

function translateToGerman() {
  translatePresentation('German');
}


// Add menu items to trigger translations
function onOpen() {
  SlidesApp.getUi()
    .createMenu('Translation')
    .addItem('Translate to English', 'translateToEnglish')
    .addItem('Translate to French', 'translateToFrench')
    .addItem('Translate to Spanish', 'translateToSpanish')
    .addItem('Translate to Korean', 'translateToKorean')
    .addItem('Translate to Chinese', 'translateToChinese')
    .addItem('Translate to German', 'translateToGerman')
    .addToUi();
}