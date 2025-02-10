// This nifty script was generated by Claude from Anthropic.
// Copy and paste this simple AppScript into your Google slides along with an API key and translate to French, German, etc.
// Edit the bottom to add more languages.

function translatePresentation(targetLanguage, mode = 'all') {
  //debug
  if(targetLanguage == undefined) targetLanguage = "Spanish";

  // Get the active presentation and slides
  const presentation = SlidesApp.getActivePresentation();
  let slides;
  
  if (mode === 'current') {
    // Get only the current slide
    const currentSlide = presentation.getSelection().getCurrentPage();
    if (!currentSlide) {
      Logger.log('No slide selected');
      return;
    }
    slides = [currentSlide];
  } else if (mode === 'current_to_end') {
    // Get current slide and all following slides
    const currentSlide = presentation.getSelection().getCurrentPage();
    if (!currentSlide) {
      Logger.log('No slide selected');
      return;
    }
    const allSlides = presentation.getSlides();
    const currentIndex = allSlides.findIndex(slide => slide.getObjectId() === currentSlide.getObjectId());
    if (currentIndex === -1) {
      Logger.log('Could not determine current slide position');
      return;
    }
    slides = allSlides.slice(currentIndex);
  } else {
    // Get all slides
    slides = presentation.getSlides();
  }
  
  // Replace these with your API endpoint and key
  //const API_ENDPOINT = 'YOUR_API_ENDPOINT';
  //const API_KEY = 'YOUR_API_KEY';
  
  //Gemini
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); 
  if (!API_KEY) {

      const ui = SlidesApp.getUi();
      ui.alert(
          'Missing API Key',
          `The API key for this script is missing or has not been set.`,
          ui.ButtonSet.OK
        );

    Logger.log('API key is missing or not set.');
    return;
  }
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-002:generateContent?key=${API_KEY}`;
  
  
  // Iterate through each slide
  slides.forEach((slide, arrayIndex) => {
    // Calculate the actual slide index based on mode
    let slideIndex;
    if (mode === 'current') {
      slideIndex = presentation.getSlides().findIndex(s => s.getObjectId() === slide.getObjectId());
    } else if (mode === 'current_to_end') {
      const currentSlide = presentation.getSelection().getCurrentPage();
      const startIndex = presentation.getSlides().findIndex(s => s.getObjectId() === currentSlide.getObjectId());
      slideIndex = startIndex + arrayIndex;
    } else {
      slideIndex = arrayIndex;
    }
    // Get all elements on the slide
    const elements = slide.getPageElements();
    
    elements.forEach((element, elementIndex) => {
      try {
        // Check if the element is a shape and has text
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE && element.asShape().getText()) {
          const shape = element.asShape();
          const textRange = shape.getText();
          const originalText = textRange.asString();
          
          // Skip empty text, spaces, or non-printable characters
          if (!originalText.trim() || originalText.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
            Logger.log(`Skipping special field that may cause trash returns, like automated slide numbers.`);
            return;
          }
          try {
            // Make API request to translate text
            const translatedText = translateText(originalText, targetLanguage, API_ENDPOINT, API_KEY);
            
            // Update the text in the shape
            textRange.setText(translatedText);
            
            // Add small delay to avoid rate limits
            Utilities.sleep(100);
            
            // Log progress
            Logger.log(`Translated text to ${targetLanguage} on slide ${slideIndex + 1}, element ${elementIndex + 1}`);
          } catch (error) {
            Logger.log(`Error translating text on slide ${slideIndex + 1}, element ${elementIndex + 1}: ${error}`);
          }
        }
      } catch (error) {
        Logger.log(`Error processing element on slide ${slideIndex + 1}, element ${elementIndex + 1}: ${error}`);
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
          
          switch(cell.getMergeState()) {
            case SlidesApp.CellMergeState.HEAD: 
              Logger.log(`Table HEAD merged cell found`); 
              break;
            case SlidesApp.CellMergeState.MERGED:
              Logger.log(`Table MERGED, not head cell found`); 
              continue; // Skip if not the top-left cell
              break;
            case SlidesApp.CellMergeState.NORMAL:
              Logger.log(`Table NORMAL cell found`); 
              break;
            default:
              Logger.log(`defaulting`); 
          }

          //if (cell.isPartOfMerge()) {
            // Only process the top-left cell of the merge range
          //  if (row !== cell.getRow() || col !== cell.getColumn()) {
          //    continue; // Skip if not the top-left cell
          //  }
          //}
          
          const originalText = cell.getText().asString();
          
          // Skip empty cells, spaces, or non-printable characters
          if (!originalText.trim() || originalText.replace(/[\s\u0000-\u001F\u007F-\u009F\uE907]/g, '') === '') {
            Logger.log(`Skipping special field that may cause trash returns, like automated slide numbers.`);
            continue;
          }
          
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
  
  // Create a time-driven trigger to show the copy prompt after 2 seconds
  const triggerId = "copyPromptTrigger_" + new Date().getTime();
  PropertiesService.getScriptProperties().setProperty('LAST_TRANSLATION_NAME', presentation.getName());
  PropertiesService.getScriptProperties().setProperty('LAST_TRANSLATION_LANGUAGE', targetLanguage);
  
  ScriptApp.newTrigger('showCopyPrompt')
    .timeBased()
    .after(2000) // 2 seconds
    .create();
}

// Function to show copy prompt and handle the response
function showCopyPrompt() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const originalName = PropertiesService.getScriptProperties().getProperty('LAST_TRANSLATION_NAME');
  const targetLanguage = PropertiesService.getScriptProperties().getProperty('LAST_TRANSLATION_LANGUAGE');
  
  const response = ui.alert(
    'Translation Complete',
    'Would you like to make a copy of this presentation?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const copy = presentation.copy(`${originalName} (${targetLanguage} Translation)`);
    ui.alert(
      'Copy Created',
      'A copy of the presentation has been created with the translated content.',
      ui.ButtonSet.OK
    );
  }
  
  // Clean up properties
  PropertiesService.getScriptProperties().deleteProperty('LAST_TRANSLATION_NAME');
  PropertiesService.getScriptProperties().deleteProperty('LAST_TRANSLATION_LANGUAGE');
  
  // Delete all triggers for this function to clean up
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === 'showCopyPrompt') {
      ScriptApp.deleteTrigger(trigger);
    }
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

    //Harden
    //if (!json["candidates"][0]["content"]["parts"][0]["text"]) {
    //  throw new Error('Unexpected Gemini API response structure: ' + JSON.stringify(json));
    //}

    //text = json["candidates"][0]["content"]["parts"][0]["text"];
    //return text;

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
  translatePresentation('French', 'all');
}

function translateToSpanish() {
  translatePresentation('Spanish', 'all');
}

function translateToEnglish() {
  translatePresentation('English', 'all');
}

function translateToKorean() {
  translatePresentation('Korean', 'all');
}

function translateToChinese() {
  translatePresentation('Simplified Chinese', 'all');
}

function translateToGerman() {
  translatePresentation('German', 'all');
}

// Functions for translating current slide only
function translateCurrentToEnglish() {
  translatePresentation('English', 'current');
}

function translateCurrentToFrench() {
  translatePresentation('French', 'current');
}

function translateCurrentToSpanish() {
  translatePresentation('Spanish', 'current');
}

function translateCurrentToKorean() {
  translatePresentation('Korean', 'current');
}

function translateCurrentToChinese() {
  translatePresentation('Simplified Chinese', 'current');
}

function translateCurrentToGerman() {
  translatePresentation('German', 'current');
}

// Functions for translating from current slide to end
function translateCurrentToEndEnglish() {
  translatePresentation('English', 'current_to_end');
}

function translateCurrentToEndFrench() {
  translatePresentation('French', 'current_to_end');
}

function translateCurrentToEndSpanish() {
  translatePresentation('Spanish', 'current_to_end');
}

function translateCurrentToEndKorean() {
  translatePresentation('Korean', 'current_to_end');
}

function translateCurrentToEndChinese() {
  translatePresentation('Simplified Chinese', 'current_to_end');
}

function translateCurrentToEndGerman() {
  translatePresentation('German', 'current_to_end');
}

// Add menu items to trigger translations
function onOpen() {
  const ui = SlidesApp.getUi();
  const menu = ui.createMenu('Translation')
    .addSubMenu(ui.createMenu('Translate All Slides')
      .addItem('to English', 'translateToEnglish')
      .addItem('to French', 'translateToFrench')
      .addItem('to Spanish', 'translateToSpanish')
      .addItem('to Korean', 'translateToKorean')
      .addItem('to Chinese', 'translateToChinese')
      .addItem('to German', 'translateToGerman'))
    .addSubMenu(ui.createMenu('Translate Current Slide')
      .addItem('to English', 'translateCurrentToEnglish')
      .addItem('to French', 'translateCurrentToFrench')
      .addItem('to Spanish', 'translateCurrentToSpanish')
      .addItem('to Korean', 'translateCurrentToKorean')
      .addItem('to Chinese', 'translateCurrentToChinese')
      .addItem('to German', 'translateCurrentToGerman'))
    .addSubMenu(ui.createMenu('Translate Current to End')
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
  const ui = SlidesApp.getUi();
  ui.alert(
    'Translation Features',
    'Google slides doesn\'t make it possible to detect page number fields.  So translations often come back with a note that the model doesn\'t understand or can\'t see the content you are passing.  You\'ll need to manually put page numbers back to the page number fields after translation is done.\n',
    ui.ButtonSet.OK
  );
}

function testGeminiAPI() {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-002:generateContent?key=${API_KEY}`;
  
  const options = {
    method: 'POST',
    headers: {
      //'Authorization': `Bearer ${API_KEY}`,
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
