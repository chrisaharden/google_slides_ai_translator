# Google Slides AI Translator

A Google Apps Script that enables automatic translation of Google Slides presentations using Google's Gemini AI. This script adds a translation menu to Google Slides, allowing you to translate your presentations into multiple languages with just a few clicks.

## Features

- Translate entire presentations or specific slides
- Multiple translation modes:
  - Translate all slides
  - Translate current slide only
  - Translate from current slide to end
- Support for multiple languages:
  - English
  - French
  - Spanish
  - Korean
  - Chinese (Simplified)
  - German
- Handles text in:
  - Shapes
  - Text boxes
  - Tables (including merged cells)
- Built-in retry mechanism with exponential backoff for API reliability
- User-friendly menu integration in Google Slides

## Installation

1. Open your Google Slides presentation
2. Go to `Extensions > Apps Script`
3. Delete any existing code in the script editor
4. Copy the entire content of `TranslationScript.js` and paste it into the script editor
5. Save the script (File > Save)

## Setup

1. Get a Google Gemini API key:
   - Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Create or select a project
   - Generate an API key

2. Add the API key to the script:
   - In the Apps Script editor, go to `Project Settings`
   - Under `Script Properties`, click `Add Script Property`
   - Set the property name as `GEMINI_API_KEY`
   - Set the value to your Gemini API key
   - Click `Save`

3. Authorize the script:
   - Return to your Google Slides presentation
   - Refresh the page
   - Click on the new `Translation` menu
   - Follow the authorization prompts

## Usage

1. Open your Google Slides presentation
2. Use the `Translation` menu to select your desired translation option:
   - `Translate All Slides`: Translates the entire presentation
   - `Translate Current Slide`: Translates only the currently selected slide
   - `Translate Current to End`: Translates from the current slide to the end of the presentation

3. Select your target language from the submenu

The script will process each text element and table cell, translating the content while maintaining the original formatting.

## About Translation Features

Access the `About Translation` menu for quick help and information about the script's features. This includes:
- Link to the latest version of the script
- Setup instructions
- Information about bug reporting and contributions

## Error Handling

The script includes robust error handling features:
- Exponential backoff for API rate limiting
- Automatic retries for failed API calls
- Skipping of empty text and special fields
- Detailed logging for troubleshooting

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## Contributing

Feel free to submit bugs and pull requests in the GitHub repository.

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/chrisaharden/google_slides_ai_translator).
