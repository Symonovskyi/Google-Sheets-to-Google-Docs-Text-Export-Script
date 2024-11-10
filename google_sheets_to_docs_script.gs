/**
 * Google Sheets to Google Docs Text Export Script
 * 
 * This script allows users to export text from a selected range of Google Sheets into a Google Docs document while preserving text formatting.
 * 
 * --- Summary ---
 * The script transfers text along with its formatting (bold, italic, underline, color, etc.) from Google Sheets to a new Google Doc.
 * It provides an easy-to-use menu in Google Sheets and handles both simple and complex formatting scenarios.
 * 
 * --- Detailed Overview ---
 * The script performs the following steps:
 * 1. Adds a custom menu called 'Custom Scripts' to Google Sheets, allowing the user to select a menu option to start the export process.
 * 2. Once initiated, the script takes the currently selected range in Google Sheets and copies the text into a new Google Docs document.
 * 3. The script handles rich text formatting from Google Sheets, meaning that if the text in a cell has different styles (e.g., part of the text is bold, another part is underlined), the exported text retains these styles.
 * 4. Text formatting is preserved, including attributes like bold, italic, underline, strikethrough, font color, background color, font family, and font size.
 * 5. Logging is included to help in debugging the script, but it can be toggled on or off using the `LOGGING_ENABLED` flag.
 * 
 * --- How to Use ---
 * 1. Open the Google Sheet where you want to use this script.
 * 2. Paste this script into the Script Editor (Extensions > Apps Script).
 * 3. Save the script and refresh the Google Sheet.
 * 4. After refreshing, you will see a new custom menu called 'Custom Scripts'.
 * 5. Select the range of cells you want to export to Google Docs.
 * 6. From the 'Custom Scripts' menu, click on 'Export to Google Doc'.
 * 7. A new Google Document will be created with the selected content, preserving the formatting from the Google Sheet.
 * 
 * --- Key Features ---
 * - Adds a user-friendly menu to Google Sheets to initiate the export process.
 * - Transfers selected range content to Google Docs while preserving rich text formatting.
 * - Handles multiple formatting attributes on the same cell text (e.g., mixed bold and italic).
 * - Supports different font families, font sizes, and colors.
 * - Built-in logging functionality that can be toggled for debugging purposes.
 * 
 * --- Important Notes ---
 * - The script currently supports up to 1000 cells for export at a time. Attempting to export more may result in performance issues.
 * - Logging can be enabled by setting `LOGGING_ENABLED` to `true` at the top of the script. This helps track the progress of the script for debugging.
 * - The script is designed for Google Sheets and Google Docs integration, so make sure you have the necessary permissions to run it.
 */

const LOGGING_ENABLED = false; // Set to true to enable logging, false to disable it

// Function to handle logging based on the LOGGING_ENABLED flag
function log(message) {
  if (LOGGING_ENABLED) {
    Logger.log(message);
  }
}

// Adding a custom menu for easy script execution
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Export to Google Doc', 'transferFormattedTextToGoogleDoc')
    .addToUi();
}

// Main function for transferring formatted text from Google Sheets to Google Docs
function transferFormattedTextToGoogleDoc() {
  var range = getSelectedRange();
  if (!range) return;

  var values = range.getRichTextValues();
  if (!validateValues(values)) return;

  var doc = DocumentApp.create("Formatted Export");
  var body = doc.getBody();

  log("Starting to transfer values to the document");
  transferValuesToDoc(values, body);

  log("Document created: " + doc.getUrl());
  SpreadsheetApp.getUi().alert("Document successfully created: " + doc.getUrl());
}

// Getting the selected range with validation
function getSelectedRange() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    SpreadsheetApp.getUi().alert("Unable to access the spreadsheet. Please try again.");
    return null;
  }
  var range = spreadsheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert("No cells selected. Please select a range for export.");
    return null;
  }
  if (range.getNumRows() * range.getNumColumns() > 1000) {
    SpreadsheetApp.getUi().alert("Selected range is too large. Please choose a smaller range to avoid performance issues.");
    return null;
  }
  log("Range successfully obtained: " + range.getA1Notation());
  return range;
}

// Validating the values in the selected range
function validateValues(values) {
  if (!values || values.length === 0) {
    SpreadsheetApp.getUi().alert("The selected range contains no data. Please select a populated range for export.");
    return false;
  }
  log("Range data successfully validated and obtained.");
  return true;
}

// Transferring values from the range to the Google Document
function transferValuesToDoc(values, body) {
  values.forEach((rowValues, rowIndex) => {
    processRow(rowValues, body, rowIndex);
  });
}

// Processing a single row of values
function processRow(rowValues, body, rowIndex) {
  var line = null;
  var hasContent = false;
  log("Processing row: " + rowIndex);

  rowValues.forEach((richText, colIndex) => {
    if (hasValidText(richText)) {
      // Append a new paragraph if this is the first cell with content in the row
      if (!line) {
        line = body.appendParagraph("");
        log("Added new paragraph for row: " + rowIndex);
      } else {
        line.appendText(" "); // Add a space between texts from neighboring cells
      }

      // Append the text from the cell and apply formatting
      var richTextContent = richText.getText();
      var text = line.appendText(richTextContent);
      applyRichTextFormatting(richText, text);
      hasContent = true;
    }
  });

  // Remove empty paragraph if no text was added, and it's not the last paragraph in the document
  if (!hasContent && line && body.getNumChildren() > 1) {
    try {
      body.removeChild(line);
      log("Removed empty paragraph for row: " + rowIndex);
    } catch (e) {
      log("Error while removing paragraph: " + e.message);
    }
  }
}

// Applying formatting from RichTextValue to the Google Document text
function applyRichTextFormatting(richText, text) {
  try {
    var length = richText.getText().length;
    if (length === 0) return; // Do nothing if the text length is zero

    var runs = richText.getRuns(); // Get all formatting runs in the text
    if (runs && runs.length > 0) {
      var currentIndex = 0;
      runs.forEach(run => {
        var runText = run.getText();
        var runLength = runText.length;
        var runEndIndex = currentIndex + runLength - 1;

        // Apply formatting to the current text segment (run)
        var textStyle = run.getTextStyle();
        if (textStyle.isBold()) {
          text.setBold(currentIndex, runEndIndex, true);
          log("Applied bold formatting to segment: " + runText);
        } else {
          text.setBold(currentIndex, runEndIndex, false);
        }
        if (textStyle.isItalic()) {
          text.setItalic(currentIndex, runEndIndex, true);
          log("Applied italic formatting to segment: " + runText);
        } else {
          text.setItalic(currentIndex, runEndIndex, false);
        }
        if (textStyle.isUnderline()) {
          text.setUnderline(currentIndex, runEndIndex, true);
          log("Applied underline formatting to segment: " + runText);
        } else {
          text.setUnderline(currentIndex, runEndIndex, false);
        }
        if (textStyle.isStrikethrough()) {
          text.setStrikethrough(currentIndex, runEndIndex, true);
          log("Applied strikethrough formatting to segment: " + runText);
        } else {
          text.setStrikethrough(currentIndex, runEndIndex, false);
        }
        if (textStyle.getForegroundColor()) {
          text.setForegroundColor(currentIndex, runEndIndex, textStyle.getForegroundColor());
          log("Applied text color - " + textStyle.getForegroundColor() + " to segment: " + runText);
        } else {
          text.setForegroundColor(currentIndex, runEndIndex, null);
        }
        if (textStyle.getBackgroundColor && textStyle.getBackgroundColor()) {
          text.setBackgroundColor(currentIndex, runEndIndex, textStyle.getBackgroundColor());
          log("Applied background color - " + textStyle.getBackgroundColor() + " to segment: " + runText);
        } else {
          text.setBackgroundColor(currentIndex, runEndIndex, null);
        }
        if (textStyle.getFontFamily()) {
          text.setFontFamily(currentIndex, runEndIndex, textStyle.getFontFamily());
          log("Applied font family - " + textStyle.getFontFamily() + " to segment: " + runText);
        } else {
          text.setFontFamily(currentIndex, runEndIndex, null);
        }
        if (textStyle.getFontSize()) {
          text.setFontSize(currentIndex, runEndIndex, textStyle.getFontSize());
          log("Applied font size - " + textStyle.getFontSize() + " to segment: " + runText);
        }

        // Move the current index forward by the length of the current run
        currentIndex += runLength;
      });
    }
  } catch (e) {
    log("Error while applying text formatting: " + e.message);
  }
}

// Checking if richText contains valid text
function hasValidText(richText) {
  return richText != null && richText.getText() !== "";
}
