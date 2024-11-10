# Google Sheets to Google Docs Text Export Script

## Overview

This script is designed to export text from a selected range of Google Sheets into a Google Docs document, while retaining all text formatting. It provides an easy-to-use custom menu that allows for a seamless transition of data, preserving complex formatting attributes such as bold, italic, underline, strikethrough, font color, and more.

## Features

- **Rich Text Formatting**: Exports text from Google Sheets to Google Docs with full formatting support, including:
  - Bold, italic, underline, strikethrough.
  - Font color, background color, font family, and font size.
  - Supports multiple formatting styles within a single cell.
- **User-Friendly Menu**: Adds a custom menu called `Custom Scripts` to Google Sheets to easily initiate the export process.
- **Configurable Logging**: Includes optional logging functionality for debugging purposes, which can be toggled on or off by setting `LOGGING_ENABLED`.

## How It Works

The script follows a simple flow to ensure that data is correctly exported:

1. **Custom Menu Creation**: After installation, a custom menu named `Custom Scripts` appears in Google Sheets.
2. **Range Selection**: Users select a range of cells that they wish to export to a new Google Docs document.
3. **Data Export**: The script transfers the selected data, preserving the formatting.

## Installation

1. **Open Google Sheet**: Open the Google Sheet where you want to use this script.
2. **Access Script Editor**: Go to `Extensions > Apps Script` to open the Google Apps Script editor.
3. **Paste the Script**: Copy and paste the provided code into the editor.
4. **Save & Refresh**: Save the script and refresh the Google Sheet.
5. **Use the Custom Menu**: After refreshing, a new menu item named `Custom Scripts` will appear in the Google Sheets menu bar.

## Usage Instructions

1. **Select the Range**: Highlight the range of cells in Google Sheets that you wish to export.
2. **Open the Custom Menu**: Click on `Custom Scripts` in the menu bar.
3. **Export the Data**: Choose `Export to Google Doc`. A new Google Docs document will be created, containing the selected range with all text formatting preserved.

## Key Configuration

- **Logging**: The script contains a `LOGGING_ENABLED` variable, which can be set to `true` to enable detailed logging of the export process. This is particularly helpful for debugging or verifying the correct operation of the script.
- **Range Size Limit**: The script currently supports exporting up to **1000 cells** at a time. Exporting larger ranges may result in performance issues or errors.

## Example Use Case

Imagine you have a table in Google Sheets that contains a mixture of formatted texts, such as bold headers, underlined rows, and cells with different background colors. You can use this script to transfer all of this content into a Google Docs document, ensuring that every style and format remains intact.

## Limitations

- **Performance**: The script may struggle with very large data sets (over 1000 cells). It is optimized for smaller ranges to avoid exceeding Google Apps Script limitations.
- **Formatting Support**: While the script supports a wide array of formatting options, some edge cases involving highly complex formatting may not be transferred perfectly.

## Troubleshooting

- **Logging for Debugging**: If something doesn’t work as expected, enable logging by setting `LOGGING_ENABLED = true` in the script. Logs can be viewed using the `Logger.log()` method in Google Apps Script.
- **Error Messages**: If you encounter an error regarding range size, try selecting a smaller range and re-running the script.

## Licensing

This project is distributed under a dual license: [MIT License](./LICENSE_MIT) and [Creative Commons Attribution 4.0 International (CC BY 4.0)](./LICENSE_CC_BY_4.0). This means that when using, copying, modifying, and distributing the project, you must comply with the terms of both licenses. In particular, you must credit authorship in accordance with the requirements of CC BY while also adhering to MIT.
