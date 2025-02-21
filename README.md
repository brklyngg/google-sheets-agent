# Google Sheets AI Agent

An AI-powered agent that analyzes Google Sheets workbooks and integrates with QuickBooks Online.

## Features

- Workbook analysis and data type detection
- Formula dependency tracking
- Interactive chat interface for queries
- QuickBooks Online integration
- Automated report generation

## Installation

1. Open your Google Sheets workbook
2. Go to Extensions > Apps Script
3. Copy the contents of `Code.gs` into the script editor
4. Save and refresh your spreadsheet

## QuickBooks Setup

1. Create a QuickBooks Online developer account
2. Register your application to get API credentials
3. Use the 'Connect QuickBooks' menu item to authorize

## Usage

1. Access the AI Agent menu in your spreadsheet
2. Choose 'Analyze Workbook' for a full analysis
3. Use 'Chat Interface' to ask questions
4. Use 'Connect QuickBooks' to set up data integration

## Query Examples

- "Summarize all sheets in this workbook"
- "What cells affect the total in B15?"
- "Pull QuickBooks P&L for Jan-Dec 2024"

## Security Note

This script requires authorization to:
- Read/write spreadsheet data
- Access QuickBooks data
- Display dialogs and sidebars

## License

MIT