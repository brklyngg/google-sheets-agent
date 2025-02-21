// Google Sheets AI Agent
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Agent')
    .addItem('Analyze', 'analyzeWorkbook')
    .addToUi();
}