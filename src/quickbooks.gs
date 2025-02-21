// QuickBooks integration functions

function setupQuickBooks() {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    const authorizationUrl = service.getAuthorizationUrl();
    const template = HtmlService.createTemplate(
      '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>'
    );
    template.authorizationUrl = authorizationUrl;
    const page = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(page, 'Authorize QuickBooks');
  }
}

function fetchProfitAndLoss(startDate, endDate) {
  const token = getOAuthService().getAccessToken();
  const url = `${CONFIG.QUICKBOOKS_API_BASE}/reports/ProfitAndLoss?start_date=${startDate}&end_date=${endDate}`;
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  });
  
  return JSON.parse(response.getContentText());
}

function handleQuickBooksCallback(request) {
  const service = getOAuthService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Failed to authorize.');
  }
}