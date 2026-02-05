/**
 * Currency Dashboard - Live exchange rates with frankfurter.app API
 * Built with VS Code + clasp + GitHub (way better than the UI!)
 */

/**
 * Runs when the add-on is installed or the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Currency Dashboard')
    .addItem('Refresh Rates', 'fetchCurrencyRates')
    .addItem('Update Converter', 'updateConverter')
    .addItem('Add Historical Entry', 'addHistoricalEntry')
    .addSeparator()
    .addItem('Launch Sidebar', 'showSidebar')
    .addItem('Reset Dashboard', 'setupCurrencyDashboard')
    .addToUi();
  
  // Auto-setup dashboard on first load
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Currency Dashboard');
  
  if (!sheet) {
    setupCurrencyDashboard();
  } else {
    fetchCurrencyRates();
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Currency Dashboard ready!',
    'Welcome',
    3
  );
}

/**
 * Shows the sidebar with interactive tools
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Currency Dashboard')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Sets up the currency dashboard in the sheet
 */
function setupCurrencyDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Currency Dashboard');
  
  // Create or clear dashboard sheet
  if (!sheet) {
    sheet = ss.insertSheet('Currency Dashboard');
  } else {
    sheet.clear();
  }
  
  // HEADER SECTION
  sheet.getRange('A1:H1').merge()
    .setValue('LIVE EXCHANGE RATES')
    .setFontSize(24)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 45);
  
  // CURRENT RATES SECTION
  sheet.getRange('A2:H2').merge()
    .setValue('Major Currency Pairs vs EUR')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Currency pair headers
  const currencies = ['USD', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD'];
  sheet.getRange('A3').setValue('Currency').setFontWeight('bold');
  sheet.getRange('B3').setValue('Rate').setFontWeight('bold');
  sheet.getRange('C3').setValue('Updated').setFontWeight('bold');
  sheet.getRange('A3:C3')
    .setBackground('#E8F0FE')
    .setHorizontalAlignment('center');
  
  // Set up currency rows
  for (let i = 0; i < currencies.length; i++) {
    const row = 4 + i;
    sheet.getRange(row, 1).setValue(currencies[i]);
    sheet.getRange(row, 2).setNumberFormat('0.0000');
    
    // Alternating colors
    if (i % 2 === 0) {
      sheet.getRange(row, 1, 1, 3).setBackground('#F8F9FA');
    }
  }
  
  // CONVERTER SECTION
  sheet.getRange('E3:H3').merge()
    .setValue('Quick Converter')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#FBBC04')
    .setFontColor('#000000')
    .setHorizontalAlignment('center');
  
  sheet.getRange('E4').setValue('Amount:');
  sheet.getRange('F4').setValue(1000).setNumberFormat('#,##0');
  
  sheet.getRange('E5').setValue('From:');
  const fromCell = sheet.getRange('F5');
  fromCell.setValue('EUR');
  const allCurrencies = ['EUR', 'USD', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD'];
  const fromRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allCurrencies, true)
    .setAllowInvalid(false)
    .build();
  fromCell.setDataValidation(fromRule);
  fromCell.setBackground('#FFFEF0').setFontWeight('bold').setHorizontalAlignment('center');
  
  sheet.getRange('E6').setValue('Converts to:');
  
  const converterCurrencies = ['USD', 'GBP', 'JPY', 'CHF'];
  for (let i = 0; i < converterCurrencies.length; i++) {
    const row = 7 + i;
    sheet.getRange(row, 5).setValue(converterCurrencies[i] + ':');
    sheet.getRange(row, 6).setValue('Loading...');
    sheet.getRange(row, 6).setNumberFormat('#,##0.00');
    sheet.getRange(row, 5, 1, 2)
      .setBackground('#FFFEF0')
      .setFontWeight('bold');
  }
  
  // HISTORICAL DATA SECTION
  sheet.getRange('A12:H12').merge()
    .setValue('Rate History')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  sheet.getRange('A13:F13')
    .setValues([['Date', 'USD', 'GBP', 'JPY', 'CHF', 'CAD']])
    .setBackground('#E8F0FE')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Column widths
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 30);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  
  // Freeze headers
  sheet.setFrozenRows(13);
  
  // Borders
  sheet.getRange('A4:C10').setBorder(true, true, true, true, true, true);
  sheet.getRange('E4:F10').setBorder(true, true, true, true, true, true);
  
  // Activate the dashboard
  ss.setActiveSheet(sheet);
  
  // Fetch initial data
  fetchCurrencyRates();
  
  // Update converter
  updateConverter();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Dashboard ready! Use menu: Currency Dashboard > Refresh Rates',
    'Setup Complete',
    3
  );
}

/**
 * Updates the converter calculations based on selected currency
 */
function updateConverter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Currency Dashboard');
  
  if (!sheet) return;
  
  const amountValue = sheet.getRange('F4').getValue();
  const fromCurrency = sheet.getRange('F5').getValue();
  
  // Convert amount to number
  const amount = Number(amountValue);
  
  // Check if amount is valid
  if (amountValue === '' || amountValue === null || isNaN(amount) || amount <= 0) {
    for (let i = 0; i < 4; i++) {
      sheet.getRange(7 + i, 6).setValue('-');
    }
    return;
  }
  
  // Check if currency is valid
  if (!fromCurrency || fromCurrency === '' || fromCurrency === null) {
    for (let i = 0; i < 4; i++) {
      sheet.getRange(7 + i, 6).setValue('-');
    }
    return;
  }
  
  // Get rates from table (all are vs EUR)
  const rates = {
    'EUR': 1,
    'USD': sheet.getRange('B4').getValue() || 0,
    'GBP': sheet.getRange('B5').getValue() || 0,
    'JPY': sheet.getRange('B6').getValue() || 0,
    'CHF': sheet.getRange('B7').getValue() || 0,
    'CAD': sheet.getRange('B8').getValue() || 0,
    'AUD': sheet.getRange('B9').getValue() || 0
  };
  
  const targetCurrencies = ['USD', 'GBP', 'JPY', 'CHF'];
  const fromRate = rates[fromCurrency];
  
  if (!fromRate || fromRate === 0) {
    for (let i = 0; i < 4; i++) {
      sheet.getRange(7 + i, 6).setValue('-');
    }
    return;
  }
  
  for (let i = 0; i < targetCurrencies.length; i++) {
    const toCurrency = targetCurrencies[i];
    const toRate = rates[toCurrency];
    const row = 7 + i;
    
    if (fromCurrency === toCurrency) {
      sheet.getRange(row, 6).setValue(amount);
    } else if (toRate && toRate !== 0) {
      // Convert: amount in FROM currency → EUR → TO currency
      const converted = (amount / fromRate) * toRate;
      sheet.getRange(row, 6).setValue(converted);
    } else {
      sheet.getRange(row, 6).setValue('-');
    }
  }
}

/**
 * Triggered when user edits a cell - auto-updates converter
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // Auto-update converter when amount or currency changes
  if (sheet.getName() === 'Currency Dashboard' && 
      range.getColumn() === 6 && 
      (range.getRow() === 4 || range.getRow() === 5)) {
    updateConverter();
  }
}

/**
 * Fetches current exchange rates from frankfurter.app API (free, unlimited)
 */
function fetchCurrencyRates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Currency Dashboard');
  
  if (!sheet) {
    setupCurrencyDashboard();
    return;
  }
  
  try {
    // Frankfurter.app - Free, unlimited, no API key needed!
    const response = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=EUR');
    const data = JSON.parse(response.getContentText());
    
    const timestamp = new Date();
    const currencies = ['USD', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD'];
    
    // Update current rates
    for (let i = 0; i < currencies.length; i++) {
      const row = 4 + i;
      const rate = data.rates[currencies[i]];
      sheet.getRange(row, 2).setValue(rate);
      sheet.getRange(row, 3).setValue(timestamp).setNumberFormat('hh:mm:ss');
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Rates updated successfully!',
      'Success',
      2
    );
    
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error fetching rates: ' + error.message,
      'Error',
      3
    );
  }
}

/**
 * Adds current rates to historical data
 */
function addHistoricalEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Currency Dashboard');
  
  if (!sheet) {
    setupCurrencyDashboard();
    return;
  }
  
  const timestamp = new Date();
  
  // Get current rates
  const usd = sheet.getRange('B5').getValue();
  const gbp = sheet.getRange('B6').getValue();
  const jpy = sheet.getRange('B7').getValue();
  const chf = sheet.getRange('B8').getValue();
  const cad = sheet.getRange('B9').getValue();
  
  // Find next empty row in history
  const lastRow = sheet.getLastRow();
  const nextRow = Math.max(14, lastRow + 1);
  
  // Add entry
  sheet.getRange(nextRow, 1, 1, 6).setValues([[
    timestamp,
    usd,
    gbp,
    jpy,
    chf,
    cad
  ]]);
  
  // Format
  sheet.getRange(nextRow, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(nextRow, 2, 1, 5).setNumberFormat('0.0000');
  
  // Alternate row colors
  if (nextRow % 2 === 0) {
    sheet.getRange(nextRow, 1, 1, 6).setBackground('#F8F9FA');
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Snapshot added to history',
    'Success',
    2
  );
}

/**
 * Helper function to get data for the sidebar
 */
function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Currency Dashboard');
  
  if (!sheet) {
    return {
      sheetName: 'No Dashboard',
      usd: 0,
      gbp: 0,
      jpy: 0,
      lastUpdate: 'N/A'
    };
  }
  
  return {
    sheetName: sheet.getName(),
    usd: sheet.getRange('B5').getValue() || 0,
    gbp: sheet.getRange('B6').getValue() || 0,
    jpy: sheet.getRange('B7').getValue() || 0,
    lastUpdate: sheet.getRange('C5').getValue() || 'Never'
  };
}

/**
 * Gets live currency rates for sidebar display
 */
function getLiveCurrencyRates() {
  try {
    const response = UrlFetchApp.fetch('https://api.frankfurter.app/latest?from=EUR');
    const data = JSON.parse(response.getContentText());
    
    return {
      success: true,
      rates: {
        USD: data.rates.USD,
        GBP: data.rates.GBP,
        JPY: data.rates.JPY,
        CHF: data.rates.CHF
      },
      lastUpdate: new Date().toLocaleString()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}
