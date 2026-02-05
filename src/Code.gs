/**
 * Data Analyzer Pro - A powerful Google Sheets add-on
 * Built with VS Code + clasp + GitHub (way better than the UI!)
 */

/**
 * Runs when the add-on is installed or the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Analyzer Pro')
    .addItem('Launch Dashboard', 'showSidebar')
    .addSeparator()
    .addItem('Analyze Selection', 'analyzeSelection')
    .addItem('Auto-Format Data', 'autoFormatData')
    .addItem('Create Summary Report', 'createSummaryReport')
    .addSeparator()
    .addItem('Fetch Currency Rates', 'fetchCurrencyRates')
    .addItem('Clean Data', 'cleanData')
    .addItem('Remove Duplicates', 'removeDuplicates')
    .addToUi();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Data Analyzer Pro loaded! Check the menu above.',
    'Welcome',
    5
  );
}

/**
 * Shows the sidebar with interactive tools
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Data Analyzer Pro')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Analyzes selected data and shows statistics
 */
function analyzeSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  
  // Flatten and filter numeric values
  const numbers = values.flat().filter(val => typeof val === 'number' && !isNaN(val));
  
  if (numbers.length === 0) {
    SpreadsheetApp.getUi().alert('No numeric data found in selection!');
    return;
  }
  
  // Calculate statistics
  const sum = numbers.reduce((a, b) => a + b, 0);
  const avg = sum / numbers.length;
  const min = Math.min(...numbers);
  const max = Math.max(...numbers);
  const sorted = numbers.sort((a, b) => a - b);
  const median = sorted.length % 2 === 0 
    ? (sorted[sorted.length/2 - 1] + sorted[sorted.length/2]) / 2 
    : sorted[Math.floor(sorted.length/2)];
  
  // Show results
  const message = `Data Analysis Results\n\n` +
    `Count: ${numbers.length}\n` +
    `Sum: ${sum.toFixed(2)}\n` +
    `Average: ${avg.toFixed(2)}\n` +
    `Median: ${median.toFixed(2)}\n` +
    `Min: ${min.toFixed(2)}\n` +
    `Max: ${max.toFixed(2)}\n` +
    `Range: ${(max - min).toFixed(2)}`;
  
  SpreadsheetApp.getUi().alert('Analysis Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Auto-formats the selected data range with professional styling
 */
function autoFormatData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  // Header row styling
  const headerRow = range.offset(0, 0, 1, range.getNumColumns());
  headerRow.setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Data rows - alternating colors
  if (range.getNumRows() > 1) {
    for (let i = 1; i < range.getNumRows(); i++) {
      const row = range.offset(i, 0, 1, range.getNumColumns());
      row.setBackground(i % 2 === 0 ? '#F8F9FA' : '#FFFFFF');
    }
  }
  
  // Add borders
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // Auto-resize columns
  for (let i = 1; i <= range.getNumColumns(); i++) {
    sheet.autoResizeColumn(range.getColumn() + i - 1);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Data formatted successfully!', 'Success', 3);
}

/**
 * Creates a summary report in a new sheet
 */
function createSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const data = sourceSheet.getDataRange().getValues();
  
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('Need at least a header row and one data row!');
    return;
  }
  
  // Create new sheet
  const reportSheet = ss.insertSheet('Summary Report - ' + new Date().toLocaleDateString());
  
  // Add title
  reportSheet.getRange('A1').setValue('SUMMARY REPORT')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');
  reportSheet.getRange('A1:D1').merge();
  
  // Add metadata
  reportSheet.getRange('A3').setValue('Source Sheet:').setFontWeight('bold');
  reportSheet.getRange('B3').setValue(sourceSheet.getName());
  reportSheet.getRange('A4').setValue('Generated:').setFontWeight('bold');
  reportSheet.getRange('B4').setValue(new Date());
  reportSheet.getRange('A5').setValue('Total Rows:').setFontWeight('bold');
  reportSheet.getRange('B5').setValue(data.length - 1);
  
  // Column analysis
  reportSheet.getRange('A7').setValue('Column Statistics')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');
  reportSheet.getRange('A7:D7').merge();
  
  reportSheet.getRange('A8:D8')
    .setValues([['Column', 'Type', 'Unique Values', 'Empty Cells']])
    .setBackground('#E8F0FE')
    .setFontWeight('bold');
  
  const headers = data[0];
  let row = 9;
  
  for (let col = 0; col < headers.length; col++) {
    const columnData = data.slice(1).map(r => r[col]);
    const uniqueValues = new Set(columnData.filter(v => v !== '')).size;
    const emptyCells = columnData.filter(v => v === '').length;
    const type = typeof columnData.find(v => v !== '');
    
    reportSheet.getRange(row, 1, 1, 4).setValues([[
      headers[col],
      type,
      uniqueValues,
      emptyCells
    ]]);
    row++;
  }
  
  // Auto-resize
  reportSheet.autoResizeColumns(1, 4);
  
  // Activate the new sheet
  ss.setActiveSheet(reportSheet);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Summary report created!',
    'Success',
    3
  );
}

/**
 * Cleans data by trimming whitespace and removing empty rows
 */
function cleanData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  
  let cleanedRows = 0;
  let trimmedCells = 0;
  
  // Clean the data
  const cleanedData = values.map(row => {
    const isEmptyRow = row.every(cell => cell === '' || cell === null);
    if (isEmptyRow && row.length > 0) {
      cleanedRows++;
      return null; // Mark for removal
    }
    
    return row.map(cell => {
      if (typeof cell === 'string') {
        const trimmed = cell.trim();
        if (trimmed !== cell) trimmedCells++;
        return trimmed;
      }
      return cell;
    });
  }).filter(row => row !== null);
  
  // Update the range
  if (cleanedData.length > 0) {
    range.clearContent();
    sheet.getRange(range.getRow(), range.getColumn(), cleanedData.length, cleanedData[0].length)
      .setValues(cleanedData);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Cleaned ${trimmedCells} cells and removed ${cleanedRows} empty rows!`,
    'Data Cleaned',
    4
  );
}

/**
 * Removes duplicate rows from the selection
 */
function removeDuplicates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  
  // Track unique rows
  const seen = new Set();
  const uniqueRows = [];
  let duplicateCount = 0;
  
  values.forEach(row => {
    const key = JSON.stringify(row);
    if (!seen.has(key)) {
      seen.add(key);
      uniqueRows.push(row);
    } else {
      duplicateCount++;
    }
  });
  
  if (duplicateCount === 0) {
    SpreadsheetApp.getUi().alert('No duplicates found!');
    return;
  }
  
  // Update the range
  range.clearContent();
  sheet.getRange(range.getRow(), range.getColumn(), uniqueRows.length, uniqueRows[0].length)
    .setValues(uniqueRows);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Removed ${duplicateCount} duplicate rows!`,
    'Success',
    4
  );
}

/**
 * Helper function to get data for the sidebar
 */
function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getDataRange();
  
  return {
    sheetName: sheet.getName(),
    totalSheets: ss.getSheets().length,
    rows: range.getNumRows(),
    columns: range.getNumColumns(),
    lastModified: ss.getLastUpdated()
  };
}

/**
 * Fetches current currency exchange rates - Perfect for poker earnings tracking!
 * Uses exchangerate-api.com free tier
 */
function fetchCurrencyRates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  try {
    // Fetch EUR, GBP, USD rates (common for poker tournaments)
    const response = UrlFetchApp.fetch('https://api.exchangerate-api.com/v4/latest/USD');
    const data = JSON.parse(response.getContentText());
    
    // Create header if needed
    sheet.getRange('A1:D1').setValues([['Currency', 'Rate to USD', 'Last Updated', 'Provider']]);
    sheet.getRange('A1:D1').setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
    
    // Insert data
    const currencies = ['EUR', 'GBP', 'CAD', 'AUD', 'JPY', 'CHF', 'MXN', 'BRL'];
    const timestamp = new Date(data.time_last_updated * 1000);
    
    const rows = currencies.map(curr => [
      curr,
      data.rates[curr],
      timestamp,
      'exchangerate-api.com'
    ]);
    
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    sheet.autoResizeColumns(1, 4);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Fetched ${currencies.length} currency rates successfully!`,
      'API Success',
      4
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('API Error: ' + error.message);
  }
}

/**
 * Gets live currency rates for sidebar display
 */
function getLiveCurrencyRates() {
  try {
    const response = UrlFetchApp.fetch('https://api.exchangerate-api.com/v4/latest/USD');
    const data = JSON.parse(response.getContentText());
    
    return {
      success: true,
      rates: {
        EUR: data.rates.EUR,
        GBP: data.rates.GBP,
        CAD: data.rates.CAD
      },
      lastUpdate: new Date(data.time_last_updated * 1000).toLocaleString()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}
