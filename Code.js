function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  sheet.getRange('B2:B').clearContent();
  
  // Get stock symbols from column A
  var tickers = sheet.getRange('A2:A').getValues();
  
  // Loop through tickers and fetch stock prices using GOOGLEFINANCE
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i][0];
    if (ticker) {
      var formula = '=GOOGLEFINANCE("' + ticker + '","price")';
      sheet.getRange(i + 2, 2).setFormula(formula);
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('B2:B').getValues();
  sheet.getRange('B2:B').setValues(prices);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addToUi();
}