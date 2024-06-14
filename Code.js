function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  sheet.getRange('B2:G').clearContent();
  
  // Get stock symbols from column A
  var tickers = sheet.getRange('A2:A').getValues();
  
  // Loop through tickers and fetch stock data using GOOGLEFINANCE
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i][0];
    if (ticker) {
      var fullTicker = 'NSE:' + ticker;
      var priceFormula = '=GOOGLEFINANCE("' + fullTicker + '","price")';
      var changeFormula = '=GOOGLEFINANCE("' + fullTicker + '","changepct")';
      var volumeFormula = '=GOOGLEFINANCE("' + fullTicker + '","volume")';
      var highFormula = '=GOOGLEFINANCE("' + fullTicker + '","high")';
      var lowFormula = '=GOOGLEFINANCE("' + fullTicker + '","low")';
      var openFormula = '=GOOGLEFINANCE("' + fullTicker + '","open")';
      
      sheet.getRange(i + 2, 2).setFormula(priceFormula); // Column B for Price
      sheet.getRange(i + 2, 3).setFormula(changeFormula); // Column C for Change
      sheet.getRange(i + 2, 4).setFormula(volumeFormula); // Column D for Volume
      sheet.getRange(i + 2, 5).setFormula(highFormula); // Column E for High
      sheet.getRange(i + 2, 6).setFormula(lowFormula); // Column F for Low
      sheet.getRange(i + 2, 7).setFormula(openFormula); // Column G for Open
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('B2:B').getValues();
  var changes = sheet.getRange('C2:C').getValues();
  var volumes = sheet.getRange('D2:D').getValues();
  var highs = sheet.getRange('E2:E').getValues();
  var lows = sheet.getRange('F2:F').getValues();
  var opens = sheet.getRange('G2:G').getValues();
  
  sheet.getRange('B2:B').setValues(prices);
  sheet.getRange('C2:C').setValues(changes);
  sheet.getRange('D2:D').setValues(volumes);
  sheet.getRange('E2:E').setValues(highs);
  sheet.getRange('F2:F').setValues(lows);
  sheet.getRange('G2:G').setValues(opens);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addToUi();
}