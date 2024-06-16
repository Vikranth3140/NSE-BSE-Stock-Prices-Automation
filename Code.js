// This is Code.gs but GitHub doesn't support .gs, so have used .js (JavaScript file format).

function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  clearStockParameters();
  
  // Get stock symbols from column A, starting from A3
  var tickers = sheet.getRange('A3:A').getValues();
  
  // Loop through tickers and fetch stock data using GOOGLEFINANCE
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i][0];
    if (ticker) {
      var fullTicker = 'NSE:' + ticker;
      var priceFormula = '=GOOGLEFINANCE("' + fullTicker + '","price")';
      var changePctFormula = '=GOOGLEFINANCE("' + fullTicker + '","changepct")';
      var volumeFormula = '=GOOGLEFINANCE("' + fullTicker + '","volume")';
      var highFormula = '=GOOGLEFINANCE("' + fullTicker + '","high")';
      var lowFormula = '=GOOGLEFINANCE("' + fullTicker + '","low")';
      var openFormula = '=GOOGLEFINANCE("' + fullTicker + '","priceopen")';
      var marketCapFormula = '=GOOGLEFINANCE("' + fullTicker + '","marketcap")';
      var volumeAvgFormula = '=GOOGLEFINANCE("' + fullTicker + '","volumeavg")';
      var peFormula = '=GOOGLEFINANCE("' + fullTicker + '","pe")';
      var epsFormula = '=GOOGLEFINANCE("' + fullTicker + '","eps")';
      var high52Formula = '=GOOGLEFINANCE("' + fullTicker + '","high52")';
      var low52Formula = '=GOOGLEFINANCE("' + fullTicker + '","low52")';
      var closeYestFormula = '=GOOGLEFINANCE("' + fullTicker + '","closeyest")';
      var sharesFormula = '=GOOGLEFINANCE("' + fullTicker + '","shares")';
      var tradeTimeFormula = '=GOOGLEFINANCE("' + fullTicker + '","tradetime")';
      var dataDelayFormula = '=GOOGLEFINANCE("' + fullTicker + '","datadelay")';
      
      sheet.getRange(i + 3, 5).setFormula(priceFormula); // Column E for Price
      sheet.getRange(i + 3, 6).setFormula(changePctFormula); // Column F for Change Percentage
      sheet.getRange(i + 3, 7).setFormula(volumeFormula); // Column G for Volume
      sheet.getRange(i + 3, 8).setFormula(highFormula); // Column H for High
      sheet.getRange(i + 3, 9).setFormula(lowFormula); // Column I for Low
      sheet.getRange(i + 3, 10).setFormula(openFormula); // Column J for Open
      sheet.getRange(i + 3, 11).setFormula(marketCapFormula); // Column K for Market Cap
      sheet.getRange(i + 3, 12).setFormula(volumeAvgFormula); // Column L for Volume Avg
      sheet.getRange(i + 3, 13).setFormula(peFormula); // Column M for PE Ratio
      sheet.getRange(i + 3, 14).setFormula(epsFormula); // Column N for EPS
      sheet.getRange(i + 3, 15).setFormula(high52Formula); // Column O for 52-Week High
      sheet.getRange(i + 3, 16).setFormula(low52Formula); // Column P for 52-Week Low
      sheet.getRange(i + 3, 17).setFormula(closeYestFormula); // Column Q for Close (Previous Day)
      sheet.getRange(i + 3, 18).setFormula(sharesFormula); // Column R for Shares
      sheet.getRange(i + 3, 19).setFormula(tradeTimeFormula); // Column S for Trade Time
      sheet.getRange(i + 3, 20).setFormula(dataDelayFormula); // Column T for Data Delay
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('E3:E').getValues();
  var changesPct = sheet.getRange('F3:F').getValues();
  var volumes = sheet.getRange('G3:G').getValues();
  var highs = sheet.getRange('H3:H').getValues();
  var lows = sheet.getRange('I3:I').getValues();
  var opens = sheet.getRange('J3:J').getValues();
  var marketCaps = sheet.getRange('K3:K').getValues();
  var volumeAvgs = sheet.getRange('L3:L').getValues();
  var peRatios = sheet.getRange('M3:M').getValues();
  var eps = sheet.getRange('N3:N').getValues();
  var high52s = sheet.getRange('O3:O').getValues();
  var low52s = sheet.getRange('P3:P').getValues();
  var closeYests = sheet.getRange('Q3:Q').getValues();
  var shares = sheet.getRange('R3:R').getValues();
  var tradeTimes = sheet.getRange('S3:S').getValues();
  var dataDelays = sheet.getRange('T3:T').getValues();
  
  sheet.getRange('E3:E').setValues(prices);
  sheet.getRange('F3:F').setValues(changesPct);
  sheet.getRange('G3:G').setValues(volumes);
  sheet.getRange('H3:H').setValues(highs);
  sheet.getRange('I3:I').setValues(lows);
  sheet.getRange('J3:J').setValues(opens);
  sheet.getRange('K3:K').setValues(marketCaps);
  sheet.getRange('L3:L').setValues(volumeAvgs);
  sheet.getRange('M3:M').setValues(peRatios);
  sheet.getRange('N3:N').setValues(eps);
  sheet.getRange('O3:O').setValues(high52s);
  sheet.getRange('P3:P').setValues(low52s);
  sheet.getRange('Q3:Q').setValues(closeYests);
  sheet.getRange('R3:R').setValues(shares);
  sheet.getRange('S3:S').setValues(tradeTimes);
  sheet.getRange('T3:T').setValues(dataDelays);
}

function clearStockParameters() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('E2:T').clearContent();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addItem('Clear Stock Details', 'clearStockParameters')
      .addToUi();
}