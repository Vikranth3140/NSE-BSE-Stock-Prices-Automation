// This is Code.gs but GitHub doesn't support .gs, so have used .js (JavaScript file format)

function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  sheet.getRange('B2:Q').clearContent();
  
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
      
      sheet.getRange(i + 3, 2).setFormula(priceFormula); // Column B for Price
      sheet.getRange(i + 3, 3).setFormula(changePctFormula); // Column C for Change Percentage
      sheet.getRange(i + 3, 4).setFormula(volumeFormula); // Column D for Volume
      sheet.getRange(i + 3, 5).setFormula(highFormula); // Column E for High
      sheet.getRange(i + 3, 6).setFormula(lowFormula); // Column F for Low
      sheet.getRange(i + 3, 7).setFormula(openFormula); // Column G for Open
      sheet.getRange(i + 3, 8).setFormula(marketCapFormula); // Column H for Market Cap
      sheet.getRange(i + 3, 9).setFormula(volumeAvgFormula); // Column I for Volume Avg
      sheet.getRange(i + 3, 10).setFormula(peFormula); // Column J for PE Ratio
      sheet.getRange(i + 3, 11).setFormula(epsFormula); // Column K for EPS
      sheet.getRange(i + 3, 12).setFormula(high52Formula); // Column L for 52-Week High
      sheet.getRange(i + 3, 13).setFormula(low52Formula); // Column M for 52-Week Low
      sheet.getRange(i + 3, 14).setFormula(closeYestFormula); // Column N for Close (Previous Day)
      sheet.getRange(i + 3, 15).setFormula(sharesFormula); // Column O for Shares
      sheet.getRange(i + 3, 16).setFormula(tradeTimeFormula); // Column P for Trade Time
      sheet.getRange(i + 3, 17).setFormula(dataDelayFormula); // Column Q for Data Delay
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('B3:B').getValues();
  var changesPct = sheet.getRange('C3:C').getValues();
  var volumes = sheet.getRange('D3:D').getValues();
  var highs = sheet.getRange('E3:E').getValues();
  var lows = sheet.getRange('F3:F').getValues();
  var opens = sheet.getRange('G3:G').getValues();
  var marketCaps = sheet.getRange('H3:H').getValues();
  var volumeAvgs = sheet.getRange('I3:I').getValues();
  var peRatios = sheet.getRange('J3:J').getValues();
  var eps = sheet.getRange('K3:K').getValues();
  var high52s = sheet.getRange('L3:L').getValues();
  var low52s = sheet.getRange('M3:M').getValues();
  var closeYests = sheet.getRange('N3:N').getValues();
  var shares = sheet.getRange('O3:O').getValues();
  var tradeTimes = sheet.getRange('P3:P').getValues();
  var dataDelays = sheet.getRange('Q3:Q').getValues();
  
  sheet.getRange('B3:B').setValues(prices);
  sheet.getRange('C3:C').setValues(changesPct);
  sheet.getRange('D3:D').setValues(volumes);
  sheet.getRange('E3:E').setValues(highs);
  sheet.getRange('F3:F').setValues(lows);
  sheet.getRange('G3:G').setValues(opens);
  sheet.getRange('H3:H').setValues(marketCaps);
  sheet.getRange('I3:I').setValues(volumeAvgs);
  sheet.getRange('J3:J').setValues(peRatios);
  sheet.getRange('K3:K').setValues(eps);
  sheet.getRange('L3:L').setValues(high52s);
  sheet.getRange('M3:M').setValues(low52s);
  sheet.getRange('N3:N').setValues(closeYests);
  sheet.getRange('O3:O').setValues(shares);
  sheet.getRange('P3:P').setValues(tradeTimes);
  sheet.getRange('Q3:Q').setValues(dataDelays);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addToUi();
}