// This is Code.gs but GitHub doesn't support .gs, so have used .js (JavaScript file format).

function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  clearStockParameters();
  
  // Get exchange and stock symbols from columns A and B, starting from row 3
  var exchanges = sheet.getRange('A3:A').getValues();
  var tickers = sheet.getRange('B3:B').getValues();
  
  // Loop through tickers and fetch stock data using GOOGLEFINANCE
  for (var i = 0; i < tickers.length; i++) {
    var exchange = exchanges[i][0];
    var ticker = tickers[i][0];
    if (exchange && ticker) {
      var fullTicker = exchange + ':' + ticker;

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
      
      sheet.getRange(i + 3, 6).setFormula(priceFormula); // Column F for Price
      sheet.getRange(i + 3, 7).setFormula(changePctFormula); // Column G for Change Percentage
      sheet.getRange(i + 3, 8).setFormula(volumeFormula); // Column H for Volume
      sheet.getRange(i + 3, 9).setFormula(highFormula); // Column I for High
      sheet.getRange(i + 3, 10).setFormula(lowFormula); // Column J for Low
      sheet.getRange(i + 3, 11).setFormula(openFormula); // Column K for Open
      sheet.getRange(i + 3, 12).setFormula(marketCapFormula); // Column L for Market Cap
      sheet.getRange(i + 3, 13).setFormula(volumeAvgFormula); // Column M for Volume Avg
      sheet.getRange(i + 3, 14).setFormula(peFormula); // Column N for PE Ratio
      sheet.getRange(i + 3, 15).setFormula(epsFormula); // Column O for EPS
      sheet.getRange(i + 3, 16).setFormula(high52Formula); // Column P for 52-Week High
      sheet.getRange(i + 3, 17).setFormula(low52Formula); // Column Q for 52-Week Low
      sheet.getRange(i + 3, 18).setFormula(closeYestFormula); // Column R for Close (Previous Day)
      sheet.getRange(i + 3, 19).setFormula(sharesFormula); // Column S for Shares
      sheet.getRange(i + 3, 20).setFormula(tradeTimeFormula); // Column T for Trade Time
      sheet.getRange(i + 3, 21).setFormula(dataDelayFormula); // Column U for Data Delay
      
      // Calculate Total Amount Invested (E column) = C (Quantity) * D (Purchase Price)
      var totalInvestedFormula = '=C' + (i + 3) + '*D' + (i + 3);
      sheet.getRange(i + 3, 5).setFormula(totalInvestedFormula); // Column E for Total Amount Invested
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('F3:F').getValues();
  var changesPct = sheet.getRange('G3:G').getValues();
  var volumes = sheet.getRange('H3:H').getValues();
  var highs = sheet.getRange('I3:I').getValues();
  var lows = sheet.getRange('J3:J').getValues();
  var opens = sheet.getRange('K3:K').getValues();
  var marketCaps = sheet.getRange('L3:L').getValues();
  var volumeAvgs = sheet.getRange('M3:M').getValues();
  var peRatios = sheet.getRange('N3:N').getValues();
  var eps = sheet.getRange('O3:O').getValues();
  var high52s = sheet.getRange('P3:P').getValues();
  var low52s = sheet.getRange('Q3:Q').getValues();
  var closeYests = sheet.getRange('R3:R').getValues();
  var shares = sheet.getRange('S3:S').getValues();
  var tradeTimes = sheet.getRange('T3:T').getValues();
  var dataDelays = sheet.getRange('U3:U').getValues();
  
  sheet.getRange('F3:F').setValues(prices);
  sheet.getRange('G3:G').setValues(changesPct);
  sheet.getRange('H3:H').setValues(volumes);
  sheet.getRange('I3:I').setValues(highs);
  sheet.getRange('J3:J').setValues(lows);
  sheet.getRange('K3:K').setValues(opens);
  sheet.getRange('L3:L').setValues(marketCaps);
  sheet.getRange('M3:M').setValues(volumeAvgs);
  sheet.getRange('N3:N').setValues(peRatios);
  sheet.getRange('O3:O').setValues(eps);
  sheet.getRange('P3:P').setValues(high52s);
  sheet.getRange('Q3:Q').setValues(low52s);
  sheet.getRange('R3:R').setValues(closeYests);
  sheet.getRange('S3:S').setValues(shares);
  sheet.getRange('T3:T').setValues(tradeTimes);
  sheet.getRange('U3:U').setValues(dataDelays);
}

function clearStockParameters() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('E3:V').clearContent();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addItem('Clear Stock Details', 'clearStockParameters')
      .addToUi();
}