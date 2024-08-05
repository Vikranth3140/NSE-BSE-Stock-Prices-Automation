function getStockPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  clearStockParameters();
  
  // Get exchange and stock symbols from columns A and B, starting from row 3
  var exchanges = sheet.getRange('A4:A').getValues(); // Adjusted to start from row 4
  var tickers = sheet.getRange('B4:B').getValues(); // Adjusted to start from row 4
  
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
      
      sheet.getRange(i + 4, 7).setFormula(priceFormula); // Column G for Price (Adjusted to start from row 4)
      sheet.getRange(i + 4, 8).setFormula(changePctFormula); // Column H for Change Percentage
      sheet.getRange(i + 4, 9).setFormula(volumeFormula); // Column I for Volume
      sheet.getRange(i + 4, 10).setFormula(highFormula); // Column J for High
      sheet.getRange(i + 4, 11).setFormula(lowFormula); // Column K for Low
      sheet.getRange(i + 4, 12).setFormula(openFormula); // Column L for Open
      sheet.getRange(i + 4, 13).setFormula(marketCapFormula); // Column M for Market Cap
      sheet.getRange(i + 4, 14).setFormula(volumeAvgFormula); // Column N for Volume Avg
      sheet.getRange(i + 4, 15).setFormula(peFormula); // Column O for PE Ratio
      sheet.getRange(i + 4, 16).setFormula(epsFormula); // Column P for EPS
      sheet.getRange(i + 4, 17).setFormula(high52Formula); // Column Q for 52-Week High
      sheet.getRange(i + 4, 18).setFormula(low52Formula); // Column R for 52-Week Low
      sheet.getRange(i + 4, 19).setFormula(closeYestFormula); // Column S for Close (Previous Day)
      sheet.getRange(i + 4, 20).setFormula(sharesFormula); // Column T for Shares
      sheet.getRange(i + 4, 21).setFormula(tradeTimeFormula); // Column U for Trade Time
      
      // Calculate Total Amount Invested (E column) = C (Quantity) * D (Purchase Price)
      var totalInvestedFormula = '=C' + (i + 4) + '*D' + (i + 4);
      sheet.getRange(i + 4, 5).setFormula(totalInvestedFormula); // Column E for Total Amount Invested
    }
  }
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
  
  // Optional: Convert formulas to values if you do not want to keep the GOOGLEFINANCE formula
  var prices = sheet.getRange('G4:G').getValues();
  var changesPct = sheet.getRange('H4:H').getValues();
  var volumes = sheet.getRange('I4:I').getValues();
  var highs = sheet.getRange('J4:J').getValues();
  var lows = sheet.getRange('K4:K').getValues();
  var opens = sheet.getRange('L4:L').getValues();
  var marketCaps = sheet.getRange('M4:M').getValues();
  var volumeAvgs = sheet.getRange('N4:N').getValues();
  var peRatios = sheet.getRange('O4:O').getValues();
  var eps = sheet.getRange('P4:P').getValues();
  var high52s = sheet.getRange('Q4:Q').getValues();
  var low52s = sheet.getRange('R4:R').getValues();
  var closeYests = sheet.getRange('S4:S').getValues();
  var shares = sheet.getRange('T4:T').getValues();
  var tradeTimes = sheet.getRange('U4:U').getValues();
  
  sheet.getRange('G4:G').setValues(prices);
  sheet.getRange('H4:H').setValues(changesPct);
  sheet.getRange('I4:I').setValues(volumes);
  sheet.getRange('J4:J').setValues(highs);
  sheet.getRange('K4:K').setValues(lows);
  sheet.getRange('L4:L').setValues(opens);
  sheet.getRange('M4:M').setValues(marketCaps);
  sheet.getRange('N4:N').setValues(volumeAvgs);
  sheet.getRange('O4:O').setValues(peRatios);
  sheet.getRange('P4:P').setValues(eps);
  sheet.getRange('Q4:Q').setValues(high52s);
  sheet.getRange('R4:R').setValues(low52s);
  sheet.getRange('S4:S').setValues(closeYests);
  sheet.getRange('T4:T').setValues(shares);
  sheet.getRange('U4:U').setValues(tradeTimes);
}

function clearStockParameters() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('E4:U').clearContent(); // Adjusted to clear from row 4
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addItem('Clear Stock Details', 'clearStockParameters')
      .addToUi();
}