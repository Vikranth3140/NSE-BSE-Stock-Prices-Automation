// This is Code.gs but GitHub doesn't support .gs, so have used .js (JavaScript file format).

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
  
  // Fetch historical data and create charts
  fetchHistoricalDataAndCreateCharts();
  
  // Wait for formulas to calculate
  SpreadsheetApp.flush();
}

function fetchHistoricalDataAndCreateCharts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var tickers = sheet.getRange('A3:A').getValues();
  var today = new Date();
  var threeMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 3, today.getDate());
  
  for (var i = 0; i < tickers.length; i++) {
    var ticker = tickers[i][0];
    if (ticker) {
      var fullTicker = 'NSE:' + ticker;
      var historicalDataFormula = '=GOOGLEFINANCE("' + fullTicker + '", "close", DATE(' + threeMonthsAgo.getFullYear() + ',' + (threeMonthsAgo.getMonth() + 1) + ',' + threeMonthsAgo.getDate() + '), TODAY(), "DAILY")';
      var historicalDataRange = 'S' + (i * 100 + 3) + ':T' + (i * 100 + 100);
      
      // Set the formula to fetch historical data
      sheet.getRange(historicalDataRange.split(':')[0]).setFormula(historicalDataFormula);
      
      // Create a new chart
      var chart = sheet.newChart()
          .setChartType(Charts.ChartType.LINE)
          .addRange(sheet.getRange(historicalDataRange))
          .setPosition(i + 2, 18, 0, 0) // Adjust the chart position
          .setOption('title', ticker + ' - Last 3 Months Performance')
          .setOption('hAxis.title', 'Date')
          .setOption('vAxis.title', 'Price')
          .build();
      
      // Insert the chart into the sheet
      sheet.insertChart(chart);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Stock Prices')
      .addItem('Update Prices', 'getStockPrices')
      .addToUi();
}