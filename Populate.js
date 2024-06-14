// This is Code.gs but GitHub doesn't support .gs, so we have used .js (JavaScript file format).
// Please note that the below stocks have been randomly picked for the population of the NSE stocks
// and we do not represent or endorse any specific stock or company here.

function populateStocks() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var stocks = [
      "RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK", "KOTAKBANK", "HINDUNILVR", "ITC", 
      "HCLTECH", "LT", "MARUTI", "BHARTIARTL", "SBIN", "AXISBANK", "ASIANPAINT", "BAJFINANCE", 
      "HDFCBANK", "ADANIGREEN", "ADANIPORTS", "TITAN", "WIPRO", "DMART", "NTPC", "POWERGRID", 
      "ULTRACEMCO", "NESTLEIND", "CIPLA", "SUNPHARMA", "M&M", "BAJAJFINSV", "HEROMOTOCO", 
      "TATAMOTORS", "JSWSTEEL", "COALINDIA", "BPCL", "GAIL", "ONGC", "IOC", "GRASIM", 
      "SHREECEM", "BRITANNIA", "DIVISLAB", "TECHM", "DRREDDY", "UPL", "EICHERMOT", 
      "TATACONSUM", "ADANIENT", "INDUSINDBK", "BANDHANBNK", "DABUR", "GODREJCP", 
      "HINDALCO", "JINDALSTEL", "MOTHERSON", "PEL", "PIDILITIND", "TATACHEM", 
      "TATAPOWER", "VOLTAS", "ZEEL", "APOLLOHOSP", "LUPIN", "BIOCON", "AUROPHARMA", 
      "AMBUJACEM", "ACC", "DMART", "SBILIFE", "HDFCLIFE", "ICICIPRULI", "AXISBANK", 
      "BANDHANBNK", "FEDERALBNK", "IDFCFIRSTB", "PNB", "RBLBANK", "YESBANK", 
      "CHOLAFIN", "M&MFIN", "HDFCAMC", "ICICIGI", "MFSL", "NAM-INDIA", 
      "HONAUT", "NESTLEIND", "PGHH"
    ];
    for (var i = 0; i < stocks.length; i++) {
      sheet.getRange(i + 3, 1).setValue(stocks[i]); // Start from A3
    }
  }