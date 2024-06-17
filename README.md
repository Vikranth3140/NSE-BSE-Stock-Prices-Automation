# NSE and BSE Stock Prices Automation

This Google Sheets script fetches real-time stock prices, changes percentage, volume, high, low, open prices, market capitalization, average daily trading volume, P/E ratio, earnings per share, 52-week high, 52-week low, previous day's closing prices, number of outstanding shares, trade time, and data delay for NSE and BSE stocks using the `GOOGLEFINANCE` function. Simply enter the stock symbols with the appropriate exchange prefix (`NSE:` or `BSE:`), and the script will handle the rest.

Read our Dev Blog [here](https://dev.to/vikranth3140/automate-nse-stock-prices-in-google-sheets-with-ease-3mop)

![Working Example](assets/image.png)

## Features

- Fetches real-time stock prices for NSE and BSE listed stocks.
- Retrieves additional stock data such as:
  - Price
  - Percentage Change
  - Volume
  - High
  - Low
  - Open
  - Market Capitalization
  - Average Daily Trading Volume
  - P/E Ratio
  - Earnings Per Share
  - 52-Week High
  - 52-Week Low
  - Previous Day's Closing Price
  - Number of Outstanding Shares
  - Trade Time
  - Data Delay
- Automatically appends the appropriate exchange prefix to stock symbols.
- Calculate the total amount invested in a particular company by specifying the number of stocks held and purchase price.

## Video Tutorial

For a detailed video tutorial on how to use the script, [click here](https://drive.google.com/file/d/1IUSCFHQpC6hRwfGvGHgsxXfry0T5IXRh/view?usp=sharing).

## Setup

1. Open your Google Sheets document.
2. Go to `Extensions` > `Apps Script`.
3. In the Apps Script editor, delete any existing code and replace it with the code from the `Code.js` file in this repository. *(Note: This file is named `Code.js` because GitHub does not support the `.gs` file extension used by Google Apps Script. However, you can copy the contents directly into your Apps Script project.)*
    - Copy the entire content of the `Code.js` file.
    - Paste the copied content into the Apps Script editor.
    - **Note**: If there are any specific lines you do not want to include, you can comment them out by adding `//` at the beginning of the line.
4. Save the script by clicking the disk icon or pressing `Ctrl + S`.
5. Close the Apps Script editor and refresh your Google Sheets document.
6. Enter the exchange (`NSE:` or `BSE:`) in column A and stock symbols starting from cell B3.
7. Enter the number of stocks held in column C and the purchase price in column D to calculate the total amount invested in column E.
8. Click on `Stock Prices` > `Update Prices` to fetch and display stock data.
9. Click on `Stock Prices` > `Clear Stock Details` to clear all the stock details.

## View and Copy

If you prefer to use a pre-configured Google Sheet with the App Script already set up, you can view and make a copy of [this Google Sheet](https://docs.google.com/spreadsheets/d/1lFifrj-Tz-uy5HfSLb8w6gkfa67wKtm-XMkU9gA29qk/edit?usp=sharing). The App Script is included, so you don't need to worry about setting up the code.

## Optional: Populate Stock Symbols Automatically

Instead of manually entering stock symbols, you can use the `Populate.js` script to add a predefined list of stocks.

### Steps to Use Populate.js

1. Open your Google Sheets document.
2. Go to `Extensions` > `Apps Script`.
3. In the Apps Script editor, delete any existing code and replace it with the code from the `Populate.js` file in this repository. *(Note: This file is named `Populate.js` because GitHub does not support the `.gs` file extension used by Google Apps Script. However, you can copy the contents directly into your Apps Script project.)*
    - Copy the entire content of the `Populate.js` file.
    - Paste the copied content into the Apps Script editor.
4. Save the script by clicking the disk icon or pressing `Ctrl + S`.
5. Close the Apps Script editor and refresh your Google Sheets document.
6. In Google Sheets, go to `Extensions` > `Apps Script` > `populateStocks` and run the function.
7. The predefined list of NSE stocks will be automatically populated starting from cell A3.

## Setting Up Automatic Updates

To set up the script to automatically update stock prices every minute:

1. Open your Google Sheets document.
2. Go to `Extensions` > `Apps Script`.
3. Ensure you have the following functions in your script:

    ```javascript
    // Set a time-driven trigger to update stock prices every minute
    function createTimeDrivenTrigger() {
      ScriptApp.newTrigger('getStockPrices')
        .timeBased()
        .everyMinutes(1)
        .create();
    }

    // Delete all existing time-driven triggers
    function deleteTriggers() {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    ```

4. Save the script by clicking the disk icon or pressing `Ctrl + S`.
5. Close the Apps Script editor and refresh your Google Sheets document.
6. Go back to the Apps Script editor and run the `deleteTriggers` function to clear any existing triggers.
7. Run the `createTimeDrivenTrigger` function to set up a time-driven trigger to update stock prices every minute.
8. Verify the trigger by clicking on the clock icon (Triggers) in the left sidebar of the Apps Script editor. Ensure there is a trigger listed for `getStockPrices` to run every minute.

## License

This project is licensed under the [MIT License](LICENSE).