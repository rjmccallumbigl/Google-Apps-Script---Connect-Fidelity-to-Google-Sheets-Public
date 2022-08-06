/****************************************************************************************************************************************
 * 
 * Update sheet from Fidelity Full View transactions. 
 *
 * @param {String} token The Bearer token grabbed from loading Fidelity Full View.
 * 
 * Background
 * I looked at how Mohito worked for Mint http://b3devs.blogspot.com/p/about-mojito.html and tried the following:
 *    Signed in to Fidelity Full View.
 *    Open Chrome Dev Tools and went to the Network tab.
 *    Loaded Transactions for the last 30 days.
 *    Exported the HAR trace.
 *    Imported the HAR trace to https://toolbox.googleapps.com/apps/har_analyzer/
 *   Searched for data that would probably be loaded in a transaction (e.g. unique account identifier).
 * 
 * This gave me the API link and returned a JSON full of transaction info. This was a good enough start 
 * to navigate to Amit Agarwal's tool here to generate the HTTP fetch request from Google Apps Script:
 * https://www.labnol.org/apps/urlfetch.html
 * 
 * Instructions
 *  1. Set up your Fidelity Full View account. 
 *  2. Open Chrome Dev Tools (F12). Go to the Network tab. 
 *  3. Filter for "GetFilteredTransactions" (no quotes).
 *  4. Navigate to the Spending tab in Fidelity Full View. If "GetFilteredTransactions" didn't return anything, it should now.
 *  5. You should find an API request matching the headers below with 2 differences: apikey and authorization.
 *  6. Fill out var token and var apikey.
 *  7. Run onOpen(). Now when you refresh your Sheet you can make run the script from the menu.
 *  8. Run the script 'Update Fidelity Sheets' [makeFidelityAPIRequest()]. It will return up to 2000 of your transactions.
 *  9. [Optional] Some of the HTML encoded parameters I've seen you can try: 
 *      from=<Start date of search> [Optional]
 *      to=<Last date of search> [Optional]
 *      descriptionSearchTerm=<Enter search term here for specific query spending> [Optional]
 *  
 * Sources
 * http://b3devs.blogspot.com/p/about-mojito.html
 * 
 * Version
 * 0.4.1
 *
 ****************************************************************************************************************************************/

function makeFidelityAPIRequest(token) {

  // Declare variables  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var startDate = encodeURIComponent("1/1/2019");
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var endDate = encodeURIComponent(today);
  var startOfMonth = ((new Date()).getMonth() + 1).toString() + "/1/" + ((new Date()).getFullYear()).toString();
  var queryParams = "?from=" + startDate + "&to=" + endDate;
  var apiResponse = 0;

  // Backup sheet before proceeding
  cloneGoogleSheet(spreadsheet);

  // API URLs to grab
  var baseURL = "https://api.emoneyadvisor.com/snb-api/api";
  var transactionUrl = baseURL + "/values/GetFilteredTransactions";
  var categoriesURL = baseURL + "/values/GetCategories";
  var accountsURL = baseURL + "/values/GetAccounts?shouldExcludeHiddenAccounts=False";
  var transactionRulesURL = baseURL + "/values/GetBankTransactionRules";
  // var budgetsURL = baseURL + "/budgets/GetBudgets?dateRangeType=Last+30+days&startDate=7%2F9%2F2021&endDate=8%2F8%2F2021";
  var budgetsURL = baseURL + "/budgets/GetBudgets?dateRangeType=Last+30+days&startDate=" + startDate + "&endDate=" + endDate;
  // var otherExpensesURL = baseURL + "/values/GetOtherExpenses?fromDate=2021-08-01T04:00:00.000Z&toDate=2021-08-08T15:27:04.959Z";
  var otherExpensesURL = baseURL + "/values/GetOtherExpenses?fromDate=" + startDate + "&toDate=" + endDate;
  var overallBudgetURL = baseURL + "/budgets/GetOverallBudget";

  // Examples if you wanted to modify the URL. A malformed URL will probably return Response Code 404
  // var transactionUrl = baseURL + "/values/GetFilteredTransactions?from=7%2F6%2F2021&to=8%2F5%2F2021&dateRangeType=Last+30+days";
  // var transactionUrl = baseURL + "/values/GetFilteredTransactions?accountIds[]=12345678&accountIds[]=10101010&descriptionSearchTerm=PIZZA&from=1%2F1%2F2021&to=8%2F6%2F2021";
  // var transactionUrl = baseURL + "/values/GetFilteredTransactions?descriptionSearchTerm=&from=1%2F1%2F2021&to=8%2F6%2F2021";

  // or using queryParams instead at the end of the URL
  // var queryParams = "";
  // var queryParams = "?dateRangeType=Last+30+days";
  // var queryParams = "?descriptionSearchTerm=BURGER";

  // The token will probably change every time you login, please update following the directions above when they expire (Response Code 401)
  var apikey = "<<INSERT-HERE>>";
  var token = token || "<<INSERT-HERE>>";

  // Build options for API request
  var options = {
    "method": "GET",
    "headers": {
      "Authorization": "Bearer " + token,
      "Accept": "application/json, text/plain, */*",
      "Accept-Encoding": "gzip, deflate, br",
      "Accept-Language": "en-US,en;q=0.9",
      "Connection": "keep-alive",
      "DNT1": "1",
      // "Host": "api.emoneyadvisor.com",
      "Origin": "https://wealth.emaplan.com",
      "Referer": "https://wealth.emaplan.com/",
      "Sec-Fetch-Dest": "empty",
      "Sec-Fetch-Mode": "cors",
      "Sec-Fetch-Site": "cross-site",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
      "apikey": apikey,
      "sec-ch-ua": "\"Chromium\";v=\"92\", \" Not A;Brand\";v=\"99\", \"Google Chrome\";v=\"92\"",
      "sec-ch-ua-mobile": "?0"
    },
    "muteHttpExceptions": true,
    "followRedirects": true,
    "validateHttpsCertificates": true,
  };

  // Send requests
  apiResponse = buildSheetFromAPIRequest(transactionUrl + queryParams, options, "Fidelity Transactions", spreadsheet);

  // Test success of API call prior to continuing script
  if (apiResponse == 401) {
    console.log("Unsuccessful, follow instructions and retry script after updating var token");
    //   var ui = SpreadsheetApp.getUi();
    //   var response = ui.prompt('Enter token from Dev Tools -> Network -> Filter for "GetFilteredTransactions"');

    //   // Process response with entered token
    //   if (response.getSelectedButton() == ui.Button.OK) {
    //     makeFidelityAPIRequest(response.getResponseText());
    //   } else {
    //     return;
    //   }

    var htmlSource = HtmlService.createHtmlOutput(buildHTMLDialog());
    SpreadsheetApp.getUi().showModalDialog(htmlSource, "Authenticate");
    return;
  }

  // Update Fidelity sheets after authenticating
  buildSheetFromAPIRequest(categoriesURL, options, "Fidelity Categories", spreadsheet);
  buildSheetFromAPIRequest(accountsURL, options, "Fidelity Accounts", spreadsheet);
  buildSheetFromAPIRequest(transactionRulesURL, options, "Fidelity Transaction Rules", spreadsheet);
  buildSheetFromAPIRequest(budgetsURL, options, "Fidelity Budgets", spreadsheet);
  buildSheetFromAPIRequest(otherExpensesURL, options, "Fidelity Other Expenses", spreadsheet);

  // Update Options for Overall Budget request, which has a couple different/new params
  options.method = "POST";
  options.headers["Content-Type"] = "application/json;charset=UTF-8";
  options.payload = JSON.stringify({ startDate: startOfMonth, endDate: today });
  options.convertArray = true;
  buildSheetFromAPIRequest(overallBudgetURL, options, "Fidelity Overall Budget", spreadsheet);

  // Update categories in Transactions sheet
  replaceCategoryIDWithName(spreadsheet);

  // Format sheets a little by deleting empty columns + rows
  removeEmptyColumns(spreadsheet);
  removeEmptyRows(spreadsheet);
}

/****************************************************************************************************************************************
 * 
 * Send off API call and create a Google Sheet out of the results.
 * 
 * @param {String} url The GET URL we are contacting.
 * @param {Object} options The API options we built in our first function.
 * @param {String} sheetName The name of our sheet.
 * @param {Object} spreadsheet The source spreadsheet
 * @return {Number} The response code of the API request. 200 is successful, anything else is a fail.
 *  
 ****************************************************************************************************************************************/

function buildSheetFromAPIRequest(url, options, sheetName, spreadsheet) {

  // Send API request
  var response = UrlFetchApp.fetch(url, options);

  // Parse response
  if (response.getResponseCode() == 200) {
    console.log("Successfully grabbed " + sheetName);
    var responseJSON = JSON.parse(response.getContentText());

    // Convert 1D array to 2D if necessary
    if (options.convertArray) {
      responseJSON = [responseJSON];
    }

    // Print to Google Sheet
    setArraySheet(responseJSON, sheetName, spreadsheet);
    console.log("Using sheet " + sheetName);
    console.log(response.getContentText());
  } else {
    console.log("Failure in grabbing " + sheetName);
    console.log(response.getResponseCode());
    console.log(response.getContentText());
  }

  // Return the response code
  return response.getResponseCode();
}

/****************************************************************************************************************************************
*
* Update Google Sheet menu allowing script to be run from the spreadsheet.
*
****************************************************************************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('Update Fidelity Sheets', 'makeFidelityAPIRequest')
    .addToUi();
}

/****************************************************************************************************************************************
 * 
 * Convert array into sheet
 * 
 * @param {Array} array The array that we need to map to a sheet
 * @param {String} sheetName The name of the sheet the array is being mapped to
 * @param {Object} spreadsheet The source spreadsheet
 * 
 ****************************************************************************************************************************************/

function setArraySheet(array, sheetName, spreadsheet) {

  // Declare variables
  var keyArray = [];
  var memberArray = [];
  var sheetRange = "";

  // Define an array of all the returned object's keys to act as the Header Row
  keyArray.length = 0;
  if (sheetName == "Mortgage") {
    keyArray = Object.keys(array);
  } else {
    keyArray = Object.keys(array[0]);
  }
  memberArray.length = 0;
  memberArray.push(keyArray);

  //  Capture values from returned data
  if (sheetName == "Mortgage") {
    memberArray.push(keyArray.map(function (key) {
      if (key == "date") {
        return Utilities.formatDate(new Date(array[x][key]), 'America/New_York', 'yyyy-MM-dd H:mm:ss');
      } else {
        if (sheetName == "Mortgage") {
          return array[key];
        } else {
          return array[x][key];
        }

      }
    }));
    var transposed = [];
    memberArray = memberArray[0].map(function (col, c) {
      // For each column, iterate all rows
      return memberArray.map(function (row, r) {
        return memberArray[r][c];
      });
    });
  } else {
    for (var x = 0; x < array.length; x++) {
      memberArray.push(keyArray.map(function (key) {
        if (key == "date") {
          return Utilities.formatDate(new Date(array[x][key]), 'America/New_York', 'yyyy-MM-dd H:mm:ss');
        } else {
          if (sheetName == "Mortgage") {
            return array[key];
          } else {
            return array[x][key];
          }

        }
      }));
    }
  }

  // Select or create the sheet
  try {
    sheet = spreadsheet.insertSheet(sheetName);
  } catch (e) {
    sheet = spreadsheet.getSheetByName(sheetName).clear();
  }

  // Set values  
  sheetRange = sheet.getRange(1, 1, memberArray.length, memberArray[0].length);
  sheetRange.setValues(memberArray);

  // Pretty up sheet
  if (sheetName == "Mortgage") {
    // sheet.autoResizeColumns(sheetRange.getColumn(), sheetRange.getLastColumn());
  } else {
    sheet.setFrozenRows(1);
    try {
      sheet.setFrozenColumns(3);
    } catch (e) {
      sheet.setFrozenColumns(1);
    }

    if (!sheet.getFilter()) {
      sheetRange.createFilter();
    }
    SpreadsheetApp.flush();
    sheet.autoResizeColumns(sheetRange.getColumn(), sheetRange.getLastColumn());
  }
}

  /****************************************************************************************************************************************
   * 
   * Replace the categoryId in the transaction sheet by the category name for easier parsing.
   *  
   * @param {Object} spreadsheet The source spreadsheet
   *  
  
   ****************************************************************************************************************************************/

  function replaceCategoryIDWithName(spreadsheet) {

    //  Declare variables
    var transactionSheet = spreadsheet.getSheetByName("Fidelity Transactions");
    var transactionSheetHeaderRange = transactionSheet.getRange(1, 1, 1, transactionSheet.getLastColumn());
    var transactionSheetHeaderRangeValues = transactionSheetHeaderRange.getDisplayValues();
    var categoryIdHeader = transactionSheetHeaderRangeValues[0].indexOf("categoryId");
    var flatCategoriesRange = transactionSheet.getRange(1, categoryIdHeader + 1, transactionSheet.getLastRow(), 1);
    var flatCategoriesArray = flatCategoriesRange.getDisplayValues().join().split(",");
    var categorySheet = spreadsheet.getSheetByName("Fidelity Categories");
    var categorySheetRange = categorySheet.getDataRange();
    var categorySheetRangeValues = categorySheetRange.getDisplayValues();
    var categoryJSON = getJsonArrayFromSheet(categorySheetRangeValues);
    var matchingCategory = {};
    var matchingCategoryString = "";
    var updatedArray = [["categoryId"]];

    // Parse through category IDs on Fidelity Transactions sheet
    for (var x = 0; x < flatCategoriesArray.length; x++) {
      matchingCategory = {};
      matchingCategoryString = "";

      // if ID detected, replace with category name (+ parent category if found)
      if (flatCategoriesArray[x] != "categoryId") {
        if (flatCategoriesArray[x]) {
          matchingCategory = findCategory(flatCategoriesArray[x], categoryJSON);
          matchingCategoryString = matchingCategory.name;

          // Prepend with Parent category if there is one
          if (matchingCategory.parentId) {
            matchingCategoryString = findCategory(matchingCategory.parentId, categoryJSON).name + " | " + matchingCategoryString;
          }

          // Update replacement array with Category name
          updatedArray.push([matchingCategoryString]);
        } else {
          updatedArray.push([""]);
        }
      }
    }

    // Update Fidelity Transactions sheet
    flatCategoriesRange.setValues(updatedArray);
  }

  /****************************************************************************************************************************************
  *
  * Find the category by the ID.
  *
  * @param {Number} categoryID Category ID we need to match.
  * @param {Array} categoryJSON The array of category objects with names and associated IDs.
  * @return {Object} Return the matching categories.
  *
  * Sources
  * https://usefulangle.com/post/3/javascript-search-array-of-objects
  *
  ****************************************************************************************************************************************/

  function findCategory(categoryID, categoryJSON) {

    // Search the object array for matching category IDs
    var category = categoryJSON.find(function (categoryJSONObject, index) {
      if (categoryJSONObject.id == categoryID)
        return true;
    });

    // Return the matching category object
    return category;
  }

  /****************************************************************************************************************************************
   * 
   * Convert Google Sheet data as a 2D array to an array full of JSON objects.
   * 
   * @param {Array} data The 2D array we are converting.
   * @return {Array} The JSON objects resulting from the conversion stored in an array.
   * 
   * Source
   * https://stackoverflow.com/a/47555577/7954017
   *  
   ****************************************************************************************************************************************/

  function getJsonArrayFromSheet(data) {

    // Declare variables
    var obj = {};
    var result = [];
    var row = [];
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {

      // Get a row to fill the object
      row = data[i];

      // Clear object
      obj = {};

      // Fill object with new values
      for (var col = 0; col < headers.length; col++) {
        obj[headers[col]] = row[col];
      }

      // Add object in a final result
      result.push(obj);
    }

    return result;
  }

  /****************************************************************************************************************************************
   * 
   * Delete empty columns
   * 
   * @param {Object} spreadsheet The source spreadsheet or sheet
   * 
   /****************************************************************************************************************************************/

  function removeEmptyColumns(spreadsheet) {
    try {
      // For spreadsheets
      var allsheets = spreadsheet.getSheets();
      for (var s in allsheets) {
        removeEmptyColumnsInSheet(allsheets[s]);
      }
    } catch (e) {
      // For sheets
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet.getName());
      removeEmptyColumnsInSheet(sheet);
    }
  }

  function removeEmptyColumnsInSheet(sheet) {
    var maxColumns = sheet.getMaxColumns();
    var lastColumn = sheet.getLastColumn();
    if (maxColumns - lastColumn != 0) {
      try {
        sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
      } catch (e) {
        console.log(sheet.getName() + ": " + e);
      }
    }
  }

  /****************************************************************************************************************************************
   * 
   * Delete empty rows
   * 
   * @param {Object} spreadsheet The source spreadsheet or sheet
   *
   /****************************************************************************************************************************************/

   function removeEmptyRows(spreadsheet) {
    try {
      // For spreadsheets
      var allsheets = spreadsheet.getSheets();
      for (var s in allsheets) {
        removeEmptyRowsInSheet(allsheets[s]);
      }
    } catch (e) {
      // For sheets
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet.getName());
      removeEmptyRowsInSheet(sheet);
    }
  }

function removeEmptyRowsInSheet(sheet) {
    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    if (maxRows - lastRow > 1) {
      try {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      } catch (e) {
        console.log(sheet.getName() + ": " + e);
      }
    }
  }
  