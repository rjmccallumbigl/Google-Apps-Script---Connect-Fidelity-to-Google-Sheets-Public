/*********************************************************************************************************
*
* Update Google Sheet menu allowing script to be run from the spreadsheet.
*
*********************************************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('Update Fidelity Sheets', 'makeFidelityAPIRequest')
    .addItem('Update Mint Sheet', 'cloneMohitoSheet')
    .addToUi();
}

/*********************************************************************************************************
 * 
 * Update sheet from Fidelity Full View transactions. 
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
 *********************************************************************************************************/

function makeFidelityAPIRequest() {

  // Declare variables  
  var transactionUrl = "https://api.emoneyadvisor.com/snb-api/api/values/GetFilteredTransactions"

  // Examples if you wanted to modify the URL. A malformed URL will probably return Response Code 404
  // var transactionUrl = "https://api.emoneyadvisor.com/snb-api/api/values/GetFilteredTransactions?from=7%2F6%2F2021&to=8%2F5%2F2021&dateRangeType=Last+30+days";
  // var transactionUrl = "https://api.emoneyadvisor.com/snb-api/api/values/GetFilteredTransactions?accountIds[]=12345678&accountIds[]=10101010&descriptionSearchTerm=PIZZA&from=1%2F1%2F2021&to=8%2F6%2F2021";
  // var transactionUrl = "https://api.emoneyadvisor.com/snb-api/api/values/GetFilteredTransactions?descriptionSearchTerm=&from=1%2F1%2F2021&to=8%2F6%2F2021";

  // or using queryParams instead
  var startDate = encodeURIComponent("1/1/2019");
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy');
  var endDate = encodeURIComponent(today);
  var queryParams = "?from=" + startDate + "&to=" + endDate;
  // var queryParams = "";
  // var queryParams = "?dateRangeType=Last+30+days";
  // var queryParams = "?descriptionSearchTerm=BURGER";

  // API URL to grab Categories
  var categoriesUrl = "https://api.emoneyadvisor.com/snb-api/api/values/GetCategories";

  // These will probably change every time you login, please update following the directions above when they expire (Response Code 401)
  var apikey = "<<INSERT-HERE>>"
  var token = "<<INSERT-HERE>>"

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
  }

  // Send requests
  buildSheetFromAPIRequest(transactionUrl + queryParams, options, "Fidelity Transactions");
  buildSheetFromAPIRequest(categoriesUrl, options, "Fidelity Categories")
}

/*********************************************************************************************************
 * 
 * Send off API call and create a Google Sheet out of the results.
 * 
 * @param {String} url The GET URL we are contacting.
 * @param {Object} options The API options we built in our first function.
 * @param {String} sheetName The name of our sheet.
 *  
 *********************************************************************************************************/

function buildSheetFromAPIRequest(url, options, sheetName) {

  // Send API request
  var response = UrlFetchApp.fetch(url, options);

  // Parse response
  if (response.getResponseCode() == 200) {
    console.log("Successfully grabbed " + sheetName);
    var responseJSON = JSON.parse(response.getContentText());

    // Print to Google Sheet
    setArraySheet(responseJSON, sheetName);
    console.log("Using sheet " + sheetName);
    console.log(response.getContentText());
  } else {
    console.log("Failure in grabbing " + sheetName);
    console.log(response.getResponseCode());
    console.log(response);
  }
}

/******************************************************************************************************
 * Convert array into sheet
 * 
 * @param {Array} array The array that we need to map to a sheet
 * @param {String} sheetName The name of the sheet the array is being mapped to
 * 
 ******************************************************************************************************/

function setArraySheet(array, sheetName) {

  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var keyArray = [];
  var memberArray = [];
  var sheetRange = "";

  // Define an array of all the returned object's keys to act as the Header Row
  keyArray.length = 0;
  keyArray = Object.keys(array[0]);
  memberArray.length = 0;
  memberArray.push(keyArray);

  //  Capture players from returned data
  for (var x = 0; x < array.length; x++) {
    memberArray.push(keyArray.map(function (key) { return array[x][key] }));
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
  sheet.setFrozenRows(1);
  if (!sheet.getFilter()) {
    sheetRange.createFilter();
  }
  sheet.autoResizeColumns(sheetRange.getColumn(), sheetRange.getLastColumn());
}
