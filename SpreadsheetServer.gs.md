/*

## Basic Instructions for *Spreadsheet with Server* (Desktop version)

1. Make your own copy of the spreadsheet with one of your Google accounts.
  a. Log into the Google account that will own your spreadsheet.
  b. Open [Spreadsheet with Server (template)](https://docs.google.com/spreadsheets/d/1kU2IiLpKKVM_Zb3BzlB_b3I9ww1Rio81olDnzu6avWg/).
  c. `File` > `Make a Copy`
  d. Click the copy button.
2. In your new copy, open Apps Script.
  a. `Extensions` > `Apps Script`
  b. (optional) Read the code to see how it adds new rows to the spreadsheet.
3. Deploy the Apps Script as a Web App.
  a. `Deploy` *in the top-right corner* > `New Deployments`
  b. `Select Type` > `Web App`
  c. Change "Who has access" to `Anyone` (NOT "Anyone with Google Account")
  d. Click `Deploy`
4. Grant permissions (for the script to modify only the one spreadsheet it is attached to).
  a. `Authorize Access`
  b. Select the Google account that owns your spreadsheet.
  c. Click `Allow` for "View and manage spreadsheets that this applicaion has been installed in".
5. Save the Web App URL.  
  a. `Copy` the Web App URL. (It starts with https://script.google.com/macros) 
  b. Save the Web App URL (e.g. Email it yourself or paste it in a document.)

You can now close the Apps Script tab.

Congratulations: You now have a way to programatically insert rows to your spreadsheet by requesting the URL from above.
  
Anyone who has that link will be able to add rows with any content they put in the url.

Note: The script runs as ***you*** -- Not as the person who sends the url request.
  - That is why you do not need to share the spreadsheet with anyone... You already have permission to read&write the spreadsheet.
  - The script can't do anything outside the permission you gave it, which was to access that one spreadsheet using your access.

Tip: You can check and revoke Third-Party permissions on your account at [My Account](https://myaccount.google.com/permissions?continue=https%3A%2F%2Fmyaccount.google.com%2Fsecurity).


## Advanced Instructions

### Using the Web App URL
Once you have the URL of a deployed Web App, it will looks something like
  - https://script.google.com/macros/s/AKfycbwIeR6hGK_NgF22d896q............................XdSnZX41Ew/exec
If you lost the URL, go to `Deploy` > `Manage deployments` and `Copy` the URL under `Web App`.

This Web App uses [query parameters](shorturl.at/lvwGU) to pass in the information for a new row in the spreadsheet.

The only special parameter keys are
 - sheetname: The preferred name of the sheet into which the row is added.
 - server_time: The time that this row is processed.
 - secret: Optional secret word if you want to allow formulas.
 - password: Eliminated from the query parameters, as a precaution.

 For all other query parameters, the key is used as the column heading, and the value is what shows up in the newly inserted row.

 Try pasting your URL into the address bar of a browser and add ?firstName=Big&lastName=Bird.
 For example: 
 ```
 https://script.google.com/macros/s/AKfycb....replace_with_your_url........nZX41Ew/exec?firstName=Big&lastName=Bird&sheetName=data
 ```

This Spreadsheet with Server is designed to let apps and websites programatically insert rows to the spreadsheet
by creating urls like the one above, and requesting that url (with a GET request).

Developer Tip: CORS Policy does not allow any data to be returned by the server, so set the `mode='no-cors'`.


### Enabling formulas for inserted spreadsheet rows.
If you want to be able to send formulas to your spreadsheet through the Web App URL, follow these steps.
1. Pick a secret.
  a. Scroll a few lines down and find `function secret()`.
  b. Change the word in single-quotes to a secret
    - This is not an account password.
    - Minimum 6 characters long.
    - Case sensitive.
2. Follow the Basic Instructions above, starting at Step 3, to create a `New Deployment`.
  - You may not need to re-authorize, so step 4 will be skipped.

##The Code:
```

*/

// Offered under the [MIT License](https://en.wikipedia.org/wiki/MIT_License) (2023)
version=202303260622;

// The comment below indicates that this script will only be given permission
// to the spreadsheet it is attached to.
/**
 * @OnlyCurrentDoc
 */

/*
 * A "secret" string that allows you to enable inserting formulas in the spreadsheet.
 */
function secret() {
  // TODO: Change this to a secret (min 6 characters) to allow you
  //       to insert formulas into the spreadsheet.
  // DO NOT use a password from any of your accounts.
  // If you don't want to have this option, leave it as "Nope"
  return 'Nope';
}


/*
 * NOTE: Throughout this code, ñ is prepended to varibale names that have been normalized.
 *       This helps keep track when comparing the values of two variables.
 * This is a helper function that takes a value and normalizes it by:
 *   1. Removing any spaces, underscores, dashes.  E.G. " First name - given_or_nickName  " -> "FirstnamegivenornickName"
 *   2. Changing all the letters to lowercase.  E.G. "FirstnamegivenornickName" -> "firstnamegivenornickname"
 * You can check the regular expression /[\s_-]\/g in https://regexr.com/79llt
 */
function normalize(key) {
  if (typeof key != 'string') {
    return undefined;
  }
  let ñkey = key;
  ñkey = ñkey.replace(/[\s_-]/g, '');
  ñkey = ñkey.toLowerCase();
  return ñkey;
}


/*
 * THIS IS THE MAIN FUNCTION.  It is the starting point that gets called whenever
 * someone sends a request to the Web App URL (from step 5d of the instructions).
 * The values that this function uses are the "query parameters".  That's all
 * the stuff after the '?' (assuming there is a '?').  The paramters take the form
 * key=value, and they are separated with '&'.  For example:
 * 
 * If the url is https://script.google.com/macros/e/d/ABcd12_34/exec
 * then query parameters are added by making the url something like this:
 * https://script.google.com/macros/e/d/ABcd12_34/exec?firstName=Snow&lastName=White&numFriends=7
 * 
 * Dealing with spaces gets a bit tricky, but for the most part, they are dealt with.
 * 
 * Even though you type the query parameters into the browser with plain text, the browser
 * encrypts them before sending them over the internet to this function.  Once they arrive,
 * they are decrypted so the function can deal with the plain text that was in the query parameters. 
 */
function doGet(e) {
  // This script is only allowed to view/edit the spreadsheet to which it is attached.
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Get the sheet that has the sheet_name specified in the request.
  // See comments in getSheet for fallbacks if no sheet by that name exists.

  // columnName2value is a map from column name to the value that should be inserted into that column.
  // Though it is named in the singular (i.e. not columnNames2values), it contains multiple pairings of {columnName: value}.
  let columnName2value = sanitizeQueryParameters(e.parameter);

  // If there is no 'sheetname', then sheetName will be undefined, and that's OK.
  let sheetName = columnName2value['sheetname'];
  let sheet = chooseSheet(spreadsheet, sheetName);

  // From https://support.google.com/docs/thread/181288162/whats-the-maximum-amount-of-rows-in-google-sheets?hl=en
  // The maximum number of columns is 18,278 and the maximum cells is 10,000,000
  // If the sheet is already within a factor of 2 for either of these, do not do anything more.
  let numRows = sheet.getMaxRows();
  let numColumns = sheet.getMaxColumns();
  if (numColumns > (18278 / 2)) {
    throw `Sheet ${sheet.getName()} has too many columns (${numColumns} > ${18278/2})`;
  }
  let numCells = numRows * numColumns;
  if (numCells > (10000000 / 2)) {
    throw `Sheet ${sheet.getName()} has too many cells (${numCells} > ${10000000/2})`;
  }

  // record the time on the server that this request came in. In JavaScript, "new Date()" gives the date-time for 'now'.
  columnName2value.server_time = JSON.stringify(new Date());
  
  // Add any missing columns that the request recorded in column2values.
  // Returns a map of the sheet's header, from column name to column index.
  appendValuesToSheet(columnName2value, sheet);

  return ContentService.createTextOutput("OK");
}



/*
 * Tries to return the sheet with the name that matches soughtSheetName.
 * The function only compares the normalized names of the sheets.
 *   - So if a sheet is named "Cute cats", and the requested sheet_name is "cute  Cats ", they will match.
 * If it can't (because the sheet doesn't exist),
 *   try a sheet named "errors",
 *   and then try the last sheet in the spreadsheet
 */
function chooseSheet(spreadsheet, soughtSheetName) {
  // This funciton will only compare the normalized versions of sheet names.
  let ñsoughtSheetName = normalize(soughtSheetName);

  // Get a list of all the sheets in the spreadsheet.
  let sheets = spreadsheet.getSheets();
  let chosenSheet = undefined;
  let errorSheet = undefined;
  for (let sheet of sheets) {
    let ñproposedSheetName = normalize(sheet.getName());
    if (ñproposedSheetName == ñsoughtSheetName) {
      // if a sheet with a matching name is found, set it and break out of the loop
      chosenSheet = sheet;
      break;
    }
    if (ñproposedSheetName == 'errors') {
      // if a sheet with the special name 'errors' (or 'ERRORS' or 'Errors', etc.) is found, remember it for later.
      errorSheet = sheet;
    }
  }

  // If the chosen_sheet is still undefined, and there is an error_sheet (i.e. it is not undefined), use the error_sheet.
  if (chosenSheet === undefined && errorSheet !== undefined) {
    chosenSheet = errorSheet;
  }
  // If the chosedn_sheet is still undefined, use the last sheet in the spreadsheet.  (length-1 is the way to index the last sheet.)
  if (chosenSheet === undefined) {
    chosenSheet = sheets[sheets.length-1];
  }

  return chosenSheet;
}

/*
 * Terminology: columnName comes from the incoming request, while headerName comes from what's already in the spreadsheet.
 *
 * Adds any columnName that isn't already in the sheet's header row.
 * Comparison is based on the normalizedKey() version of both the columnName and the headerName.
 * Returns ñheaderName2columnIndex, which is a mapping from normalize header name to the column index.
 * 
 * Discussion about stray values (optional reading):
 *  Each column has a value in the header row (the top row).  That value is called a headerName.
 *  It is possible that someone puts a value somewhere in a column that does not have a headerName (the top cell in the column is empty).
 *  In that case, this code will treat that as a real column with values that should be preserved.  The value of that headerName is "" (the empty string).
 *    This happens because the call to sheet.getLastColumn() returns the position (1-based) of the last column that has content.
 *
 *  The same goes for stray values in a row way down, below where all the "real" data is:
 *    If there are many empty rows, and then a row has some value in a cell, all the rows up to that are treated as valid rows.
 *    Their values happen to all be "" (the empty string).
 *    New rows will be added below the last row with any values in it.
 *    This happens because of how sheet.appendRow() inserts a row after all the existing content.
 */
function addNovelColumns(sheet, columnName2value) {
  // How many columns have a value in them already
  let numHeaderNames = sheet.getLastColumn();

  // Count how many column names are in the incoming data.
  let numColumnNames = 0;
  for (let columnName in columnName2value) {
    numColumnNames++;
  }

  // If none of the columnNames match the headerNames, the final spreadsheet will have all the orginal headers plus all the new columnNames.
  let maxHeaderNames = numHeaderNames + numColumnNames;
  // Get the first maxHeaderNames cells in the first row.
  let headerRange = sheet.getRange(startRow=1, startColumn=1, numRows=1, maxHeaderNames);  // note: The startRow and startColumn for getRange() are 1-based.
  // Detail: getValues() returns a 2D array. It is an array of rows, each row being an array of column values.
  // After the values are updated in-place, rangeValues will be used to set the values of that headerRange back in the spreadsheet.
  let rangeValues = headerRange.getValues();
  // headerNames is an array of the header names.  The last numColumnName values are empty.
  // Detail: The [0] is because we only want to work with the first row.
  let headerNames = rangeValues[0];

  //  This will store the column index for each header name.  The headerNames will be normalized; hence the ñ prefix.
  let ñheaderName2columnIndex = {};

  // Go over the list of headerNames and record their column index.
  for (let columnIndex in headerNames) {
    let headerName = headerNames[columnIndex];
    let ñheaderName = normalize(headerName);
    // Some of the headerNames might normalize to the same value.  In that case, the last one wins.
    ñheaderName2columnIndex[ñheaderName] = columnIndex;
  }

  // Optional: See detailed discussion about stray values in the spreadsheet for corner-case details.
  
  // Now it's time to add more ñcolumnName -> columnIndex pairings to ñheaderName2columnIndex,
  // for the new column names that are  not in the spreadsheet's header names.

  // When we find a novel column name, it will go at nextColumnIndex.
  let nextColumnIndex = numHeaderNames;
  // Go over each column name and check if it matches an existing header name.
  // If it is novel, add it to the header names, and update the ñheaderName2columnIndex mapping.
  for (let columnName in columnName2value) {
    let ñcolumnName = normalize(columnName);
    if (ñcolumnName in ñheaderName2columnIndex) {
      // That column name already matches an existing header name.  Continue to the next column name.
      continue;
    }

    // Add the novel column name (not the normalized version) to the end of the headerNames.
    headerNames[nextColumnIndex] = columnName;
    // Let's remember that this new header name now exists at nextColumnIndex.
    ñheaderName2columnIndex[ñcolumnName] = nextColumnIndex;
    // increment so we are ready to use the next integer value for nextColumnIndex, if another novel column name is found.
    nextColumnIndex++;
  }

  // if the value of nextColumnIndex never changed from its original value of numHeaderNames (see about 20 lines above),
  //  there are no new header names to update.
  if (nextColumnIndex == numHeaderNames) {
    // nothing to do since we didn't add any new column names
  } else {
    // Detail: Since JavaScript is a by-reference language, when we updated headerNames, we were also updating the values in headerValues.
    headerRange.setValues(rangeValues);
  }
  return ñheaderName2columnIndex;
}


/*
 * Creates a simple 1D array with all the values form columnName2value.
 * The order of the values in the array is governed by the mapping of ñheaderName2columnIndex.
 * The row probably won't be dense, meaning some of the values will be undefined.
 */
function createRowToInsert(ñheaderName2columnIndex, columnName2value) {
  let errorMessages = [];
  let rowToInsert = [];

  // For each value, find the index for its associated columnName by looking at the mapping ñheaderName2columnIndex.
  for (columnName in columnName2value) {
    // This is the value that will be stored in the spreadsheet.
    let value = columnName2value[columnName];
    let ñcolumnName = normalize(columnName);
    let columnIndex = ñheaderName2columnIndex[ñcolumnName];

    // Do a sanity-check and capture an errorMessage in case the lookup for columnIndex failed.
    if (columnIndex === undefined) {  // JS: indexing into a non-existent key gives undefined
      // This should not happen because all the column names should already exist in ñheaderName2columnIndex, thanks to addMissingColumns()
      let errorMessage = `ERROR: ${ñcolumnName} not found for value = ${value}`;
      errorMessages.push(errorMessage);
      continue;
    }

    // Detail: This is specific to JavaScript, where array items can be referenced without allocating them.
    // Put the value at the correct column index.
    rowToInsert[columnIndex] = value;
  }
  // In case there are any errorMessages, stick them on the end.
  rowToInsert.concat(errorMessages);
  return rowToInsert;
}

/*
 * Makes sure the column names all exist (creating new columns with header names as needed).
 * Appends the values to the end of the sheet, respecting the order of the columns already in the sheet.
 */
function appendValuesToSheet(columnName2value, sheet) {
  // Add any novel columns that may be in columnName but aren't in the sheet's header row yet.
  let headerName2columnIndex = addNovelColumns(sheet, columnName2value);
  // Create the row with all the values in the right order to match the order of the sheet's header names.
  let rowToInsert = createRowToInsert(headerName2columnIndex, columnName2value);
  // appendRow() is safe to call even if there aren't enough columns in the sheet.  (They will be added as needed.)
  sheet.appendRow(rowToInsert);
}


/**** The functions below this line may be a bit harder to read for non-programmers. ****/

/* 
 * Go through the queryParameters, and copy them into a columnName2value mapping.
 * An arbitrary max of 64 columnNames will cut off any excessivly long list.
 * An arbitrary max of 128 characters for the column name will reject longer names.
 * An arbitrary max of 4096 bytes per value will reject any large values.
 */
function sanitizeQueryParameters(queryParameters) {

  // if a secret has been set to a string at least 6 characters long,
  // and the request has a "secret" that matches, formulas are allowed. 
  let allowFormulas = false;
  if (secret() && typeof secret() == 'string' && secret().length >= 6) {
    if (queryParameters.secret = secret()) {
      allowFormulas = true;
    }
  }

  let errors = {
    numParametersIgnored: 0,  // max 64
    numColumnsTooLong: 0,     // max 128
    numValuesTooLarge: 0,     // max 4096
    numFormulasExcluded: 0,   // if contains = sign and allowFormulas=false
  };

  let numColumns = 0;
  let columnName2value = {};
  for (let key in queryParameters) {
    // increment the value every time we are in the loop
    numColumns++;

    let value = queryParameters[key];
    // This check is done first because even if the key is 'sheetName', we don't want to use such a long value.
    if (value.length > 4096) {
      errors.numValuesTooLarge++;
      continue;
    }

    // Even if allowFormulas is on, there are some functions which can be particularly dangerous
    // in their ability to call out to a url and thus send the spreadsheet's information to a 3rd party.
    let valueMightBeDangrousFormula = doesValueHaveDangerousFormula(value);
    if (valueMightBeDangrousFormula) {
      errors.numFormulasExcluded++;
      continue;
    }

    // if value might have a formula and that is not allowed, continue to the next key/value
    let valueMightHaveFormula = doesValueHaveFormula(value);
    if (valueMightHaveFormula && !allowFormulas) {
      errors.numFormulasExcluded++;
      continue;
    }

    // Things like 'sheet_name', 'Sheet Name', 'sheetName' will all normalize to 'sheetname'
    // We're looking to find a sheet name, even if it's past the max number of columns allowed.
    let ñkey = normalize(key);
    if (ñkey == 'sheetname') {
      columnName2value['sheetname'] = queryParameters[key];  // This is the only key that we alter into the normalized one.
      continue;
    }

    // We don't want to store the secret (or Secret) in the spreadsheet.
    if (ñkey == 'secret') {
      continue;
    }

    // For that matter, let's not store anything called 'password'
    if (ñkey == 'password') {
      continue;
    }

    if (numColumns > 64) {
      errors.numParametersIgnored++;
      continue;
    }
    if (key.length > 128) {
      errors.numColumnsTooLong++;
      continue;
    }
    columnName2value[key] = value;
  }

  // add in any errors we encountered
  for (let error in errors) {
    let numErrors = errors[error];
    if (numErrors > 0) {
      columnName2value[error] = numErrors;
    }
  }

  return columnName2value;
}


/*
 * An conservative approximation that checks if the value might be a formula.
 */
function doesValueHaveFormula(value) {
  // Some manual testing showed that sending escape sequences for an equal sign does not create formulas.
  // Only the actual = character initiates a formula.
  let hasEqualSign = value.indexOf('=') >= 0;
  return hasEqualSign;
}

/*
 * Heuristic that checks if the value might have a dangerous formula that
 * could potentially contain a url argument which can leak information from
 * the spreadsheet.
 */
function doesValueHaveDangerousFormula(value) {
  // Based on https://support.google.com/docs/search?q=url, these are the functions that treat
  // one of their parameters as a url, so that is a potential leak for sending out data.
  let urlFunctions = ["HYPERLINK", "IMPORTDATA", "IMPORTHTML", "IMPORTRANGE", "IMAGE", "IMPORTFEED", "IMPORTXML"];
  // Looking at the above, I'll condense it a bit:
  let dangerousSubstrings = ["LINK", "IMPORT", "HTML", "IMAGE"];
  for (let dangerousSubstring of dangerousSubstrings) {
    if (value.indexOf(dangerousSubstring) >= 0) {
      return true;
    }
  }
  return false;
}

// ```
