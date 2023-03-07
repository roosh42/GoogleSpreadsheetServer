/**
 * @OnlyCurrentDoc
 */
// The comment above indicates that this script should only be given permission to the spreadsheet it is attached to.

// Offered under the MIT License (2023)

/*
 * If you put in your own secret inside the quotes, that will allow you to
 * send formulas to your spreadsheet as well.
 * Leave this as-is if you don't want the possibility of formulas.
 */
function secret() {
  // TODO: Change this to a secret (min 6 characters) to allow you
  //       to insert formulas into the spreadsheet.
  // DO NOT use a password from any of your accounts.
  // If you don't want to have this option, leave it as "Nope"
  return "Nope";
}

/*****************************************************************************
 *          INSTRUCTIONS                                                     *
 * 0. (optional) Read this code and satisfy yourself that you are willing    *
 *    to run it on a specifically chosen spreadsheet in your Google account. *
 *    There are about 130 lines of actual code here, and they are not dense. *
 *                                                                           *
 * 1. In a new Google Spreadsheet, use the menu Extensions > Apps Script     *
 * 2. Copy all of this code, and paste it into the editor (for Code.gs).     *
 * 3. Change the secret() return value above to something only you know.     *
 * 4. Click Run.                                                             *
 *    - click Review Permissions                                             *
 *    - Choose your Google Account                                           *
 *    - You will be prompted to give the script permission to modify ONLY    *
 *      the spreadsheet that you just created: Approve it(^).                *
 * 5. Deploy > New Deployments                                               *
 *    - Select Type > Web App                                                *
 *    - Fill in any description you like. E.G. "My Spreadsheet Server".      *
 *    - IMPORTANT! fill in these two fields as follows:                      *
 *      Execute as:     [ Me (youremail@gmail.com) ]           !!!!          *
 *      Who has access: [ Anyone                   ]           !!!!          *
 *                      (NOT "Anyone with Google Account")                   *
 *    - Click Deploy                                                         *
 *    - copy the Web App URL                                                 *
 *     (E.G. https://script.google.com/macros/s/AKfycb.....P7_ISg/exec)      *
 *                                                                           *
 * Congratulations: You now have a way to insert rows to your spreadsheet    *
 *   by requesting the website above.  Anyone who has that link will be able *
 *   to add rows with any content they put in the url.                       *
 *                                                                           *
 * The script runs as YOU -- Not as the person who sends the url request.    *
 *   - That is why you do not need to share the spreadsheet with anyone...   *
 *     YOU already have permission to read&write to the spreadsheet.         *
 *   - The script can't do anything outside the permission you gave it,      *
 *     which was to access that one spreadsheet using your access.           *
 *                                                                           *
 * (^) You can check and revoke Third-Party permissions on your account at:  *
 *   https://myaccount.google.com/permissions?continue=https%3A%2F%2Fmyaccount.google.com%2Fsecurity
 *****************************************************************************/

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


/*
 * An conservative approximation that checks if the value might be a formula.
 */
function doesValueHaveFormula(value) {
  // Some manual testing showed that sending escape sequences for an equal sign does not create formulas.
  // Only the actual = character initiates a formula.
  let hasEqualSign = value.indexOf('=') >= 0;
  return hasEqualSign;
}

// Go through the queryParameters, and copy them into a columnName2value mapping.
// An arbitrary max of 64 columnNames will cut off any excessivly long list.
// An arbitrary max of 128 characters for the column name will reject longer names.
// An arbitrary max of 4096 bytes per value will reject any large values.
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

// Using get in order to make it transparent.
// The user can hover over the "send" link in the app,  to see what request will get sent.
// The call is over https, so it's already end-to-end encrypted
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

