/**
 * @OnlyCurrentDoc
 */
// The comment above indicates that this script should only be given permission to the spreadsheet it is attached to.

/*
 * NOTE: Throughout this code, ñ is prepended to varibale names that have been normalized.
 *       This helps keep track when comparing the values of two variables.
 * This is a helper function that takes a value and normalizes it by:
 *   1. Removing any whitespace.  E.G. " First name  " -> "Firstname"
 *   2. Changing all the letters to lowercase.  E.G. "Firstname" -> "firstname"
 */
function normalize(key) {
  let ñkey = key;
  ñkey = ñkey.replace(/\w/g, '');
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
    if (ñproposedSheetName == "errors") {
      // if a sheet with the special name "errors" (or "ERRORS" or "Errors", etc.) is found, remember it for later.
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
  let numHeaderNames = sheet.getLastColumn();
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
  
  // Now it's time to add more ñcolumnName -> columnIndex pairings to ñheaderName2columnIndex, for the new column names that are  not in the spreadsheet's header names.

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

  // if the value of nextColumnIndex never changed from its original value of numHeaderNames (see about 20 lines above), there are no new header names to update.
  if (nextColumnIndex == numHeaderNames) {
    // nothing to do since we didn't add any new column names
  } else {
    // Detail: Since JavaScript is a by-reference language, when we updated headerNames, we were also updating the values in headerValues.
    header.setValues(headerValues);
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
  for (columnName in columnName2value) {
    let value = columnName2value[columnName];
    let ñcolumnName = normalize(columnName);
    let columnIndex = ñheaderName2columnIndex[ñcolumnName];
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
  let rowToInsert = createRowToInsert(headerName2columnIndex, columnName2values);
  // appendRow() is safe to call even if there aren't enough columns in the sheet.  (They will be added as needed.)
  sheet.appendRow(rowToInsert);
}

// Using get in order to make it transparent.
// The user can hover over the "send" link in the app,  to see what request will get sent.
// The call is over https, so it's already end-to-end encrypted
function doGet(e) {
  // columnName2value is a map from column name to the value that should be inserted into that column.
  // Though it is named in the singular (i.e. not columnNames2values), it contains multiple pairings of {columnName: value}.
  let columnName2value = e.parameters;

  // This script is only allowed to view/edit the spreadsheet to which it is attached.
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Get the sheet that has the sheet_name specified in the request.
  // See comments in getSheet for fallbacks if no sheet by that name exists.
  let sheet = chooseSheet(spreadsheet, columnName2value.sheetName);

  // record the time on the server that this request came in. In JavaScript, "new Date()" gives the date-time for 'now'.
  columnName2value["server time"] = new Date();
  
  // Add any missing columns that the request recorded in column2values.
  // Returns a map of the sheet's header, from column name to column index.
  appendValuesToSheet(columnName2value, sheet);
}
