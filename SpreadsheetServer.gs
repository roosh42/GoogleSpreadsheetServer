// For instructions, see https://anerb.github.io/GoogleSpreadsheetServer

// Offered under the [MIT License](https://en.wikipedia.org/wiki/MIT_License) (2023)

version=20230455160100;

// TODO: Change all references to sheet/row/column into table/record/field.  This code should use general database terms, even if it's implemented in a spreadsheet.

// The comment below indicates that this script will only be given permission
// to the spreadsheet it is attached to.
/**
 * @OnlyCurrentDoc
 */

/*
 * Some field names have special meaning for the server and the server reserves the privelege to do more (or less) than directly store them.
 */
const SERVER_FIELDS = {
  id: 'server_id',
  v: 'server_v',
  sheetname: 'server_sheetname',
  created_us: 'server_created_us',
  modified_us: 'server_modified_us',
  secret: 'server_secret',
}

/*
 * TODO: Let the person deploying this server choose which of the server fields are always writen for every row.
 */
function getMandatoryColumnNames() {
  let mandatoryColumnNames = [];
  for (let serverColumnName in SERVER_FIELDS) {
    mandatoryColumnNames.push(SERVER_FIELDS[serverColumnName]);
  }
  return mandatoryColumnNames;
}

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
 *   1. Changing all the letters to lowercase.  E.G. " First name - given_or_nickName  " -> " first name - given_or_nickname  "
 *   2. Removing any spaces, underscores, dashes.  E.G. " first name - given_or_nickname  " -> "firstnamegivenornickname"
 * You can check the regular expression inside replace() in https://regexr.com/7chdp
 * 
 * The special reserved values which start with 'server_' are treated differently and allowed to keep their underscores (and other non-letter characters)
 */
function normalize(key) {
  if (typeof key != 'string') {
    return undefined;
  }
  let ñkey = key;
  ñkey = ñkey.toLowerCase();
  if (ñkey.indexOf('server_') == 0) {
    // Any key that starts with 'server_' (case insensitive) will only be lowercased for normalization.
    return ñkey;
  }
  ñkey = ñkey.replace(/[\s-=;:'"`~@!#$%^&*()_+{}\[\]<>?,.|\/\\]/g, '');
  return ñkey;
}

/*
 * If the processing of a request on this server wants to communicate anything, it can add it via a message.
 * If any server_messages exist by the time processing is ready to write the row, they will go under a column named 'server_messages'.
 */
function addServerMessage(columnName2value, message) {
  if (!(SERVER_FIELDS.messages in columnName2value)) {
    columnName2value[SERVER_FIELDS.messages] = message;
  } else {
    columnName2value[SERVER_FIELDS.messages] += '\n' + message;
  }
}

/*
 * The standard time for computers to start counting is midnight Jan 1 1970.  This is called "the Epoch".
 * The current time depends on the timezone, but the number of seconds (or μseconds) since the epoch is the same globally if we only consider UTC.
 * Javascript only provides millisecond precision, but that's no reason to avoid working in microseconds.
 */
function getEpochUTCUs() {
  // Reference: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/now
  return Date.now() * 1000;
}

/*
fetch("https://script.google.com/macros/s/AKfycbxkG5hM6MMswwHdzWSJKwutMYsOZRT3zjC7jFti0sDvJ47bWB4BTsHPhvbyEVGSsSc5/exec", {
    method: 'POST',
    body: data,
    headers: {
        'Content-Type': 'text/plain;charset=utf-8',
    }
}).then(response => {
    console.log("success:", response);
}).catch(err => {
    console.log("Error:" + err);
})

fetch(URL, {
      redirect: "follow",
      method: "POST",
      body: JSON.stringify(DATA),
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
    })
*/

/*
 * The way a POST request is handled, the data passed in is stored in e.postData.contents.
 * e.postData.type indicates the mime-type.  Only text/json is supported here.
 * It must be parsed as JSON to create a JavaScript object that maps each key (column name) to the desired value. 
 * 
 * Reference: https://developers.google.com/apps-script/guides/web
 */
function doPost(e) {
  // JSON.parse does some simple type conversions such as converting '42' to the number 42.
  let jsonData = JSON.parse(e.postData.contents);
  // We want all the keys and values passed in to be strings so they can be treated the same.
  let stringKey2stringValue = {};
  for (let key in jsonData) {
    let stringKey = undefined;
    if (typeof key == typeof "some string") {
      stringKey = key;
    } else {
      stringKey = JSON.stringify(key);
    }

    let stringValue = undefined;
    let value = jsonData[key];
    if (typeof value == typeof "some string") {
      stringValue = value;
    } else {
      stringValue = JSON.stringify(value);
    }
    stringKey2stringValue[stringKey] = stringValue;
  }
  return handleRequest(stringKey2stringValue);
}

function testPost() {
  let datas = [
    {
     "value" : "99 luftballoons",
     "server_sheetname": "data"
    },
    {
     "server_v" : 2,
     "value" : "99 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 2,
     "value" : "42 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 3,
     "value" : "41 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 3,
     "value" : "41444 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 1000000,
     "value" : "40 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 1 << 20 - 1,
     "value" : "39 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 1 << 20,
     "value" : "38 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 1 << 20 + 1,
     "value" : "37 balloons",
     "server_sheetname": "data"
    },
    {
     "server_id": 234,
     "server_v" : 88,
     "value" : "36 balloons",
     "server_sheetname": "data"
    },  
   ];
   for (let i = 3; i < datas.length; i++) {
     let e = {postData: {contents: JSON.stringify(datas[i])}};
     doPost(e);
     let ii = 42;
   }
}

/*
 * THIS IS THE MAIN FUNCTION.  It is the starting point that gets called whenever
 * someone sends a request to the Web App URL (from step 5 of the Basic Instructions).
 * The values that this function uses are the [query parameters](https://shorturl.at/lvwGU)
 * For example:
 * 
 * If the url is https://script.google.com/macros/e/d/ABcd12_34/exec
 * then query parameters are added by making the url something like this:
 * https://script.google.com/macros/e/d/ABcd12_34/exec?firstName=Snow&lastName=White&numFriends=7
 */

/*
 * Not Recommended: Sending requests to this server via GET is fine for testing, but it is not recommended in an application/browser.
 * The main reason it's not recommended is that GET requests might get cached on the client, and will not actually call out the server if the same request is sent a second time.
 *
 * The way a GET request is handled, the data passed in is stored in e.parameter as a string.
 * Each key (column name) maps to the desired value.

 * These key=value pairs are obtained from the queryString in the request. For example:
 * 
 * If the url is https://script.google.com/macros/e/d/ABcd12_34/exec
 * then query parameters are added by making the url something like this:
 * https://script.google.com/macros/e/d/ABcd12_34/exec?firstName=Snow&lastName=White&numFriends=7
 *
 * Reference: https://developers.google.com/apps-script/guides/web
 */
function doGet(e) {
  let key2value = e.parameter;
  return handleRequest(key2value);
}

function handleRequest(parameterObject) {
  if (parameterObject[SERVER_FIELDS.force_error]) {
    throw new Error("Error requested by server_force_error.");
  }

  // columnName2value is a map from column name to the value that should be inserted into that column.
  // Though it is named in the singular (i.e. not columnNames2values), it contains multiple pairings of {columnName: value}.
  let columnName2value = sanitizeQueryParameters(parameterObject);

  // This script is only allowed to view/edit the spreadsheet to which it is attached.
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
  // Get the sheet that has the sheet_name specified in the request.
  // If there is no 'server_sheetname' or 'sheet_name', then sheetName will be undefined, and that's OK.
  let sheetName = columnName2value[SERVER_FIELDS.sheetname];
  // See comments in chooseSheet for fallbacks if no sheet by that name exists.
  let sheet = chooseSheet(spreadsheet, sheetName);

  let ii1 = sheet.getLastColumn();
  let ii2 =sheet.getLastRow();
  let ii3 = sheet.getName();
  let ii = 42;

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
  
  // Add any missing columns that the request recorded in column2values.
  // Returns a map of the sheet's header, from column name to column index.
  appendValuesToSheet(columnName2value, sheet);

  return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
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

  let mandatoryColumnNames = getMandatoryColumnNames();

  // If none of the columnNames match the headerNames, the final spreadsheet will have all the orginal headers plus all the new columnNames.
  let maxHeaderNames = numHeaderNames + numColumnNames + mandatoryColumnNames.length;
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
  for (let columnIndex = 0; columnIndex < headerNames.length; columnIndex++) {
    let headerName = headerNames[columnIndex];
    let ñheaderName = normalize(headerName);

    // Sometimes, after normalization, there is nothing of substance left.  E.G. " -- "
    if (ñheaderName.length == 0) {
      continue;
    }
    // Some of the headerNames might normalize to the same value.  In that case, the last one wins.
    ñheaderName2columnIndex[ñheaderName] = columnIndex;
  }

  // Optional: See detailed discussion about stray values in the spreadsheet for corner-case details.
  
  // Now it's time to add more ñcolumnName -> columnIndex pairings to ñheaderName2columnIndex,
  // for the new column names that are not in the spreadsheet's header names.

  // Create an array of all the column names needed for this row.
  let columnNames = Object.keys(columnName2value);
  // Also add the mandatory column names.  It's OK if some of them are already in the columnNames. The
  // code below handles the corner-case where column names are duplicated.
  columnNames = columnNames.concat(mandatoryColumnNames);

  // When we find a novel column name, it will go at nextColumnIndex.
  let nextColumnIndex = numHeaderNames;
  // Go over each column name and check if it matches an existing header name.
  // If it is novel, add it to the header names, and update the ñheaderName2columnIndex mapping.

  for (let columnName of columnNames) {
    let ñcolumnName = normalize(columnName);
    if (ñcolumnName in ñheaderName2columnIndex) {
      // That column name already matches an existing header name.  Continue to the next column name.
      // Corner-case: If a column name is repeated, this will also find it already exists.
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
 * Creates a simple 1D array with a value for each headerName (from ñheaderName2columnIndex).
 * Any value that appears in columnName2value will be used.  Otherwise, it is set to undefined.
 * The order of the values in the array is governed by the mapping of ñheaderName2columnIndex.
 * The row probably won't be dense, meaning some of the values will be undefined.
 */
function createProposedRow(ñheaderName2columnIndex, columnName2value) {
  let proposedRow = [];

  // This ensures that the proposedRow will have a cell for every possible column index associated with a header name.
  for (let ñheaderName in ñheaderName2columnIndex) {
    let columnIndex = ñheaderName2columnIndex[ñheaderName];
    proposedRow[columnIndex] = undefined;
  }

  // For each value, find the index for its associated columnName by looking at the mapping ñheaderName2columnIndex.
  for (columnName in columnName2value) {
    // This is the value that will be stored in the spreadsheet.
    let value = columnName2value[columnName];
    let ñcolumnName = normalize(columnName);
    let columnIndex = ñheaderName2columnIndex[ñcolumnName];

    // Do a sanity-check and capture an errorMessage in case the lookup for columnIndex failed.
    if (columnIndex === undefined) {  // JS: indexing into a non-existent key gives undefined
      // This should not happen because all the column names should already exist in ñheaderName2columnIndex, thanks to addMissingColumns()
      addServerMessage(columnName2value, `ERROR: ${ñcolumnName} not found for value = ${value}`);
      continue;
    }

    // Detail: This is specific to JavaScript, where array items can be referenced without allocating them.
    // Put the value at the correct column index.
    proposedRow[columnIndex] = value;
  }

  // Add in the server-based timestamps as μseconds since epoch in UTC.
  let server_now_us = getEpochUTCUs();
  proposedRow[ñheaderName2columnIndex[SERVER_FIELDS.created_us]] = server_now_us;
  proposedRow[ñheaderName2columnIndex[SERVER_FIELDS.modified_us]]= server_now_us;

  return proposedRow;
}

/*
 * Look through the id of every row, and try to find one that matches the given id.
 * If one is found, return that row's index.
 * Otherwise, return undefined.
 */
function findRowIndexById(sheet, ñheaderName2ColumnIndex, id) {

  let ii1 = sheet.getLastColumn();
  let ii2 =sheet.getLastRow();
  let ii3 = sheet.getName();
  let ii = 42;



  let idIndex = ñheaderName2ColumnIndex[SERVER_FIELDS.id];
  // Get the column of ids.
  let idRange = sheet.getRange(1,  idIndex + 1, sheet.getLastRow(), 1);
  // This values array is a Nx1 2D array.
  let idValues = idRange.getValues();
  let rowIndex = undefined;
  // We will search starting at the last row, and work our way up to the 0th row.
  // This is an optimization that assumes the row we are looking for is more likely to be found near the bottom.
  for (let r = idValues.length - 1; r >= 0; r--) {
    let idValue = idValues[r][0];
    if (idValue == id) {
      rowIndex = r;
      break;
    }
  }
  return rowIndex;  // might be undefined;
}

/*
 * A parsing with some rules:
 *  1. must already be a number, or parse as a number
 *  2. the number must be an int
 *  3. the value must be >= minValue
 *  4. the value must be <= maxValue
 *
 *  Returns undefined to indicate parsing failed.
 */
function parseIntWithLimits(rawNumber, minValue, maxValue) {
  let parsed = parseInt(String(rawNumber)); 
  // If rules 1 or 2 fail, this is NaN
  if (isNaN(parsed)) {
    return undefined;
  }
  if (!(minValue <= parsed && parsed <= maxValue)) {
    return undefined;
  }
  return parsed;
}

/*
 * This function could benefit from some rewriting to make it more clear.
 * Primarily, rather than working in the  space of arrays (which require non-trivial indexing),
 * all of the comparisons and calculations should be on JS objects, which are more readable.
 * 
 * I think it would be cleaner to try and find the existing row before the proposedRow is created.
 */
function possiblyMergeRowUpdate(sheet, rowIndex, ñheaderName2columnIndex, proposedRow) {
  let rowRange = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn());
  let rowValues = rowRange.getValues();  // A 1xN 2D array.
  // Get a 1D array of the existing row.
  let existingRow = rowValues[0];

  let idIndex = ñheaderName2columnIndex[SERVER_FIELDS.id];
  let vIndex = ñheaderName2columnIndex[SERVER_FIELDS.v];

  if (proposedRow[idIndex] != existingRow[idIndex]) {
    return false;  // This shouldn't happen, but if the ids don't match, this update shouldn't happen.
  }
  
    // Handle rules for server_v:
    // server_v must be an integer in the range [0, 2^20).  Anything outside that range is an indication that server_v is not being used as a version
    // for the record, and therefore will be ignored.  If the user placed useful data in server_v, it will be lost.
    //  1. If existing_v does not exist or is not an integer or is outside [0, 2^20), treat it as undefined.
    //  2. If proposed_v does not exist or is not an integer or is outside [0, 2^20], treat it as undefined.
    //  3. Compare proposed_v to existing_v  (Note that rejecting the update is not an error.  It is correct, well-defined behavior.)
    //     - If neither is defined, accept the update.
    //     - If existing_v is undefined and proposed_v is defined, accept the update.
    //     - If existing_v is defined and proposed_v is undefined, reject the update.
    //     - If both are defined, only accept the update if proposed_v > existing_v.

  let existing_v = existingRow[vIndex];
  let proposed_v = proposedRow[vIndex];
  let ñexisting_v = parseIntWithLimits(existing_v, 0, 1<<20 - 1);
  let ñproposed_v = parseIntWithLimits(proposed_v, 0, 1<<20);
  if (ñexisting_v != undefined && ñproposed_v == undefined) {
    return false;  // Can't allow a non-versioned update overwrite a properly versioned row that already exists.
  }
  if (ñexisting_v != undefined && ñproposed_v <= ñexisting_v) {
    return false;  // Both version exist, but the proposed one is not strictly greater.
  }

  // The only value that is retained is server_created_us.
  let createdµsIndex = ñheaderName2columnIndex[SERVER_FIELDS.created_us];
  proposedRow[createdµsIndex] = existingRow[createdµsIndex];

  rowValues[0] = proposedRow;
  // This actually sends the data to the spreadsheet and updates the row.
  rowRange.setValues(rowValues);
}


/*
 * Makes sure the column names all exist (creating new columns with header names as needed).
 * Appends the values to the end of the sheet, respecting the order of the columns already in the sheet.
 */
function appendValuesToSheet(columnName2value, sheet) {
  // Add any novel columns that may be in columnName but aren't in the sheet's header row yet.
  let ñheaderName2columnIndex = addNovelColumns(sheet, columnName2value);

  let ii1 = sheet.getLastColumn();
  let ii2 =sheet.getLastRow();
  let ii3 = sheet.getName();
  let ii = 42;


  // Create the row with all the values in the right order to match the order of the sheet's header names.
  let proposedRow = createProposedRow(ñheaderName2columnIndex, columnName2value);

  // Determine if we are going to insert a new row, or update an existing one
  let rowIndex = findRowIndexById(sheet, ñheaderName2columnIndex, columnName2value[SERVER_FIELDS.id]);
  if (rowIndex != undefined) {
    possiblyMergeRowUpdate(sheet, rowIndex, ñheaderName2columnIndex, proposedRow);
    return;
  }

  // appendRow() is safe to call even if there aren't enough columns in the sheet.  (They will be added as needed.)
  // Note that server_version might be missing, or might be some invalid value.
  // That's OK.  If this row ever gets an update request, the missing/bad value will be dealt with then.
  sheet.appendRow(proposedRow);
}


/*
 * Generates a random id out of 2^64 possible ids.
 * Uses the characters from plus-codes so the id will not form unpleasant words by accident.
 */
function createRandomId() {
  // There are 20 plus-code characters, so each one takes about 4.321 bits to encode.
  let plusCodeChars = '23456789CFGHJMPQRVWX';
  let id = '';
  for (let c = 0; c < Math.floor(64/4.321); c++) {
    id += plusCodeChars[Math.floor(20*Math.random())];
  }
  return id;
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
    if (queryParameters[SERVER_FIELDS.secret] = secret()) {
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
    // This check is done first because even if the key is 'server_sheetname', we don't want to use such a long value.
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

    let ñkey = normalize(key);

    // We don't want to store the secret in the spreadsheet.
    if (ñkey.indexOf('secret') >= 0) {
      continue;
    }
    // For that matter, let's not store anything with 'password' in the name 
    if (ñkey.indexOf('password') >= 0) {
      continue;
    }

    // The server_ keys are important enough that we look for them past the max number of columns allowed.
    if (ñkey.indexOf('server_') == 0) {
      // Other column names are left unnormalized, but the server_ column names use the normalized version even in the final spreadsheet header row.
      columnName2value[ñkey] = queryParameters[key];
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
      addServerMessage(columnName2value, `{error}={numErrors}`);
    }
  }

  if (!('server_id' in columnName2value)) {
    addServerMessage(columnName2value, 'Creating an id on the server.');
    columnName2value['server_id'] = createRandomId();
    if ('server_v' in columnName2value) {
      addServerMessage(columnName2value, 'Overriding v along with id creation.');
    }
    // Even if there was a 'server_v', it doesn't matter -- creating a new id also means setting the v=0
    columnName2value['server_v'] = 0;
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
