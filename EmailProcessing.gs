/******************************************************************************
 *             QUICK INSTRUCTIONS                                             *
 * 0. Have a Google/Gmail account.  (Maybe create a new one just for this).   *
 * 1. Visit <insert link> this Spreadsheet with your Google account signed in *
 * 2. File > Make a Copy                                                      *
 ******************************************************************************/

/*
 * Allow anyone to send an email that will end up as a row in a spreadsheet.
 * 
 * Email is used as the deliver system because it is a robust queue that will eventually deliver.
 *   - Email can be processed in a device without internet connectivity, and it will be sent later.
 *   - Email can be recieved by robust systems (e.g. Google's Gmail servers) and stored for later processing.
 *   - In short, the world has provided us with email as a reliable pub-sub system.
 * 
 * Email is secure
 *   - By carrying the communication from the device in an email, an existing trusted channel is used.
 *   - The app on the device does not need to send out any data on the network.
 *     - Once the app is cached, it does not need any network connectivity.
 * 
 * The body of the email is JSON (with or without the surrounding {}).
 *   - spreadsheet_id: indicates which spreadsheet to write to
 *   - sheet_name: indicates which sheet to write to.
 *   - all other values are treated as column headings.
 *     - column headings are NOT case sensitive.  So a JSON property 'favorite color' will go into a column with header 'Favorite Color'.
 *     - if a column does not exist, it will be appended as the new final column.
 *
 * Each email appends a row after the final row in the sheet.
 *   - errors are also appended as a row, but only the special columns error_message email_text get filled in.
 *   - if the error is too severe to extact the spreadsheet_id, a hardcoded spreadsheet_id is used.
 *   - if the error is too severe to extact the sheet_name (but spreadsheet_id is determined), a special sheet named ERROR_sheet is used.
 * 
 * This processing owns the is:starred attribute of the email, and sets it after processing an email message (even with errors).
 * I would have preferred to use a special label, but labels are for threads, not messages.
 *
 * Hierarchy of invented protocol
 *  - The "to" field is used only to deliver the message.
 *    - A suffix such as +email2sheet or -email2sheet can be used.  E.G. johndoe+email2sheet@example.com
 *  - The "cc" and "bcc" fields are ignored.
 *  - The "subject" is used as a comment field and is only for human consumption.  It is recorded, but not parsed.
 *  - The "body" contains all of the data, and cannot contain any comments or other info because the whole thing is parsed as JSON.
 *    - Automated signature (e.g. "sent from my iPhone") need to be turned off for the body to be parsed correctly.
 */

/*
 * Philosopy of assembly-line vs. staged batches.
 * 
 * TL;DR: Assembly-line is the default unless there's a good reason to avoid it.
 * 
 * An assembly line with 1 stage is equivalent to Staged batches.
 * Staged batches with batchsize = 1 is equivalent to assembly line.
 * 
 * The simplest model is assembly line, where each item is processed through all stages without pausing.
 * Benefits for batching:
 * 1. Shallower code.  This can make it easier to debug and test.
 * 2. Efficiency in processing multiple items against pre-existing batched interfaces.
 *   2.5. Less overhead for batching once a connection or other resource is acquired.
 * 3. If there is filtering, the batches have the filtering logic already baked in.
 * 
 * If efficiency isn't a major consideration, assembly-line seems more intuitive and has little downside.
 */

var JSON_PARSER_FN_ = JSON.parse;
var DEFAULT_SPREADSHEET_ID_ = undefined;
var DEFAULT_SHEET_NAME_ = undefined;
var MAX_MESSAGES_TO_PROCESS_ = 500;
// This gets reset to MAX_MESSAGES_TO_PROCESS at the start of every run.
var kMessageProcessingQuotaRemaining_ = 0;

// If the email was composed using HTML, this is a heuristic to remove all the HTML tags.
// Note that every call to .replace() replaces some text with " ", so this only removes information.
function removeTags_(html) {
  let html_no_head = html.replace(/<head>.*<\/head>/, " ");
  let html_no_style = html_no_head.replace(/<style>.*<\/style>/, " ");
  let html_no_tags = html_no_style.replace(/[<][^>]+[>]/g, " ");
  return html_no_tags;
}

// I don't remember why, but decodeURIComponent wasn't enough, and I had to add the decoding of space and quote.
function decodeText_(text) {
  let cleaner = text;
  let sequences = {" ": /&nbsp;/g, '"': /&quot;/g};                 
  for (let sequence in sequences) {
    cleaner = cleaner.replace(sequences[sequence], sequence);
  }
  return cleaner;
}

// Warning: fragile chaining
function cleanBody_(body) {
  let plain_body = removeTags_(body);
  let clean_text = decodeURIComponent(plain_body);
  let decoded_text = decodeText_(clean_text);
  // condense whitespace (after &nbsp; has been converted).  Replace multiple whitespace with a single space.
  let single_whitespace = decoded_text.replace(/\s[\s]*/g, " ");
  return single_whitespace;
}

// If the text isn't surrounded by {}, add the braces around it.
function jsonBraces_(text) {
  text = text.trim();
  text = text.replace(/^([^{].*[^}])$/, '{$1}');
  return text;
}

// Parse the text as JSON (or JSON5) into a JS object.
// JSON_PARSER_FN_ is the builtin JSON.parse, or the JSON5.parse if use_json5 == true.
function parseJSON_(text) {
  let object = {};
  let text_with_braces = jsonBraces_(text);
  try {
      object = JSON_PARSER_FN_(text_with_braces);
  } catch (e){
    // FRAGILE: other parts of the code rely on these constants, and even the prefix.
    object = {"error_message": e.message};
  }
  return object;
}

/*
 * TODO: use some shared name of an image (or other file) to "get an attachement" indirectly.
 * Or send the image as base64 encoded in the body.
 * 
 * Returns the item with keys normalized according to normalizeKey.
 */
function getMessageItem_(message) {
  let metadata = {};
  metadata.email_date = message.getDate();
  metadata.email_from = message.getFrom();
  metadata.email_subject = message.getSubject();
  metadata.email_permalink = message.getThread().getPermalink();
  let body = message.getPlainBody();
  metadata.email_text = body;

  let data = {};
  try {
    let text = cleanBody_(body);
    metadata.email_text = text;
    data = parseJSON_(text);
  } catch (e) {
    data = {"error_message": e.message};
  }
  // NOTE: data will overwrite metadata, if it contains one the email_.* properties.
  let item = {...metadata, ...data};
  return item;
}

// Get a reference to the sheet that will get the new row.
// In this function, I have to wrap many calls with try{} because they don't simply return undefined or null -- they throw. :(
function getSheet_(spreadsheet_id, sheet_name) {

  let spreadsheet = undefined;

  if (!spreadsheet) {
    if (spreadsheet_id !== undefined) {
      try {
        spreadsheet = SpreadsheetApp.openById(spreadsheet_id);
      } catch(e) {Logger.log(e);}
    }
  }

  if (!spreadsheet) {  // SLOPPY: JS implicit boolean conversion should be avoided.
    if (DEFAULT_SPREADSHEET_ID_ !== undefined) {
      try {
        spreadsheet = SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID_);
      } catch(e) {Logger.log(e);}
    }
  }
  
  if (!spreadsheet) {
    // We have a real problem because this email can't be parsed.
    let final_fallback_spreadsheet_name = "Email Processing Library ERROR sheet for " + ScriptApp.getScriptId();
    // Cannot reuse a previous fallback spreadsheet because spreadsheets can only be opened by url or id.
    spreadsheet = SpreadsheetApp.create(final_fallback_spreadsheet_name);

    // For the rest of this run, we'll reuse the same newly-created spreadsheet.
    DEFAULT_SPREADSHEET_ID_ = spreadsheet.getId();
  }

  let sheet = undefined;

  if (!sheet) {  // this is always true, but it helps keep the code pattern consistent.
    if (sheet_name !== undefined) {
      try {
        sheet = spreadsheet.getSheetByName(sheet_name);
      } catch(e) {Logger.log(e);}
    }
  }
  if (!sheet) {  // SLOPPY: This library returns null, but undefined would be more JS-like.
    if (DEFAULT_SHEET_NAME_ !== undefined) {
      try {
        sheet = spreadsheet.getSheetByName(DEFAULT_SHEET_NAME_);
      } catch(e) {Logger.log(e);}
    }
  }
  if (!sheet) {  // If all else fails, we'll use the first sheet
     sheet = spreadsheet.getSheets()[0];
  }

  return sheet;
}


/*
 * Extract just the column names from the item, and omit any "special" properties.
 */

// FRAGILE: Using strings as keys that are assumed the same throughout the code.
function reduceItem_(item) {
  let reduced_item = {};
  if ("error_message" in item) {
    reduced_item = {"error_message": item["error_message"],
                   "email_text": item["email_text"]
                  };
    return reduced_item;
  }

  let blocked_keys = ['spreadsheet_id', 'sheet'];
  for (let key in item) {
    if (blocked_keys.includes(key)) {
      continue;
    }
    let normalized_key = normalizeKey_(key);
    if (normalized_key in reduced_item) {
      reduced_item = {"error_message": `Two keys normalize as the same ${normalized_key}.`,
                      "email_text": item["email_text"],
                     };
      return reduced_item;
    }
    reduced_item[normalized_key] = item[key];
  }
  return reduced_item;
}

// LESSON: For myself: Name variables with the none last, so they can be pluralized.  (I.E. key_normalized doesn't pluralize well.)
function normalizeKey_(key) {
  let normalized_key = key.trim().toLowerCase();
  return normalized_key;
}

/*
 * Adds any column_names that aren't already in the sheet's header row.
 * Comparison is based on the normalizedKey() version of both the column_names and the header_values.
 * If there is a match, the existing header_value will be kept, and the incoming column_name is assumed to have a capitalization mistake.
 */
function addMissingColumns_(sheet, reduced_item) {
  let num_valued_columns = sheet.getLastColumn();
  let num_keys = 0;
  for (let key in reduced_item) {
    num_keys++;
  }
  let max_num_columns = num_valued_columns + num_keys;
  let num_rows = 1;
  let header = sheet.getRange(1, 1, num_rows, max_num_columns);
  let header_values = header.getValues();

  // Note: Originally, I had an alias variable for header_values[0], but the by-reference/by-value assignment wasn't obvious, so I took it out.
  // header_map is the return value of this function but it could theoretically be built up after any new columns have been added.
  // However, since we're key-ing off of normalized keys, it's helpful to already have the normalizedKey() version of each header value in [LABEL:A].
  let header_map = {};

  // There might be multiple empty header values here.
  for (let c = 0; c < header_values[0].length; c++) {
    let header_value = header_values[0][c];
    let header_value_normalized = normalizeKey_(header_value);
    header_map[header_value_normalized] = c;
  }

  // If the sheet has some stray values in columns that have no header value (past the last labeled column), this will add columns after that.
  let next_column = num_valued_columns;
  for (let column_name in reduced_item) {
    if (column_name in header_map) {  // LABEL:A
      continue;
    }
    header_values[0][next_column] = column_name;  // HACKY: This uses the JS trick that arrays can use object-like indexing into non-existent cells.
    // Earlier Error, fixed: I didn't update header_map when I updated the header_values.
    header_map[column_name] = next_column;
    next_column++;
  }

  // OPT: This whole if-else-statement is an efficiency optimization and it's correct to always call setValues here.
  if (next_column == num_valued_columns) {
    // nothing to do since we didn't add any new column names
  } else {
    header.setValues(header_values);
  }
  return header_map;
}

/*
 * header_map is a normalized(string)->int map.  The integers aren't necessarily dense.
 */
function createRowToInsert_(header_map, reduced_item) {
  let last_column_index = 0;
  for (const [key, value] of Object.entries(header_map)) {
    last_column_index = Math.max(last_column_index, value);
  }
  let row = [];
  for (let c = 0; c <= last_column_index; c++) {
    row.push("");
  }

  for (column_name in reduced_item) {
    let column_index = header_map[column_name];
    if (column_index === undefined) {  // JS: indexing into a non-existent key gives undefined
      // This is actually an error since all the column_name should already exist in header_map.
      // TODO: Do something useful with this error condition.
      continue;
    }
    // Note: column_index is valid within row because we just built up dense row using max_column_index.
    row[column_index] = reduced_item[column_name];
  }
  return row;
}

function insertRowInSheet_(row, sheet) {
  // ASSUMPTION: the sheet is set up to hold a row of length row.length.
  let range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
  let values = [row];
  range.setValues(values);
}

function processMessage_(message) {
  // Just because the thread passed "-is:starred" doesn't guarantee each message !isStarred().
  if (message.isStarred()) {
    return;
  }
  // I considered having -is:draft in the filter, but I can't be confident that I know what two negative items in the filter mean at the thread vs. message level.
  // Since having a draft to the designated address is rare and only happens for testing reasons, filtering here feels more appropriate.
  // TESTING - this is needed because testing generates drafts.
  if (message.isDraft()) {
    return;
  }

  let item = getMessageItem_(message);
  let sheet = getSheet_(item.spreadsheet_id, item.sheet_name);
  let row_to_insert = [[message.getThread().getPermalink()]];
  try {
    // Business Logic to reduce the item to the keys that will be stored.
    let reduced_item = reduceItem_(item);

    let header_map = addMissingColumns_(sheet, reduced_item);
    row_to_insert = createRowToInsert_(header_map, reduced_item);
  } catch (e) {
    row_to_insert = [[message.getThread().getPermalink()], [e.message]];
  }
  insertRowInSheet_(row_to_insert, sheet);

  // It's really important to try and get here so each run gets through more emails.
  // Otherwise,  a single bad email can stop the whole chain.
  message.star();  // Record that this message has been processed.
}

function processThread_(thread) {
  let messages = thread.getMessages();
  for (let m = 0; m < messages.length && kMessageProcessingQuotaRemaining_ > 0; m++) {
    try {
      processMessage_(messages[m]);
      kMessageProcessingQuotaRemaining_--;
    } catch(e) {Logger.log(e);}
  }
}

/**
 * Process emails that match the filter, and write them to a Google sheet.
 * The body of the email should be parsable as JSON.
 * Each email can specify 'spreadsheet_id' and 'sheet_name' in the body (as keys of the JSON).
 * If those are not specified, or not available, the defaults can be passed in to this function.
 * Hardcoded behavior that only unstarred emails are processed, and a message is starred once it is processed.
 *
 * This function takes a plain object as the only argument.  All values are optional.  The recognized keys are:
 *   - filter: The string filter to use for restricting processing to the emails you care about.
 *   - default_spreadsheet_id: a string with the fallback spreadsheet_id 
 *   - default_sheet_name: a string with th fallback sheet_name
 *   - max_messages_to_process: a number indicating how many messages to process each time this function is called.
 *   - use_json5: do you want to load and use the JSON5 library? 
 */
function processUnstarredEmails(args) {
  // If this was called from a script attached to a spreadsheet, we default to using that spreadsheet as a fallback.
  let active_spreadsheet_id = undefined;
  try {
    let active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (active_spreadsheet) { // SLOPPY
      active_spreadsheet_id = active_spreadsheet.getId();
    }
  } catch(e) {Logger.log(e);}

  let default_args = {
    filter: "",
    default_spreadsheet_id: active_spreadsheet_id,
    default_sheet_name: undefined,
    max_messages_to_process: MAX_MESSAGES_TO_PROCESS_,
    use_json5: true,
  }

  if (args === undefined) {
    args = {};
  }

  // Early exit if this function was called with extra/unknown/mistyped arguments.
  for (let arg in args) {
    if (!(arg in default_args)) {
      throw(`${arg} is not a recognized option in args`);
    }
  }

  // Merge, with args overwriting default_args.
  args = {...default_args, ...args};

  // upgrade three of the args to be accessible "globally"
  DEFAULT_SPREADSHEET_ID_ = args.default_spreadsheet_id;
  DEFAULT_SHEET_NAME_ = args.default_sheet_name;
  MAX_MESSAGES_TO_PROCESS_ = Math.min(args.max_messages_to_process, MAX_MESSAGES_TO_PROCESS_);
  // upgrade another args property to create a global JSON parser
  if (args.use_json5) {
    const json5_url = "https://unpkg.com/json5@2/dist/index.min.js"
    // Creates a global JSON5 object with JSON5.parse().
    eval(UrlFetchApp.fetch(json5_url).getContentText());
    JSON_PARSER_FN_ = JSON5.parse;
  }

  // This will find all matching threads that have at least one unstarred message.
  // Some messaged may be starred.
  // The only args option that didn't get upgraded to a global variable because it is only used here.
  let full_filter = args.filter + " -is:starred";
  let threads = GmailApp.search(full_filter);
  kMessageProcessingQuotaRemaining_ = MAX_MESSAGES_TO_PROCESS_;
  for (let t = 0; t < threads.length && kMessageProcessingQuotaRemaining_ > 0; t++) {
    processThread_(threads[t]);
  }
}
