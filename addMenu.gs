let email_processing_library_permalink = "https://raw.githubusercontent.com/anerb/GAS_email_processing/4933505d3a57884ee11b452dbc6042512c9608f4/EmailProcessing.gs";

function getPlusEmail() {
  let email = Session.getActiveUser().getEmail();
  let token = SpreadsheetApp.getActiveSpreadsheet().getId();
  // ASSUMPTION: This is urlencodedsafe, TODO: Use plus-codes alphabet
  token = token.slice(token.length - 16);  // Keep only the last 16 characters
  let plus_email = email.replace("@", `+${token}@`);
  return plus_email;
}

function allThePermissions() {
  ScriptApp.getScriptTriggers();
  GmailApp.getUserLabels();
  SpreadsheetApp.getActiveSpreadsheet();
  UrlFetchApp.fetch('https://www.wikipedia.org');
}

function triggerExists() {
  try {
    let trigger_id = ScriptProperties.getProperty('trigger_id');
    let triggers = ScriptApp.getScriptTriggers();
    for (let trigger of triggers) {
      if (trigger.getUniqueId() == trigger_id) {
        return true;
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

function permissionsExist() {
  try {
    allThePermissions();
    return true;
  } catch(e) {
    return false;
  }
  return false;
}

function processNow() {
  try {
    let content = UrlFetchApp.fetch(email_processing_library_permalink).getContentText();
    eval(content);
    processUnstarredEmails({
      filter: "to:" + getPlusEmail(),
      max_messages_to_process: 5
    })
  } catch(e) {
    SpreadsheetApp.getUi().alert('processing had an error: ' + e.message);
  }
  SpreadsheetApp.getUi().alert('Processed emails succesfully');
}

function doNothing() {}

function revokePermissions() {
  deactivateProcessing();
  ScriptApp.invalidateAuth();
  addMenu();
}

function activateProcessing() {
  if (triggerExists()) {
    return;
  }
  let trigger_id = ScriptApp.newTrigger('processNow').timeBased().everyMinutes(15).create();
  ScriptProperties.setProperty('trigger_id', trigger_id.getUniqueId());
  addMenu();
}

function deactivateProcessing() {
  if (!triggerExists()) {
    return;
  }
  try {
    let trigger_id = ScriptProperties.getProperty('trigger_id');
    let triggers = ScriptApp.getScriptTriggers();
    for (let trigger of triggers) {
      if (trigger.getUniqueId() == trigger_id) {
        ScriptApp.deleteTrigger(trigger);
        return;
      }
    }
  } catch (e) {
    // pass
  } finally {
    addMenu();
  }

}

function onOpen() {
  // When addMenu is called from withing onOpen, none of the permissions are active.
  addMenu();
}

function showHelp() {
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href=""https://docs.google.com/presentation/d/e/2PACX-1vRgkfjkMLKcrTka9Jsk3Ww2_YfuOut6_MleS30O4wRR79a5RgYpSBC1yaiO9w3ebIebkeIdnlT1wAgp/pub?start=true&loop=false&delayms=10000">A Slideshow explaining the permissions dialog boxs.</a>')
    .setWidth(250)
    .setHeight(300);
SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Processing with Gmail and Google Sheets');
}

// Calling this multiple times will replace the previous menu with the new one.
function addMenu() {
  var ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('Email Processing');

  if (!permissionsExist()) {
    menu.addItem('Check Permissions (required)', 'addMenu');
    menu.addItem('Process emails once now (requires permisssions)', 'processNow');
    menu.addItem('Revoke Permissions', 'revokePermissions');
    menu.addItem('Help', 'showHelp');
  } else if (!triggerExists()) {
    menu.addItem('Activate processing for ' + getPlusEmail(), 'activateProcessing');
    menu.addItem('Process emails once now', 'processNow');
    menu.addItem('Revoke permissions', 'revokePermissions');
    menu.addItem('Help', 'showHelp');
  } else {
    // Both permissions and Trigger are exist
    menu.addItem('Deactiate processing for ' + getPlusEmail(), 'deactivateProcessing');
    menu.addItem('Process emails once now', 'processNow');
    menu.addItem('Revoke permissions (and deactivate)', 'revokePermissions');
    menu.addItem('Help', 'showHelp');
  }
  menu.addToUi();
}
