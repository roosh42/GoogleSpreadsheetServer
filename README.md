## 💻+📈 Add a Web App Server to any Google spreadsheet.

The server you deploy has a public URL that can be used to insert new rows into any of the sheets in your chosen spreadsheet.

## It's your code, running as you
  - [See and revoke access](https://myaccount.google.com/permissions?continue=https%3A%2F%2Fmyaccount.google.com%2Fsecurity%3Fpli%3D1%26nlr%3D1) to scripts in your Google Account.
  - [See and delete the triggers](https://script.google.com/home/triggers) (time-based and otherwise) in your Google Account.
  - The server only has permission to operate on the one spreadsheet that it is attached to.

## 📜 Terminology
| description                                                   | Google equivalent             | Excel equivalent               | Numbers equivalent            |
|---------------------------------------------------------------|-------------------------------|--------------------------------|-------------------------------|
| An individual tab within the spreadsheet file.                | sheet                         | worksheet                      | sheet                         |
| A file in Google Sheets. Can contain many sheets.             | spreadsheet (contains sheets) | workbook (contains worksheets) | spreadsheet (contains sheets) |
| The program/service that Google provides for grid-based data. | Google Sheets                 | Microsoft Excel                | Apple Numbers                 |

## 😀 Basic Instructions for *Spreadsheet with Server* (Desktop version)

1. Make your own copy of the spreadsheet with one of your Google accounts.
   1. Log into the Google account that will own your spreadsheet.
   1. Open [Spreadsheet with Server (template)](https://docs.google.com/spreadsheets/d/1kU2IiLpKKVM_Zb3BzlB_b3I9ww1Rio81olDnzu6avWg/).
   1. `File` > `Make a Copy`
   1. Click the `Make a Copy` button.
2. In your new copy, open Apps Script with `Extensions` > `Apps Script`
3. Deploy the Apps Script as a Web App.
   1. `Deploy` *in the top-right corner* > `New Deployments`
   1. `Select Type` > `Web App`
   1. Change "Who has access" to `Anyone` (NOT "Anyone with Google Account")
   1. Click `Deploy`
4. Grant permissions (for the script to modify only the one spreadsheet it is attached to).
   1. `Authorize Access`
   1. Select the Google account that owns your spreadsheet.
   1. Click `Allow` for "View and manage spreadsheets that this applicaion has been installed in".
5. Save the Web App URL.  
   1. `Copy` the Web App URL. (It starts with https://script.google.com/macros) 
   1. Save the Web App URL (e.g. Email it yourself or paste it in a document.)

You can now close the Apps Script tab.

**🎉Congratulations🎉** You now have a way to programatically insert rows to your spreadsheet by requesting the URL from above.
  
Anyone who has that link will be able to add rows with any content they put in the url.

Note: The script runs as ***you*** -- Not as the person/program who sends the url request.
That is why you do not need to share the spreadsheet with anyone -- You already have permission to read&write the spreadsheet.

## 🎛 Advanced Instructions

### 🐧 Open Source instructions (for programmers)
If you don't want to copy the template spreadsheet with its code already included, you can get the code yourself.

1. Create a new Google spreadsheet.
2. Open the attached Apps Script by going to `Extensions` > `Apps Script`.
3. Replace the code in `Code.gs` with the [permalinked code for version=202303260837](https://raw.githubusercontent.com/anerb/GoogleSpreadsheetServer/5a19d8dfd050db4d2158c224dca5de91edffaff9/SpreadsheetServer.gs)
  - This code is frozen in time.  If you or anyone else has reviewed it, Github ensures that it will be unchanged via the permalink.
4. Proceed with the Basic Instructions, starting at step 3.

### 🖧 Using the Web App URL
Once you have the URL of a deployed Web App, it will looks something like
  - https://script.google.com/macros/s/AKfycbwIeR6hGK_NgF22d896q............................XdSnZX41Ew/exec
If you lost the URL, go to `Deploy` > `Manage deployments` and `Copy` the URL under `Web App`.

This Web App uses [query parameters](https://shorturl.at/lvwGU) to pass in the information for a new row in the spreadsheet.

The only special parameter keys are
 - spreadsheetid: Only used when a demo server is set up to change anyone-can-edit Spreadsheets.
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

### ∑ Enabling formulas for inserted spreadsheet rows.
If you want to be able to send formulas to your spreadsheet through the Web App URL, follow these steps.
1. Pick a secret.
  a. Scroll a few lines down and find `function secret()`.
  b. Change the word in single-quotes to a secret
    - This is not an account password.
    - Minimum 6 characters long.
    - Case sensitive.
2. Follow the Basic Instructions above, starting at Step 3, to create a `New Deployment`.
  - You may not need to re-authorize, so step 4 will be skipped.

## ↖ Using the collected data outside Google Sheets
To use the data outside of Google Sheets, you can publish the sheet as an online CSV source and consume it in your favorite program.
  - For working in Microsoft Excel, here is a good [YouTube video by Marc Ursell](https://www.youtube.com/watch?v=vAdJrUIhS8o).
  - For working in Apple Numbers, I couldn't find an off-the-shelf solution, but AppleScript is powerful enough to implement something like [this Apple community post](https://discussions.apple.com/thread/8126136)

## 🔐 Security matters

When the row contents is sent over the internet to the Web App server (either from a browser or an app/website), it is encrypted.
That's what the `https` at the strt of the URL means.
Once they arrive, they are decrypted so the function can deal with the plain text that was in the query parameters.

Only people you give access to will be able to see/edit the spreadsheet.
However, anyone with the Web App URL can insert new rows into the spreadsheet.

