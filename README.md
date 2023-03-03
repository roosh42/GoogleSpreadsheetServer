# GAS_email_processing
Google Apps Script code for email processing.

Include this in your Google Apps Script so you can run processUnstarredEmails on a time-based trigger.

This code will run as "you" on whatever Google/Gmail account you use it.

[Here is an explanation](https://docs.google.com/presentation/d/e/2PACX-1vRgkfjkMLKcrTka9Jsk3Ww2_YfuOut6_MleS30O4wRR79a5RgYpSBC1yaiO9w3ebIebkeIdnlT1wAgp/pub?start=true&loop=false&delayms=10000) of the permissions you will need to Allow.

# It's your code, running as you.

## To see and revoke access to scripts in your Google Account
https://myaccount.google.com/permissions?continue=https%3A%2F%2Fmyaccount.google.com%2Fsecurity%3Fpli%3D1%26nlr%3D1

## To see and delete the triggers (time-based and otherwise) in your Google Account
https://script.google.com/home/triggers

## To see the code running when you open the spreadsheet
In the spreadsheet's menu, Extensions > Apps Script
  - Take a look that under `Files`, the only one listed is `Code.gs`
  - Skim through the file.  This is the code that is generating the `Email Processing` menu.
