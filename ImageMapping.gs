function doGet(e)  {
  let file_not_found = DriveApp.getFileById("1IgW1Tam8XtzAA1na717C76EdiywO4Uaz");
  let folder_not_found = DriveApp.getFolderById("1yvK5f9K6nNqLic91yFybHWDmMEWo2edV");

  let parts = e.pathInfo.split('/') || "";
  let folder_id = parts[0];
  let folder = folder_not_found;
  let the_file = file_not_found;
  try {
    folder = DriveApp.getFolderById(folder_id);
    if (folder.getOwner().getEmail() == "email2sheets@gmail.com") {
      // Never serve out of folders owned by this script account.
      folder = folder_not_found;
    }
  } catch(e) {};
  for (let p = 1; p < parts.length; p++) {
    if (p == parts.length-1) {
      // treat it like a file.
      // If multiple files have the same name, return the most recently modified.
      let files = folder.getFilesByName(parts[p]);
      let newest_file = file_not_found;
      while (files.hasNext()) {
        let file = files.next();
        if (file.getLastUpdated() > newest_file.getLastUpdated()) {
          newest_file = file;
        }
      }
      the_file = newest_file;
    } else {
      // Treat it like a folder
      let folders = folder.getFoldersByName(parts[p]);
      let newest_folder = folder_not_found;
      while (folders.hasNext()) {
        let f = folders.next();
        if (f.getLastUpdated() > newest_folder.getLastUpdated()) {
          newest_folder = f;
        }
      }
      folder = f;
    }
  }
  let mimeType = the_file.getMimeType();
  if (! mimeType.match(/image\/.*/)) {
    the_file = file_not_found;
  }
  let url = the_file.getUrl();
  let name = the_file.getName();
  let previewUrl = url.replace(/\/view.*/, "/preview");
  let response = UrlFetchApp.fetch(previewUrl);
  let text = response.getContentText();
  let realUrlMatch = text.match(name + '.{1,20}' + '(https://.*?googleusercontent[^,\\\\]+)');
  let realUrl = realUrlMatch[1];
  let output = ContentService.createTextOutput(realUrl).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}
