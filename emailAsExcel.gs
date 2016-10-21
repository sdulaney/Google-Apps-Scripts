// Source: http://stackoverflow.com/questions/31809987/google-app-scripts-email-a-spreadsheet-as-excel
// Requirements for each GAS project: 1. Resources -> Advanced Google Services... -> Drive API v2 -> on, 2. Enable Google Drive API in Google Developers Console for the project as well
// Documentation: https://developers.google.com/apps-script/reference/gmail/gmail-app

/**
 * Thanks to a few answers that helped me build this script
 * Explaining the Advanced Drive Service must be enabled: http://stackoverflow.com/a/27281729/1385429
 * Explaining how to convert to a blob: http://ctrlq.org/code/20009-convert-google-documents
 * Explaining how to convert to zip and to send the email: http://ctrlq.org/code/19869-email-google-spreadsheets-pdf
 */
function emailAsExcel(config) {
  if (!config || !config.to || !config.subject || !config.body) {
    throw new Error('Configure "to", "subject" and "body" in an object as the first parameter');
  }

  var spreadsheet   = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId()
  var file          = Drive.Files.get(spreadsheetId);
  var url           = file.exportLinks[MimeType.MICROSOFT_EXCEL];
  var token         = ScriptApp.getOAuthToken();
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });

  var fileName = (config.fileName || spreadsheet.getName()) + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];
  if (config.zip) {
    blobs = [Utilities.zip(blobs).setName(fileName + '.zip')];
  }

  GmailApp.sendEmail(
    config.to,
    config.subject,
    config.body,
    {
      attachments: blobs
    }
  );
}

// Example usage
// function sendDailyCallLog() {
//   emailAsExcel( { to:"svita@concord-re.com,lbanner@concord-re.com,sdulaney@concord-re.com", subject:"Daily Call Report", body:"This spreadsheet was sent from Google Sheets." } );
// }
