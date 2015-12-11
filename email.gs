/* Function 1: creates a Menu when the script loads */

function onOpen() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Adds a menu item with a single drop-down 'Email report'
  activeSpreadsheet.addMenu(
      "Email this report", [{
        name: "Email report", functionName: "emailAsPDF"
      }]);
}

/* Function 2: sends Spreadsheet in an email as a PDF */

// reworked from ctrlq.org/code/19869-email-google-spreadsheets-pdf //

function emailAsPDF() {

  // Send the PDF of the spreadsheet to this email address
  var email = "someone@somewhere.gov.uk,someone@elsewhere.gov.uk";

  // Gets the URL of the currently active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl();
  url = url.replace(/edit$/,'');

  // Subject of email message
  // The date time string can be formatted using Utilities.formatDate method
  var subject = "Metrics mailing - " + Utilities.formatDate(new Date(), "GMT", "dd-MMM-yyyy");

  // Body of email message
  var body = "\n\nHello\n\nThis is a mailing of a Google Sheet.\n \n";

  /* Specify PDF export parameters
  // Taken from: code.google.com/p/google-apps-script-issues/issues/detail?id=3579
    exportFormat = pdf / csv / xls / xlsx
    gridlines = true / false
    printtitle = true (1) / false (0)
    size = A4 / letter /legal
    fzr (repeat frozen rows) = true / false
    portrait = true (1) / false (0)
    fitw (fit to page width) = true (1) / false (0)
    add gid if to export a particular sheet - 0, 1, 2,..
  */

  var url_ext = 'export?exportFormat=pdf' // export as pdf
                + '&format=pdf'           // export as pdf
                + '&size=A4'              // paper size
                + '&portrait=true'        // page orientation
                + '&fitw=true'            // fits width; false for actual size
                + '&sheetnames=false'     // hide optional headers and footers
                + '&printtitle=false'     // hide optional headers and footers
                + '&pagenumbers=false'    // hide page numbers
                + '&gridlines=false'      // hide gridlines
                + '&fzr=false'            // do not repeat row headers
                + '&gid=0';               // the sheet's Id

  var token = ScriptApp.getOAuthToken();

  // Convert worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext)

  //convert the response to a blob
  file = response.getBlob().setName('mailing.pdf');

  // Send the email with the PDF attachment. Google sets limits on the number of emails you can send: https://docs.google.com/macros/dashboard
  if (MailApp.getRemainingDailyQuota() > 0)
     GmailApp.sendEmail(email, subject, body, {attachments:[file]});

}
