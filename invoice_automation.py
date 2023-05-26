function generateAndSendInvoice() {
  var spreadsheetId = '1JsVusJbKU3v7g6OrCPIaThfEOlZvXmkQf4Y8A67eOdM'; // Replace with the ID of your Google Sheet
  var sheetName = 'responses'; // Replace with the name of your sheet containing the form responses
  var templateDocId = '1ifSI1FmP0kTT2wS0dQq2P3I-DiCXlfqaBktiQK9A7Lg'; // Replace with the ID of your Google Docs invoice template
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastProcessedRow = parseInt(scriptProperties.getProperty('lastProcessedRow')) || 0;

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();

  if (lastRow > lastProcessedRow) {
    var dataRange = sheet.getRange(lastProcessedRow + 1, 1, lastRow - lastProcessedRow, sheet.getLastColumn());
    var data = dataRange.getValues();
    var headers = data[0];

    // Start from the second row to skip the header row
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      var invoiceData = {};

      // Map the header values to the corresponding form response values
      for (var j = 0; j < headers.length; j++) {
        invoiceData[headers[j]] = rowData[j];
      }

      // Generate the invoice from the template
      var invoice = generateInvoiceFromTemplate(invoiceData, templateDocId);

      // Send the invoice to the email address
      var recipientEmail = invoiceData['Email'];
      var recipientName = invoiceData['Fullname'];

      if (recipientEmail) {
        sendInvoiceByEmail(invoice, recipientEmail, recipientName);
      } else {
        Logger.log('Email address not found for entry in row ' + (lastProcessedRow + i + 1));
      }
    }

    // Update the last processed row in script properties
    scriptProperties.setProperty('lastProcessedRow', lastRow.toString());
  } else {
    Logger.log('No new entries found.');
  }
}
function generateInvoiceFromTemplate(invoiceData, templateDocId) {
  var templateDoc = DocumentApp.openById(templateDocId);
  var templateBody = templateDoc.getBody();
  var invoiceBody = templateBody.getText();
  
  // Replace the placeholder variables in the template with the actual values
  for (var key in invoiceData) {
    var placeholder = '{{' + key + '}}';
    var value = invoiceData[key];
    invoiceBody = invoiceBody.replace(new RegExp(placeholder, 'g'), value);
  }
  
  // Create a new Google Doc for the invoice
  var newInvoiceDoc = DocumentApp.create('Invoice');
  var newInvoiceBody = newInvoiceDoc.getBody();
  newInvoiceBody.setText(invoiceBody);
  
  // Save and close the new invoice document
  newInvoiceDoc.saveAndClose();
  
  // Return the new invoice document as a Google Drive file object
  return DriveApp.getFileById(newInvoiceDoc.getId());
}

function sendInvoiceByEmail(invoice, recipientEmail, recipientName) {
  var subject = 'Invoice for Conference Registration';
  var body = 'Dear ' + recipientName + ',\n\nPlease find attached the invoice for registering to attend Ubuntunet-Connect 2023 Kampala, Uganda.';
  
  // Attach the invoice document to the email
  var attachment = invoice.getAs(MimeType.PDF);
  
  // Send the email
  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    body: body,
    attachments: [attachment]
  });
}
