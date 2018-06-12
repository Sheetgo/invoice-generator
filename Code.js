/*================================================================================================================*
  Invoice Generator
  ================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/Sheetgo/invoice-generator
  Copyright:    (c) 2018 by Sheetgo

  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  ----------------------------------------------------------------------------------------------------------------
  Changelog:

  1.0.0  Initial release
 *================================================================================================================*/

/**
 * Project Settings
 * @type {JSON}
 */
SETTINGS = {

    // Spreadsheet ID (template filled with your data)
    // The id is a unique set of characters that can be found in the spreadsheet url
    spreadsheetId: '<your_spreadsheet_id>',

    // Spreadsheet name
    sheetName: "Data",

    // Document Id
    documentId: "<your_document_id>",

    // Folder id
    folderId: '<your_folder_id>'

};


/**
 * This funcion will run when you open the spreadsheet. It creates a Spreadsheet menu option to run the spript
 */
function onOpen() {
    // Adds a custom menu to the spreadsheet.
    SpreadsheetApp.getUi()
        .createMenu('Invoice Generator')
        .addItem('Generate Invoices', 'sendInvoice')
        .addToUi();
}


/**
 * Reads the spreadsheet data and creates the PDF invoice
 */
function sendInvoice() {
    // Opens the spreadsheet and access the tab containing the data
    var ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
    var dataSheet = ss.getSheetByName(SETTINGS.sheetName);

    // Gets all values from the instanciated tab
    var sheetValues = dataSheet.getDataRange().getValues();

    // Gets the PDF url column index
    var pdfIndex = sheetValues[0].indexOf("PDF Url");

    // Gets the user's name (will be used as the PDF file name)
    var clientNameIndex = sheetValues[0].indexOf("client_name");

    var counter, invoiceNumCount, pdfInvoice, invoiceId, key, values, pdfName, invoiceNumber, invoiceDate;
    for (var i = 1; i < sheetValues.length; i++) {

        // Creates the Invoice
        if (!sheetValues[i][pdfIndex]) {

            // Duplicate teh template on Google Drive to manipulate the data
            invoiceId = DriveApp.getFileById(SETTINGS.documentId).makeCopy('Template Copy').getId();

            // Instantiate the document
            var docBody = DocumentApp.openById(invoiceId).getBody();

            // Iterates over the spreadsheet columns to get the values used to write the document
            for (var j = 0; j < sheetValues[i].length; j++) {

                // Key and Values to be replaced
                key = sheetValues[0][j];
                values = sheetValues[i][j];

                if (key === "date") {
                    // Invoice Date
                    invoiceDate = (values.getMonth() + 1) + "/" + values.getDate() + "/" + values.getFullYear();
                    replace('%date%', invoiceDate, docBody);    // Write data

                } else if (values) {
                    // Everything else appart from date values
                    if (key.indexOf("price") > -1 || key === "discount" || key.indexOf("total") > -1) {
                        replace('%' + key + '%', '$' + values.toFixed(2), docBody); // Replace values
                    } else if (key === "tax_id") {
                        replace('%' + key + '%', "Tax ID: " + values, docBody); // Replace the tax_id
                    } else {
                        replace('%' + key + '%', values, docBody)
                    }

                } else {
                    replace('%' + key + '%', '', docBody) // Replace empty string
                }
            }

            // Get last invoice count from the tab 'Count'
            counter = ss.getSheetByName('Count').getRange('A2');
            invoiceNumCount = counter.getValue() + 1;

            invoiceNumber = invoiceNumCount.padLeft(4, '0') + "/" + invoiceDate.split("/")[2];
            replace('%number%', invoiceNumber, docBody);

            // Rename the invoice document
            pdfName = sheetValues[i][clientNameIndex] + " " + invoiceNumber;
            DocumentApp.openById(invoiceId).setName(pdfName).saveAndClose();

            // Convert the Invoice Document into a PDF file
            pdfInvoice = convertPDF(invoiceId);

            // Set the PDF url into the spreadsheet
            dataSheet.getRange(i + 1, pdfIndex + 1).setValue(pdfInvoice[0]);

            // Update invoice counter
            counter.setValue(invoiceNumCount);

            // Delete the original document (will leave only the PDF)
            Drive.Files.remove(invoiceId);
        }
    }
}


/**
 * Convert a Google Docs into a PDF file
 * @param {string} id - File Id
 * @returns {*[]}
 */
function convertPDF(id) {
    var doc = DocumentApp.openById(id);
    var docBlob = DocumentApp.openById(id).getAs('application/pdf');
    docBlob.setName(doc.getName() + ".pdf"); // Add the PDF extension
    var invFolder = SETTINGS.folderId;
    var file = DriveApp.getFolderById(invFolder).createFile(docBlob);
    var url = file.getUrl();
    id = file.getId();
    return [url, id];
}


/**
 * Replace the document key/value
 * @param {String} key - The document key to be replaced
 * @param {String} text - The document text to be inserted
 * @param {Body} body - the active document's Body.
 * @returns {Element}
 */
function replace(key, text, body) {
    return body.editAsText().replaceText(key, text);
}


/**
 * Returns a new string that right-aligns the characters in this instance by padding them with any string on the left,
 * for a specified total length.
 * @param {Number} n - Number of characters to pad
 * @param {String} str - The string to be padded
 * @returns {string}
 */
Number.prototype.padLeft = function (n, str) {
    return Array(n - String(this).length + 1).join(str || '0') + this;
};