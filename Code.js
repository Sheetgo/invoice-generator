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
1.1.0  Auto configuration
*================================================================================================================*/

/**
* Project Settings
* @type {JSON}
*/
SETTINGS = {

    // Spreadsheet name
    sheetName: "Data",

    // Document Url
    documentUrl: null,

    // Template Url
    templateUrl: '14oTfL_zUbBdRD4VXY8U0NAJjQ4cKNxHGBax-bfH5NDs',

    // Set name spreadsheet
    spreadsheetName: 'Invoice data',

    //Set name document
    documentName: 'Invoice Template',

    // Sheet Settings
    sheetSettings: "Settings",

    // Column Settings
    col: {
        templateId: "B1",
        count: "B2",
        folderId: "B3",
        systemCreated: "B4"
    }
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
* This function Create system
*/
function createSystem() {

    try {

        var ss = SpreadsheetApp.getActiveSpreadsheet();


        // Get name tab
        var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

        // Checks function createSystem is run
        var systemCreated = sheetSettings.getRange(SETTINGS.col.systemCreated);
        if (!systemCreated.getValue()) {
            systemCreated.setValue('True');
        } else {
            showUiDialog('Warnning', 'Solution has already been created!');
            return;
        }

        // Checks if cell Count exists
        var count = sheetSettings.getRange(SETTINGS.col.count);
        if (!count.getValue()) {
            count.setValue(0);
        }

        // Create the Solution folder on users Drive
        var invoiceFolder = DriveApp.createFolder('Invoice Folder');
        var folder = invoiceFolder.createFolder('Invoices');

        // Set URL Invoice Folder in tab Instructions
        ss.getSheetByName('Instructions').getRange('C15').setValue(invoiceFolder.getUrl());

        // Move the current Dashboard spreadsheet into the Solution folder
        var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        file.setName(SETTINGS.spreadsheetName);

        // Move the sheet for invoice folder
        moveFile(file, invoiceFolder);

        // Move the current Dashboard template into the Solution folder
        var doc = DriveApp.getFileById(SETTINGS.templateUrl);
        var docCopy = doc.makeCopy(SETTINGS.documentName);

        // Set tab settings document ID
        sheetSettings.getRange(SETTINGS.col.templateId).setValue(docCopy.getId());

        // Move an copy for invoice folder
        moveFile(docCopy, invoiceFolder);

        // Set folder ID 
        sheetSettings.getRange(SETTINGS.col.folderId).setValue(folder.getId());


        // End process
        showUiDialog('Success', 'Your solution is ready');

        return true;
    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}


/**
* Reads the spreadsheet data and creates the PDF invoice
*/
function sendInvoice() {

    try {

        // Opens the spreadsheet and access the tab containing the data
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName(SETTINGS.sheetName);
        var sheetSettings = ss.getSheetByName(SETTINGS.sheetSettings);

        // Checks if cell Count exists
        var count = sheetSettings.getRange(SETTINGS.col.count).getValue();
        if (!count) {
            sheetSettings.getRange(SETTINGS.col.count).setValue(0);
        }

        // Checks function createSystem is run
        var control = sheetSettings.getRange(SETTINGS.col.systemCreated).getValue();
        if (!control) {
            showUiDialog('Warnning', 'Run "Install Solution" in tab Instructions');
            return;
        }

        // Gets all values from the instanciated tab
        var sheetValues = dataSheet.getDataRange().getValues();

        // Gets the PDF url column index
        var pdfIndex = sheetValues[0].indexOf("PDF Url");

        // Gets the user's name (will be used as the PDF file name)
        var clientNameIndex = sheetValues[0].indexOf("client_name");

        var counter, invoiceNumCount, pdfInvoice, invoiceId, key, values, pdfName, invoiceNumber, invoiceDate;

        // Duplicate teh template on Google Drive to manipulate the data
        var docId = sheetSettings.getRange(SETTINGS.col.templateId).getValue();

        for (var i = 1; i < sheetValues.length; i++) {

            // Creates the Invoice
            if (!sheetValues[i][pdfIndex]) {

                // Copy tamplate document
                var invoiceId = DriveApp.getFileById(docId).makeCopy('Template Copy').getId();

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
                counter = sheetSettings.getRange(SETTINGS.col.count);
                invoiceNumCount = counter.getValue() + 1;

                // Format invoice name pdf
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
    } catch (e) {

        // Show the error
        showUiDialog('Something went wrong', e.message)

    }
}

/**
* Move a file from one folder into another
* @param {Object} file A file object in Google Drive
* @param {Object} dest_folder A folder object in Google Drive 
*/
function moveFile(file, dest_folder, isFolder) {

    if (isFolder === true) {
        dest_folder.addFolder(file)
    } else {
        dest_folder.addFile(file);
    }
    var parents = file.getParents();
    while (parents.hasNext()) {
        var folder = parents.next();
        if (folder.getId() != dest_folder.getId()) {
            if (isFolder === true) {
                folder.removeFolder(file)
            } else {
                folder.removeFile(file)
            }

        }
    }
}

/**
* Convert a Google Docs into a PDF file
* @param {string} id - File Id
* @returns {*[]}
*/
function convertPDF(id) {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS.sheetSettings);
    var doc = DocumentApp.openById(id);
    var docBlob = DocumentApp.openById(id).getAs('application/pdf');
    docBlob.setName(doc.getName() + ".pdf"); // Add the PDF extension
    var invFolder = ss.getRange(SETTINGS.col.folderId).getValue();
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

/**
* Loads the showDialog
*/
function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('iframe.html')
        .setWidth(200)
        .setHeight(150)
    SpreadsheetApp.getUi().showModalDialog(html, 'Creating Solution..')
}

/**
* Show an UI dialog
* @param {string} title - Dialog title
* @param {string} message - Dialog message
*/
function showUiDialog(title, message) {
    try {
        var ui = SpreadsheetApp.getUi()
        ui.alert(title, message, ui.ButtonSet.OK)
    } catch (e) {
        // pass
    }
}