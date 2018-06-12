# Invoice Generator


## Introduction

At Sheetgo we have a database of paid clients in Google Sheets and we need to send them invoicing. Pretty soon after launching we realized this would be a monumental task to do manually and the existing software out there did not meet our needs. We created this Google Apps Script to automatically generate PDF invoices based on data in Google Sheets. See below for more info on how to configure and use.


## How to configure

1. To begin using first create a copy of our invoice document template at this <a href="https://docs.google.com/document/d/14oTfL_zUbBdRD4VXY8U0NAJjQ4cKNxHGBax-bfH5NDs/copy" target="_blank">link</a>, and a copy of our invoice spreadsheet template at this <a href="https://docs.google.com/spreadsheets/d/1uaHmsl_-R2wJyt6HE3NpUSKII_Cuk5D90BTnK-YDbYY/copy" target="_blank">link</a>.
2. In the document, add your logo and your company data anywhere you see these brackets ‘< >’, the words / values contained between the percent symbols (%) will be automatically substitutes based on the data in the spreadsheet.
3. Now copy and save the ID of the document.
4. Now fill out the requisite data in the spreadsheet. Note: Do not change the column headers in the spreadsheet as they correspond to tags in the document. Copy the ID of the spreadsheet.
5. On the tab “count” in the spreadsheet, you can set the invoice number for the first invoice by swapping out the 0 that is in there. For each new invoice the number will increment by 1.
6. Now, create a folder on Google Drive to store newly created invoices. You’ll also need to copy this ID for future use.
7. Go back to the spreadsheet and click on Tools and then Script Editor to edit the script.
8. Now just substitute your unique Spreadsheet and Folder IDs. Paste the spreadsheet ID that you copied in step 3 in between the quotes, substituting any existing values. Do the same things with the ID of the folder that you copied in step 5 and the document ID in step 2.
9. Congrats! You have configured your script. Now to use it go back to the spreadsheet and select the menu item that you just created, Invoice Generator and then create invoices. The document(s) will be created if the column “Sent?” is blank. After creating the value is changed automatically to “Yes” and a link with a PDF is automatically saved.



## Version 

+ 1.0.0 - Initial release.



## More information

For detailed information please visit the <a href="https://blog.sheetgo.com/google-cloud-solutions/invoice-generator/" target="_blank">blog</a>.