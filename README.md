# Invoice Generator


## Introduction

At Sheetgo we have a database of paid clients in Google Sheets and we need to send them invoicing. Pretty soon after launching we realized this would be a monumental task to do manually and the existing software out there did not meet our needs. We created this Google Apps Script to automatically generate PDF invoices based on data in Google Sheets. See below for more info on how to configure and use.


## How to configure

1. Create a copy of our invoice spreadsheet template at <a href="https://docs.google.com/spreadsheets/d/1uaHmsl_-R2wJyt6HE3NpUSKII_Cuk5D90BTnK-YDbYY/copy" target="_blank">link</a>. Now the files are saved to your Google Drive.
2. Inside the spreadsheet, on the instruction tab, simply follow the instructions provided. First click on **‘Install Solution’** to configure your Invoice Generator Template. You will need to authorize the script to run and login using your Google account.
3. Open the Invoice folder that has been created inside your Google Drive by clicking the link on the instruction tab. Open the ‘Invoice Template’ document, add your logo and your company data anywhere you see these brackets ‘< >’ (see on the image below), the words/values contained between the percent symbols (%) will be automatically substituted based on the data in the spreadsheet.
4. Back inside the Invoice data spreadsheet, follow the next step and thus go to the ‘Data’ tab to overwrite the dummy data and add your invoices.
5. To generate the Invoices PDFs, go to the spreadsheet Menu bar -> ‘Invoice Generator’ -> ‘Generate Invoices’. Your invoice PDFs are now automatically created in the Invoice Folder inside your Google Drive.


## More information

For detailed information please visit the <a href="https://blog.sheetgo.com/google-cloud-solutions/invoice-generator/" target="_blank">blog</a>.