# Google Sheets PDF Automation
This repository contains Google Apps Script functions designed to automate the batch export of specific ranges from a Google Sheet into individual PDF files, and then combine these PDFs into a single merged document. This is particularly useful for generating multiple forms or reports from a template sheet based on a list of unique identifiers.

## Features
- Batch PDF Export (batchExportAndSplitMergePDF):
  - Iterates through a list of "Form Numbers" (or similar identifiers) from a designated data sheet.
  - For each form number, it updates a specific cell in a template sheet, causing the sheet's content to dynamically change.
  - Exports a predefined range of this template sheet as an individual PDF file.
  - Saves each generated PDF into a specified Google Drive folder.
  - Includes basic rate limit handling to pause execution if Google's API limits are hit.
- Combine Recently Exported PDFs (combinePDFs):
  - Automatically called after batchExportAndSplitMergePDF completes.
  - Collects all PDF files created within a recent time window (default: 30 minutes) from the individual PDF folder.
  - Merges these PDFs into a single, combined PDF document.
  - Names the merged PDF dynamically using "Form Numbers" from the data sheet.
  - Saves the combined PDF to a specified merged PDF folder.
  - Automatically opens the newly created merged PDF in a new browser tab.

## Setup
To use these scripts, you'll need a Google Account and access to Google Sheets and Google Drive.

1. Open Google Apps Script:
   - Open your Google Sheet (the one containing "Pemindahan AT" and "Sheet1").
   - Go to Extensions > Apps Script. This will open the Google Apps Script editor.

2. Create Script Files:
   - In the Apps Script editor, you'll see a Code.gs file by default. Paste all the provided code into this file.

3. Enable Advanced Google Services:
   - In the Apps Script editor, on the left sidebar, click on "Services" (the + icon next to "Services").
   - Search for and add the Google Drive API service.
   - Click "Add".

4. Prepare Your Google Sheet:
   - "Pemindahan AT" Sheet: This sheet acts as your template. Ensure it has a cell (e.g., L10 as per the script) where the "Form Number" can be set to dynamically update the content you wish to export.
   - "Sheet1" Sheet: This sheet should contain the list of "Form Numbers" in column B, starting from B2. It also contains the firstNumber and lastNumber for the combined PDF's filename in column U (index 20), starting from row 2.

5. Create Google Drive Folders:
   - Create a Google Drive folder where the individual exported PDFs will be saved (e.g., "Individual Exports").
   - Create another Google Drive folder where the final combined PDF will be saved (e.g., "Merged PDFs").
   - Note the Folder IDs for both (from the URL: https://drive.google.com/drive/folders/YOUR_FOLDER_ID).

## Configuration
Before running the scripts, you need to update the IDs and sheet names within the code:
- batchExportAndSplitMergePDF function:
  - INDIVIDUAL_FOLDER_ID: Replace "12nHUSvxAM24JjpGf_nD2V1DASplvJUHt" with the ID of your folder for individual PDFs.
  - MERGED_FOLDER_ID: Replace "1UbtEipzy_U-_IXKdWCH49k7iZNUoMwCr" with the ID of your folder for merged PDFs.
  - sheet.getRange("L10"): This is the cell in the "Pemindahan AT" sheet that gets updated with the form number. Adjust if your template uses a different cell.
  - range=B2:M37: This defines the print range for the PDF export. Adjust this to match the specific area of your "Pemindahan AT" sheet you want to export.

- combinePDFs function:
  - sourceFolderId: Replace '12nHUSvxAM24JjpGf_nD2V1DASplvJUHt' with the ID of the folder containing the individual PDFs to be combined (this should be the same as INDIVIDUAL_FOLDER_ID in batchExportAndSplitMergePDF).
  - targetFolderId: Replace '1UbtEipzy_U-_IXKdWCH49k7iZNUoMwCr' with the ID of the folder where the final combined PDF will be saved.
  - firstNumber = data[1][20] and lastNumber = data[data.length - 1][20]: These lines retrieve the first and last numbers for the merged PDF's filename. They are hardcoded to column U (index 20) of "Sheet1". Adjust the column index 20 if these numbers are in a different column.

## Usage
To run the entire process (batch export and then combine PDFs):

- In the Apps Script editor, select the batchExportAndSplitMergePDF function from the dropdown menu at the top.
- Click the "Run" button (play icon).
- The first time you run it, you will be prompted to authorize the script. Follow the on-screen instructions to grant the necessary permissions.

## Important Notes
- Permissions: When you first run these scripts, Google will ask for permissions to access your Google Sheets and Google Drive. Grant these permissions for the scripts to function correctly.

- PDF-lib: The combinePDFs function uses the pdf-lib library, which is loaded directly from a CDN using UrlFetchApp.fetch().getContentText(). This is a common practice in Google Apps Script for external libraries but be aware of potential security implications if the CDN source is compromised (though cdnjs.cloudflare.com is generally reliable).

- Rate Limits: The batchExportAndSplitMergePDF function includes basic handling for HTTP 429 (Too Many Requests) errors by pausing execution. For very large batches, you might still encounter limits or need more sophisticated error handling/retry mechanisms.

- Execution Time Limit: Google Apps Script has an execution time limit (typically 6 minutes for free accounts). For very large datasets, the script might time out. Consider breaking down the task or using installable triggers.

- Time Window for Combining: The combinePDFs function only merges PDFs that were created within the last timeLimit (30 minutes by default). This prevents old, irrelevant PDFs from being included in the merge.

- No Automatic Cleanup of Individual PDFs: Unlike some other mail merge scripts, this version does not automatically delete the individual PDFs from the INDIVIDUAL_FOLDER_ID after they are merged. You will need to manually clean up that folder if desired.

## License
This project is open-source and available under the MIT License.
