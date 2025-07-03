function batchExportAndSplitMergePDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pemindahan AT");
  const dataSheet = ss.getSheetByName("Sheet1");

  const INDIVIDUAL_FOLDER_ID = "CHANGE_YOUR_ID"; 
  const MERGED_FOLDER_ID = "CHANGE_YOUR_ID";    
  const individualFolder = DriveApp.getFolderById(INDIVIDUAL_FOLDER_ID);
  const mergedFolder = DriveApp.getFolderById(MERGED_FOLDER_ID);

  const formNos = dataSheet.getRange("B2:B").getValues().flat().filter(String);

  const pdfBlobs = [];

  // Ambil Waktu Execution
  const startTime = new Date();
  PropertiesService.getScriptProperties().setProperty("LAST_BATCH_EXPORT_TIME", startTime.toISOString());

  for (let i = 0; i < formNos.length; i++) {
    const formNo = formNos[i];
    sheet.getRange("L10").setValue(formNo);
    
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    try {
      const url = ss.getUrl().replace(/edit$/, '');
      const exportUrl = url + 'export?format=pdf' +
                        '&exportFormat=pdf' +
                        '&gid=' + sheet.getSheetId() +
                        '&range=B2:M37' +
                        '&portrait=false' +
                        '&size=A4' +
                        '&top_margin=0.5&bottom_margin=0' +
                        '&left_margin=0&right_margin=0' +
                        '&sheetnames=false&printtitle=false&pagenumbers=false' +
                        '&gridlines=false&fzr=false';

      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 429) {
        Logger.log("Rate limit hit. Waiting...");
        Utilities.sleep(5000);
        i--;
        continue;
      }

      const blob = response.getBlob().setName(formNo + ".pdf");
      individualFolder.createFile(blob);
      pdfBlobs.push(blob);

      Utilities.sleep(3000);
    } catch (err) {
      Logger.log("Gagal ekspor form: " + formNo + " Error: " + err);
    }
  }
  combinePDFs();

}
async function combinePDFs() {
  const sourceFolderId = 'CHANGE_YOUR_ID'; // Folder PDF hasil mail merge
  const targetFolderId = 'CHANGE_YOUR_ID'; // Folder hasil gabungan PDF

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const files = sourceFolder.getFiles();

  // Ambil nomor urut awal dan akhir dari spreadsheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const firstNumber = data[1][20]; // baris kedua (index 1), kolom pertama (index 0)
  const lastNumber = data[data.length - 1][20]; // baris terakhir

  // Gunakan nama sesuai format diminta
  const mergedFileName = `${lastNumber} Form Serah Terima HP Samsung ${firstNumber} - ${lastNumber}.pdf`;

  const now = new Date();
  const timeLimit = 30 * 60 * 1000; // 5 menit dalam ms

  // Load PDF-lib dari CDN
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText());
  const setTimeout = function (f, t) {
    Utilities.sleep(t);
    return f();
  };

  const pdfDoc = await PDFLib.PDFDocument.create();

  while (files.hasNext()) {
    const file = files.next();
    const created = file.getDateCreated();

    if (
      file.getMimeType() === 'application/pdf' &&
      (now - created) <= timeLimit
    ) {
      const pdfData = await PDFLib.PDFDocument.load(new Uint8Array(file.getBlob().getBytes()));
      const pages = await pdfDoc.copyPages(pdfData, [...Array(pdfData.getPageCount())].map((_, i) => i));
      pages.forEach(page => pdfDoc.addPage(page));
    }
  }

  const bytes = await pdfDoc.save();
  const mergedFile = targetFolder.createFile(
    Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, mergedFileName)
  );

  // Buka otomatis 
  const mergedUrl = mergedFile.getUrl();
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${mergedUrl}', '_blank');google.script.host.close();</script>`
  ).setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, "Membuka File Hasil Merge...");
}
