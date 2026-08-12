function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const spreadsheetId = 'REPLACE_WITH_YOUR_SPREADSHEET_ID';
    const sheetName = 'Responses';

    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'Saved At',
        'Payload Timestamp',
        'Recipient Name',
        'Decision',
        'Anger Value',
        'Selected Promises',
        'Custom Promise',
        'Page URL',
      ]);
    }

    sheet.appendRow([
      new Date(),
      payload.timestamp || '',
      payload.recipientName || '',
      payload.decision || '',
      payload.angerValue != null ? payload.angerValue : '',
      Array.isArray(payload.selectedPromises) ? payload.selectedPromises.join(' | ') : '',
      payload.customPromise || '',
      payload.pageUrl || '',
    ]);

    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(
      ContentService.MimeType.JSON,
    );
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: error.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ success: true, message: 'Apps Script is running.' }),
  ).setMimeType(ContentService.MimeType.JSON);
}
