function setCorsHeaders(response) {
  return response
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function parsePayload(e) {
  if (e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }

  const raw = e.postData && e.postData.contents ? e.postData.contents : '';
  if (raw) {
    // Try JSON first, then fallback to URL-encoded form body.
    try {
      return JSON.parse(raw);
    } catch (error) {
      const parsed = {};
      raw.split('&').forEach((pair) => {
        const [key, value] = pair.split('=');
        if (key) {
          parsed[decodeURIComponent(key)] = decodeURIComponent(value || '');
        }
      });
      return parsed;
    }
  }

  return {};
}

function doPost(e) {
  try {
    const payload = parsePayload(e);
    const spreadsheetId = '1LPZ6KKibSS7TdwqUWi7SQmoUVAVYnB7_Jg4p1Cw56a4';
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

    const selectedPromises = Array.isArray(payload.selectedPromises)
      ? payload.selectedPromises.join(' | ')
      : payload.selectedPromises || '';

    sheet.appendRow([
      new Date(),
      payload.timestamp || '',
      payload.recipientName || '',
      payload.decision || '',
      payload.angerValue != null ? payload.angerValue : '',
      selectedPromises,
      payload.customPromise || '',
      payload.pageUrl || '',
    ]);

    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ success: true })));
  } catch (error) {
    return setCorsHeaders(
      ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message })),
    );
  }
}

function doGet() {
  return setCorsHeaders(
    ContentService.createTextOutput(
      JSON.stringify({ success: true, message: 'Apps Script is running.' }),
    ),
  );
}
