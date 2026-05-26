const SHEET_NAME = 'LicenseKeys';
const SS = SpreadsheetApp.getActiveSpreadsheet();
const API_KEY = 'mdtech-secret-2026'; // Simple auth to prevent abuse

function setupSheet() {
  const sheet = SS.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const newSheet = SS.insertSheet(SHEET_NAME);
    newSheet.appendRow(['LicenseKey', 'Email', 'Status', 'Created', 'DeviceId', 'LastUsed']);
    newSheet.setFrozenRows(1);
  }
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    setupSheet();
    const data = e.parameter || {};
    const action = data.action || '';
    const auth = data.apiKey || '';

    if (action === 'verify') {
      return verifyKey(data.licenseKey, data.deviceId);
    }
    if (action === 'generate' && auth === API_KEY) {
      return generateKey(data.email);
    }
    if (action === 'revoke' && auth === API_KEY) {
      return revokeKey(data.licenseKey);
    }
    if (action === 'status' && auth === API_KEY) {
      return getStatus(data.licenseKey);
    }

    return respond({ error: 'Invalid action or auth' }, 400);
  } catch (err) {
    return respond({ error: err.toString() }, 500);
  }
}

function generateKey(email) {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const key = generateLicenseKey();
  const now = new Date();
  sheet.appendRow([key, email, 'active', now.toISOString(), '', '']);
  return respond({ success: true, licenseKey: key });
}

function verifyKey(licenseKey, deviceId) {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === licenseKey) {
      if (data[i][2] === 'active') {
        const row = i + 1;
        sheet.getRange(row, 5).setValue(deviceId || '');
        sheet.getRange(row, 6).setValue(new Date().toISOString());
        return respond({ valid: true });
      } else {
        return respond({ valid: false, reason: 'revoked' });
      }
    }
  }
  return respond({ valid: false, reason: 'not_found' });
}

function revokeKey(licenseKey) {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === licenseKey) {
      sheet.getRange(i + 1, 3).setValue('revoked');
      return respond({ success: true });
    }
  }
  return respond({ success: false, error: 'Key not found' });
}

function getStatus(licenseKey) {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === licenseKey) {
      return respond({
        success: true,
        licenseKey: data[i][0],
        email: data[i][1],
        status: data[i][2],
        created: data[i][3]
      });
    }
  }
  return respond({ success: false, error: 'Key not found' });
}

function generateLicenseKey() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let key = '';
  for (let i = 0; i < 4; i++) {
    for (let j = 0; j < 4; j++) {
      key += chars[Math.floor(Math.random() * chars.length)];
    }
    if (i < 3) key += '-';
  }
  const sheet = SS.getSheetByName(SHEET_NAME);
  const existing = sheet.getDataRange().getValues().map(r => r[0]);
  if (existing.includes(key)) return generateLicenseKey();
  return key;
}

function respond(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  if (statusCode) {
    const range = `ResponseCode${statusCode}`; // not directly supported, skipping
  }
  return output;
}
