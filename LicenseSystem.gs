const SHEET_NAME = 'LicenseKeys';
const API_KEY = 'mdtech-secret-2026';

function getSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    try { ss = SpreadsheetApp.openById('1vBzAfimOCv3pdGc7SkhyVjZSOSUqeBk-9DxfkPbmjZwoac5s5yOPzTYQ'); } catch (e) {}
  }
  if (!ss) {
    ss = SpreadsheetApp.create('MDTechnology License Keys');
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['LicenseKey','Email','Status','Created','DeviceId','LastUsed']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    getSheet();
    const d = e.parameter || {};
    const a = d.action || '';
    const k = d.apiKey || '';
    if (a === 'verify') return verifyKey(d.licenseKey, d.deviceId);
    if (a === 'generate' && k === API_KEY) return generateKey(d.email);
    if (a === 'revoke' && k === API_KEY) return revokeKey(d.licenseKey);
    if (a === 'status' && k === API_KEY) return getStatus(d.licenseKey);
    return respond({ error:'Invalid action or auth' });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

function generateKey(email) {
  const s = getSheet();
  const key = generateLicenseKey();
  s.appendRow([key, email, 'active', new Date().toISOString(), '', '']);
  return respond({ success: true, licenseKey: key });
}

function verifyKey(licenseKey, deviceId) {
  const s = getSheet();
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === licenseKey) {
      if (d[i][2] === 'active') {
        s.getRange(i+1,5).setValue(deviceId||'');
        s.getRange(i+1,6).setValue(new Date().toISOString());
        return respond({ valid: true });
      } else {
        return respond({ valid: false, reason: 'revoked' });
      }
    }
  }
  return respond({ valid: false, reason: 'not_found' });
}

function revokeKey(licenseKey) {
  const s = getSheet();
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === licenseKey) {
      s.getRange(i+1,3).setValue('revoked');
      return respond({ success: true });
    }
  }
  return respond({ success: false, error:'Key not found' });
}

function getStatus(licenseKey) {
  const s = getSheet();
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === licenseKey) {
      return respond({ success:true, licenseKey:d[i][0], email:d[i][1], status:d[i][2], created:d[i][3] });
    }
  }
  return respond({ success: false, error:'Key not found' });
}

function generateLicenseKey() {
  const c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let k = '';
  for (let i = 0; i < 4; i++) {
    for (let j = 0; j < 4; j++) {
      k += c[Math.floor(Math.random()*c.length)];
    }
    if (i < 3) k += '-';
  }
  const s = getSheet();
  const existing = s.getDataRange().getValues().map(r => r[0]);
  if (existing.includes(k)) return generateLicenseKey();
  return k;
}

function respond(data) {
  const o = ContentService.createTextOutput(JSON.stringify(data));
  o.setMimeType(ContentService.MimeType.JSON);
  return o;
}
