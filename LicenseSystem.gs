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
    const a = e.parameter.action || '';
    const k = e.parameter.apiKey || '';
    if (a === 'verify') return verifyKey(e.parameter.licenseKey, e.parameter.deviceId);
    if (a === 'generate' && k === API_KEY) return generateKey(e.parameter.email, e.parameter.ref || '');
    if (a === 'revoke' && k === API_KEY) return revokeKey(e.parameter.licenseKey);
    if (a === 'status' && k === API_KEY) return getStatus(e.parameter.licenseKey);
    if (a === 'lookup' && k === API_KEY) return lookupKeyByEmail(e.parameter.email);
    if (a === 'gumroadWebhook' && k === API_KEY) return handleGumroadWebhook(e);
    return respond({ error:'Invalid action or auth' });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

function handleGumroadWebhook(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }
    const licenseKey = data.license_key || data.licenseKey || '';
    const email = data.buyer_email || data.email || '';
    if (!licenseKey || !email) {
      Logger.log('Gumroad webhook missing data: ' + JSON.stringify(data));
      return respond({ success: false, error: 'Missing license_key or buyer_email' });
    }
    const s = getSheet();
    const existing = s.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0] === licenseKey) {
        return respond({ success: true, message: 'Key already exists' });
      }
    }
    s.appendRow([licenseKey, email, 'active', new Date().toISOString(), 'gumroad', '']);
    return respond({ success: true, message: 'Key stored from Gumroad' });
  } catch (err) {
    Logger.log('Gumroad webhook error: ' + err.toString());
    return respond({ success: false, error: err.toString() });
  }
}

function lookupKeyByEmail(email) {
  if (!email) return respond({ success: false, error: 'Email required' });
  const s = getSheet();
  const d = s.getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < d.length; i++) {
    if (d[i][1] === email && d[i][2] === 'active') {
      keys.push({ licenseKey: d[i][0], created: d[i][3] });
    }
  }
  if (keys.length > 0) {
    return respond({ success: true, keys: keys });
  }
  return respond({ success: false, error: 'No active keys found for this email' });
}

function generateKey(email, ref) {
  const s = getSheet();
  const key = generateLicenseKey();
  s.appendRow([key, email, 'active', new Date().toISOString(), ref || '', '']);
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
