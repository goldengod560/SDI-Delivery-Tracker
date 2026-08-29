// SDI Delivery Tracker v2 — Backend
// Sheet tab: "Orders"
// Columns: ID | Date | Name | Driver | Status | TimeOut | TimeBack | PhotoOutId | PhotoBackId

const SHEET_NAME = 'Orders';
const PHOTO_FOLDER_NAME = 'SDI Delivery Photos';

function doGet(e) {
  try {
    const action = (e.parameter || {}).action;
    if (action === 'list') return listOrders(e.parameter.date);
    if (action === 'listAll') return listAllOrders();
    return json({ error: 'Unknown action' });
  } catch (err) { return json({ error: err.message }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const a = body.action;
    if (a === 'addOut') return addOutgoing(body);
    if (a === 'addBack') return addReturned(body);
    if (a === 'delete') return deleteOrder(body);
    return json({ error: 'Unknown action' });
  } catch (err) { return json({ error: err.message }); }
}

// ===== HELPERS =====

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID','Date','Name','Driver','Status','TimeOut','TimeBack','PhotoOutId','PhotoBackId']);
  }
  return sheet;
}

function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function savePhoto(base64Data, fileName) {
  if (!base64Data) return '';
  try {
    // Remove data:image/...;base64, prefix
    const parts = base64Data.split(',');
    const raw = parts.length > 1 ? parts[1] : parts[0];
    const blob = Utilities.newBlob(Utilities.base64Decode(raw), 'image/jpeg', fileName);
    const folder = getPhotoFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getId();
  } catch (err) {
    Logger.log('Photo save error: ' + err.message);
    return '';
  }
}

function parseDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  if (!str) return '';
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  const d = new Date(str);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return str;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getAllOrders() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    obj.Date = parseDateVal(obj.Date);
    return obj;
  });
}

// ===== ACTIONS =====

function listOrders(dateFilter) {
  const orders = getAllOrders();
  const filtered = dateFilter ? orders.filter(o => o.Date === dateFilter) : orders;
  // Add photo URLs
  filtered.forEach(o => {
    if (o.PhotoOutId) o.PhotoOutUrl = 'https://drive.google.com/thumbnail?id=' + o.PhotoOutId + '&sz=w400';
    if (o.PhotoBackId) o.PhotoBackUrl = 'https://drive.google.com/thumbnail?id=' + o.PhotoBackId + '&sz=w400';
  });
  return json({ orders: filtered });
}

function listAllOrders() {
  return listOrders(null);
}

function addOutgoing(body) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateStr = body.date || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'h:mm a');

  let photoId = '';
  if (body.photo) {
    photoId = savePhoto(body.photo, 'out_' + id + '.jpg');
  }

  sheet.appendRow([
    id,
    dateStr,
    body.name || '',
    body.driver || '',
    'out',
    timeStr,
    '',
    photoId,
    ''
  ]);

  return json({ success: true, id: id });
}

function addReturned(body) {
  const sheet = getSheet();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const timeStr = Utilities.formatDate(now, tz, 'h:mm a');

  let photoId = '';
  if (body.photo) {
    photoId = savePhoto(body.photo, 'back_' + (body.id || Utilities.getUuid()) + '.jpg');
  }

  // If an existing order ID is provided, update it
  if (body.id) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('Status');
    const timeBackCol = headers.indexOf('TimeBack');
    const photoBackCol = headers.indexOf('PhotoBackId');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === body.id) {
        const row = i + 1;
        sheet.getRange(row, statusCol + 1).setValue('returned');
        sheet.getRange(row, timeBackCol + 1).setValue(timeStr);
        if (photoId) sheet.getRange(row, photoBackCol + 1).setValue(photoId);
        return json({ success: true });
      }
    }
  }

  // Otherwise add as a new returned entry
  const id = Utilities.getUuid();
  const dateStr = body.date || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  sheet.appendRow([
    id,
    dateStr,
    body.name || '',
    body.driver || '',
    'returned',
    '',
    timeStr,
    '',
    photoId
  ]);

  return json({ success: true, id: id });
}

function deleteOrder(body) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('ID');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      sheet.deleteRow(i + 1);
      return json({ success: true });
    }
  }
  return json({ error: 'Not found' });
}
