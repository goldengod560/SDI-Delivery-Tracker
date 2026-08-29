// SDI Delivery Tracker — Backend
// Orders columns:
// ID | Date | Name | Driver | Status | TimeOut | TimeBack | PhotoOutId | PhotoBackId | PhotoDeliveredId

const SHEET_NAME = 'Orders';
const DRIVERS_SHEET_NAME = 'Drivers';
const PHOTO_FOLDER_NAME = 'SDI Delivery Photos';

function doGet(e) {
  try {
    const action = (e.parameter || {}).action;
    if (action === 'list') return listOrders(e.parameter.date);
    if (action === 'listAll') return listAllOrders();
    if (action === 'drivers') return listDrivers();
    return json({ error: 'Unknown action' });
  } catch (err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'addOut') return addOutgoing(body);
    if (action === 'addBack') return addReturned(body);
    if (action === 'addDeliveredPhoto') return addDeliveredPhoto(body);
    if (action === 'delete') return deleteOrder(body);
    return json({ error: 'Unknown action' });
  } catch (err) {
    return json({ error: err.message });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID','Date','Name','Driver','Status','TimeOut','TimeBack','PhotoOutId','PhotoBackId','PhotoDeliveredId']);
    return sheet;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (headers.indexOf('PhotoDeliveredId') === -1) {
    sheet.getRange(1, lastColumn + 1).setValue('PhotoDeliveredId');
  }
  return sheet;
}

function getDriversSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DRIVERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DRIVERS_SHEET_NAME);
    sheet.appendRow(['Name']);
  }
  return sheet;
}

function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function savePhoto(base64Data, fileName) {
  if (!base64Data) throw new Error('Photo is required');
  const parts = base64Data.split(',');
  const raw = parts.length > 1 ? parts[1] : parts[0];
  const blob = Utilities.newBlob(Utilities.base64Decode(raw), 'image/jpeg', fileName);
  const folder = getPhotoFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function parseDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  if (!str) return '';
  const match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + '-' + match[2] + '-' + match[3];
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
  return data.slice(1).filter(row => row[0] !== '').map(row => {
    const obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    obj.Date = parseDateVal(obj.Date);
    return obj;
  });
}

function cleanDriverName(name) {
  return String(name || '').trim().replace(/\s+/g, ' ');
}

function saveDriver(name) {
  const cleaned = cleanDriverName(name);
  if (!cleaned) return '';

  const sheet = getDriversSheet();
  const values = sheet.getDataRange().getValues();
  const target = cleaned.toLowerCase();
  for (let i = 1; i < values.length; i++) {
    const existing = cleanDriverName(values[i][0]);
    if (existing.toLowerCase() === target) return existing;
  }
  sheet.appendRow([cleaned]);
  return cleaned;
}

function listDrivers() {
  const sheet = getDriversSheet();
  const values = sheet.getDataRange().getValues();
  const drivers = [];
  const seen = {};

  for (let i = 1; i < values.length; i++) {
    const name = cleanDriverName(values[i][0]);
    if (!name) continue;
    const key = name.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      drivers.push(name);
    }
  }

  getAllOrders().forEach(order => {
    const name = cleanDriverName(order.Driver);
    if (!name) return;
    const key = name.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      drivers.push(name);
      sheet.appendRow([name]);
    }
  });

  drivers.sort((a, b) => a.localeCompare(b));
  return json({ drivers: drivers });
}

function listOrders(dateFilter) {
  const orders = getAllOrders();
  const filtered = dateFilter ? orders.filter(order => order.Date === dateFilter) : orders;
  filtered.forEach(order => {
    if (order.PhotoOutId) order.PhotoOutUrl = 'https://drive.google.com/thumbnail?id=' + order.PhotoOutId + '&sz=w800';
    if (order.PhotoBackId) order.PhotoBackUrl = 'https://drive.google.com/thumbnail?id=' + order.PhotoBackId + '&sz=w800';
    if (order.PhotoDeliveredId) order.PhotoDeliveredUrl = 'https://drive.google.com/thumbnail?id=' + order.PhotoDeliveredId + '&sz=w800';
  });
  return json({ orders: filtered });
}

function listAllOrders() {
  return listOrders(null);
}

function addOutgoing(body) {
  const name = String(body.name || '').trim();
  const driver = saveDriver(body.driver);
  if (!name) return json({ error: 'Order name is required' });
  if (!driver) return json({ error: 'Delivery person is required' });

  const sheet = getSheet();
  const id = Utilities.getUuid();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const dateStr = body.date || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'h:mm a');
  const photoId = savePhoto(body.photo, 'out_' + id + '.jpg');

  sheet.appendRow([
    id,
    dateStr,
    name,
    driver,
    'out',
    timeStr,
    '',
    photoId,
    '',
    ''
  ]);

  return json({ success: true, id: id, driver: driver });
}

function addReturned(body) {
  if (!body.id) return json({ error: 'Order ID is required' });

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return json({ error: 'Order not found' });

  const headers = data[0];
  const idCol = headers.indexOf('ID');
  const statusCol = headers.indexOf('Status');
  const timeBackCol = headers.indexOf('TimeBack');
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(body.id)) {
      if (data[i][statusCol] === 'returned') return json({ success: true });
      const row = i + 1;
      sheet.getRange(row, statusCol + 1).setValue('returned');
      sheet.getRange(row, timeBackCol + 1).setValue(timeStr);
      return json({ success: true });
    }
  }

  return json({ error: 'Order not found' });
}

function addDeliveredPhoto(body) {
  if (!body.id) return json({ error: 'Order ID is required' });
  if (!body.photo) return json({ error: 'Delivery photo is required' });

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return json({ error: 'Order not found' });

  const headers = data[0];
  const idCol = headers.indexOf('ID');
  const statusCol = headers.indexOf('Status');
  const deliveredPhotoCol = headers.indexOf('PhotoDeliveredId');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(body.id)) {
      if (data[i][statusCol] === 'returned') return json({ error: 'This order is marked not delivered' });
      const photoId = savePhoto(body.photo, 'delivered_' + body.id + '.jpg');
      sheet.getRange(i + 1, deliveredPhotoCol + 1).setValue(photoId);
      return json({ success: true });
    }
  }

  return json({ error: 'Order not found' });
}

function deleteOrder(body) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return json({ error: 'Not found' });
  const idCol = data[0].indexOf('ID');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(body.id)) {
      sheet.deleteRow(i + 1);
      return json({ success: true });
    }
  }
  return json({ error: 'Not found' });
}
