// SDI Delivery Tracker — Secure Backend (v2)
// Uses PropertiesService for token persistence across script restarts
// Sheet: ID, Date, Name, Driver, OrderNumber, Notes, Status, Reason, Payment, PayMethod, ResolvedTime, AddedTime, PayTime, ReceivedBy

const SHEET_NAME = 'Stops';
const REPORT_EMAIL = 'sadadelivery1@gmail.com';
const APP_PASSWORD = 'sdi2026'; // CHANGE THIS to your own password
const TOKEN_EXPIRY_HOURS = 24;

function getSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'list') {
    const token = e.parameter.token;
    if (!isValidToken(token)) return json({ error: 'Unauthorized' });
    return json({ stops: getAllStops() });
  }
  return json({ error: 'Unknown action' });
}

function doPost(e) {
  let body; try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ error: 'Invalid JSON' }); }
  const a = body.action;
  
  if (a === 'login') return doLogin(body);
  
  if (!isValidToken(body.token)) return json({ error: 'Unauthorized' });
  
  if (a === 'add') return addStop(body);
  if (a === 'update') return updateStop(body);
  if (a === 'delete') return deleteStop(body);
  if (a === 'copyRoute') return copyRoute(body);
  if (a === 'sendReport') return sendReport(7, 'weekly');
  if (a === 'sendMonthlyReport') return sendReport(30, 'monthly');
  return json({ error: 'Unknown action' });
}

// ===== AUTH =====
function doLogin(body) {
  if (body.password === APP_PASSWORD) {
    const token = Utilities.getUuid();
    const props = PropertiesService.getScriptProperties();
    const tokens = JSON.parse(props.getProperty('SDI_TOKENS') || '{}');
    tokens[token] = Date.now() + (TOKEN_EXPIRY_HOURS * 60 * 60 * 1000);
    props.setProperty('SDI_TOKENS', JSON.stringify(tokens));
    return json({ success: true, token: token });
  }
  return json({ error: 'Invalid password' });
}

function isValidToken(token) {
  if (!token) return false;
  const props = PropertiesService.getScriptProperties();
  const tokens = JSON.parse(props.getProperty('SDI_TOKENS') || '{}');
  if (!tokens[token]) return false;
  if (Date.now() > tokens[token]) {
    delete tokens[token];
    props.setProperty('SDI_TOKENS', JSON.stringify(tokens));
    return false;
  }
  return true;
}

function cleanupTokens() {
  const props = PropertiesService.getScriptProperties();
  const tokens = JSON.parse(props.getProperty('SDI_TOKENS') || '{}');
  const now = Date.now();
  let changed = false;
  Object.keys(tokens).forEach(k => { if (now > tokens[k]) { delete tokens[k]; changed = true; } });
  if (changed) props.setProperty('SDI_TOKENS', JSON.stringify(tokens));
}

// ===== DATA ACCESS =====
function getHeaders() { return getSheet().getDataRange().getValues()[0]; }
function getAllStops() {
  const data = getSheet().getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).filter(r => r[0] !== '').map(r => { const obj = {}; headers.forEach((h, i) => obj[h] = r[i]); return obj; });
}
function findRow(id) {
  const data = getSheet().getDataRange().getValues();
  const idCol = data[0].indexOf('ID');
  for (let i = 1; i < data.length; i++) if (data[i][idCol] === id) return i + 1;
  return -1;
}

// ===== ACTIONS =====
function addStop(body) {
  const now = new Date();
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');
  const dateStr = body.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  getSheet().appendRow([Utilities.getUuid(), dateStr, body.name || '', body.address || '', body.orderNumber || '', body.notes || '', 'pending', '', '', '', '', timeStr, '', '']);
  return json({ success: true });
}

function updateStop(body) {
  const rowNum = findRow(body.id); if (rowNum === -1) return json({ error: 'Stop not found' });
  const headers = getHeaders();
  const now = new Date(); const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');
  Object.keys(body.fields).forEach(key => {
    const col = headers.indexOf(key); if (col !== -1) getSheet().getRange(rowNum, col + 1).setValue(body.fields[key]);
  });
  if (body.fields.Status) { const col = headers.indexOf('ResolvedTime'); if (col !== -1) getSheet().getRange(rowNum, col + 1).setValue(timeStr); }
  if (body.fields.Payment) { const col = headers.indexOf('PayTime'); if (col !== -1) getSheet().getRange(rowNum, col + 1).setValue(timeStr); }
  return json({ success: true });
}

function deleteStop(body) {
  const rowNum = findRow(body.id); if (rowNum === -1) return json({ error: 'Stop not found' });
  getSheet().deleteRow(rowNum); return json({ success: true });
}

function copyRoute(body) {
  const now = new Date(); const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm a');
  const targetDate = body.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let copied = 0;
  for (const s of (body.stops || [])) {
    if (!s.name) continue;
    getSheet().appendRow([Utilities.getUuid(), targetDate, s.name || '', s.address || '', '', s.notes || '', 'pending', '', '', '', '', timeStr, '', '']);
    copied++;
  }
  return json({ success: true, copied });
}

// ===== REPORTS =====
function parseDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const str = String(val).trim();
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const serial = parseFloat(str);
  if (!isNaN(serial) && serial > 30000 && serial < 100000) {
    const d = new Date(1899, 11, 30); d.setDate(d.getDate() + serial);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  const d = new Date(str); if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return str;
}
function sendReport(daysBack, type) {
  try {
    const data = getReportData(daysBack);
    const html = buildReportHtml(data, type, daysBack);
    const now = new Date(); const start = new Date(now); start.setDate(now.getDate() - daysBack);
    const subject = `SDI ${type === 'monthly' ? 'Monthly' : 'Weekly'} Report — ${Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMM d')} to ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d')}`;
    MailApp.sendEmail({ to: REPORT_EMAIL, subject: subject, htmlBody: html, name: 'SDI Delivery Tracker' });
    return json({ success: true, message: `${type} report sent to ${REPORT_EMAIL}` });
  } catch (err) { return json({ error: String(err.message || err) }); }
}

function getReportData(daysBack) {
  const stops = getAllStops();
  const now = new Date(); now.setHours(12,0,0,0);
  const start = new Date(now); start.setDate(now.getDate() - daysBack); start.setHours(0,0,0,0);
  const filtered = stops.filter(s => {
    const dateStr = parseDateVal(s.Date);
    if (!dateStr) return false;
    const d = new Date(dateStr + 'T12:00:00');
    return !isNaN(d.getTime()) && d >= start && d <= now;
  });
  return {
    stops: filtered,
    delivered: filtered.filter(s => s.Status === 'delivered'),
    notDelivered: filtered.filter(s => s.Status === 'not-delivered'),
    pending: filtered.filter(s => s.Status === 'pending'),
    completed: filtered.filter(s => s.Status === 'completed'),
    paid: filtered.filter(s => s.Payment === 'paid'),
    unpaid: filtered.filter(s => s.Payment === 'unpaid'),
    missing: filtered.filter(s => {
      if (s.Status !== 'pending') return false;
      const dateStr = parseDateVal(s.Date);
      if (!dateStr) return false;
      const d = new Date(dateStr + 'T12:00:00');
      return ((now - d) / (1000*60*60)) >= 48;
    })
  };
}

function buildReportHtml(data, type, daysBack) {
  if (!data) return '<p>No data available.</p>';
  const { stops, delivered, notDelivered, pending, completed, paid, unpaid, missing } = data;
  const now = new Date(); const start = new Date(now); start.setDate(now.getDate() - daysBack);
  const range = `${Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMM d')} — ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy')}`;
  const title = type === 'monthly' ? 'SDI Monthly Delivery Report' : 'SDI Weekly Delivery Report';
  const total = stops.length;
  let html = `<div style="font-family:Inter,system-ui,sans-serif;max-width:640px;margin:0 auto;color:#1c1f1b"><h1 style="font-size:22px;margin-bottom:4px">${title}</h1><p style="color:#6b6f66;font-size:13px;margin:0 0 20px">${range}</p>`;
  html += `<div style="background:#fff;border:1px solid #dcdfd9;border-radius:14px;padding:16px;margin-bottom:20px">`;
  html += `<h2 style="font-size:15px;margin:0 0 12px;font-weight:700">Summary</h2>`;
  html += `<table style="width:100%;border-collapse:collapse;font-size:13px">`;
  const sumRow = (label, count, color) => `<tr><td style="padding:8px 0;border-bottom:1px solid #eee;font-weight:600">${label}</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;font-weight:700;color:${color}">${count}</td></tr>`;
  html += sumRow('Total Stops', total, '#1c1f1b');
  html += sumRow('Delivered', delivered.length, '#1d9e75');
  html += sumRow('Unpaid', unpaid.length, '#d23b3b');
  html += sumRow('Not Delivered', notDelivered.length, '#d23b3b');
  html += sumRow('Pending', pending.length, '#c9821a');
  html += sumRow('Missing', missing.length, '#7c3aed');
  html += `</table></div>`;
  const nameList = (arr) => { if (!arr.length) return ''; return arr.map(s => `${s.Name} (${String(s.Date || '').slice(0,10)})${s.PayMethod ? ' — ' + s.PayMethod : ''}`).join(', '); };
  const section = (title, count, color, items) => { if (!items.length) return ''; return `<div style="background:#fff;border:1px solid #dcdfd9;border-radius:14px;padding:16px;margin-bottom:12px;border-left:4px solid ${color}"><h3 style="font-size:14px;margin:0 0 8px;font-weight:700">${title} (${count})</h3><p style="font-size:12px;color:#6b6f66;margin:0;line-height:1.5">${nameList(items)}</p></div>`; };
  let paidHtml = '';
  if (paid.length) {
    paidHtml = `<div style="background:#fff;border:1px solid #dcdfd9;border-radius:14px;padding:16px;margin-bottom:12px;border-left:4px solid #1d9e75"><h3 style="font-size:14px;margin:0 0 8px;font-weight:700">Paid Orders (${paid.length})</h3>`;
    paidHtml += `<table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:#f4f5f3;font-weight:600;color:#6b6f66;text-align:left"><th style="padding:6px 10px">Date</th><th style="padding:6px 10px">Customer</th><th style="padding:6px 10px">Order #</th><th style="padding:6px 10px">Payment</th></tr></thead><tbody>`;
    paidHtml += paid.map(s => `<tr><td style="padding:6px 10px;border-bottom:1px solid #eee">${String(s.Date || '').slice(0,10)}</td><td style="padding:6px 10px;border-bottom:1px solid #eee">${s.Name}</td><td style="padding:6px 10px;border-bottom:1px solid #eee">${s.OrderNumber || '—'}</td><td style="padding:6px 10px;border-bottom:1px solid #eee">${s.PayMethod || '—'}</td></tr>`).join('');
    paidHtml += `</tbody></table></div>`;
  }
  html += paidHtml;
  html += section('Not Paid', unpaid.length, '#d23b3b', unpaid);
  html += section('Not Delivered', notDelivered.length, '#d23b3b', notDelivered);
  html += section('Missing', missing.length, '#7c3aed', missing);
  html += section('Pending', pending.length, '#c9821a', pending);
  html += section('Completed', completed.length, '#1d9e75', completed);
  html += `<p style="font-size:12px;color:#9a9d94;margin-top:24px;text-align:center">Generated automatically from SDI Delivery Tracker</p></div>`;
  return html;
}

// ===== AUTO-TRIGGERS =====
function weeklyReport() { sendReport(7, 'weekly'); }
function monthlyReport() { sendReport(30, 'monthly'); }
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'weeklyReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('weeklyReport').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).nearMinute(0).create();
  Logger.log('Weekly trigger: every Monday at 8:00 AM');
}
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'monthlyReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monthlyReport').timeBased().onMonthDay(1).atHour(8).nearMinute(0).create();
  Logger.log('Monthly trigger: 1st of every month at 8:00 AM');
}

// ===== AUTO-MOVE OVERDUE STOPS =====
function dailyAutoMove() {
  const sheet = getSheet(); const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0]; const dateCol = headers.indexOf('Date'); const statusCol = headers.indexOf('Status'); const notesCol = headers.indexOf('Notes'); const nameCol = headers.indexOf('Name');
  const today = new Date(); today.setHours(12, 0, 0, 0); const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let moved = 0, missing = 0; const movedNames = [], missingNames = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; const status = String(row[statusCol] || '');
    if (status !== 'pending' && status !== 'not-delivered') continue;
    const dateVal = String(row[dateCol] || '').trim();
    const stopDate = new Date(dateVal + 'T12:00:00'); if (isNaN(stopDate.getTime())) continue;
    const daysDiff = Math.round((today - stopDate) / (1000 * 60 * 60 * 24));
    if (daysDiff < 1) continue;
    const rowNum = i + 1; const name = String(row[nameCol] || 'Unknown');
    if (daysDiff === 1) { sheet.getRange(rowNum, dateCol + 1).setValue(todayStr); moved++; movedNames.push(name); }
    else { sheet.getRange(rowNum, statusCol + 1).setValue('missing'); const notes = String(row[notesCol] || ''); sheet.getRange(rowNum, notesCol + 1).setValue((notes ? notes + ' | ' : '') + 'Auto-marked missing after ' + daysDiff + ' days'); missing++; missingNames.push(name); }
  }
  if (moved || missing) {
    let body = `<div style="font-family:Inter,sans-serif;max-width:480px;margin:0 auto"><h2 style="font-size:18px">🔄 Daily Auto-Move</h2><p style="color:#6b6f66;font-size:13px">${Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMM d, yyyy')}</p>`;
    if (moved) { body += `<div style="background:#fbeed9;border-radius:8px;padding:12px;margin-bottom:12px"><p style="font-weight:600;font-size:13px;color:#6b4408">📦 ${moved} moved to today</p><ul style="font-size:12px">${movedNames.map(n => `<li>${n}</li>`).join('')}</ul></div>`; }
    if (missing) { body += `<div style="background:#f0e7fe;border-radius:8px;padding:12px;margin-bottom:12px"><p style="font-weight:600;font-size:13px;color:#3b1a7a">⚠️ ${missing} marked MISSING</p><ul style="font-size:12px">${missingNames.map(n => `<li>${n}</li>`).join('')}</ul></div>`; }
    body += `</div>`;
    MailApp.sendEmail({ to: REPORT_EMAIL, subject: `Daily Auto-Move: ${moved} moved, ${missing} missing`, htmlBody: body, name: 'SDI Delivery Tracker' });
  }
  Logger.log(`Auto-move: ${moved} moved, ${missing} missing`);
}
function setupDailyAutoMoveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailyAutoMove') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyAutoMove').timeBased().everyDays(1).atHour(1).nearMinute(0).create();
  Logger.log('Daily auto-move: every day at 1:00 AM');
}

// ===== MONTHLY RESET =====
function monthlyReset() {
  const sheet = getSheet(); const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0]; const dateCol = headers.indexOf('Date'); const statusCol = headers.indexOf('Status'); const notesCol = headers.indexOf('Notes'); const nameCol = headers.indexOf('Name');
  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDayOfPrevMonth = new Date(now.getFullYear(), now.getMonth(), 0);
  const firstDayStr = Utilities.formatDate(firstOfMonth, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lastDayStr = Utilities.formatDate(lastDayOfPrevMonth, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let carriedOver = 0; const carriedNames = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; const status = String(row[statusCol] || ''); const dateVal = parseDateVal(row[dateCol]);
    if (status === 'pending' && dateVal === lastDayStr) {
      sheet.getRange(i + 1, dateCol + 1).setValue(firstDayStr);
      const notes = String(row[notesCol] || '');
      sheet.getRange(i + 1, notesCol + 1).setValue((notes ? notes + ' | ' : '') + 'Carried over from ' + lastDayStr);
      carriedOver++; carriedNames.push(String(row[nameCol] || 'Unknown'));
    }
  }
  if (carriedOver > 0) {
    let emailBody = `<div style="font-family:Inter,sans-serif;max-width:480px;margin:0 auto"><h2 style="font-size:18px">📅 Monthly Reset</h2><p style="color:#6b6f66;font-size:13px">${Utilities.formatDate(firstOfMonth, Session.getScriptTimeZone(), 'MMMM d, yyyy')}</p>`;
    emailBody += `<div style="background:#fbeed9;border-radius:8px;padding:12px;margin-bottom:12px"><p style="font-weight:600;font-size:13px;color:#6b4408">📦 ${carriedOver} order(s) carried over from ${lastDayStr}</p><ul style="font-size:12px">${carriedNames.map(n => `<li>${n}</li>`).join('')}</ul></div></div>`;
    MailApp.sendEmail({ to: REPORT_EMAIL, subject: `Monthly Reset: ${carriedOver} orders carried over`, htmlBody: emailBody, name: 'SDI Delivery Tracker' });
  }
  Logger.log(`Monthly reset: ${carriedOver} carried over from ${lastDayStr} to ${firstDayStr}`);
}
function setupMonthlyResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'monthlyReset') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monthlyReset').timeBased().onMonthDay(1).atHour(2).nearMinute(0).create();
  Logger.log('Monthly reset trigger: 1st of every month at 2:00 AM');
}

function setupAllTriggers() {
  setupWeeklyTrigger(); setupMonthlyTrigger(); setupDailyAutoMoveTrigger(); setupMonthlyResetTrigger();
  Logger.log('All triggers set up!');
}

// Test email function (run once to authorize MailApp)
function testEmail() {
  MailApp.sendEmail({ to: 'sadadelivery1@gmail.com', subject: 'SDI Test Email', body: 'If you received this, email is working!' });
}
