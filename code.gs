/**
 * KARIM N TRUCKING LTD — POD SYSTEM BACKEND
 * ------------------------------------------
 * This is a drop-in replacement for your existing Apps Script backend.
 * It keeps doing everything it already did — writing each POD to the
 * Google Sheet and saving photos to a linked Drive folder — and adds
 * ONE new thing: after a POD is saved, it automatically emails a copy
 * to dispatch AND cc's the driver, so a copy stays with both sides.
 *
 * HOW TO APPLY THIS UPDATE (2 minutes):
 * 1. Open the Google Sheet this system already writes to.
 * 2. Extensions > Apps Script.
 * 3. Select all the existing code and delete it.
 * 4. Paste this entire file in its place.
 * 5. Near the top, set DISPATCH_EMAIL to the address dispatch wants
 *    to receive these at.
 * 6. Save (Ctrl+S / Cmd+S).
 * 7. Deploy > Manage deployments > click the pencil/edit icon on the
 *    existing deployment > Version: "New version" > Deploy.
 *    (Using "New version" on the SAME deployment keeps the exact same
 *    Web App URL — you do NOT need to change anything in the HTML file.)
 * 8. The first time it tries to send mail, Google may ask you to
 *    re-authorize — that's expected, click through and allow it.
 */
 
const SHEET_NAME = 'PODs';
const FOLDER_NAME = 'Karim N Trucking PODs — Photos';
const DISPATCH_EMAIL = 'karimpods@gmail.cm';
 
// ===================== ROUTES =====================
 function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'ping') return jsonOut({ success: true, message: 'POD backend is live' });
    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}
 
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return submitPOD(payload);
  } catch (err) {
    return jsonOut({ success: false, message: String(err) });
  }
}
 
// ===================== CORE =====================
 
const HEADERS = [
  'Ticket No', 'Driver Name', 'Driver Email', 'Load Start Date', 'Delivery Date',
  'Reference Type', 'Reference Number', 'Location(s)', 'Notes', 'Photo URLs', 'Submitted At'
];
 
// Turns a "YYYY-MM-DD" load start date into a sheet tab name like "July 2026".
function getMonthSheetName(dateStr) {
  const parts = String(dateStr).split('-'); // ['2026','07','05'] — parsed manually to avoid timezone shift
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const monthIndex = parseInt(parts[1], 10) - 1;
  const year = parts[0];
  return monthNames[monthIndex] + ' ' + year;
}
 
// Gets (or creates) the sheet tab for the month a POD's Load Start Date falls in.
// Every monthly tab gets the exact same header row as row 1, so columns always
// line up with the data underneath them.
function getSheetForMonth(loadStartDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = getMonthSheetName(loadStartDate);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    const headerRange = sh.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#10233B');
    headerRange.setFontColor('#FFFFFF');
    sh.autoResizeColumns(1, HEADERS.length);
  }
  return sh;
}
 
function getFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}
 
function submitPOD(payload) {
  const sh = getSheetForMonth(payload.loadStartDate);
  const folder = getFolder();
 
  // Upload photos to Drive, collect view URLs
  const photoUrls = [];
  (payload.photos || []).forEach(function (p) {
    const bytes = Utilities.base64Decode(p.base64);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', p.name || (payload.ticketNo + '.jpg'));
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    photoUrls.push('https://drive.google.com/uc?export=view&id=' + file.getId());
  });
 
  const now = new Date().toISOString();
  const locations = (payload.locations || []).join(' | ');
 
  sh.appendRow([
    payload.ticketNo,
    payload.driverName,
    payload.driverEmail,
    payload.loadStartDate,
    payload.deliveryDate,
    payload.refType,
    payload.refNumber,
    locations,
    payload.notes || '',
    photoUrls.join(' | '),
    now
  ]);
 
  // Email a copy to dispatch, with the driver CC'd so it stays with them too.
  let emailSent = false;
  let emailError = '';
  try {
    if (DISPATCH_EMAIL.indexOf('PASTE_DISPATCH_EMAIL_HERE') > -1) {
      throw new Error('DISPATCH_EMAIL is still set to the placeholder value — edit it near the top of the script.');
    }
    sendConfirmationEmail(payload, photoUrls);
    emailSent = true;
  } catch (mailErr) {
    // The POD is already safely saved to the Sheet at this point — a mail
    // hiccup (e.g. missing permission, bad address, daily quota) should
    // never make the submission look like it failed outright. But we DO
    // report it back so the app can tell the driver honestly, instead of
    // failing silently.
    emailError = String(mailErr);
    Logger.log('Email send failed: ' + emailError);
  }
 
  return jsonOut({ success: true, ticketNo: payload.ticketNo, emailSent: emailSent, emailError: emailError });
}
 
function sendConfirmationEmail(payload, photoUrls) {
  const refLabel = { TL: 'TL / Transplace #', BOL: 'BOL / Waybill #', RM: 'RM #', OTHER: 'Reference #' }[payload.refType] || 'Reference #';
  const locations = (payload.locations || []).join(', ');
 
  const subject = 'POD ' + payload.ticketNo + ' — ' + payload.driverName + ' — ' + locations;
 
  const lines = [
    'NEW POD SUBMISSION',
    '',
    'Ticket #: ' + payload.ticketNo,
    'Driver: ' + payload.driverName,
    'Driver Email: ' + payload.driverEmail,
    'Load Start Date: ' + payload.loadStartDate,
    'Delivery Date: ' + payload.deliveryDate,
    refLabel + ': ' + payload.refNumber,
    'Location(s): ' + locations,
    'Notes: ' + (payload.notes || '(none)'),
    '',
    'Photos (' + photoUrls.length + '):',
  ].concat(photoUrls.length ? photoUrls : ['(none attached)']);
 
  const plainBody = lines.join('\n');
 
  const htmlPhotos = photoUrls.map(function (u) {
    return '<div style="margin-bottom:10px;"><img src="' + u + '" style="max-width:320px;border-radius:8px;border:1px solid #ddd;display:block;"></div>';
  }).join('');
 
  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:520px;">' +
    '<h2 style="margin:0 0 12px;">POD ' + escapeHtml_(payload.ticketNo) + '</h2>' +
    '<table style="border-collapse:collapse;width:100%;margin-bottom:16px;">' +
    row_('Driver', payload.driverName) +
    row_('Driver Email', payload.driverEmail) +
    row_('Load Start Date', payload.loadStartDate) +
    row_('Delivery Date', payload.deliveryDate) +
    row_(refLabel, payload.refNumber) +
    row_('Location(s)', locations) +
    row_('Notes', payload.notes || '(none)') +
    '</table>' +
    (photoUrls.length ? '<h3 style="margin:0 0 8px;">Photos</h3>' + htmlPhotos : '<p>(no photos attached)</p>') +
    '</div>';
 
  GmailApp.sendEmail(DISPATCH_EMAIL, subject, plainBody, {
    cc: payload.driverEmail,
    htmlBody: htmlBody,
    name: 'Karim N Trucking LTD — POD System'
  });
}
 
function row_(label, value) {
  return '<tr>' +
    '<td style="padding:6px 10px;border-bottom:1px solid #eee;color:#666;white-space:nowrap;">' + escapeHtml_(label) + '</td>' +
    '<td style="padding:6px 10px;border-bottom:1px solid #eee;font-weight:600;">' + escapeHtml_(String(value)) + '</td>' +
    '</tr>';
}
 
function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
 
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
 
/**
 * RUN THIS ONCE, MANUALLY, FROM THE APPS SCRIPT EDITOR — not from the app.
 *
 * Web app calls (from the HTML file) can't trigger Google's permission
 * popup. If email isn't sending, this is almost always why: the script
 * has never been explicitly granted permission to send Gmail on your
 * behalf, even though it's already authorized for the Sheet and Drive.
 *
 * How to run it:
 * 1. In the Apps Script editor, use the function dropdown at the top
 *    (next to the "Run" button) and select "authorizeGmailAccess".
 * 2. Click "Run".
 * 3. Google will show a permission screen — click through it and
 *    allow Gmail access. (You may see an "unverified app" warning
 *    since this is your own private script — click "Advanced" then
 *    "Go to [project name] (unsafe)" to proceed; this is normal and
 *    expected for scripts you wrote yourself.)
 * 4. Check your inbox — you should get a test email within a minute.
 * 5. Once that works, submissions from the app will start emailing
 *    correctly too.
 */
function authorizeGmailAccess() {
  if (DISPATCH_EMAIL.indexOf('PASTE_DISPATCH_EMAIL_HERE') > -1) {
    throw new Error('Set DISPATCH_EMAIL near the top of this file before running this.');
  }
  GmailApp.sendEmail(DISPATCH_EMAIL, 'POD system — test email', 'If you got this, Gmail sending is authorized correctly and the app is ready to email submissions.');
}
