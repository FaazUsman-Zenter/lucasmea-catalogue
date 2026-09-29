/**
 * @OnlyCurrentDoc
 *
 * Lucas online catalogues: Request for Price handler (Google Apps Script)
 * -------------------------------------------------------------------------
 * What it does, for every request sent from lucasmeacatalogue.com/cart/:
 *   1. Emails the sales team, with the parts list and the Excel file attached.
 *   2. Emails the customer a confirmation with their reference number and parts.
 *   3. Saves the Excel file into a master Google Drive folder:
 *        RFQ Inquiries (master) / Lucas online catalogue / 2026-09 / LMEA-260929-AB12 - Company.xlsx
 *   4. Adds a row to the "Requests" tab of this Google Sheet (a running log with a link
 *      to each Excel file). testSetup moves this Sheet into the master folder too.
 *
 * Setup (about 10 minutes, once):
 *   1. Sign in to the Google account that should send the emails.
 *   2. Create a new Google Sheet, e.g. "Lucas RFQ log".
 *   3. Extensions > Apps Script. Delete everything in Code.gs and paste this whole file.
 *   4. Check the CONFIG block below (sales email address etc.) and click Save.
 *   5. In the function menu at the top choose "testSetup" and click Run.
 *      Approve the permissions Google asks for. A test email arrives at SALES_EMAIL.
 *   6. Deploy > New deployment > type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *      Click Deploy and copy the Web app URL (ends in /exec).
 *   7. Send that URL to whoever maintains the website. It goes into cart/index.html
 *      as RFQ_ENDPOINT.
 *
 * If you edit this code later: Deploy > Manage deployments > edit (pencil) >
 * Version: New version > Deploy. The URL stays the same.
 */

var CONFIG = {
  SALES_EMAIL: 'info@lucasmeaparts.com',   // where requests go; separate several with commas
  REPLY_TO: 'info@lucasmeaparts.com',      // where customers' replies to the confirmation go
  FROM_NAME: 'Lucas Online Catalogue',     // sender name shown in both emails
  SITE_URL: 'https://lucasmeacatalogue.com',
  LOGO_URL: 'https://lucasmeacatalogue.com/assets/lucas-logo.png',
  SEND_CUSTOMER_EMAIL: true,
  ATTACH_EXCEL_TO_CUSTOMER: true,          // false = customer gets the email without the Excel file
  MAX_REQUESTS_PER_EMAIL_PER_HOUR: 5,      // simple protection against abuse
  LOG_SHEET_NAME: 'Requests',
  MASTER_FOLDER_NAME: 'RFQ Inquiries (master)', // created in My Drive on first run
  MASTER_FOLDER_ID: '',                    // optional: paste the ID of an existing (e.g. shared) folder to use instead
  SOURCE_FOLDER: 'Lucas online catalogue'  // sub-folder for this website (Stravik can get its own later)
};

var GREEN = '#00954C', BLACK = '#231F20', GREY = '#6D6E71', LIGHT = '#EDEDED';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function doPost(e) {
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.botcheck) return json_({ success: true }); // honeypot ticked: drop silently

    var req = {
      ref: clean_(d.ref, 40),
      name: clean_(d.name, 120),
      company: clean_(d.company, 160),
      email: clean_(d.email, 160).toLowerCase(),
      phone: clean_(d.phone, 60),
      country: clean_(d.country, 80),
      type: clean_(d.type, 60),
      message: clean_(d.message, 3000),
      page: clean_(d.page, 200),
      link: /^https:\/\/lucasmeacatalogue\.com\/rfq\/#[\w-]+$/.test(String(d.link || '')) ? String(d.link).slice(0, 60000) : ''
    };
    if (!/^LMEA-\d{6}-[A-Z0-9]{4}$/.test(req.ref)) throw new Error('Invalid reference');
    if (!req.name) throw new Error('Name is required');
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(req.email)) throw new Error('A valid email is required');

    var items = (Array.isArray(d.items) ? d.items : []).slice(0, 500).map(function (it) {
      return {
        no: clean_(it.no, 40), catalogue: clean_(it.catalogue, 60), make: clean_(it.make, 80),
        app: clean_(it.app, 300), oe: clean_(it.oe, 80),
        qty: Math.max(1, Math.min(99999, parseInt(it.qty, 10) || 1))
      };
    }).filter(function (it) { return it.no; });
    if (!items.length && !req.message) throw new Error('Nothing to send');

    // Throttle and de-duplicate (double clicks, refreshes)
    var cache = CacheService.getScriptCache();
    if (cache.get('ref:' + req.ref)) return json_({ success: true, ref: req.ref, duplicate: true });
    var key = 'n:' + req.email, n = Number(cache.get(key) || 0);
    if (n >= CONFIG.MAX_REQUESTS_PER_EMAIL_PER_HOUR) throw new Error('Too many requests from this email address. Please try again later or email us directly.');

    var attachments = [];
    if (d.excel && d.excel.base64) {
      var bytes = Utilities.base64Decode(String(d.excel.base64));
      if (bytes.length > 3 * 1024 * 1024) throw new Error('Attachment too large');
      var fname = (clean_(d.excel.name, 80) || ('Lucas-RFQ-' + req.ref + '.xlsx')).replace(/[^\w.\-]/g, '_');
      attachments.push(Utilities.newBlob(bytes, XLSX_MIME, fname));
    }

    var totalQty = items.reduce(function (s, it) { return s + it.qty; }, 0);
    var who = req.company || req.name;

    // Save the Excel file to the master Drive folder (never blocks the emails)
    req.fileUrl = '';
    if (attachments.length) {
      try { req.fileUrl = saveToDrive_(req, attachments[0]); } catch (err) { console.warn('Drive save failed: ' + err); }
    }

    // 1. Sales team
    MailApp.sendEmail({
      to: CONFIG.SALES_EMAIL,
      subject: (items.length ? 'Price request ' : 'Enquiry ') + req.ref + ' | ' + items.length + ' part' + (items.length === 1 ? '' : 's') + ' | ' + who,
      htmlBody: salesHtml_(req, items, totalQty, attachments.length > 0),
      body: plain_(req, items, totalQty),
      replyTo: req.email,
      name: CONFIG.FROM_NAME,
      attachments: attachments
    });

    // 2. Customer confirmation
    var customerSent = false;
    if (CONFIG.SEND_CUSTOMER_EMAIL) {
      try {
        MailApp.sendEmail({
          to: req.email,
          subject: 'We have received your request, reference ' + req.ref,
          htmlBody: customerHtml_(req, items, totalQty),
          body: 'Thank you ' + req.name + '.\n\nWe have received your ' + (items.length ? 'request for price' : 'enquiry') +
                '. Your reference number is ' + req.ref + '.\nA sales agent will contact you shortly.\n\n' + plain_(req, items, totalQty),
          replyTo: CONFIG.REPLY_TO,
          name: CONFIG.FROM_NAME,
          attachments: CONFIG.ATTACH_EXCEL_TO_CUSTOMER ? attachments : []
        });
        customerSent = true;
      } catch (err) { console.warn('Customer email failed: ' + err); }
    }

    cache.put('ref:' + req.ref, '1', 21600);
    cache.put(key, String(n + 1), 3600);
    try { log_(req, items, totalQty, customerSent); } catch (err) { console.warn('Log failed: ' + err); }

    return json_({ success: true, ref: req.ref, confirmation: customerSent });
  } catch (err) {
    console.error(err);
    return json_({ success: false, message: String(err && err.message || err) });
  }
}

function doGet() { return json_({ ok: true, service: 'Lucas RFQ handler' }); }

/** Run this once from the editor to grant permissions and send yourself a test email. */
function testSetup() {
  var master = masterFolder_();
  child_(master, CONFIG.SOURCE_FOLDER);
  try { sheet_(); } catch (err) {}
  try { // keep the log Sheet inside the master folder
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) DriveApp.getFileById(ss.getId()).moveTo(master);
  } catch (err) { Logger.log('Could not move the Sheet into the folder: ' + err); }
  MailApp.sendEmail({
    to: CONFIG.SALES_EMAIL, name: CONFIG.FROM_NAME,
    subject: 'Lucas RFQ handler: test email',
    htmlBody: '<p>The request-for-price handler is set up and can send email.</p>' +
      '<p>Master folder: <a href="' + master.getUrl() + '">' + esc_(CONFIG.MASTER_FOLDER_NAME) + '</a></p>' +
      '<p>Remaining daily email quota: ' + MailApp.getRemainingDailyQuota() + '</p>'
  });
  Logger.log('Test email sent to ' + CONFIG.SALES_EMAIL + '. Master folder: ' + master.getUrl() + ' . Remaining quota today: ' + MailApp.getRemainingDailyQuota());
}

/* ---------------- master folder ---------------- */

function masterFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = CONFIG.MASTER_FOLDER_ID || props.getProperty('MASTER_FOLDER_ID');
  if (id) { try { var f0 = DriveApp.getFolderById(id); if (!f0.isTrashed()) return f0; } catch (err) {} }
  var it = DriveApp.getFoldersByName(CONFIG.MASTER_FOLDER_NAME);
  var f = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.MASTER_FOLDER_NAME);
  props.setProperty('MASTER_FOLDER_ID', f.getId());
  return f;
}
function child_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function saveToDrive_(req, blob) {
  var month = Utilities.formatDate(new Date(), 'Asia/Dubai', 'yyyy-MM');
  var folder = child_(child_(masterFolder_(), CONFIG.SOURCE_FOLDER), month);
  var who = (req.company || req.name).replace(/[\\/:*?"<>|#%]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  var file = folder.createFile(blob.copyBlob().setName(req.ref + ' - ' + who + '.xlsx'));
  file.setDescription('Request for price ' + req.ref + ' from ' + req.name + ' <' + req.email + '>');
  return file.getUrl();
}

/* ---------------- helpers ---------------- */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 500); }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function nl2br_(s) { return esc_(s).replace(/\n/g, '<br>'); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Dubai', 'dd/MM/yyyy HH:mm') + ' (UAE)'; }

function plain_(req, items, totalQty) {
  var lines = items.map(function (it, i) {
    return (i + 1) + '. ' + it.no + '  x ' + it.qty + '  |  ' + it.catalogue + '  |  ' + [it.make, it.app].filter(String).join(' ') + (it.oe ? '  |  OE ' + it.oe : '');
  });
  return 'Reference: ' + req.ref + '\nName: ' + req.name + '\nCompany: ' + req.company + '\nEmail: ' + req.email +
    '\nPhone: ' + req.phone + '\nCountry: ' + req.country + '\nCustomer type: ' + req.type +
    '\n\nParts (' + items.length + ' lines, ' + totalQty + ' pcs):\n' + (lines.join('\n') || '(none, general enquiry)') +
    '\n\nMessage:\n' + (req.message || '-');
}

function partsTable_(items, totalQty) {
  if (!items.length) return '';
  var th = 'style="background:' + GREEN + ';color:#fff;font:bold 12px Arial,sans-serif;text-align:left;padding:8px 10px;"';
  var td = 'style="border-bottom:1px solid #DCDCDC;font:13px Arial,sans-serif;color:' + BLACK + ';padding:8px 10px;vertical-align:top;"';
  var rows = items.map(function (it, i) {
    return '<tr' + (i % 2 ? ' style="background:#F7F7F7"' : '') + '>' +
      '<td ' + td + '>' + (i + 1) + '</td>' +
      '<td ' + td + '><b style="color:' + GREEN + '">' + esc_(it.no) + '</b><br><span style="color:' + GREY + ';font-size:12px">' + esc_(it.catalogue) + '</span></td>' +
      '<td ' + td + '>' + esc_(it.make) + (it.app ? '<br><span style="color:' + GREY + ';font-size:12px">' + esc_(it.app) + '</span>' : '') + '</td>' +
      '<td ' + td + '>' + esc_(it.oe || '-') + '</td>' +
      '<td ' + td + ' align="center"><b>' + it.qty + '</b></td></tr>';
  }).join('');
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #DCDCDC;margin:8px 0 4px">' +
    '<tr><th ' + th + '>#</th><th ' + th + '>Lucas no.</th><th ' + th + '>Vehicle / application</th><th ' + th + '>OE no.</th><th ' + th + ' align="center">Qty</th></tr>' +
    rows +
    '<tr><td colspan="4" style="background:#E6F4ED;font:bold 13px Arial,sans-serif;padding:8px 10px;text-align:right;color:' + BLACK + '">Total</td>' +
    '<td align="center" style="background:#E6F4ED;font:bold 13px Arial,sans-serif;padding:8px 10px;color:' + BLACK + '">' + totalQty + '</td></tr></table>';
}

function detailsTable_(req) {
  var rows = [['Reference', '<b>' + esc_(req.ref) + '</b>'], ['Name', esc_(req.name)], ['Company', esc_(req.company || '-')],
    ['Email', '<a href="mailto:' + esc_(req.email) + '" style="color:' + GREEN + '">' + esc_(req.email) + '</a>'],
    ['Phone / WhatsApp', esc_(req.phone || '-')], ['Country', esc_(req.country || '-')], ['Customer type', esc_(req.type || '-')],
    ['Message', nl2br_(req.message || '-')]];
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;margin:8px 0">' +
    rows.map(function (r) {
      return '<tr><td style="width:150px;background:' + LIGHT + ';font:bold 12px Arial,sans-serif;color:' + GREY + ';padding:7px 10px;border:1px solid #DCDCDC;vertical-align:top">' + r[0] +
        '</td><td style="font:13px Arial,sans-serif;color:' + BLACK + ';padding:7px 10px;border:1px solid #DCDCDC">' + r[1] + '</td></tr>';
    }).join('') + '</table>';
}

function shell_(inner) {
  return '<div style="background:#F4F4F4;padding:20px 0">' +
    '<table role="presentation" align="center" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:#fff;border-collapse:collapse">' +
    '<tr><td style="background:' + GREEN + ';padding:0;height:42px"><img src="' + CONFIG.LOGO_URL + '" width="208" height="42" alt="Lucas" style="display:block;border:0"></td></tr>' +
    '<tr><td style="padding:24px 24px 8px;font:14px/1.5 Arial,sans-serif;color:' + GREY + '">' + inner + '</td></tr>' +
    '<tr><td style="padding:16px 24px 20px;border-top:4px solid ' + GREEN + ';font:12px Arial,sans-serif;color:' + GREY + '">' +
    '<a href="' + CONFIG.SITE_URL + '" style="color:' + GREEN + ';font-weight:bold;text-decoration:none">' + CONFIG.SITE_URL.replace(/^https?:\/\//, '') + '</a>' +
    '<span style="float:right;color:' + GREEN + ';font-weight:bold;font-size:13px">OUR RANGE GOES FURTHER.</span></td></tr>' +
    '<tr><td style="background:' + GREEN + ';height:14px;line-height:14px;font-size:0">&nbsp;</td></tr></table></div>';
}

function salesHtml_(req, items, totalQty, hasExcel) {
  return shell_(
    '<h1 style="margin:0 0 6px;font:bold 22px Arial,sans-serif;color:' + GREEN + '">' + (items.length ? 'New request for price' : 'New enquiry') + '</h1>' +
    '<p style="margin:0 0 14px">Received ' + now_() + ' from the online catalogue. Reply to this email to answer the customer directly.' +
    (hasExcel ? ' The parts list is attached as an Excel file, with columns for price and lead time.' : '') + '</p>' +
    (req.fileUrl ? '<p style="margin:0 0 14px">Saved to the master folder: <a href="' + esc_(req.fileUrl) + '" style="color:' + GREEN + '">open the Excel file in Google Drive</a>.</p>' : '') +
    (req.link ? '<p style="margin:0 0 14px"><a href="' + esc_(req.link) + '" style="display:inline-block;background:' + GREEN + ';color:#fff;font:bold 14px Arial,sans-serif;text-decoration:none;padding:10px 18px;border-radius:4px">Open request &amp; download Excel</a></p>' : '') +
    detailsTable_(req) +
    (items.length ? '<h2 style="margin:18px 0 4px;font:bold 16px Arial,sans-serif;color:' + GREEN + '">Parts (' + items.length + ' lines, ' + totalQty + ' pcs)</h2>' + partsTable_(items, totalQty) : ''));
}

function customerHtml_(req, items, totalQty) {
  var first = esc_(req.name.split(' ')[0]);
  return shell_(
    '<h1 style="margin:0 0 8px;font:bold 24px Arial,sans-serif;color:' + GREEN + '">Thank you, ' + first + '</h1>' +
    '<p style="margin:0 0 14px">We have received your ' + (items.length ? 'request for price' : 'enquiry') + '. A sales agent will contact you shortly.</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 16px"><tr>' +
    '<td style="background:#E6F4ED;border:1px dashed ' + GREEN + ';padding:12px 16px;font:12px Arial,sans-serif;color:' + GREY + '">YOUR REFERENCE NUMBER<br>' +
    '<span style="font:bold 22px Arial,sans-serif;color:' + BLACK + ';letter-spacing:1px">' + esc_(req.ref) + '</span></td></tr></table>' +
    '<p style="margin:0 0 14px">Please quote this reference number in any correspondence.</p>' +
    (items.length ? '<h2 style="margin:18px 0 4px;font:bold 16px Arial,sans-serif;color:' + GREEN + '">Your parts</h2>' + partsTable_(items, totalQty) : '') +
    '<h2 style="margin:18px 0 4px;font:bold 16px Arial,sans-serif;color:' + GREEN + '">Your details</h2>' + detailsTable_(req) +
    '<p style="margin:16px 0 8px">If anything is wrong, just reply to this email.</p>' +
    '<p style="margin:0 0 8px;font-size:12px">Dimensions are for reference only. Always confirm against the physical part before ordering.</p>');
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sh = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sh.appendRow(['Received (UAE)', 'Reference', 'Name', 'Company', 'Email', 'Phone / WhatsApp', 'Country', 'Customer type',
      'Lines', 'Total qty', 'Parts', 'Excel file', 'Message', 'Confirmation sent', 'Status', 'Assigned to', 'Notes']);
    sh.getRange(1, 1, 1, 17).setFontWeight('bold').setBackground(GREEN).setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(11, 420);
  }
  return sh;
}

function log_(req, items, totalQty, customerSent) {
  var sh = sheet_();
  if (!sh) return;
  var parts = items.map(function (it) { return it.no + ' x' + it.qty; }).join(', ');
  // Store customer text as plain text so nothing typed into the form can run as a formula
  var t = function (v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; };
  sh.appendRow([now_(), req.ref, t(req.name), t(req.company), t(req.email), t(req.phone), t(req.country), t(req.type),
    items.length, totalQty, t(parts), '', t(req.message), customerSent ? 'Yes' : 'No', 'New', '', '']);
  if (req.fileUrl) sh.getRange(sh.getLastRow(), 12).setFormula('=HYPERLINK("' + req.fileUrl.replace(/"/g, '') + '","Open Excel")');
}
