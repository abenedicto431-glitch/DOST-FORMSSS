// ─────────────────────────────────────────────────────────────────────────────
// DOST-3 2025 Assessment and Qualifying Form — Google Apps Script Backend
// ─────────────────────────────────────────────────────────────────────────────

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // <-- replace with your actual ID
var NOTIFY_EMAIL   = 'abenedicto431@gmail.com';
var EDIT_BASE_URL  = 'https://abenedicto431-glitch.github.io/DOST-FORMS/edit.html';
var GAS_BASE_URL   = 'https://script.google.com/macros/s/AKfycbwH6DtPKnDmijeDshGmhk24A1MKrHcCABpdYTP-uZSP_VBlHMzgly998EwnfZu_OMWn/exec';

// ─────────────────────────────────────────────────────────────────────────────
// doPost — handles new submissions and edits
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();

    // ── UPDATE existing submission ──────────────────────────────────────────
    if (data._action === 'update' && data.token) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][5] === data.token) {
          sheet.getRange(i + 1, 7).setValue(JSON.stringify(data));
          sendEmail(data, data.token, true);
          return ContentService
            .createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Token not found.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── NEW submission ──────────────────────────────────────────────────────
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Applicant', 'Programs', 'Contact', 'Email', 'Token', 'Data']);
    }

    var token = Utilities.getUuid();

    sheet.appendRow([
      new Date().toLocaleString('en-PH'),
      data.applicant || '',
      data.programs  || '',
      data.contact   || '',
      data.email     || '',
      token,
      JSON.stringify(data)
    ]);

    sendEmail(data, token, false);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// doGet — handles view and getjson actions
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  var token  = e.parameter.token;
  var action = e.parameter.action || 'view';

  if (!token) {
    return ContentService.createTextOutput('Invalid request.').setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  var rows  = sheet.getDataRange().getValues();
  var found = null;

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][5] === token) { found = rows[i]; break; }
  }

  if (!found) {
    return ContentService.createTextOutput('Submission not found.').setMimeType(ContentService.MimeType.TEXT);
  }

  var data = JSON.parse(found[6]);

  if (action === 'view') {
    return HtmlService.createHtmlOutput(buildViewPage(data, token));
  }

  if (action === 'getjson') {
    var jsonData = JSON.stringify(data);
    var html = '<script>window.parent.postMessage(\'' +
      jsonData.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\n/g, ' ') +
      '\', \'*\');<\/script>';
    return HtmlService.createHtmlOutput(html);
  }

  return ContentService.createTextOutput('Done.').setMimeType(ContentService.MimeType.TEXT);
}

// ─────────────────────────────────────────────────────────────────────────────
// sendEmail — sends a nicely formatted email to DOST staff
// ─────────────────────────────────────────────────────────────────────────────
function sendEmail(data, token, isEdit) {
  var viewLink = GAS_BASE_URL + '?token=' + token + '&action=view';
  var editLink = EDIT_BASE_URL + '?token=' + token;

  var applicant = data.applicant || 'Applicant';
  var programs  = data.programs  || '';
  var submitted = data.timestamp || new Date().toLocaleString('en-PH');

  var subject = isEdit
    ? 'DOST-3 Updated Submission — ' + applicant + ' (' + programs + ')'
    : 'DOST-3 New Submission — '     + applicant + ' (' + programs + ')';

  var summaryHTML = buildSummaryHTML(data);

  var body = '';
  body += '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>';
  body += '<body style="margin:0;padding:20px;background:#e8edf5;font-family:Arial,sans-serif;font-size:12px;">';
  body += '<div style="max-width:900px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);">';

  // Header
  body += '<div style="background:#1a3a6b;padding:20px 28px;text-align:center;">';
  body += '<h2 style="color:#fff;margin:0;font-size:16px;">DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>';
  body += '<p style="color:#a0c4ff;margin:4px 0 0;font-size:12px;">2025 Assessment and Qualifying Form — ' + (isEdit ? 'Updated Submission' : 'New Submission') + '</p>';
  body += '</div>';

  // Meta
  body += '<div style="background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;">';
  body += '<p style="margin:0;color:#333;"><strong>Applicant:</strong> ' + applicant + '</p>';
  body += '<p style="margin:4px 0 0;color:#333;"><strong>Programs:</strong> ' + programs + '</p>';
  body += '<p style="margin:4px 0 0;color:#333;"><strong>' + (isEdit ? 'Updated' : 'Submitted') + ':</strong> ' + submitted + '</p>';
  body += '</div>';

  // Action buttons
  body += '<div style="padding:16px 28px;text-align:center;border-bottom:1px solid #eee;">';
  body += '<a href="' + viewLink + '" style="display:inline-block;background:#1a3a6b;color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">View and Download as PDF</a>';
  body += '<a href="' + editLink + '" style="display:inline-block;background:#27ae60;color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">Edit Submission</a>';
  body += '<p style="font-size:11px;color:#888;margin-top:8px;">To save as PDF: Click View &rarr; Print &rarr; Change destination to &ldquo;Save as PDF&rdquo;</p>';
  body += '</div>';

  // Summary content
  body += '<div style="padding:20px 28px;">';
  body += summaryHTML;
  body += '</div>';

  // Footer
  body += '<div style="background:#1a3a6b;padding:14px 28px;text-align:center;">';
  body += '<p style="color:#a0c4ff;font-size:11px;margin:0;">DOST-3 2025 Assessment and Qualifying Form System</p>';
  body += '</div>';

  body += '</div></body></html>';

  GmailApp.sendEmail(NOTIFY_EMAIL, subject, '', { htmlBody: body });
}

// ─────────────────────────────────────────────────────────────────────────────
// buildSummaryHTML — builds the full summary used in both email and view page
// ─────────────────────────────────────────────────────────────────────────────
function buildSummaryHTML(data) {
  var progColors = {
    'APP - Agricultural Productivity Program' : '#27ae60',
    'MPP - Manufacturing Productivity Program': '#2980b9',
    'EMP - Energy Management Program'         : '#e67e22',
    'Food Safety Enrollment Form'             : '#9b59b6'
  };

  var programs = (data.programs || '').split(', ');
  var html = '';

  // ── Program pills ─────────────────────────────────────────────────────────
  html += '<div style="background:#f0f4ff;border-left:4px solid #1a3a6b;padding:10px 14px;margin-bottom:16px;border-radius:0 4px 4px 0;">';
  html += '<strong style="color:#333;">Programs Selected:</strong> ';
  programs.forEach(function (p) {
    p = p.trim();
    if (!p) return;
    var c = progColors[p] || '#1a3a6b';
    html += '<span style="background:' + c + ';color:#fff;padding:2px 10px;border-radius:10px;font-size:11px;margin-left:4px;">' + p + '</span>';
  });
  html += '</div>';

  // ── Helpers ───────────────────────────────────────────────────────────────
  function secHead(title, color) {
    return '<div style="background:' + color + ';color:#fff;padding:7px 12px;font-weight:bold;font-size:12px;border-radius:4px;margin:16px 0 8px;">' + title + '</div>';
  }
  function subLabel(title, color) {
    return '<div style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:' + color + ';">' + title + '</div>';
  }
  function tableStart() {
    return '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
  }
  function row2col(label, value) {
    return '<tr>' +
      '<td style="padding:4px 8px;border:1px solid #eee;font-weight:bold;color:#555;width:35%;background:#fafafa;">' + label + '</td>' +
      '<td style="padding:4px 8px;border:1px solid #eee;color:#333;">' + (value || '—') + '</td>' +
      '</tr>';
  }
  function thCell(label, color) {
    return '<th style="background:' + color + ';color:#fff;padding:5px 8px;text-align:left;border:1px solid #aaa;font-size:11px;">' + label + '</th>';
  }
  function tdCell(value) {
    return '<td style="padding:4px 8px;border:1px solid #eee;color:#333;font-size:11px;">' + (value || '—') + '</td>';
  }
  function emptyNote() {
    return '<p style="font-size:11px;color:#aaa;font-style:italic;margin-bottom:12px;">No information filled in for this section.</p>';
  }
  function get(key) { return data[key] || ''; }

  // ── I. Basic Information ──────────────────────────────────────────────────
  html += secHead('I. Basic Information', '#1a3a6b');
  html += tableStart();

  var entityName = get('Name of Farm / Firm / Organization:') || get('Name of Farm:') || get('Name of Firm / Organization:');
  if (entityName)           html += row2col('Name of Farm / Firm / Organization:', entityName);
  if (get('Area of Farm (ha):'))         html += row2col('Area of Farm (ha):',         get('Area of Farm (ha):'));
  if (get('Year Established:'))          html += row2col('Year Established:',           get('Year Established:'));
  if (get('No. of Workers:'))            html += row2col('No. of Workers:',             get('No. of Workers:'));
  if (get('Province:'))                  html += row2col('Province:',                   get('Province:'));

  var ownerName = get('Owner / Chairman:') || get('Owner:');
  if (ownerName)                         html += row2col('Owner / Chairman:',           ownerName);
  if (get('Special Classification'))     html += row2col('Special Classification',      get('Special Classification'));
  if (get('Contact Person:'))            html += row2col('Contact Person:',             get('Contact Person:'));
  if (get('Sex (M/F):'))                 html += row2col('Sex (M/F):',                  get('Sex (M/F):'));
  if (get('Age:'))                       html += row2col('Age:',                         get('Age:'));
  if (get('Position:'))                  html += row2col('Position:',                   get('Position:'));
  if (get('Birthdate:'))                 html += row2col('Birthdate:',                  get('Birthdate:'));

  var address = get('Complete Address:') || get('Farm Complete Address:');
  if (address)                           html += row2col('Complete Address:',           address);
  if (get('Contact Number / Mobile No.:')) html += row2col('Contact Number / Mobile No.:', get('Contact Number / Mobile No.:'));
  if (get('E-mail Address:'))            html += row2col('E-mail Address:',             get('E-mail Address:'));

  html += '</table>';

  // ── II. Previous Consultations ────────────────────────────────────────────
  html += secHead('II. Previous Consultations', '#1a3a6b');
  var prevCon = get('Previous Consultations');
  if (!prevCon) {
    html += emptyNote();
  } else if (prevCon.indexOf('Yes') === 0) {
    // Parse pipe-delimited consultation rows
    var detail = prevCon.replace(/^Yes\s*[—-]?\s*/i, '');
    var conRows = detail.split(' || ');
    var hasRows = false;
    var conTable = tableStart();
    conTable += '<thead><tr>';
    conTable += thCell('Agency / Company', '#1a3a6b');
    conTable += thCell('Date of Assistance', '#1a3a6b');
    conTable += thCell('Type of Assistance / Consultancy', '#1a3a6b');
    conTable += '</tr></thead><tbody>';
    conRows.forEach(function (row) {
      if (!row.trim()) return;
      var cols = row.split(' | ');
      if (cols.some(function (c) { return c.trim(); })) {
        hasRows = true;
        conTable += '<tr>' + tdCell(cols[0]) + tdCell(cols[1]) + tdCell(cols[2]) + '</tr>';
      }
    });
    conTable += '</tbody></table>';
    html += hasRows ? conTable : '<p style="font-size:11px;color:#555;margin-bottom:12px;">Yes — but no details provided.</p>';
  } else {
    html += '<p style="font-size:11px;color:#555;margin-bottom:12px;"><strong>No prior consultation.</strong> ' +
      prevCon.replace(/^No\.?\s*(Reason:\s*)?/i, 'Reason: ') + '</p>';
  }

  // ── III. Program-Specific ─────────────────────────────────────────────────
  programs.forEach(function (prog) {
    prog = prog.trim();
    if (!prog) return;
    var c = progColors[prog] || '#1a3a6b';

    // ── APP ──────────────────────────────────────────────────────────────
    if (prog === 'APP - Agricultural Productivity Program') {
      html += secHead('APP — Agricultural Productivity Program', c);
      html += tableStart();
      if (get('Brief Farm Background:'))  html += row2col('Brief Farm Background:',  get('Brief Farm Background:'));
      if (get('Commodity 1:'))            html += row2col('Commodity 1:',            get('Commodity 1:'));
      if (get('Commodity 2:'))            html += row2col('Commodity 2:',            get('Commodity 2:'));
      if (get('Commodity 3:'))            html += row2col('Commodity 3:',            get('Commodity 3:'));
      html += '</table>';

      if (get('Farm Data (D)')) {
        html += subLabel('D. Farm Data', '#1a3a6b');
        html += tableStart();
        html += '<thead><tr>';
        ['Commodity','Variety/Breed','Area','Avg. Yield/Cropping Season','Avg. Income/Cropping Season','No. of Cropping/Year','Other Info']
          .forEach(function (h) { html += thCell(h, c); });
        html += '</tr></thead><tbody>';
        get('Farm Data (D)').split(' || ').forEach(function (rowStr) {
          var vals = rowStr.replace(/^Row \d+:\s*/, '').split(' | ');
          if (vals.some(function (v) { return v.trim(); })) {
            html += '<tr>';
            for (var i = 0; i < 7; i++) html += tdCell(vals[i] || '');
            html += '</tr>';
          }
        });
        html += '</tbody></table>';
      }
    }

    // ── MPP ──────────────────────────────────────────────────────────────
    if (prog === 'MPP - Manufacturing Productivity Program') {
      html += secHead('MPP — Manufacturing Productivity Program', c);
      html += tableStart();
      var mppKeys = [
        'Brief Firm Background:', 'Year Firm Established:', 'Initial Capital (Php):',
        'Company Registration No.:', 'Year Registered:',
        'Annual Volume of Production:', 'Estimated Value (Php):',
        'Existing Foreign Market:', 'Existing Local Market:', 'Target Additional Market:',
        'Specific products/services offered to customers:',
        "Firm's plan (next 5 years):", "Firm's plan (next 10 years):",
        'Organizational Chart:'
      ];
      mppKeys.forEach(function (k) { if (get(k)) html += row2col(k, get(k)); });

      // Selected checkboxes / radios stored under program key
      var mppSel = get('MPP - Manufacturing Productivity Program - Selected');
      if (mppSel) html += row2col('Selected Options:', mppSel);
      html += '</table>';
    }

    // ── EMP ──────────────────────────────────────────────────────────────
    if (prog === 'EMP - Energy Management Program') {
      html += secHead('EMP — Energy Management Program', c);
      html += tableStart();
      var empKeys = [
        'Brief Firm Background:', 'Year Firm Established:', 'Initial Capital (Php):',
        'Company Registration No.:', 'Year Registered:',
        'Annual Volume of Production:', 'Estimated Value (Php):',
        'Specific products/services offered to customers:',
        'Technical Problems/Needs in regards to Energy Consumption:',
        'Total kWh (1 year):', 'Total Fuel Cost, Php (1 year):', 'Combined Total (Php/year):'
      ];
      empKeys.forEach(function (k) { if (get(k)) html += row2col(k, get(k)); });
      html += '</table>';
    }

    // ── Food Safety ───────────────────────────────────────────────────────
    if (prog === 'Food Safety Enrollment Form') {
      html += secHead('Food Safety — Enrollment Form', c);

      // Registrations
      var regKeys = [
        ['DOLE/CDA/SEC Registration Reg. No.', 'DOLE/CDA/SEC Registration'],
        ['DTI Registration Reg. No.',           'DTI Registration'],
        ['Business Permit Reg. No.',            'Business Permit'],
        ['FDA License to Operate Reg. No.',     'FDA License to Operate'],
        ['FDA CPR — Product 1',                 'FDA CPR — Product 1'],
        ['FDA CPR — Product 2',                 'FDA CPR — Product 2'],
        ['Other Registration',                  'Other Registration']
      ];
      var hasReg = regKeys.some(function (r) { return !!get(r[0]); });
      if (hasReg) {
        html += subLabel('Business Registration &amp; Certifications', c);
        html += tableStart();
        regKeys.forEach(function (r) { if (get(r[0])) html += row2col(r[1] + ':', get(r[0])); });
        html += '</table>';
      }

      // General FS fields
      html += tableStart();
      var fsKeys = [
        'No. of Employees (M):', 'No. of Employees (F):',
        'Local Market Details:', 'Foreign Market Details:', 'Target Additional Market:',
        'Products/Services 1:', 'Products/Services 2:', 'Products/Services 3:',
        'Remarks:'
      ];
      fsKeys.forEach(function (k) { if (get(k)) html += row2col(k, get(k)); });
      var fsSel = get('Food Safety Enrollment Form - Selected');
      if (fsSel) html += row2col('Selected Options:', fsSel);
      html += '</table>';
    }
  });

  // ── IV. Pre-Assessment ────────────────────────────────────────────────────
  html += secHead('IV. Pre-Assessment (DOST Staff)', '#555');
  var pa = get('pre_assessment');
  if (pa && pa.trim()) {
    html += tableStart();
    html += '<thead><tr>';
    html += thCell('Possible Areas of Interventions', '#555');
    html += thCell('Technical Needs / Problems', '#555');
    html += thCell('Initial Recommendations / Remarks', '#555');
    html += '</tr></thead><tbody>';
    pa.split('\n').forEach(function (line) {
      if (!line.trim()) return;
      var parts = line.split(' | ');
      html += '<tr>' + tdCell(parts[0]) + tdCell(parts[1]) + tdCell(parts[2]) + '</tr>';
    });
    html += '</tbody></table>';
  } else {
    html += '<p style="font-size:11px;color:#aaa;font-style:italic;">Not yet filled by DOST Staff.</p>';
  }

  return html;
}

// ─────────────────────────────────────────────────────────────────────────────
// buildViewPage — the HTML page shown when clicking "View and Download as PDF"
// ─────────────────────────────────────────────────────────────────────────────
function buildViewPage(data, token) {
  var editLink    = EDIT_BASE_URL + '?token=' + token;
  var summaryHTML = buildSummaryHTML(data);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<title>DOST-3 Form Submission</title>';
  html += '<style>';
  html += 'body{font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;font-size:12px;}';
  html += '.container{max-width:900px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);}';
  html += '.hdr{background:#1a3a6b;padding:20px 28px;text-align:center;}';
  html += '.hdr h2{color:#fff;margin:0;font-size:16px;}';
  html += '.hdr p{color:#a0c4ff;margin:4px 0 0;font-size:12px;}';
  html += '.meta{background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;}';
  html += '.meta p{margin:0;color:#333;} .meta p+p{margin-top:4px;}';
  html += '.actions{padding:16px 28px;text-align:center;border-bottom:1px solid #eee;}';
  html += '.btn{display:inline-block;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;color:#fff;}';
  html += '.btn-blue{background:#1a3a6b;} .btn-green{background:#27ae60;}';
  html += '.content{padding:20px 28px;}';
  html += '.footer{background:#1a3a6b;padding:14px 28px;text-align:center;}';
  html += '.footer p{color:#a0c4ff;font-size:11px;margin:0;}';
  html += '@media print{.no-print{display:none!important;} body{background:#fff;} .container{box-shadow:none;}}';
  html += '</style></head><body>';

  html += '<div class="container">';

  html += '<div class="hdr"><h2>DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>';
  html += '<p>2025 Assessment and Qualifying Form — Submission Details</p></div>';

  html += '<div class="meta">';
  html += '<p><strong>Applicant:</strong> ' + (data.applicant || '—') + '</p>';
  html += '<p><strong>Programs:</strong> '  + (data.programs  || '—') + '</p>';
  html += '<p><strong>Submitted:</strong> ' + (data.timestamp || '—') + '</p>';
  html += '</div>';

  html += '<div class="actions no-print">';
  html += '<button class="btn btn-blue" onclick="window.print()">Print / Download as PDF</button>';
  html += '<a class="btn btn-green" href="' + editLink + '">Edit Submission</a>';
  html += '<p style="font-size:11px;color:#888;margin-top:8px;">To save as PDF: Click Print &rarr; Change destination to &ldquo;Save as PDF&rdquo;</p>';
  html += '</div>';

  html += '<div class="content">' + summaryHTML + '</div>';

  html += '<div class="footer no-print"><p>DOST-3 2025 Assessment and Qualifying Form System</p></div>';
  html += '</div></body></html>';

  return html;
}
