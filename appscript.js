// ============================================================
//  DOST-3 2025 Assessment and Qualifying Form — Code.gs
//  Improved: buildSummaryHTML & sendEmail match email preview
// ============================================================

var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // <-- replace with your actual ID
var NOTIFY_EMAIL   = 'abenedicto431@gmail.com';
var SCRIPT_URL     = 'https://script.google.com/macros/s/AKfycbwH6DtPKnDmijeDshGmhk24A1MKrHcCABpdYTP-uZSP_VBlHMzgly998EwnfZu_OMWn/exec';
var EDIT_BASE_URL  = 'https://abenedicto431-glitch.github.io/DOST-FORMS/edit.html';

// ── Color palette (matches the HTML previews) ───────────────
var COLORS = {
  blue : '#1a3a6b',
  app  : '#27ae60',
  mpp  : '#2980b9',
  emp  : '#e67e22',
  fs   : '#9b59b6',
  grey : '#555555'
};

var PROG_LABELS = {
  app : 'APP - Agricultural Productivity Program',
  mpp : 'MPP - Manufacturing Productivity Program',
  emp : 'EMP - Energy Management Program',
  fs  : 'Food Safety Enrollment Form'
};

var PROG_SHORT = {
  'APP - Agricultural Productivity Program' : 'APP',
  'MPP - Manufacturing Productivity Program': 'MPP',
  'EMP - Energy Management Program'         : 'EMP',
  'Food Safety Enrollment Form'             : 'Food Safety'
};

var PROG_COLOR = {
  'APP - Agricultural Productivity Program' : COLORS.app,
  'MPP - Manufacturing Productivity Program': COLORS.mpp,
  'EMP - Energy Management Program'         : COLORS.emp,
  'Food Safety Enrollment Form'             : COLORS.fs
};

// ── Shared basic-info keys (rendered in "Basic Information") ─
var BASIC_KEYS = [
  'Name of Farm / Firm / Organization:',
  'Name of Farm:',
  'Name of Firm / Organization:',
  'Area of Farm (ha):',
  'Year Established:',
  'No. of Workers:',
  'Province:',
  'Owner / Chairman:',
  'Owner:',
  'Contact Person:',
  'Sex (M/F):',
  'Age:',
  'Position:',
  'Birthdate:',
  'Farm Complete Address:',
  'Complete Address:',
  'Contact Number / Mobile No.:',
  'E-mail Address:',
  'Special Classification'
];

// ── Keys that are NOT rendered in the per-program tables ─────
var SKIP_KEYS = [
  'applicant','programs','contact','email','timestamp',
  'rawData','summaryHTML','pre_assessment','token','_action',
  'Previous Consultations'
].concat(BASIC_KEYS);


// ============================================================
//  doPost — receives form submissions & edits
// ============================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById('1Kg37heNWOuBulPfkXBiZkEi15fociUv3adnJTMExBKo').getActiveSheet();

    // ── Handle edit/update ───────────────────────────────────
    if (data._action === 'update' && data.token) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][5] === data.token) {
          // Rebuild summary to make sure it's fresh
          data.summaryHTML = buildSummaryHTML(data);
          sheet.getRange(i + 1, 7).setValue(JSON.stringify(data));
          sendEmail(data, data.token);
          return jsonResponse({ success: true });
        }
      }
      return jsonResponse({ success: false, error: 'Token not found' });
    }

    // ── New submission ───────────────────────────────────────
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Applicant','Programs','Contact','Email','Token','Data']);
    }

    var token = Utilities.getUuid();

    // Always rebuild summary from raw data (not the client-side HTML)
    data.summaryHTML = buildSummaryHTML(data);

    sheet.appendRow([
      new Date().toLocaleString(),
      data.applicant || '',
      data.programs  || '',
      data.contact   || '',
      data.email     || '',
      token,
      JSON.stringify(data)
    ]);

    sendEmail(data, token);

    return jsonResponse({ success: true });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
//  doGet — view page & JSON bridge for edit page
// ============================================================
function doGet(e) {
  var token  = e.parameter.token;
  var action = e.parameter.action || 'view';

  if (!token) {
    return ContentService.createTextOutput('Invalid request.').setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.openById('1Kg37heNWOuBulPfkXBiZkEi15fociUv3adnJTMExBKo').getActiveSheet();
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
    var safe = JSON.stringify(data).replace(/'/g, "\\'").replace(/\n/g, ' ');
    var html = '<script>window.parent.postMessage(\'' + safe + '\', \'*\');<\/script>';
    return HtmlService.createHtmlOutput(html);
  }

  return ContentService.createTextOutput('Done.').setMimeType(ContentService.MimeType.TEXT);
}


// ============================================================
//  buildSummaryHTML — mirrors the email preview exactly
// ============================================================
function buildSummaryHTML(data) {
  var programs = (data.programs || '').split(', ').map(function(p){ return p.trim(); }).filter(Boolean);
  var html = '';

  // ── 1. Program pills bar ─────────────────────────────────
  html += '<div style="background:#f0f4ff;border-left:4px solid ' + COLORS.blue + ';padding:10px 14px;margin-bottom:16px;border-radius:0 4px 4px 0;">';
  html += '<strong>Programs Selected:</strong>';
  programs.forEach(function(prog) {
    var c = PROG_COLOR[prog] || COLORS.blue;
    var s = PROG_SHORT[prog] || prog;
    html += ' <span style="background:' + c + ';color:#fff;padding:2px 10px;border-radius:10px;font-size:11px;">' + s + '</span>';
  });
  html += '</div>';

  // ── 2. Basic Information ─────────────────────────────────
  html += secHead('Basic Information', COLORS.blue);
  html += '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
  BASIC_KEYS.forEach(function(k) {
    if (data[k] && data[k] !== '—') {
      html += row2col(k, data[k]);
    }
  });
  html += '</table>';

  // ── 3. Previous Consultations ────────────────────────────
  html += secHead('Previous Consultations', COLORS.blue);
  if (data['Previous Consultations']) {
    var pc = data['Previous Consultations'];
    if (pc.indexOf('Yes') === 0 || pc.indexOf('yes') === 0) {
      // Parse the " || " separated rows
      var detail = pc.replace(/^Yes[^:]*:?\s*/i, '');
      var rows = detail.split(' || ');
      var hasRows = rows.some(function(r){ return r.trim(); });
      if (hasRows) {
        html += '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
        html += '<thead><tr>';
        ['Agency / Company','Date of Assistance','Type of Assistance / Consultancy'].forEach(function(h) {
          html += '<th style="background:' + COLORS.blue + ';color:#fff;padding:5px 8px;text-align:left;border:1px solid #aaa;">' + h + '</th>';
        });
        html += '</tr></thead><tbody>';
        rows.forEach(function(row) {
          if (!row.trim()) return;
          var cols = row.split(' | ');
          html += '<tr>';
          for (var i = 0; i < 3; i++) {
            html += '<td style="padding:4px 8px;border:1px solid #eee;font-size:11px;">' + (cols[i] || '') + '</td>';
          }
          html += '</tr>';
        });
        html += '</tbody></table>';
      } else {
        html += '<p style="font-size:11px;color:#555;margin-bottom:12px;">Yes — no details provided.</p>';
      }
    } else {
      // No — show reason
      var reason = pc.replace(/^No\.?\s*(Reason:\s*)?/i, '');
      html += '<p style="font-size:11px;color:#555;margin-bottom:12px;"><strong>No prior consultation.</strong>' + (reason ? ' Reason: ' + reason : '') + '</p>';
    }
  } else {
    html += emptyNote();
  }

  // ── 4. Program-specific sections ────────────────────────
  programs.forEach(function(prog) {
    var c = PROG_COLOR[prog] || COLORS.blue;
    html += secHead(prog, c);

    var rows = [];

    // Render all non-skip keys that have values
    for (var key in data) {
      if (SKIP_KEYS.indexOf(key) > -1) continue;
      // Assign to only one program if key is shared — first match wins
      if (data[key] && data[key] !== '—') {
        rows.push(row2col(key, data[key]));
      }
    }

    if (rows.length) {
      html += '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
      // De-duplicate rows (same key may appear for multiple programs)
      var seen = {};
      rows.forEach(function(r) {
        if (!seen[r]) { seen[r] = true; html += r; }
      });
      html += '</table>';
    }

    // APP: Farm Data table
    if (prog === PROG_LABELS.app && data['Farm Data (D)']) {
      html += '<div style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:' + COLORS.blue + ';">D. Farm Data</div>';
      html += '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
      html += '<thead><tr>';
      ['Commodity','Variety/Breed','Area','Avg. Yield/Cropping Season','Avg. Income/Cropping Season','No. of Cropping/Year','Other Info'].forEach(function(h) {
        html += '<th style="background:' + COLORS.app + ';color:#fff;padding:4px 8px;text-align:left;font-size:10.5px;">' + h + '</th>';
      });
      html += '</tr></thead><tbody>';
      data['Farm Data (D)'].split(' || ').forEach(function(rowStr) {
        if (!rowStr.trim()) return;
        var cells = rowStr.replace(/^Row \d+:\s*/, '').split(' | ');
        html += '<tr>';
        for (var i = 0; i < 7; i++) {
          html += '<td style="padding:4px 8px;border:1px solid #eee;font-size:11px;">' + (cells[i] || '') + '</td>';
        }
        html += '</tr>';
      });
      html += '</tbody></table>';
    }

    if (!rows.length) html += emptyNote();
  });

  // ── 5. Pre-Assessment (DOST Staff) ──────────────────────
  html += secHead('Pre-Assessment (DOST Staff)', COLORS.grey);
  if (data['pre_assessment'] && data['pre_assessment'].trim()) {
    html += '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px;">';
    html += '<thead><tr>';
    ['Possible Areas of Interventions','Technical Needs / Problems','Initial Recommendations / Remarks'].forEach(function(h) {
      html += '<th style="background:' + COLORS.grey + ';color:#fff;padding:5px 8px;text-align:left;border:1px solid #aaa;">' + h + '</th>';
    });
    html += '</tr></thead><tbody>';
    data['pre_assessment'].split('\n').forEach(function(line) {
      if (!line.trim()) return;
      var parts = line.split(' | ');
      html += '<tr>';
      for (var i = 0; i < 3; i++) {
        html += '<td style="padding:4px 8px;border:1px solid #eee;font-size:11px;">' + (parts[i] || '') + '</td>';
      }
      html += '</tr>';
    });
    html += '</tbody></table>';
  } else {
    html += '<p style="font-size:11px;color:#aaa;font-style:italic;margin-bottom:12px;">Not yet filled by DOST Staff.</p>';
  }

  return html;
}

// ── HTML helpers ─────────────────────────────────────────────
function secHead(title, color) {
  return '<div style="background:' + color + ';color:#fff;padding:7px 12px;font-weight:bold;font-size:12px;border-radius:4px;margin:16px 0 8px;">' + title + '</div>';
}

function row2col(label, value) {
  return '<tr>' +
    '<td style="padding:4px 8px;border:1px solid #eee;font-weight:bold;color:#555;width:35%;background:#fafafa;">' + label + '</td>' +
    '<td style="padding:4px 8px;border:1px solid #eee;">' + value + '</td>' +
    '</tr>';
}

function emptyNote() {
  return '<p style="font-size:11px;color:#aaa;font-style:italic;margin-bottom:12px;">No information filled in for this section.</p>';
}


// ============================================================
//  sendEmail — styled exactly like the email preview HTML
// ============================================================
function sendEmail(data, token) {
  var programs = (data.programs || '').split(', ').map(function(p){ return p.trim(); }).filter(Boolean);
  var progStr  = programs.join(', ');

  var viewLink = SCRIPT_URL + '?token=' + token + '&action=view';
  var editLink = EDIT_BASE_URL + '?token=' + token;

  // Always use the server-built summary (strips client emoji / bad chars)
  var summaryHTML = buildSummaryHTML(data);

  var subject = 'DOST-3 New Submission — ' + (data.applicant || 'Applicant') + ' (' + progStr + ')';

  // Program pills for the meta block
  var pillsHTML = '';
  programs.forEach(function(prog) {
    var c = PROG_COLOR[prog] || COLORS.blue;
    var s = PROG_SHORT[prog] || prog;
    pillsHTML += '<span style="display:inline-block;background:' + c + ';color:#fff;padding:2px 10px;border-radius:10px;font-size:11px;margin-left:4px;">' + s + '</span>';
  });

  var body =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;">' +

    // Outer container
    '<div style="max-width:860px;margin:0 auto;background:#fff;border-radius:8px;' +
    'overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);">' +

    // Header
    '<div style="background:' + COLORS.blue + ';padding:20px 28px;text-align:center;">' +
    '<h2 style="color:#fff;margin:0;font-size:16px;">DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>' +
    '<p style="color:#a0c4ff;margin:4px 0 0;font-size:12px;">2025 Assessment and Qualifying Form — New Submission</p>' +
    '</div>' +

    // Meta block
    '<div style="background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;">' +
    '<p style="margin:0;"><strong>Applicant:</strong> ' + (data.applicant || '') + '</p>' +
    '<p style="margin:4px 0 0;"><strong>Programs:</strong> ' + pillsHTML + '</p>' +
    '<p style="margin:4px 0 0;"><strong>Submitted:</strong> ' + (data.timestamp || new Date().toLocaleString()) + '</p>' +
    '</div>' +

    // Action buttons
    '<div style="padding:16px 28px;text-align:center;border-bottom:1px solid #eee;">' +
    '<a href="' + viewLink + '" style="display:inline-block;background:' + COLORS.blue + ';color:#fff;' +
    'padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">' +
    'View and Download as PDF</a>' +
    '<a href="' + editLink + '" style="display:inline-block;background:' + COLORS.app + ';color:#fff;' +
    'padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">' +
    'Edit Submission</a>' +
    '<p style="font-size:11px;color:#888;margin-top:8px;">' +
    'To save as PDF: Click View &rarr; Print &rarr; Change destination to &ldquo;Save as PDF&rdquo;</p>' +
    '</div>' +

    // Summary content
    '<div style="padding:20px 28px;">' +
    summaryHTML +
    '</div>' +

    // Footer
    '<div style="background:' + COLORS.blue + ';padding:14px 28px;text-align:center;">' +
    '<p style="color:#a0c4ff;font-size:11px;margin:0;">DOST-3 2025 Assessment and Qualifying Form System</p>' +
    '</div>' +

    '</div></body></html>';
  Logger.log(showCarCodes(body));

  GmailApp.sendEmail(NOTIFY_EMAIL, subject, '', { htmlBody: body });
}


// ============================================================
//  buildViewPage — for the GAS-hosted view link
// ============================================================
function buildViewPage(data, token) {
  var editLink    = EDIT_BASE_URL + '?token=' + token;
  var summaryHTML = buildSummaryHTML(data);

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<title>DOST-3 Form Submission</title>' +
    '<style>' +
    'body{font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;}' +
    '.wrap{max-width:860px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1);}' +
    '.hdr{background:#1a3a6b;color:#fff;padding:20px 28px;text-align:center;}' +
    '.hdr h2{margin:0;font-size:16px;}' +
    '.hdr p{margin:4px 0 0;font-size:12px;color:#a0c4ff;}' +
    '.meta{background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;}' +
    '.acts{padding:16px 28px;text-align:center;border-bottom:1px solid #eee;}' +
    '.content{padding:20px 28px;}' +
    '.ftr{background:#1a3a6b;padding:14px 28px;text-align:center;}' +
    '.ftr p{color:#a0c4ff;font-size:11px;margin:0;}' +
    '.btn{display:inline-block;padding:10px 24px;border-radius:4px;font-weight:bold;font-size:13px;' +
    'text-decoration:none;color:#fff;margin:4px;cursor:pointer;border:none;}' +
    '@media print{.no-print{display:none!important;}}' +
    '</style></head><body>' +

    '<div class="wrap">' +

    '<div class="hdr"><h2>DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>' +
    '<p>2025 Assessment and Qualifying Form — Submission Details</p></div>' +

    '<div class="meta">' +
    '<strong>Applicant:</strong> ' + (data.applicant || '—') + '&nbsp;&nbsp;|&nbsp;&nbsp;' +
    '<strong>Programs:</strong> ' + (data.programs || '—') + '&nbsp;&nbsp;|&nbsp;&nbsp;' +
    '<strong>Submitted:</strong> ' + (data.timestamp || '—') +
    '</div>' +

    '<div class="acts no-print">' +
    '<button class="btn" style="background:#1a3a6b;" onclick="window.print()">Print / Download as PDF</button>' +
    '<a class="btn" style="background:#27ae60;" href="' + editLink + '">Edit Submission</a>' +
    '<p style="font-size:11px;color:#888;margin-top:8px;">To save as PDF: Click Print &rarr; Change destination to &ldquo;Save as PDF&rdquo;</p>' +
    '</div>' +

    '<div class="content">' + summaryHTML + '</div>' +

    '<div class="ftr no-print"><p>DOST-3 2025 Assessment and Qualifying Form System</p></div>' +

    '</div></body></html>';
  return html;
  function showCharCodes(str) {
  return str.split('').map(function(c) {
    return c + ' (' + c.charCodeAt(0) + ')';
  }).join(' ');
}

  return html;
}
