function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID').getActiveSheet();

    // Handle update action
    if(data._action === 'update' && data.token) {
      var rows = sheet.getDataRange().getValues();
      for(var i = 1; i < rows.length; i++) {
        if(rows[i][5] === data.token) {
          sheet.getRange(i+1, 7).setValue(JSON.stringify(data));
          // Send updated email
          sendEmail(data, data.token);
          return ContentService
            .createTextOutput(JSON.stringify({ success: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // Add headers if sheet is empty
    if(sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Applicant', 'Programs', 'Contact', 'Email', 'Token', 'Data']);
    }

    // Generate unique token for view/edit
    var token = Utilities.getUuid();

    // Save to sheet
    sheet.appendRow([
      new Date().toLocaleString(),
      data.applicant || '',
      data.programs || '',
      data.contact || '',
      data.email || '',
      token,
      JSON.stringify(data)
    ]);

    // Send email
    sendEmail(data, token);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var token = e.parameter.token;
  var action = e.parameter.action || 'view';

  if(!token) {
    return ContentService.createTextOutput('Invalid request.').setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID').getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var found = null;

  for(var i = 1; i < rows.length; i++) {
    if(rows[i][5] === token) {
      found = rows[i];
      break;
    }
  }

  if(!found) {
    return ContentService.createTextOutput('Submission not found.').setMimeType(ContentService.MimeType.TEXT);
  }

  var data = JSON.parse(found[6]);

  if(action === 'view') {
    return HtmlService.createHtmlOutput(buildViewPage(data, token));
  }

  if(action === 'getjson') {
    var jsonData = JSON.stringify(data);
    var html = '<script>window.parent.postMessage(\''+jsonData.replace(/'/g,"\\'").replace(/\n/g,' ')+'\', \'*\');<\/script>';
    return HtmlService.createHtmlOutput(html);
  }

  return ContentService.createTextOutput('Done.').setMimeType(ContentService.MimeType.TEXT);
}

function buildViewPage(data, token) {
  var progColors = {
    'APP - Agricultural Productivity Program': '#27ae60',
    'MPP - Manufacturing Productivity Program': '#2980b9',
    'EMP - Energy Management Program': '#e67e22',
    'Food Safety Enrollment Form': '#9b59b6'
  };

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<title>DOST-3 Form Submission</title>';
  html += '<style>';
  html += 'body{font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;}';
  html += '.container{max-width:800px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);}';
  html += '.header{background:#1a3a6b;color:#fff;padding:20px 28px;}';
  html += '.header h2{margin:0;font-size:16px;}';
  html += '.header p{margin:4px 0 0;font-size:12px;color:#a0c4ff;}';
  html += '.meta{background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;}';
  html += '.content{padding:24px 28px;}';
  html += '.sec-head{padding:8px 14px;font-weight:bold;font-size:13px;border-radius:4px;margin:16px 0 8px;color:#fff;}';
  html += '.blue{background:#1a3a6b;}';
  html += 'table{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px;}';
  html += 'th{background:#555;color:#fff;padding:6px 10px;text-align:left;}';
  html += 'td{padding:6px 10px;border:1px solid #eee;}';
  html += 'td:first-child{background:#f9f9f9;font-weight:bold;width:35%;}';
  html += '.btn{display:inline-block;padding:10px 24px;border-radius:4px;font-weight:bold;font-size:13px;cursor:pointer;border:none;color:#fff;margin:4px;}';
  html += '.btn-blue{background:#1a3a6b;}';
  html += '.btn-green{background:#27ae60;}';
  html += '.btn-orange{background:#e67e22;}';
  html += '.footer{background:#1a3a6b;padding:14px 28px;text-align:center;}';
  html += '.footer p{color:#a0c4ff;font-size:11px;margin:0;}';
  html += '@media print{.no-print{display:none!important;}}';
  html += '</style></head><body>';

  html += '<div class="container">';

  // Header
  html += '<div class="header">';
  html += '<h2>DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>';
  html += '<p>2025 Assessment and Qualifying Form — Submission Details</p>';
  html += '</div>';

  // Meta
  html += '<div class="meta">';
  html += '<strong>Applicant:</strong> ' + (data.applicant||'—') + '&nbsp;&nbsp;|&nbsp;&nbsp;';
  html += '<strong>Programs:</strong> ' + (data.programs||'—') + '&nbsp;&nbsp;|&nbsp;&nbsp;';
  html += '<strong>Submitted:</strong> ' + (data.timestamp||new Date().toLocaleString());
  html += '</div>';

  // Action buttons
  html += '<div class="content no-print">';
  html += '<div style="text-align:center;margin-bottom:20px;">';
  html += '<button class="btn btn-blue" onclick="window.print()">Print / Download as PDF</button>';
  html += '<a class="btn btn-green" href="https://abenedicto431-glitch.github.io/DOST-FORMS/edit.html?token=' + token + '">Edit Submission</a>';
  html += '</div>';
  html += '<p style="text-align:center;font-size:11px;color:#888;margin-bottom:20px;">To save as PDF: Click Print → Change destination to "Save as PDF"</p>';
  html += '</div>';

  // Basic Info
  html += '<div class="content">';
  html += '<div class="sec-head blue">Basic Information</div>';
  html += '<table>';
  var skip = ['applicant','programs','contact','email','timestamp','rawData'];
  var basicKeys = ['Name of Farm / Firm / Organization:','Name of Farm:','Name of Firm / Organization:',
    'Area of Farm (ha):','Year Established:','No. of Workers:','Owner / Chairman:','Owner:',
    'Contact Person:','Sex (M/F):','Age:','Position:','Complete Address:','Farm Complete Address:',
    'Contact Number / Mobile No.:','E-mail Address:','Province:','Birthdate:','Special Classification'];
  basicKeys.forEach(function(k){
    if(data[k]) html += '<tr><td>'+k+'</td><td>'+data[k]+'</td></tr>';
  });
  html += '</table>';

  // Programs
  var programs = (data.programs||'').split(', ');
  programs.forEach(function(prog){
    prog = prog.trim();
    if(!prog) return;
    var color = progColors[prog] || '#1a3a6b';
    html += '<div class="sec-head" style="background:'+color+';">'+prog+'</div>';
    html += '<table>';
    for(var key in data) {
      if(skip.indexOf(key) > -1) continue;
      if(basicKeys.indexOf(key) > -1) continue;
      if(data[key] && data[key] !== '—') {
        html += '<tr><td>'+key+'</td><td>'+data[key]+'</td></tr>';
      }
    }
    html += '</table>';
  });

  // Pre-Assessment
  html += '<div class="sec-head" style="background:#555;">Pre-Assessment (DOST Staff)</div>';
  if(data['pre_assessment']) {
    html += '<table><tr><th>Possible Areas</th><th>Technical Needs</th><th>Recommendations</th></tr>';
    html += '<tr><td colspan="3">' + data['pre_assessment'] + '</td></tr></table>';
  } else {
    html += '<p style="color:#aaa;font-style:italic;font-size:12px;">Not yet filled by DOST Staff.</p>';
  }

  html += '</div>';

  // Footer
  html += '<div class="footer no-print">';
  html += '<p>DOST-3 2025 Assessment and Qualifying Form System</p>';
  html += '</div>';

  html += '</div></body></html>';
  return html;
}

function sendEmail(data, token) {
  var viewLink = 'https://script.google.com/macros/s/AKfycbwH6DtPKnDmijeDshGmhk24A1MKrHcCABpdYTP-uZSP_VBlHMzgly998EwnfZu_OMWn/exec?token=' + token + '&action=view';
  var editLink = 'https://abenedicto431-glitch.github.io/DOST-FORMS/edit.html?token=' + token;
  var summaryHTML = (data.summaryHTML || '<p style="color:#aaa;">No summary available.</p>').replace(/[\uFFFD\u{1F000}-\u{1FFFF}\u{2600}-\u{27FF}]/gu, '');
  var subject = 'DOST-3 New Submission - ' + (data.applicant||'Applicant') + ' (' + (data.programs||'') + ')';

  var body = '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;">';
  body += '<div style="max-width:860px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);">';

  // Header
  body += '<div style="background:#1a3a6b;padding:20px 28px;text-align:center;">';
  body += '<h2 style="color:#fff;margin:0;font-size:16px;">DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>';
  body += '<p style="color:#a0c4ff;margin:4px 0 0;font-size:12px;">2025 Assessment and Qualifying Form — New Submission</p>';
  body += '</div>';

  // Meta
  body += '<div style="background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;">';
  body += '<p style="margin:0;"><strong>Applicant:</strong> '+(data.applicant||'—')+'</p>';
  body += '<p style="margin:4px 0 0;"><strong>Programs:</strong> '+(data.programs||'—')+'</p>';
  body += '<p style="margin:4px 0 0;"><strong>Submitted:</strong> '+(data.timestamp||new Date().toLocaleString())+'</p>';
  body += '</div>';

  // Action Buttons
  body += '<div style="padding:16px 28px;text-align:center;border-bottom:1px solid #eee;">';
  body += '<a href="'+viewLink+'" style="display:inline-block;background:#1a3a6b;color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">View &amp; Download as PDF</a>';
  body += '<a href="'+editLink+'" style="display:inline-block;background:#27ae60;color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:13px;margin:4px;">Edit Submission</a>';
  body += '</div>';

  // Exact summary from the form
  body += '<div style="padding:20px 28px;">';
  body += summaryHTML;
  body += '</div>';

  // Footer
  body += '<div style="background:#1a3a6b;padding:14px 28px;text-align:center;">';
  body += '<p style="color:#a0c4ff;font-size:11px;margin:0;">DOST-3 2025 Assessment and Qualifying Form System</p>';
  body += '</div>';

  body += '</div></body></html>';
  GmailApp.sendEmail('abenedicto431@gmail.com', subject, '', { htmlBody: body });
}

function buildViewPage(data, token) {
  var editLink = 'https://abenedicto431-glitch.github.io/DOST-FORMS/edit.html?token=' + token;
  var summaryHTML = data.summaryHTML || '<p style="color:#aaa;">No summary available.</p>';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<title>DOST-3 Form Submission</title>';
  html += '<style>';
  html += 'body{font-family:Arial,sans-serif;background:#e8edf5;margin:0;padding:20px;}';
  html += '.container{max-width:860px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);}';
  html += '.header{background:#1a3a6b;color:#fff;padding:20px 28px;text-align:center;}';
  html += '.header h2{margin:0;font-size:16px;}';
  html += '.header p{margin:4px 0 0;font-size:12px;color:#a0c4ff;}';
  html += '.meta{background:#f0f4ff;padding:14px 28px;border-bottom:1px solid #dde;font-size:12px;}';
  html += '.actions{padding:16px 28px;text-align:center;border-bottom:1px solid #eee;}';
  html += '.content{padding:20px 28px;}';
  html += '.footer{background:#1a3a6b;padding:14px 28px;text-align:center;}';
  html += '.footer p{color:#a0c4ff;font-size:11px;margin:0;}';
  html += '.btn{display:inline-block;padding:10px 24px;border-radius:4px;font-weight:bold;font-size:13px;text-decoration:none;color:#fff;margin:4px;cursor:pointer;border:none;}';
  html += '@media print{.no-print{display:none!important;}}';
  html += '</style></head><body>';

  html += '<div class="container">';
  html += '<div class="header"><h2>DEPARTMENT OF SCIENCE AND TECHNOLOGY</h2>';
  html += '<p>2025 Assessment and Qualifying Form — Submission Details</p></div>';

  html += '<div class="meta">';
  html += '<strong>Applicant:</strong> '+(data.applicant||'—')+'&nbsp;&nbsp;|&nbsp;&nbsp;';
  html += '<strong>Programs:</strong> '+(data.programs||'—')+'&nbsp;&nbsp;|&nbsp;&nbsp;';
  html += '<strong>Submitted:</strong> '+(data.timestamp||'—');
  html += '</div>';

  html += '<div class="actions no-print">';
  html += '<button class="btn" style="background:#1a3a6b;" onclick="window.print()">Print / Download as PDF</button>';
  html += '<a class="btn" style="background:#27ae60;" href="'+editLink+'">Edit Submission</a>';
  html += '<p style="font-size:11px;color:#888;margin-top:8px;">To save as PDF: Click Print → Change destination to "Save as PDF"</p>';
  html += '</div>';

  html += '<div class="content">';
  html += summaryHTML;
  html += '</div>';

  html += '<div class="footer no-print"><p>DOST-3 2025 Assessment and Qualifying Form System</p></div>';
  html += '</div></body></html>';
  return html;
}
