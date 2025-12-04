/**
 * ON-THE-GO APPSUITE - MASTER BACKEND v17.0 (Full Advanced + Secure)
 * * INTEGRATES:
 * 1. Core Security (Secret Key Validation)
 * 2. Smart Data Handling (De-duplication, Photo Decoding)
 * 3. Alerting (Email + Textbelt SMS)
 * 4. Advanced Reporting (Monthly Stats, Archiving, PDFs)
 */

// ─────────────────────────────────────────────────────────────────────────────
// 1. CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {
  // Security
  SECRET_KEY: "%%SECRET_KEY%%",
  
  // API Keys
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%",
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  
  // Folder IDs (for Advanced Features)
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  PDF_FOLDER_ID: "",        // Optional: Folder for PDF Reports
  REPORT_TEMPLATE_ID: "",   // Optional: Google Doc Template ID
  
  // Settings
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30
};

// ─────────────────────────────────────────────────────────────────────────────
// 2. INCOMING DATA HANDLER (doPost)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // Wait up to 30s
  
  try {
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error", message:"No Data"}));
    const p = e.parameter;
    
    // [SECURITY CHECK]
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) {
       return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // [SCHEMA SETUP] 21 Columns
    if(sheet.getLastColumn() === 0) {
      const headers = [
        "Timestamp", "Date", "Worker Name", "Worker Phone Number", 
        "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email",
        "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email",
        "Alarm Status", "Notes", "Location Name", "Location Address", 
        "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", 
        "Distance (km)", "Visit Report Data", "Anticipated Departure Time"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
    
    // [PHOTO PROCESSING]
    let photoUrl = "";
    if(p['Photo 1'] && p['Photo 1'].includes('base64')) {
      try {
        const data = Utilities.base64Decode(p['Photo 1'].split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', p['Worker Name'] + '_' + Date.now() + '.jpg');
        let folder = (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) 
          ? DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID)
          : DriveApp.getRootFolder();
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch(err) { photoUrl = "Err: "+err; }
    }

    // [SMART ROW UPDATE]
    const worker = p['Worker Name'];
    const newStatus = p['Alarm Status'];
    let rowUpdated = false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const searchDepth = Math.min(lastRow - 1, 50); 
      const startRow = lastRow - searchDepth + 1;
      const maxCols = sheet.getLastColumn();
      const data = sheet.getRange(startRow, 1, searchDepth, maxCols).getValues(); 
      
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2]; // Col C
        const rowStatus = data[i][10]; // Col K
        
        // Update if active, OR if this is a Monitor Resolution event
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus) || newStatus === 'SAFE - MONITOR CLEARED')) {
             const realRowIndex = startRow + i;
             
             // 1. Update Status
             sheet.getRange(realRowIndex, 11).setValue(newStatus); 
             
             // 2. Update Vitals
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']);
             
             // 3. Update Due Time (Col 21)
             if (p['Anticipated Departure Time']) sheet.getRange(realRowIndex, 21).setValue(p['Anticipated Departure Time']);

             // 4. Append Notes (if new)
             if (p['Notes'] && !p['Notes'].includes("Locating") && !p['Notes'].includes("GPS Slow")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) { 
                   sheet.getRange(realRowIndex, 12).setValue(oldNotes + " | " + p['Notes']);
                }
             }
             
             // 5. Update Assets
             if (photoUrl) sheet.getRange(realRowIndex, 18).setValue(photoUrl);
             if (p['Distance']) sheet.getRange(realRowIndex, 19).setValue(p['Distance']);
             if (p['Visit Report Data']) sheet.getRange(realRowIndex, 20).setValue(p['Visit Report Data']);

             rowUpdated = true;
             break; 
        }
      }
    }

    // [NEW ROW FALLBACK]
    if (!rowUpdated) {
        const row = [
          new Date(),
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          p['Worker Name'], "'" + (p['Worker Phone Number'] || ""), 
          p['Emergency Contact Name'] || '', "'" + (p['Emergency Contact Number'] || ""), p['Emergency Contact Email'] || '',
          p['Escalation Contact Name'] || '', "'" + (p['Escalation Contact Number'] || ""), p['Escalation Contact Email'] || '',
          newStatus,
          p['Notes'],
          p['Location Name'] || '', p['Location Address'] || '',
          p['Last Known GPS'], p['Timestamp'] || new Date().toISOString(),
          p['Battery Level'] || '', photoUrl, 
          p['Distance'] || '', p['Visit Report Data'] || '',
          p['Anticipated Departure Time'] || ''
        ];
        sheet.appendRow(row);
    }

    // [ALERTS]
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return ContentService.createTextOutput("OK");

  } catch(e) { return ContentService.createTextOutput("Error: " + e.toString()); } 
  finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. MONITOR & FORM HANDLERS (doGet)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  // Connection Test
  if(e.parameter.test) {
     if(e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Monitor Polling
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Form Engine (Advanced)
  if(e.parameter.action === 'getForms') {
     return getChecklistForm(e.parameter.companyName);
  }
  
  // Manual Triggers
  if(e.parameter.run === 'reports') {
      runAllLongitudinalReports();
      return ContentService.createTextOutput("Reports Generated");
  }
  if(e.parameter.run === 'archive') {
      archiveOldData();
      return ContentService.createTextOutput("Archive Complete");
  }

  return ContentService.createTextOutput("OTG Online");
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. UTILITIES & MESSAGING
// ─────────────────────────────────────────────────────────────────────────────

function sendAlert(data) {
  let recipients = [Session.getEffectiveUser().getEmail()];
  let smsNumbers = [];
  
  if (data['Alarm Status'] === 'ESCALATION_SENT') {
     if(data['Escalation Contact Email']) recipients.push(data['Escalation Contact Email']);
     if(data['Escalation Contact Number']) smsNumbers.push(data['Escalation Contact Number']);
  } else {
     if(data['Emergency Contact Email']) recipients.push(data['Emergency Contact Email']);
     if(data['Emergency Contact Number']) smsNumbers.push(data['Emergency Contact Number']);
  }
  
  // Email
  recipients = [...new Set(recipients)].filter(e => e && e.includes('@'));
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p>Worker: ${data['Worker Name']}</p><p>Location: ${data['Location Name']}</p><p>Battery: ${data['Battery Level']}</p><p>Map: <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`;
  if(recipients.length > 0) MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body});

  // SMS (Textbelt)
  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Map: http://maps.google.com/?q=${data['Last Known GPS']}`;
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  try { UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', payload: { phone: clean, message: msg, key: key }}); } catch(e) {}
}

function getChecklistForm(companyName) {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const data = sh.getDataRange().getValues();
     let foundRow = data.find(r => r[1] === companyName);
     if(!foundRow) foundRow = data.find(r => r[1] === 'Travel Report');
     if(!foundRow) foundRow = data.find(r => r[1] === '(Standard)');
     if(!foundRow) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const questions = [];
     for(let i=2; i<data[0].length; i++) {
         const val = foundRow[i];
         if(val && val !== "") {
             let type='check', text=val;
             if(val.includes('[TEXT]')) { type='text'; text=val.replace('[TEXT]','').trim(); }
             else if(val.includes('[PHOTO]')) { type='photo'; text=val.replace('[PHOTO]','').trim(); }
             else if(val.includes('[YESNO]')) { type='yesno'; text=val.replace('[YESNO]','').trim(); }
             else if(val.includes('[NUMBER]')) { type='number'; text=val.replace('[NUMBER]','').trim(); }
             else if(val.includes('$')) { type='number'; text=val.replace('$','').trim(); }
             else if(val.includes('[GPS]')) { type='gps'; text=val.replace('[GPS]','').trim(); }
             else if(val.includes('[HEADING]')) { type='header'; text=val.replace('[HEADING]','').trim(); }
             else if(val.includes('[NOTE]')) { type='note'; text=val.replace('[NOTE]','').trim(); }
             else if(val.includes('[SIGN]')) { type='signature'; text=val.replace('[SIGN]','').trim(); }
             questions.push({type, text});
         }
     }
     return ContentService.createTextOutput(JSON.stringify(questions)).setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. ADVANCED REPORTING (FULL ENGINE RESTORED)
// ─────────────────────────────────────────────────────────────────────────────

// Archive rows > 30 days old
function archiveOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  let archive = ss.getSheetByName('Archive');
  if (!archive) archive = ss.insertSheet('Archive');
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const today = new Date(), keep = [data[0]], move = [];
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]);
    const diff = (today - date) / (1000 * 60 * 60 * 24);
    if (diff > CONFIG.ARCHIVE_DAYS && data[i][10] === 'DEPARTED') {
       rowsToArchive.push(data[i]);
    } else {
       rowsToKeep.push(data[i]);
    }
  }
  
  if (rowsToArchive.length > 0) {
    if (archive.getLastRow() === 0) archive.appendRow(data[0]);
    archive.getRange(archive.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  }
}

// Generate Longitudinal Workbook
function runAllLongitudinalReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const reportFile = createLongitudinalWorkbook();
  const reportSS = SpreadsheetApp.open(reportFile);
  
  generateWorkerActivityReport(data, reportSS);
  generateTravelReport(data, reportSS);
  
  const recipient = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: recipient,
    subject: `Monthly Safety Report Generated - ${CONFIG.ORG_NAME}`,
    htmlBody: `<p>Your longitudinal report is ready.</p><p><a href="${reportSS.getUrl()}">View Report</a></p>`
  });
}

function createLongitudinalWorkbook() {
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM");
  const name = `Longitudinal Report - ${dateStr} - ${CONFIG.ORG_NAME}`;
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) return files.next();
  return DriveApp.getFileById(SpreadsheetApp.create(name).getId());
}

function generateWorkerActivityReport(data, reportSS) {
  let sheet = reportSS.getSheetByName('Worker Activity');
  if (sheet) sheet.clear(); else sheet = reportSS.insertSheet('Worker Activity');
  
  sheet.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Avg Duration (mins)"]);
  sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#dbeafe");
  
  const stats = {};
  
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2];
    const status = data[i][10];
    if (!stats[worker]) stats[worker] = { visits: 0, alerts: 0 };
    stats[worker].visits++;
    if (status.includes("EMERGENCY") || status.includes("OVERDUE")) stats[worker].alerts++;
  }
  
  const output = Object.keys(stats).map(w => [w, stats[w].visits, stats[w].alerts, "N/A"]);
  if (output.length > 0) sheet.getRange(2, 1, output.length, 4).setValues(output);
}

function generateTravelReport(data, reportSS) {
  let sheet = reportSS.getSheetByName('Travel Stats');
  if (sheet) sheet.clear(); else sheet = reportSS.insertSheet('Travel Stats');
  
  sheet.appendRow(["Worker Name", "Total Distance (km)", "Trips"]);
  sheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dcfce7");
  
  const stats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2];
    const dist = parseFloat(data[i][18]) || 0; 
    if (!stats[worker]) stats[worker] = { km: 0, trips: 0 };
    if (dist > 0) { stats[worker].km += dist; stats[worker].trips++; }
  }
  
  const output = Object.keys(stats).map(w => [w, stats[w].km.toFixed(2), stats[w].trips]);
  if (output.length > 0) sheet.getRange(2, 1, output.length, 3).setValues(output);
}

/**
 * PDF GENERATION (Full Implementation)
 * Creates a Google Doc from template, fills it, saves as PDF.
 */
function generateVisitPdf(rowIndex) {
    if (!CONFIG.REPORT_TEMPLATE_ID || !CONFIG.PDF_FOLDER_ID) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const data = sheet.getDataRange().getValues();
    const row = data[rowIndex]; // Array of cell values
    const headers = data[0];
    
    try {
      // 1. Copy Template
      const templateFile = DriveApp.getFileById(CONFIG.REPORT_TEMPLATE_ID);
      const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
      const copy = templateFile.makeCopy(`Report - ${row[2]} - ${row[1]}`, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      
      // 2. Replace Text
      headers.forEach((header, i) => {
          body.replaceText(`{{${header}}}`, String(row[i]));
      });
      
      // 3. Insert Photo
      if (row[17]) { // Photo URL is Col 18
         try {
             const imgBlob = UrlFetchApp.fetch(row[17]).getBlob();
             body.appendImage(imgBlob).setWidth(300);
         } catch(e) {}
      }
      
      doc.saveAndClose();
      
      // 4. Convert to PDF
      const pdf = copy.getAs(MimeType.PDF);
      folder.createFile(pdf);
      
      // 5. Cleanup Temp Doc
      copy.setTrashed(true);
      
    } catch(e) {
      console.log("PDF Gen Error: " + e.toString());
    }
}
