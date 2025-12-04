/**
 * ON-THE-GO APPSUITE - MASTER BACKEND v18.0 (Full Advanced + Secure)
 * * FEATURES:
 * 1. Secure Data Entry (Key Validation)
 * 2. Smart Row Updating (Prevents Duplicate Rows)
 * 3. Textbelt SMS Integration
 * 4. Global Form Serving (Action: getGlobalForms)
 * 5. Full Longitudinal Reporting Engine (Monthly Stats)
 * 6. Automated Archiving Engine
 * 7. PDF Generation Engine
 */

// ─────────────────────────────────────────────────────────────────────────────
// 1. CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {
  // --- SECURITY ---
  SECRET_KEY: "%%SECRET_KEY%%",
  
  // --- API KEYS ---
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%",
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  
  // --- FOLDER IDs (For Advanced Features) ---
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  PDF_FOLDER_ID: "",        // (Optional) ID of Drive Folder to save PDF reports
  REPORT_TEMPLATE_ID: "",   // (Optional) ID of Google Doc Template for PDFs
  
  // --- SETTINGS ---
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30
};

// ─────────────────────────────────────────────────────────────────────────────
// 2. DATA INGESTION (doPost)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // Wait up to 30s to prevent write collisions
  
  try {
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error", message:"No Data"}));
    const p = e.parameter;
    
    // [SECURITY CHECK]
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) {
       return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // [SCHEMA] Ensure 21 Columns (Matches Worker App v50+)
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
        // Use configured folder or root
        let folder = (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) 
          ? DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID)
          : DriveApp.getRootFolder();
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch(err) { photoUrl = "Err: "+err; }
    }

    // [SMART ROW UPDATE]
    // Logic: Find the last active row for this worker and update it instead of creating a new one.
    const worker = p['Worker Name'];
    const newStatus = p['Alarm Status'];
    let rowUpdated = false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const searchDepth = Math.min(lastRow - 1, 50); // Look back 50 rows
      const startRow = lastRow - searchDepth + 1;
      const maxCols = sheet.getLastColumn();
      const data = sheet.getRange(startRow, 1, searchDepth, maxCols).getValues(); 
      
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2]; // Col C (Index 2)
        const rowStatus = data[i][10]; // Col K (Index 10)
        
        // Update if row is NOT closed, OR if we are forcing a "Safe" resolution from Monitor
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus) || newStatus === 'SAFE - MONITOR CLEARED')) {
             const realRowIndex = startRow + i;
             
             // 1. Update Status
             sheet.getRange(realRowIndex, 11).setValue(newStatus); 
             
             // 2. Update Vitals (GPS, Battery)
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']); // Col Q
             
             // 3. Update Due Time (Col 21 / Index 20)
             if (p['Anticipated Departure Time']) sheet.getRange(realRowIndex, 21).setValue(p['Anticipated Departure Time']);

             // 4. Append Notes (Intelligently avoid duplicates for GPS pings)
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
          p['Worker Name'], "'" + (p['Worker Phone Number'] || ""), // Force string
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
// 3. DATA RETRIEVAL (doGet)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  // 1. Connection Test
  if(e.parameter.test) {
     if(e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. Monitor App Polling (JSONP)
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // 3. Form Engine (Standard & Global)
  
  // A. Get GLOBAL FORMS (Row A = "FORMS")
  if(e.parameter.action === 'getGlobalForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const data = sh.getDataRange().getValues();
     const globalForms = [];
     
     // Scan for 'FORMS' in Column A (Index 0)
     for(let r=1; r<data.length; r++) {
         if(String(data[r][0]).toUpperCase().trim() === 'FORMS') {
             const tplName = data[r][1]; // Name is in Col B
             const questions = parseQuestions(data[r]);
             globalForms.push({ name: tplName, questions: questions });
         }
     }
     return ContentService.createTextOutput(JSON.stringify(globalForms)).setMimeType(ContentService.MimeType.JSON);
  }

  // B. Get SPECIFIC FORM (By Company or Template Name)
  if(e.parameter.action === 'getForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const data = sh.getDataRange().getValues();
     const param = e.parameter.companyName; // Logic reuses this param name for template lookup
     
     // Search Logic:
     // 1. Try to match Template Name (Col B)
     // 2. Try to match Company Name (Col A)
     // 3. Fallback to Standard
     let foundRow = data.find(r => r[1] === param); 
     if(!foundRow) foundRow = data.find(r => r[0] === param);
     if(!foundRow) foundRow = data.find(r => r[1] === 'Travel Report');
     if(!foundRow) foundRow = data.find(r => r[1] === '(Standard)');
     
     if(!foundRow) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const questions = parseQuestions(foundRow);
     return ContentService.createTextOutput(JSON.stringify(questions)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 4. Manual Triggers (for testing)
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

// Helper: Parse Row into Question Objects
function parseQuestions(row) {
     const questions = [];
     // Questions start at Column C (Index 2)
     for(let i=2; i<row.length; i++) {
         const val = row[i];
         if(val && val !== "") {
             let type='check', text=val;
             if(val.includes('[TEXT]')) { type='text'; text=val.replace('[TEXT]','').trim(); }
             else if(val.includes('[PHOTO]')) { type='photo'; text=val.replace('[PHOTO]','').trim(); }
             else if(val.includes('[YESNO]')) { type='yesno'; text=val.replace('[YESNO]','').trim(); }
             else if(val.includes('[NUMBER]')) { type='number'; text=val.replace('[NUMBER]','').trim(); }
             else if(val.includes('$')) { type='number'; text=val.replace('$','').trim(); } // Mileage syntax
             else if(val.includes('[GPS]')) { type='gps'; text=val.replace('[GPS]','').trim(); }
             else if(val.includes('[HEADING]')) { type='header'; text=val.replace('[HEADING]','').trim(); }
             else if(val.includes('[NOTE]')) { type='note'; text=val.replace('[NOTE]','').trim(); }
             else if(val.includes('[SIGN]')) { type='signature'; text=val.replace('[SIGN]','').trim(); }
             questions.push({type, text});
         }
     }
     return questions;
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. ALERTS & MESSAGING
// ─────────────────────────────────────────────────────────────────────────────

function sendAlert(data) {
  let recipients = [Session.getEffectiveUser().getEmail()];
  let smsNumbers = [];
  
  // Determine who to notify based on escalation status
  if (data['Alarm Status'] === 'ESCALATION_SENT') {
     if(data['Escalation Contact Email']) recipients.push(data['Escalation Contact Email']);
     if(data['Escalation Contact Number']) smsNumbers.push(data['Escalation Contact Number']);
  } else {
     if(data['Emergency Contact Email']) recipients.push(data['Emergency Contact Email']);
     if(data['Emergency Contact Number']) smsNumbers.push(data['Emergency Contact Number']);
  }
  
  // Send Email
  recipients = [...new Set(recipients)].filter(e => e && e.includes('@'));
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `
    <h1 style="color:red;">${data['Alarm Status']}</h1>
    <p><strong>Worker:</strong> ${data['Worker Name']}</p>
    <p><strong>Location:</strong> ${data['Location Name'] || 'Unknown'}</p>
    <p><strong>Battery:</strong> ${data['Battery Level'] || 'Unknown'}</p>
    <p><strong>Map:</strong> <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>
    <hr><p><i>OTG Safety System</i></p>
  `;
  if(recipients.length > 0) MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body});

  // Send SMS (Textbelt)
  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Map: http://maps.google.com/?q=${data['Last Known GPS']}`;
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  try { 
    UrlFetchApp.fetch('https://textbelt.com/text', { 
      method: 'post', 
      contentType: 'application/json',
      payload: JSON.stringify({ phone: clean, message: msg, key: key }),
      muteHttpExceptions: true
    }); 
  } catch(e) { console.log("SMS Fail", e); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. ADVANCED REPORTING & MAINTENANCE (Full Engine)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * ARCHIVE OLD DATA
 * Moves rows older than CONFIG.ARCHIVE_DAYS to 'Archive' sheet.
 */
function archiveOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  let archive = ss.getSheetByName('Archive');
  if (!archive) archive = ss.insertSheet('Archive');
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; 
  
  const today = new Date();
  const rowsToKeep = [data[0]]; 
  const rowsToArchive = [];
  
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]); 
    const diff = (today - date) / (1000 * 60 * 60 * 24);
    // Archive only if old AND Completed
    if (diff > CONFIG.ARCHIVE_DAYS && (data[i][10] === 'DEPARTED' || data[i][10] === 'COMPLETED')) {
       rowsToArchive.push(data[i]);
    } else {
       rowsToKeep.push(data[i]);
    }
  }
  
  if (rowsToArchive.length > 0) {
    if (archive.getLastRow() === 0) archive.appendRow(data[0]);
    archive.getRange(archive.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    
    // Rewrite main sheet (Clear & Paste)
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
    console.log(`Archived ${rowsToArchive.length} rows.`);
  }
}

/**
 * LONGITUDINAL REPORTS
 * Creates monthly spreadsheets for data analysis.
 */
function runAllLongitudinalReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // Create Report File
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM");
  const name = `Longitudinal Report - ${dateStr} - ${CONFIG.ORG_NAME}`;
  let reportFile;
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) reportFile = files.next();
  else reportFile = DriveApp.getFileById(SpreadsheetApp.create(name).getId());
  
  const reportSS = SpreadsheetApp.open(reportFile);
  
  // 1. Activity Report
  let sheetAct = reportSS.getSheetByName('Worker Activity');
  if (sheetAct) sheetAct.clear(); else sheetAct = reportSS.insertSheet('Worker Activity');
  sheetAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Avg Duration (mins)"]);
  sheetAct.getRange(1,1,1,4).setFontWeight("bold").setBackground("#dbeafe");
  
  const stats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2];
    const status = data[i][10];
    if (!stats[worker]) stats[worker] = { visits: 0, alerts: 0 };
    stats[worker].visits++;
    if (status.includes("EMERGENCY") || status.includes("OVERDUE")) stats[worker].alerts++;
  }
  const actRows = Object.keys(stats).map(w => [w, stats[w].visits, stats[w].alerts, "N/A"]);
  if (actRows.length > 0) sheetAct.getRange(2, 1, actRows.length, 4).setValues(actRows);

  // 2. Travel Report
  let sheetTrav = reportSS.getSheetByName('Travel Stats');
  if (sheetTrav) sheetTrav.clear(); else sheetTrav = reportSS.insertSheet('Travel Stats');
  sheetTrav.appendRow(["Worker Name", "Total Distance (km)", "Trips"]);
  sheetTrav.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dcfce7");
  
  const tStats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2];
    // Distance is Col 19 (Index 18)
    const dist = parseFloat(data[i][18]) || 0; 
    if (!tStats[worker]) tStats[worker] = { km: 0, trips: 0 };
    if (dist > 0) { tStats[worker].km += dist; tStats[worker].trips++; }
  }
  const travRows = Object.keys(tStats).map(w => [w, tStats[w].km.toFixed(2), tStats[w].trips]);
  if (travRows.length > 0) sheetTrav.getRange(2, 1, travRows.length, 3).setValues(travRows);
  
  // Email Admin
  const recipient = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: recipient,
    subject: `Monthly Safety Report Generated`,
    htmlBody: `<p>Your report is ready.</p><p><a href="${reportSS.getUrl()}">View Report</a></p>`
  });
}

/**
 * PDF GENERATION
 * Full logic to create PDFs from Google Docs templates.
 */
function generateVisitPdf(rowIndex) {
    if (!CONFIG.REPORT_TEMPLATE_ID || !CONFIG.PDF_FOLDER_ID) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const data = sheet.getDataRange().getValues();
    const row = data[rowIndex];
    const headers = data[0];
    
    try {
      const templateFile = DriveApp.getFileById(CONFIG.REPORT_TEMPLATE_ID);
      const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
      const copy = templateFile.makeCopy(`Report - ${row[2]} - ${row[1]}`, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      
      // Replace {{Header}} with Value
      headers.forEach((header, i) => {
          body.replaceText(`{{${header}}}`, String(row[i]));
      });
      
      // Insert Photo if available
      if (row[17]) { // Photo URL Col 18
         try {
             const imgBlob = UrlFetchApp.fetch(row[17]).getBlob();
             body.appendImage(imgBlob).setWidth(300);
         } catch(e) {}
      }
      
      doc.saveAndClose();
      
      // PDF Convert
      const pdf = copy.getAs(MimeType.PDF);
      folder.createFile(pdf);
      copy.setTrashed(true); // Delete temp doc
      
    } catch(e) { console.log("PDF Error: " + e.toString()); }
}
