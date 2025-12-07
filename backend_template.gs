/**
 * ON-THE-GO APPSUITE - MASTER BACKEND v30.0 (The Definitive Heavy Version)
 * * MODULES INCLUDED:
 * 1. CORE: Security, Data Ingestion, Smart De-duplication.
 * 2. ASSETS: Image/Signature Decoding & Drive Storage.
 * 3. FORMS: Dynamic Form Serving (Global + Location specific).
 * 4. ALERTS: Watchdog Timer, SMS (Textbelt), & Email Alerts.
 * 5. REPORTS: Instant Form Emails, Monthly Analytics, PDF Generation.
 * 6. MAINTENANCE: Auto-Archiving.
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
  
  // Folders (Required for Photos/PDFs)
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  PDF_FOLDER_ID: "",        // Optional: Folder for PDF Reports
  REPORT_TEMPLATE_ID: "",   // Optional: Google Doc Template ID
  
  // System Settings
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30,
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%
};

// ─────────────────────────────────────────────────────────────────────────────
// 2. DATA INGESTION (doPost)
// Handles incoming data from the Worker App
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Wait up to 30 seconds to prevent data collision
  lock.tryLock(30000); 
  
  try {
    // 1. Basic Validation
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error", message:"No Data"}));
    const p = e.parameter;
    
    // 2. Security Check
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) {
       return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));
    }

    // 3. Database Setup
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // 4. Schema Definition (25 Columns)
    if(sheet.getLastColumn() === 0) {
      const headers = [
        "Timestamp", "Date", "Worker Name", "Worker Phone Number", 
        "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email",
        "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email",
        "Alarm Status", "Notes", "Location Name", "Location Address", 
        "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", 
        "Distance (km)", "Visit Report Data", "Anticipated Departure Time", 
        "Signature", "Photo 2", "Photo 3", "Photo 4"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
    
    // 5. Asset Processing (Photos & Signatures)
    const savedAssets = {};
    
    // Save Photos 1-4
    ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4'].forEach(key => {
        if(p[key] && p[key].includes('base64')) {
             savedAssets[key] = saveImageToDrive(p[key], `${p['Worker Name']}_${key}_${Date.now()}.jpg`);
        }
    });

    // Save Signature
    let sigUrl = "";
    if(p['Signature'] && p['Signature'].includes('base64')) {
      sigUrl = saveImageToDrive(p['Signature'], `${p['Worker Name']}_Sig_${Date.now()}.png`);
    }

    // 6. Smart Row Update (The "De-duplication Engine")
    const worker = p['Worker Name'];
    const newStatus = p['Alarm Status'];
    let rowUpdated = false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const searchDepth = Math.min(lastRow - 1, 50); // Scan last 50 rows
      const startRow = lastRow - searchDepth + 1;
      const data = sheet.getRange(startRow, 1, searchDepth, 25).getValues(); 
      
      // Iterate backwards to find latest visit
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2]; // Col C
        const rowStatus = data[i][10]; // Col K
        
        // Update if active OR if this is a remote resolution
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus) || newStatus === 'SAFE - MONITOR CLEARED')) {
             const realRowIndex = startRow + i;
             
             // Update Status
             sheet.getRange(realRowIndex, 11).setValue(newStatus); 
             
             // Update Vitals
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']);
             if (p['Anticipated Departure Time']) sheet.getRange(realRowIndex, 21).setValue(p['Anticipated Departure Time']);

             // Append Notes (Avoid duplicate "Locating..." logs)
             if (p['Notes'] && !p['Notes'].includes("Locating") && !p['Notes'].includes("GPS Slow")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) { 
                   sheet.getRange(realRowIndex, 12).setValue(oldNotes + " | " + p['Notes']);
                }
             }
             
             // Update Metrics
             if (p['Distance']) sheet.getRange(realRowIndex, 19).setValue(p['Distance']);
             if (p['Visit Report Data']) sheet.getRange(realRowIndex, 20).setValue(p['Visit Report Data']);
             
             // Update Assets (Only if new ones exist)
             if (savedAssets['Photo 1']) sheet.getRange(realRowIndex, 18).setValue(savedAssets['Photo 1']);
             if (sigUrl) sheet.getRange(realRowIndex, 22).setValue(sigUrl);
             if (savedAssets['Photo 2']) sheet.getRange(realRowIndex, 23).setValue(savedAssets['Photo 2']);
             if (savedAssets['Photo 3']) sheet.getRange(realRowIndex, 24).setValue(savedAssets['Photo 3']);
             if (savedAssets['Photo 4']) sheet.getRange(realRowIndex, 25).setValue(savedAssets['Photo 4']);

             rowUpdated = true;
             
             // Trigger PDF if visit is ending
             if ((newStatus === 'DEPARTED' || newStatus === 'COMPLETED') && CONFIG.REPORT_TEMPLATE_ID) {
                 generateVisitPdf(realRowIndex);
             }
             break; 
        }
      }
    }

    // 7. New Row Creation (Fallback)
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
          p['Battery Level'] || '', 
          savedAssets['Photo 1'] || '', // Col 18
          p['Distance'] || '', 
          p['Visit Report Data'] || '',
          p['Anticipated Departure Time'] || '', 
          sigUrl || '', // Col 22
          savedAssets['Photo 2'] || '',
          savedAssets['Photo 3'] || '',
          savedAssets['Photo 4'] || ''
        ];
        sheet.appendRow(row);
    }
    
    // 8. Email Dispatcher (Immediate Form Email)
    if (p['Template Name'] && p['Visit Report Data']) {
        processFormEmail(p);
    }

    // 9. Safety Alerts (SMS/Email)
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) {
        sendAlert(p);
    }

    return ContentService.createTextOutput("OK");

  } catch(e) { return ContentService.createTextOutput("Error: " + e.toString()); } 
  finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

// Save Base64 to Drive
function saveImageToDrive(base64String, filename) {
    try {
        const data = Utilities.base64Decode(base64String.split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', filename); // PNGs save fine as blobs
        
        let folder;
        if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
             try { folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); } 
             catch(err) { folder = DriveApp.getRootFolder(); }
        } else { folder = DriveApp.getRootFolder(); }
        
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return ""; }
}

// Parse Questions from Spreadsheet Row
function parseQuestions(row) {
     const questions = [];
     // Questions start at Column D (Index 3), skipping Company, Template, Email
     for(let i=3; i<row.length; i++) {
         const val = row[i];
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
     return questions;
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. API & DATA RETRIEVAL (doGet)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  // Test Connection
  if(e.parameter.test) {
     return (e.parameter.key === CONFIG.SECRET_KEY) 
        ? ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON) 
        : ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Monitor Polling (JSONP)
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Fetch Global Forms
  if(e.parameter.action === 'getGlobalForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     const data = sh.getDataRange().getValues();
     const globalForms = [];
     for(let r=1; r<data.length; r++) {
         if(String(data[r][0]).toUpperCase().trim() === 'FORMS') {
             globalForms.push({ name: data[r][1], questions: parseQuestions(data[r]) });
         }
     }
     return ContentService.createTextOutput(JSON.stringify(globalForms)).setMimeType(ContentService.MimeType.JSON);
  }

  // Fetch Specific Form (Legacy)
  if(e.parameter.action === 'getForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     const data = sh.getDataRange().getValues();
     const param = e.parameter.companyName; 
     let foundRow = data.find(r => r[1] === param); 
     if(!foundRow) foundRow = data.find(r => r[0] === param);
     if(!foundRow) foundRow = data.find(r => r[1] === 'Travel Report');
     if(!foundRow) foundRow = data.find(r => r[1] === '(Standard)');
     
     if(!foundRow) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify(parseQuestions(foundRow))).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Manual Triggers
  if(e.parameter.run === 'reports') { runAllLongitudinalReports(); return ContentService.createTextOutput("Reports Generated"); }
  if(e.parameter.run === 'archive') { archiveOldData(); return ContentService.createTextOutput("Archive Complete"); }
  if(e.parameter.run === 'watchdog') { checkOverdueVisits(); return ContentService.createTextOutput("Watchdog Run Complete"); }

  return ContentService.createTextOutput("OTG Online");
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. ALERTS & MESSAGING (Email + SMS)
// ─────────────────────────────────────────────────────────────────────────────

function processFormEmail(p) {
    try {
        const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
        const data = sh.getDataRange().getValues();
        const row = data.find(r => r[1] === p['Template Name']); 
        if (!row) return;
        
        const recipient = row[2]; // Col C is Email
        if (!recipient || !recipient.includes('@')) return;
        
        const reportData = JSON.parse(p['Visit Report Data']);
        const worker = p['Worker Name'];
        const loc = p['Location Name'] || p['Location Address'] || "Unknown";
        
        let html = `<div style="font-family: sans-serif; max-width: 600px; border: 1px solid #ddd; padding: 20px;">
            <h2 style="color: #2563eb;">${p['Template Name']}</h2>
            <p><strong>Submitted by:</strong> ${worker}<br><strong>Location:</strong> ${loc}<br><strong>Time:</strong> ${new Date().toLocaleString()}</p><hr><table style="width:100%; border-collapse: collapse;">`;
        
        for (const [key, val] of Object.entries(reportData)) {
            if (key === 'Signature_Image') continue;
            // Smart Link for GPS
            let displayVal = val;
            if (typeof val === 'string' && val.match(/^-?\d+(\.\d+)?,\s*-?\d+(\.\d+)?$/)) {
                displayVal = `<a href="https://www.google.com/maps/search/?api=1&query=${val}">${val}</a>`;
            }
            html += `<tr style="border-bottom: 1px solid #eee;"><td style="padding: 8px; font-weight: bold; color: #555;">${key}</td><td style="padding: 8px;">${displayVal}</td></tr>`;
        }
        html += `</table>`;
        
        const inlineImages = {};
        
        // Attach all photos
        ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4'].forEach((k, i) => {
             if (p[k] && p[k].includes('base64')) {
                 const b = Utilities.newBlob(Utilities.base64Decode(p[k].split(',')[1]), 'image/jpeg', `p${i}.jpg`);
                 inlineImages[`p${i}`] = b;
                 html += `<p><strong>${k}:</strong><br><img src="cid:p${i}" style="max-width:300px; border-radius:8px;"></p>`;
             }
        });
        
        // Attach Signature
        let sigData = p['Signature'] || reportData['Signature_Image'];
        if (sigData && sigData.includes('base64')) {
             const b = Utilities.newBlob(Utilities.base64Decode(sigData.split(',')[1]), 'image/png', 'sig.png');
             inlineImages['sig0'] = b;
             html += `<p><strong>Signature:</strong><br><img src="cid:sig0" style="max-width:200px; border:1px solid #ccc;"></p>`;
        }
        
        html += `</div>`;
        
        MailApp.sendEmail({
            to: recipient,
            subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${worker}`,
            htmlBody: html,
            inlineImages: inlineImages
        });
    } catch(e) { console.log("Form Email Error: " + e); }
}

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

  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Map: http://maps.google.com/?q=${data['Last Known GPS']}`;
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  try { UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: clean, message: msg, key: key }), muteHttpExceptions: true }); } catch(e) {}
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. MAINTENANCE & WATCHDOG
// ─────────────────────────────────────────────────────────────────────────────

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
    if (diff > CONFIG.ARCHIVE_DAYS && (data[i][10] === 'DEPARTED' || data[i][10] === 'COMPLETED')) {
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
    console.log(`Archived ${rowsToArchive.length} rows.`);
  }
}

function runAllLongitudinalReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM");
  const name = `Longitudinal Report - ${dateStr} - ${CONFIG.ORG_NAME}`;
  let reportFile;
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) reportFile = files.next();
  else reportFile = DriveApp.getFileById(SpreadsheetApp.create(name).getId());
  
  const reportSS = SpreadsheetApp.open(reportFile);
  
  // Activity
  let sheetAct = reportSS.getSheetByName('Worker Activity');
  if (sheetAct) sheetAct.clear(); else sheetAct = reportSS.insertSheet('Worker Activity');
  sheetAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Avg Duration (mins)"]);
  sheetAct.getRange(1,1,1,4).setFontWeight("bold").setBackground("#dbeafe");
  
  const stats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2]; const status = data[i][10];
    if (!stats[worker]) stats[worker] = { visits: 0, alerts: 0 };
    stats[worker].visits++;
    if (status.includes("EMERGENCY") || status.includes("OVERDUE")) stats[worker].alerts++;
  }
  const actRows = Object.keys(stats).map(w => [w, stats[w].visits, stats[w].alerts, "N/A"]);
  if (actRows.length > 0) sheetAct.getRange(2, 1, actRows.length, 4).setValues(actRows);

  // Travel
  let sheetTrav = reportSS.getSheetByName('Travel Stats');
  if (sheetTrav) sheetTrav.clear(); else sheetTrav = reportSS.insertSheet('Travel Stats');
  sheetTrav.appendRow(["Worker Name", "Total Distance (km)", "Trips"]);
  sheetTrav.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dcfce7");
  const tStats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2]; const dist = parseFloat(data[i][18]) || 0; 
    if (!tStats[worker]) tStats[worker] = { km: 0, trips: 0 };
    if (dist > 0) { tStats[worker].km += dist; tStats[worker].trips++; }
  }
  const travRows = Object.keys(tStats).map(w => [w, tStats[w].km.toFixed(2), tStats[w].trips]);
  if (travRows.length > 0) sheetTrav.getRange(2, 1, travRows.length, 3).setValues(travRows);
  
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: `Report: ${name}`, htmlBody: `<a href="${reportSS.getUrl()}">View Report</a>` });
}

function generateVisitPdf(rowIndex) {
    if (!CONFIG.REPORT_TEMPLATE_ID || !CONFIG.PDF_FOLDER_ID) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const rowValues = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, 21).getValues()[0];
    
    try {
      const templateFile = DriveApp.getFileById(CONFIG.REPORT_TEMPLATE_ID);
      const folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);
      const copy = templateFile.makeCopy(`Report - ${rowValues[2]} - ${rowValues[1]}`, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      
      headers.forEach((header, i) => {
          const tag = header.replace(/[^a-zA-Z0-9]/g, ""); 
          body.replaceText(`{{${tag}}}`, String(rowValues[i]));
          body.replaceText(`{{${header}}}`, String(rowValues[i]));
      });
      
      if (rowValues[17]) { 
         try { const imgBlob = UrlFetchApp.fetch(rowValues[17]).getBlob(); body.appendImage(imgBlob).setWidth(300); } catch(e) {}
      }
      
      doc.saveAndClose();
      const pdf = copy.getAs(MimeType.PDF);
      folder.createFile(pdf);
      
      const recipient = Session.getEffectiveUser().getEmail();
      MailApp.sendEmail({
        to: recipient,
        subject: `Visit Report: ${rowValues[2]}`,
        body: "Please find the visit report attached.",
        attachments: [pdf]
      });

      copy.setTrashed(true); 
      
    } catch(e) { console.log("PDF Error: " + e.toString()); }
}

// Watchdog
function checkOverdueVisits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if(!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const now = new Date().getTime();
  
  const escalationMs = (CONFIG.ESCALATION_MINUTES || 15) * 60 * 1000;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 2;
    const status = row[10];
    const dueTimeStr = row[20];
    
    if (['DEPARTED', 'COMPLETED', 'SAFE - MONITOR CLEARED', 'SAFE - MANUALLY CLEARED'].includes(status) || !dueTimeStr) {
      continue;
    }

    const dueTime = new Date(dueTimeStr).getTime();
    if (isNaN(dueTime)) continue;

    const timeOverdue = now - dueTime;

    if (timeOverdue > escalationMs) {
       if (!status.includes('EMERGENCY')) {
          const newStatus = "EMERGENCY - OVERDUE (Server Watchdog)";
          sheet.getRange(rowIndex, 11).setValue(newStatus);
          
          sendAlert({
             'Worker Name': row[2], 
             'Worker Phone Number': row[3],
             'Alarm Status': newStatus, 
             'Location Name': row[12], 
             'Last Known GPS': row[14], 
             'Notes': "Worker failed to check in. Phone may be offline.",
             'Emergency Contact Email': row[6], 
             'Emergency Contact Number': row[5],
             'Escalation Contact Email': row[9], 
             'Escalation Contact Number': row[8]
          });
          console.log(`Critical Alert triggered for ${row[2]}`);
       }
    }
    else if (timeOverdue > 0) {
       if (status === 'ON SITE') {
          sheet.getRange(rowIndex, 11).setValue("OVERDUE");
       }
    }
  }
}
