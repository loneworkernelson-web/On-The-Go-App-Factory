/**
 * OTG APPSUITE - MASTER BACKEND v51.0 (Integration Fix)
 */

const CONFIG = {
  SECRET_KEY: "%%SECRET_KEY%%",
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  // DYNAMIC PROPERTIES
  PDF_FOLDER_ID: "",        
  REPORT_TEMPLATE_ID: "",   
  // SETTINGS
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30,
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%
};

// LOAD PROPERTIES
const sp = PropertiesService.getScriptProperties();
const tid = sp.getProperty('REPORT_TEMPLATE_ID');
const pid = sp.getProperty('PDF_FOLDER_ID');
if(tid) CONFIG.REPORT_TEMPLATE_ID = tid;
if(pid) CONFIG.PDF_FOLDER_ID = pid;

// ────────────────────────────────────────────────────────────
// 1. API HANDLERS (doGet)
// ────────────────────────────────────────────────────────────
function doGet(e) {
  if(e.parameter.test) {
     if(e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if(e.parameter.action === 'sync') {
      const worker = e.parameter.worker;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Get Forms
      const tSheet = ss.getSheetByName('Templates');
      const tData = tSheet ? tSheet.getDataRange().getValues() : [];
      const forms = [];
      const cachedTemplates = {};
      
      for(let i=1; i<tData.length; i++) {
          const row = tData[i];
          if(row.length < 3) continue;
          const type = row[0]; 
          const name = row[1];
          const assign = row[2]; 
          
          if(type === 'FORM' && (assign.includes(worker) || assign === 'ALL')) {
              forms.push({ name: name, questions: parseQuestions(row) });
          }
          cachedTemplates[name] = parseQuestions(row);
      }
      
      // Get Sites
      const sSheet = ss.getSheetByName('Sites');
      const sData = sSheet ? sSheet.getDataRange().getValues() : [];
      const sites = [];
      
      for(let i=1; i<sData.length; i++) {
          const row = sData[i];
          if(row.length < 1) continue;
          const assign = row[0]; 
          if(assign.includes(worker) || assign === 'ALL') {
              sites.push({
                  template: row[1], company: row[2], siteName: row[3], address: row[4],
                  contactName: row[5], contactPhone: row[6], contactEmail: row[7], notes: row[8]
              });
          }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
          sites: sites, forms: forms, cachedTemplates: cachedTemplates
      })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    if(!sh) return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify({status:"error"})+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    const lastRun = PropertiesService.getScriptProperties().getProperty('LAST_WATCHDOG_RUN') || "Never";
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify({
        workers: rows, server_time: new Date().toISOString(), watchdog_last_run: lastRun
    })+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if(e.parameter.run === 'setupTemplate') return ContentService.createTextOutput(setupReportTemplate()); 
  if(e.parameter.run === 'reports') { runAllLongitudinalReports(); return ContentService.createTextOutput("Reports Generated"); }
  if(e.parameter.run === 'archive') { archiveOldData(); return ContentService.createTextOutput("Archive Complete"); }
  if(e.parameter.run === 'watchdog') { checkOverdueVisits(); return ContentService.createTextOutput("Watchdog Run Complete"); }

  return ContentService.createTextOutput(JSON.stringify({status: "online", message: "OTG Server Active"})).setMimeType(ContentService.MimeType.JSON);
}

// ────────────────────────────────────────────────────────────
// 2. DATA INGESTION (doPost)
// ────────────────────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 

  try {
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error", message:"No Data"}));
    const p = e.parameter;
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
      sheet.appendRow(headers); sheet.setFrozenRows(1);
    }
    
    // ASSET SAVING (v51.0 FIX)
    const assets = {};
    const assetKeys = ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4', 'Signature'];
    
    assetKeys.forEach(key => {
        if(p[key] && p[key].length > 100) { // Simple check for base64 content
             const ext = key === 'Signature' ? 'png' : 'jpg';
             // Clean filename
             const safeWorker = (p['Worker Name'] || 'Worker').replace(/[^a-z0-9]/gi, '_');
             const filename = `${safeWorker}_${key.replace(' ', '')}_${Date.now()}.${ext}`;
             assets[key] = saveImageToDrive(p[key], filename);
        } else {
             assets[key] = "";
        }
    });

    const worker = p['Worker Name'];
    const newStatus = p['Alarm Status'];
    let rowUpdated = false;
    const lastRow = sheet.getLastRow();
    
    // UPDATE EXISTING ROW
    if (lastRow > 1) {
      const searchDepth = Math.min(lastRow - 1, 50);
      const startRow = lastRow - searchDepth + 1;
      const data = sheet.getRange(startRow, 1, searchDepth, 25).getValues(); 
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2];
        const rowStatus = data[i][10];
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus) || newStatus === 'SAFE - MONITOR CLEARED' || (rowStatus === 'DATA_ENTRY_ONLY' && newStatus === 'DATA_ENTRY_ONLY'))) {
             const rIdx = startRow + i;
             sheet.getRange(rIdx, 11).setValue(newStatus); 
             if (p['Last Known GPS']) sheet.getRange(rIdx, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(rIdx, 17).setValue(p['Battery Level']);
             if (p['Anticipated Departure Time']) sheet.getRange(rIdx, 21).setValue(p['Anticipated Departure Time']);
             if (p['Notes'] && !p['Notes'].includes("Locating")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) sheet.getRange(rIdx, 12).setValue(oldNotes ? oldNotes + " | " + p['Notes'] : p['Notes']);
             }
             if (p['Distance']) sheet.getRange(rIdx, 19).setValue(p['Distance']);
             if (p['Visit Report Data'] && p['Visit Report Data'].length > 5 && p['Visit Report Data'] !== '{}') sheet.getRange(rIdx, 20).setValue(p['Visit Report Data']);
             
             // Update Assets in specific columns
             if(assets['Photo 1']) sheet.getRange(rIdx, 18).setValue(assets['Photo 1']);
             if(assets['Signature']) sheet.getRange(rIdx, 22).setValue(assets['Signature']);
             if(assets['Photo 2']) sheet.getRange(rIdx, 23).setValue(assets['Photo 2']);
             if(assets['Photo 3']) sheet.getRange(rIdx, 24).setValue(assets['Photo 3']);
             if(assets['Photo 4']) sheet.getRange(rIdx, 25).setValue(assets['Photo 4']);

             rowUpdated = true;
             break;
        }
      }
    }

    // NEW ROW
    if (!rowUpdated) {
        const row = [
            new Date(), 
            Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), 
            p['Worker Name'], "'" + (p['Worker Phone Number'] || ""), 
            p['Emergency Contact Name'], "'" + (p['Emergency Contact Number'] || ""), p['Emergency Contact Email'], 
            p['Escalation Contact Name'], "'" + (p['Escalation Contact Number'] || ""), p['Escalation Contact Email'], 
            newStatus, p['Notes'], p['Location Name'], p['Location Address'], p['Last Known GPS'], p['Timestamp'] || new Date().toISOString(), 
            p['Battery Level'], 
            assets['Photo 1'], // Col R
            p['Distance'], p['Visit Report Data'], p['Anticipated Departure Time'], 
            assets['Signature'], // Col V
            assets['Photo 2'], assets['Photo 3'], assets['Photo 4']
        ];
        sheet.appendRow(row);
    }
    
    // EMAIL TRIGGER (v51.0) - Checks for Template Name
    if (p['Template Name']) {
        processFormEmail(p, assets);
    }
    
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return ContentService.createTextOutput("OK");
  } catch(e) { return ContentService.createTextOutput("Error: " + e.toString()); } 
  finally { lock.releaseLock(); }
}

// ────────────────────────────────────────────────────────────
// 3. HELPERS (Updated for Email/Photos)
// ────────────────────────────────────────────────────────────
function processFormEmail(p, assets) {
    try {
        const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates'); 
        const data = sh.getDataRange().getValues();
        // Look for the template name in Column B (Index 1)
        const row = data.find(r => String(r[1]).trim() === String(p['Template Name']).trim());
        if (!row) return; // Template not found
        
        const recipient = row[3]; 
        if (!recipient || !String(recipient).includes('@')) return; // No email configured
        
        let reportData = {};
        try { reportData = JSON.parse(p['Visit Report Data']); } catch(e) {}
        
        const worker = p['Worker Name'];
        const loc = p['Location Name'] || "Unknown";
        let displayNotes = p['Notes'] || "";
        
        let html = `<div style="font-family: sans-serif; max-width: 600px; border: 1px solid #ddd; padding: 20px;">
            <h2 style="color: #2563eb;">${p['Template Name']}</h2>
            <p><strong>Submitted by:</strong> ${worker}<br><strong>Location:</strong> ${loc}<br><strong>Time:</strong> ${new Date().toLocaleString()}</p>
            <div style="background:#f8fafc; padding:10px; margin-bottom:15px; border-left:4px solid #3b82f6;"><strong>Notes:</strong> ${displayNotes}</div>
            <hr><table style="width:100%; border-collapse: collapse;">`;
            
        for (const [key, val] of Object.entries(reportData)) {
            if (key === 'Signature_Image') continue;
            let displayVal = val;
            if (typeof val === 'string' && val.length > 20 && !val.includes('http')) displayVal = smartScribe(val);
            html += `<tr style="border-bottom: 1px solid #eee;"><td style="padding: 8px; font-weight: bold; color: #555;">${key}</td><td style="padding: 8px;">${displayVal}</td></tr>`;
        }
        html += `</table>`;
        
        // Embed Photos
        if (assets['Photo 1']) html += `<p><strong>Photo 1:</strong> <a href="${assets['Photo 1']}">View</a></p>`;
        if (assets['Signature']) html += `<p><strong>Signature:</strong> <a href="${assets['Signature']}">View</a></p>`;
        
        html += `</div>`;
        MailApp.sendEmail({ to: recipient, subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${worker}`, htmlBody: html });
    } catch(e) { console.log("Email Error: " + e); }
}

function saveImageToDrive(base64String, filename) {
    try {
        // Strip header if present
        const base64Clean = base64String.split(',').pop();
        const data = Utilities.base64Decode(base64Clean);
        const blob = Utilities.newBlob(data, 'image/jpeg', filename); 
        let folder;
        if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
             try { folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); } catch(e){ folder = DriveApp.getRootFolder(); }
        } else {
             folder = DriveApp.getRootFolder();
        }
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return "Error Saving: " + e.message; }
}

// ... (Other helpers like parseQuestions, smartScribe, sendAlert remain the same) ...
function checkOverdueVisits() {
  PropertiesService.getScriptProperties().setProperty('LAST_WATCHDOG_RUN', new Date().toISOString());
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
    const status = row[10];
    const dueTimeStr = row[20];
    if (['DEPARTED', 'COMPLETED', 'SAFE - MONITOR CLEARED', 'SAFE - MANUALLY CLEARED'].includes(status) || !dueTimeStr) continue;
    const dueTime = new Date(dueTimeStr).getTime();
    if (isNaN(dueTime)) continue;
    const timeOverdue = now - dueTime;
    if (timeOverdue > escalationMs && !status.includes('EMERGENCY')) {
          const newStatus = "EMERGENCY - OVERDUE (Server Watchdog)";
          sheet.getRange(i + 2, 11).setValue(newStatus);
          sendAlert({ 'Worker Name': row[2], 'Worker Phone Number': row[3], 'Alarm Status': newStatus, 'Location Name': row[12], 'Last Known GPS': row[14], 'Notes': "Worker failed to check in.", 'Emergency Contact Email': row[6], 'Emergency Contact Number': row[5], 'Escalation Contact Email': row[9], 'Escalation Contact Number': row[8], 'Battery Level': row[16] });
    } else if (timeOverdue > 0 && status === 'ON SITE') {
       sheet.getRange(i + 2, 11).setValue("OVERDUE");
    }
  }
}
function parseQuestions(row) {
     const questions = [];
     for(let i=4; i<row.length; i++) {
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
function setupReportTemplate() {
    try {
        const doc = DocumentApp.create(`${CONFIG.ORG_NAME} Master Report Template`);
        const body = doc.getBody();
        body.appendParagraph(CONFIG.ORG_NAME).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.appendParagraph("VISIT REPORT").setHeading(DocumentApp.ParagraphHeading.HEADING2);
        const cells = [["Worker:", "{{WorkerName}}"], ["Location:", "{{LocationName}}"], ["Date:", "{{Date}}"], ["Status:", "{{AlarmStatus}}"]];
        cells.forEach(r => body.appendParagraph(`${r[0]} ${r[1]}`));
        body.appendHorizontalRule(); body.appendParagraph("NOTES").setHeading(DocumentApp.ParagraphHeading.HEADING3); body.appendParagraph("{{Notes}}");
        body.appendHorizontalRule(); body.appendParagraph("FORM DATA").setHeading(DocumentApp.ParagraphHeading.HEADING3); body.appendParagraph("{{VisitReportData}}");
        body.appendHorizontalRule(); body.appendParagraph("Authorized Signature:").setHeading(DocumentApp.ParagraphHeading.HEADING4); body.appendParagraph("{{Signature}}");
        doc.saveAndClose();
        PropertiesService.getScriptProperties().setProperty('REPORT_TEMPLATE_ID', doc.getId());
        return "SUCCESS: Template Created. ID: " + doc.getId();
    } catch(e) { return "ERROR: " + e.toString(); }
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
  const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p><strong>Worker:</strong> ${data['Worker Name']}</p><p><strong>Location:</strong> ${data['Location Name'] || 'Unknown'}</p><p><strong>Battery:</strong> ${data['Battery Level'] || 'Unknown'}</p><p><strong>Map:</strong> <a href="https://www.google.com/maps/search/?api=1&query=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`;
  if(recipients.length > 0) MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body});
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}`;
  smsNumbers.forEach(phone => {
       const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
       try { UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: clean, message: smsMsg, key: key }), muteHttpExceptions: true }); } catch(e) {}
  });
}
function smartScribe(text) {
  if (!CONFIG.GEMINI_API_KEY || !text || text.length < 5) return text;
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const payload = { "contents": [{ "parts": [{ "text": "Correct grammar to NZ English: " + text }] }] };
    const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim();
  } catch (e) { return text; }
}
function generateVisitPdf(rowIndex) {
    if (!CONFIG.REPORT_TEMPLATE_ID) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const rowValues = sheet.getRange(rowIndex, 1, 1, 25).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, 25).getValues()[0];
    try {
      const templateFile = DriveApp.getFileById(CONFIG.REPORT_TEMPLATE_ID);
      const copy = templateFile.makeCopy(`Report - ${rowValues[2]}`, DriveApp.getRootFolder());
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      headers.forEach((header, i) => {
          let val = String(rowValues[i]);
          if (header === 'Notes' || (header === 'Visit Report Data' && val.length > 20)) val = smartScribe(val);
          body.replaceText(`{{${header.replace(/[^a-zA-Z0-9]/g, "")}}}`, val);
      });
      doc.saveAndClose();
      MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: "Report", attachments: [copy.getAs(MimeType.PDF)] });
      copy.setTrashed(true); 
    } catch(e) {}
}
function archiveOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  let archive = ss.getSheetByName('Archive');
  if (!archive) archive = ss.insertSheet('Archive');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const today = new Date(); const rowsToKeep = [data[0]]; const rowsToArchive = [];
  for (let i = 1; i < data.length; i++) {
    const date = new Date(data[i][0]);
    const diff = (today - date) / (1000 * 60 * 60 * 24);
    if (diff > CONFIG.ARCHIVE_DAYS && (data[i][10] === 'DEPARTED' || data[i][10] === 'COMPLETED')) { rowsToArchive.push(data[i]); } else { rowsToKeep.push(data[i]); }
  }
  if (rowsToArchive.length > 0) {
    if (archive.getLastRow() === 0) archive.appendRow(data[0]);
    archive.getRange(archive.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    sheet.clearContents(); sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
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
  let sheetAct = reportSS.getSheetByName('Worker Activity');
  if (sheetAct) sheetAct.clear(); else sheetAct = reportSS.insertSheet('Worker Activity');
  sheetAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Avg Duration"]);
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
  let sheetTrav = reportSS.getSheetByName('Travel Stats');
  if (sheetTrav) sheetTrav.clear(); else sheetTrav = reportSS.insertSheet('Travel Stats');
  sheetTrav.appendRow(["Worker Name", "Total Distance (km)", "Trips"]);
  sheetTrav.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dcfce7");
  const tStats = {};
  for (let i = 1; i < data.length; i++) {
    const worker = data[i][2];
    const dist = parseFloat(data[i][18]) || 0; 
    if (!tStats[worker]) tStats[worker] = { km: 0, trips: 0 };
    if (dist > 0) { tStats[worker].km += dist; tStats[worker].trips++; }
  }
  const travRows = Object.keys(tStats).map(w => [w, tStats[w].km.toFixed(2), tStats[w].trips]);
  if (travRows.length > 0) sheetTrav.getRange(2, 1, travRows.length, 3).setValues(travRows);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: `Report: ${name}`, htmlBody: `<a href="${reportSS.getUrl()}">View Report</a>` });
}
