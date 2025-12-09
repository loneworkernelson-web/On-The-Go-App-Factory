/**
 * ON-THE-GO APPSUITE - MASTER BACKEND v35.0 (Auto-Setup Edition)
 * Status: FULL HEAVY VERSION
 * * CORE MODULES:
 * 1. Security: Secret Key Validation
 * 2. Database: Smart Row De-duplication
 * 3. Smart Scribe: AI Grammar Fixer (NZ English) for Reports ONLY
 * 4. Reporting: PDF & Email Generation
 * 5. Safety: Server-Side Watchdog
 * 6. Setup: One-Click Template Generator
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
  
  // Storage Folders
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  
  // DYNAMIC CONFIG (Managed by Script Properties)
  // These are loaded from storage if they exist
  PDF_FOLDER_ID: "",        
  REPORT_TEMPLATE_ID: "",   
  
  // Settings
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30,
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%
};

// LOAD PROPERTIES (For Template ID)
const scriptProps = PropertiesService.getScriptProperties();
const storedTemplateId = scriptProps.getProperty('REPORT_TEMPLATE_ID');
const storedPdfFolderId = scriptProps.getProperty('PDF_FOLDER_ID');
if(storedTemplateId) CONFIG.REPORT_TEMPLATE_ID = storedTemplateId;
if(storedPdfFolderId) CONFIG.PDF_FOLDER_ID = storedPdfFolderId;


// ─────────────────────────────────────────────────────────────────────────────
// 2. DATA INGESTION (doPost) - SAVES RAW DATA (Source of Truth)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 

  try {
    // 1. Validation
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error", message:"No Data"}));
    const p = e.parameter;
    
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) {
       return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));
    }

    // 2. Database Connection
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // 3. Schema (25 Columns)
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
    
    // 4. Asset Processing
    const savedAssets = {};
    ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4'].forEach(key => {
        if(p[key] && p[key].includes('base64')) {
             savedAssets[key] = saveImageToDrive(p[key], `${p['Worker Name']}_${key}_${Date.now()}.jpg`);
        }
    });
    
    let sigUrl = "";
    if(p['Signature'] && p['Signature'].includes('base64')) {
      sigUrl = saveImageToDrive(p['Signature'], `${p['Worker Name']}_Sig_${Date.now()}.png`);
    }

    // 5. Smart Row Update
    const worker = p['Worker Name'];
    const newStatus = p['Alarm Status'];
    let rowUpdated = false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const searchDepth = Math.min(lastRow - 1, 50);
      const startRow = lastRow - searchDepth + 1;
      const data = sheet.getRange(startRow, 1, searchDepth, 25).getValues(); 
      
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2];
        const rowStatus = data[i][10];
        
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus) || newStatus === 'SAFE - MONITOR CLEARED')) {
             const realRowIndex = startRow + i;
             
             // Update Status & Vitals
             sheet.getRange(realRowIndex, 11).setValue(newStatus);
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']);
             if (p['Anticipated Departure Time']) sheet.getRange(realRowIndex, 21).setValue(p['Anticipated Departure Time']);
             
             // Update Notes (Append Only - PRESERVE RAW)
             if (p['Notes'] && !p['Notes'].includes("Locating")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) {
                    sheet.getRange(realRowIndex, 12).setValue(oldNotes + " | " + p['Notes']);
                }
             }
             
             // Update Data
             if (p['Distance']) sheet.getRange(realRowIndex, 19).setValue(p['Distance']);
             if (p['Visit Report Data']) sheet.getRange(realRowIndex, 20).setValue(p['Visit Report Data']);
             
             // Update Assets
             if (savedAssets['Photo 1']) sheet.getRange(realRowIndex, 18).setValue(savedAssets['Photo 1']);
             if (sigUrl) sheet.getRange(realRowIndex, 22).setValue(sigUrl);
             if (savedAssets['Photo 2']) sheet.getRange(realRowIndex, 23).setValue(savedAssets['Photo 2']);
             if (savedAssets['Photo 3']) sheet.getRange(realRowIndex, 24).setValue(savedAssets['Photo 3']);
             if (savedAssets['Photo 4']) sheet.getRange(realRowIndex, 25).setValue(savedAssets['Photo 4']);

             rowUpdated = true;
             
             // TRIGGER: PDF Generation (On Departure)
             if ((newStatus === 'DEPARTED' || newStatus === 'COMPLETED') && CONFIG.REPORT_TEMPLATE_ID) {
                 generateVisitPdf(realRowIndex);
             }
             break;
        }
      }
    }

    // 6. New Row (Fallback)
    if (!rowUpdated) {
        const row = [
          new Date(),
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          p['Worker Name'], 
          "'" + (p['Worker Phone Number'] || ""), 
          p['Emergency Contact Name'] || '', 
          "'" + (p['Emergency Contact Number'] || ""), 
          p['Emergency Contact Email'] || '',
          p['Escalation Contact Name'] || '', 
          "'" + (p['Escalation Contact Number'] || ""), 
          p['Escalation Contact Email'] || '',
          newStatus,
          p['Notes'], // Raw Notes stored here
          p['Location Name'] || '', 
          p['Location Address'] || '',
          p['Last Known GPS'], 
          p['Timestamp'] || new Date().toISOString(),
          p['Battery Level'] || '', 
          savedAssets['Photo 1'] || '', 
          p['Distance'] || '', 
          p['Visit Report Data'] || '',
          p['Anticipated Departure Time'] || '', 
          sigUrl || '', 
          savedAssets['Photo 2'] || '',
          savedAssets['Photo 3'] || '',
          savedAssets['Photo 4'] || ''
        ];
        sheet.appendRow(row);
    }
    
    // 7. Email Dispatcher (Immediate)
    if (p['Template Name'] && p['Visit Report Data']) {
        processFormEmail(p);
    }

    // 8. Safety Alerts
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) {
        sendAlert(p);
    }

    return ContentService.createTextOutput("OK");

  } catch(e) { 
      return ContentService.createTextOutput("Error: " + e.toString());
  } finally { 
      lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. SMART SCRIBE (AI ENGINE)
// ─────────────────────────────────────────────────────────────────────────────

function smartScribe(text) {
  // If no key or short text, return original
  if (!CONFIG.GEMINI_API_KEY || !text || text.length < 5) return text;
  
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    
    const payload = {
      "contents": [{
        "parts": [{
          "text": "Correct the grammar and spelling of the following text to standard New Zealand English. Do not add any conversational filler or introductions. Only return the corrected text. Text: " + text
        }]
      }]
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.candidates && json.candidates[0].content.parts[0].text) {
       return json.candidates[0].content.parts[0].text.trim();
    }
    return text; // Fallback to raw if AI fails
  } catch (e) {
    console.log("Smart Scribe Error: " + e);
    return text; // Fallback to raw
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. OUTPUT ENGINES (EMAIL & PDF) - USES SMART SCRIBE
// ─────────────────────────────────────────────────────────────────────────────

function processFormEmail(p) {
    try {
        const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
        const data = sh.getDataRange().getValues();
        
        const row = data.find(r => String(r[1]).trim() === String(p['Template Name']).trim());
        if (!row) return;
        
        const recipient = row[2];
        if (!recipient || !String(recipient).includes('@')) return;
        
        const reportData = JSON.parse(p['Visit Report Data']);
        const worker = p['Worker Name'];
        const loc = p['Location Name'] || "Unknown";
        
        // --- SMART SCRIBE PROCESSING ---
        let displayNotes = p['Notes'] || "";
        if(displayNotes) displayNotes = smartScribe(displayNotes);
        
        // Build HTML
        let html = `<div style="font-family: sans-serif; max-width: 600px; border: 1px solid #ddd; padding: 20px;">
            <h2 style="color: #2563eb;">${p['Template Name']}</h2>
            <p><strong>Submitted by:</strong> ${worker}<br><strong>Location:</strong> ${loc}<br><strong>Time:</strong> ${new Date().toLocaleString()}</p>
            <div style="background:#f8fafc; padding:10px; margin-bottom:15px; border-left:4px solid #3b82f6;"><strong>Notes:</strong> ${displayNotes}</div>
            <hr><table style="width:100%; border-collapse: collapse;">`;
            
        for (const [key, val] of Object.entries(reportData)) {
            if (key === 'Signature_Image') continue;
            let displayVal = val;
            if (typeof val === 'string' && val.length > 10 && !val.includes('http')) {
               displayVal = smartScribe(val);
            }
            if (typeof val === 'string' && val.match(/^-?\d+(\.\d+)?,\s*-?\d+(\.\d+)?$/)) {
                displayVal = `<a href="https://www.google.com/maps/search/?api=1&query=${val}">${val}</a>`;
            }
            html += `<tr style="border-bottom: 1px solid #eee;"><td style="padding: 8px; font-weight: bold; color: #555;">${key}</td><td style="padding: 8px;">${displayVal}</td></tr>`;
        }
        html += `</table>`;
        
        // Attach Images
        const inlineImages = {};
        ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4'].forEach((k, i) => {
             if (p[k] && p[k].includes('base64')) {
                 const b = Utilities.newBlob(Utilities.base64Decode(p[k].split(',')[1]), 'image/jpeg', `p${i}.jpg`);
                 inlineImages[`p${i}`] = b;
                 html += `<p><strong>${k}:</strong><br><img src="cid:p${i}" style="max-width:300px; border-radius:8px;"></p>`;
             }
        });
        
        let sigData = p['Signature'] || reportData['Signature_Image'];
        if (sigData && sigData.includes('base64')) {
             const b = Utilities.newBlob(Utilities.base64Decode(sigData.split(',')[1]), 'image/png', 'sig.png');
             inlineImages['sig0'] = b;
             html += `<p><strong>Signature:</strong><br><img src="cid:sig0" style="max-width:200px; border:1px solid #ccc;"></p>`;
        }
        
        html += `</div>`;
        MailApp.sendEmail({ to: recipient, subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${worker}`, htmlBody: html, inlineImages: inlineImages });
    } catch(e) { console.log("Email Error: " + e); }
}

function generateVisitPdf(rowIndex) {
    if (!CONFIG.REPORT_TEMPLATE_ID) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const rowValues = sheet.getRange(rowIndex, 1, 1, 25).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, 25).getValues()[0];

    try {
      const templateFile = DriveApp.getFileById(CONFIG.REPORT_TEMPLATE_ID);
      // Use configured folder or root
      let folder;
      if (CONFIG.PDF_FOLDER_ID) try { folder = DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID); } catch(e){ folder = DriveApp.getRootFolder(); }
      else folder = DriveApp.getRootFolder();

      const copy = templateFile.makeCopy(`Report - ${rowValues[2]} - ${rowValues[1]}`, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      headers.forEach((header, i) => {
          const tag = header.replace(/[^a-zA-Z0-9]/g, ""); 
          let val = String(rowValues[i]);
          // Apply Smart Scribe to Notes
          if (header === 'Notes' || (header === 'Visit Report Data' && val.length > 20)) {
              val = smartScribe(val);
          }
          body.replaceText(`{{${tag}}}`, val);
          body.replaceText(`{{${header}}}`, val);
      });

      // Images
      const imgIndices = [{idx: 17, name: "Photo 1"}, {idx: 22, name: "Photo 2"}, {idx: 23, name: "Photo 3"}, {idx: 24, name: "Photo 4"}, {idx: 21, name: "Signature"}];
      imgIndices.forEach(item => {
          if (rowValues[item.idx]) {
             try {
                 const imgBlob = UrlFetchApp.fetch(rowValues[item.idx]).getBlob();
                 body.appendParagraph(item.name + ":");
                 body.appendImage(imgBlob).setWidth(300);
             } catch(e) {}
          }
      });
      doc.saveAndClose();
      
      const pdf = copy.getAs(MimeType.PDF);
      const recipient = Session.getEffectiveUser().getEmail();
      MailApp.sendEmail({
        to: recipient,
        subject: `Visit Report (PDF): ${rowValues[2]}`,
        body: "Please find the polished visit report attached.",
        attachments: [pdf]
      });
      copy.setTrashed(true); 
    } catch(e) { console.log("PDF Error: " + e.toString()); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

function saveImageToDrive(base64String, filename) {
    try {
        const data = Utilities.base64Decode(base64String.split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', filename); 
        let folder;
        if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
             try { folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); } 
             catch(err) { folder = DriveApp.getRootFolder(); }
        } else { 
             folder = DriveApp.getRootFolder();
        }
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return ""; }
}

function parseQuestions(row) {
     const questions = [];
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
// 6. API HANDLERS (doGet)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  if(e.parameter.test) {
     if(e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

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

  if(e.parameter.action === 'getForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     const data = sh.getDataRange().getValues();
     const param = e.parameter.companyName; 
     let foundRow = data.find(r => r[1] === param) || data.find(r => r[0] === param) || data.find(r => r[1] === 'Travel Report') || data.find(r => r[1] === '(Standard)');
     if(!foundRow) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify(parseQuestions(foundRow))).setMimeType(ContentService.MimeType.JSON);
  }
  
  // TRIGGER: One-Click Setup
  if(e.parameter.run === 'setupTemplate') { 
      const res = setupReportTemplate();
      return ContentService.createTextOutput(res); 
  }
  
  if(e.parameter.run === 'reports') { runAllLongitudinalReports(); return ContentService.createTextOutput("Reports Generated"); }
  if(e.parameter.run === 'archive') { archiveOldData(); return ContentService.createTextOutput("Archive Complete"); }
  if(e.parameter.run === 'watchdog') { checkOverdueVisits(); return ContentService.createTextOutput("Watchdog Run Complete"); }

  return ContentService.createTextOutput("OTG Online");
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. ALERTS & WATCHDOG
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
  
  recipients = [...new Set(recipients)].filter(e => e && e.includes('@'));
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p><strong>Worker:</strong> ${data['Worker Name']}</p><p><strong>Location:</strong> ${data['Location Name'] || 'Unknown'}</p><p><strong>Battery:</strong> ${data['Battery Level'] || 'Unknown'}</p><p><strong>Map:</strong> <a href="https://www.google.com/maps/search/?api=1&query=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`;
  if(recipients.length > 0) MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body});

  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Map: https://maps.google.com/?q=${data['Last Known GPS']}`;
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  try { 
      UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: clean, message: msg, key: key }), muteHttpExceptions: true });
  } catch(e) {}
}

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
    if (['DEPARTED', 'COMPLETED', 'SAFE - MONITOR CLEARED', 'SAFE - MANUALLY CLEARED'].includes(status) || !dueTimeStr) continue;
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
             'Escalation Contact Number': row[8],
             'Battery Level': row[16]
          });
       }
    }
    else if (timeOverdue > 0) {
       if (status === 'ON SITE') sheet.getRange(rowIndex, 11).setValue("OVERDUE");
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. ONE-CLICK SETUP (The Magic Button)
// ─────────────────────────────────────────────────────────────────────────────
function setupReportTemplate() {
    try {
        const doc = DocumentApp.create(`${CONFIG.ORG_NAME} Master Report Template`);
        const body = doc.getBody();
        
        // Professional Header
        body.appendParagraph(CONFIG.ORG_NAME).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph("VISIT REPORT").setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendHorizontalRule();
        
        // Metadata Table
        const cells = [
            ["Worker:", "{{WorkerName}}"],
            ["Location:", "{{LocationName}}"],
            ["Date:", "{{Date}}"],
            ["Status:", "{{AlarmStatus}}"]
        ];
        
        // Simple key-value loop
        cells.forEach(r => {
            const p = body.appendParagraph(`${r[0]} ${r[1]}`);
        });
        
        body.appendHorizontalRule();
        body.appendParagraph("NOTES").setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendParagraph("{{Notes}}");
        
        body.appendHorizontalRule();
        body.appendParagraph("FORM DATA").setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendParagraph("{{VisitReportData}}");
        
        body.appendHorizontalRule();
        body.appendParagraph("Authorized Signature:").setHeading(DocumentApp.ParagraphHeading.HEADING4);
        // Signature placeholder
        body.appendParagraph("{{Signature}}");
        
        doc.saveAndClose();
        
        // Store ID in Script Properties so the code remembers it
        PropertiesService.getScriptProperties().setProperty('REPORT_TEMPLATE_ID', doc.getId());
        
        // Notify User
        const recipient = Session.getEffectiveUser().getEmail();
        MailApp.sendEmail({
            to: recipient,
            subject: "OTG Setup: Template Created",
            body: `Success! A report template has been created in your root folder.\n\nFile Name: ${CONFIG.ORG_NAME} Master Report Template\nID: ${doc.getId()}\n\nThis ID has been automatically saved to your system configuration.`
        });
        
        return "SUCCESS: Template Created & Saved. Check your Email.";
        
    } catch(e) {
        return "ERROR: " + e.toString();
    }
}
