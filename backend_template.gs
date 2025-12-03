// ==========================================
// OTG APPSUITE - MASTER BACKEND v9.0 (Advanced + Secure)
// ==========================================

// --- CONFIGURATION ---
const CONFIG = {
  SECRET_KEY: "%%SECRET_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", // Drive Folder ID for photos
  PDF_FOLDER_ID: "", // (Optional) Drive Folder ID for PDF reports
  DOC_TEMPLATE_ID: "", // (Optional) Google Doc Template ID for reports
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%",
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  ORG_NAME: "My Organisation",
  TIMEZONE: Session.getScriptTimeZone()
};

// --- CORE HANDLERS ---
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 
  
  try {
    const p = e.parameter;
    
    // 1. SECURITY CHECK
    if (p.key !== CONFIG.SECRET_KEY) {
       return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid Key"}));
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // 2. SETUP HEADERS
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp","Date","Worker Name","Worker Phone Number","Emergency Contact Name","Emergency Contact Number","Emergency Contact Email","Escalation Contact Name","Escalation Contact Number","Escalation Contact Email","Alarm Status","Notes","Location Name","Location Address","Last Known GPS","GPS Timestamp","Battery Level","Photo 1","Distance (km)"];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
    
    // 3. PHOTO HANDLING
    let photoUrl = "";
    if(p['Photo 1'] && p['Photo 1'].includes('base64')) {
      try {
        const data = Utilities.base64Decode(p['Photo 1'].split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', p['Worker Name'] + '_' + Date.now() + '.jpg');
        let file = (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) 
          ? DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID).createFile(blob)
          : DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch(err) { photoUrl = "Err: "+err; }
    }

    // 4. SMART ROW UPDATE (De-Duplication)
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
        
        // Update if open visit found
        if (rowWorker === worker && !['DEPARTED', 'COMPLETED', 'SAFE - MANUALLY CLEARED'].includes(rowStatus)) {
             const realRowIndex = startRow + i;
             
             sheet.getRange(realRowIndex, 11).setValue(newStatus); 
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']);
             
             // Append notes
             if (p['Notes'] && !p['Notes'].includes("Locating") && !p['Notes'].includes("GPS Slow")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) sheet.getRange(realRowIndex, 12).setValue(oldNotes + " | " + p['Notes']);
             }
             
             if (photoUrl) sheet.getRange(realRowIndex, 18).setValue(photoUrl);
             if (p['Distance']) sheet.getRange(realRowIndex, 19).setValue(p['Distance']);

             rowUpdated = true;
             break; 
        }
      }
    }

    // 5. NEW ROW
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
          p['Battery Level'] || '', photoUrl, p['Distance'] || ''
        ];
        sheet.appendRow(row);
    }

    // 6. ALERTS
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return ContentService.createTextOutput("OK");

  } catch(e) {
    return ContentService.createTextOutput("Error: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  // Test
  if(e.parameter.test && e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
  
  // Monitor Polling
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Form Engine
  if(e.parameter.action === 'getForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     const data = sh.getDataRange().getValues();
     let foundRow = data.find(r => r[1] === e.parameter.companyName); // Col B is Template Name
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
  
  // Trigger Reports Manually
  if(e.parameter.run === 'reports') {
      runDailyReport();
      return ContentService.createTextOutput("Reports Triggered");
  }

  return ContentService.createTextOutput("OTG Online");
}

// --- MESSAGING ---
function sendAlert(data) {
  let emailRecipients = [Session.getEffectiveUser().getEmail()];
  let smsNumbers = [];
  
  if (data['Alarm Status'] === 'ESCALATION_SENT') {
     if(data['Escalation Contact Email']) emailRecipients.push(data['Escalation Contact Email']);
     if(data['Escalation Contact Number']) smsNumbers.push(data['Escalation Contact Number']);
  } else {
     if(data['Emergency Contact Email']) emailRecipients.push(data['Emergency Contact Email']);
     if(data['Emergency Contact Number']) smsNumbers.push(data['Emergency Contact Number']);
  }
  
  emailRecipients = [...new Set(emailRecipients)].filter(e => e && e.includes('@'));
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p>Worker: ${data['Worker Name']}</p><p>Location: ${data['Location Name']}</p><p>Battery: ${data['Battery Level']}</p><p>Map: <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`;
  if(emailRecipients.length > 0) MailApp.sendEmail({to: emailRecipients.join(','), subject: subject, htmlBody: body});

  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Map: http://maps.google.com/?q=${data['Last Known GPS']}`;
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const cleanPhone = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  try {
    UrlFetchApp.fetch('https://textbelt.com/text', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ phone: cleanPhone, message: msg, key: key }),
      muteHttpExceptions: true
    });
  } catch(e) { console.log("SMS Fail", e); }
}

// --- ADVANCED REPORTING & MAINTENANCE ---

// 1. Archive Old Data (Run this via Time-Driven Trigger -> Weekly)
function archiveOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  const archiveSheet = ss.getSheetByName('Archive') || ss.insertSheet('Archive');
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // Empty
  
  const today = new Date();
  const rowsToKeep = [data[0]]; // Keep headers
  const rowsToArchive = [];
  
  // 30 Day cutoff
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]); // Timestamp col
    const diffDays = (today - rowDate) / (1000 * 60 * 60 * 24);
    
    if (diffDays > 30 && data[i][10] === 'DEPARTED') { // Only archive completed visits
       rowsToArchive.push(data[i]);
    } else {
       rowsToKeep.push(data[i]);
    }
  }
  
  if (rowsToArchive.length > 0) {
    // Setup Archive Headers if new
    if (archiveSheet.getLastRow() === 0) archiveSheet.appendRow(data[0]);
    
    // Bulk append to Archive
    const lastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(lastRow + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    
    // Clear and rewrite Visits (safest way to delete non-contiguous rows)
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  }
}

// 2. Daily Summary Report (Run via Time-Driven Trigger -> Daily)
function runDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  const data = sheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
  
  let visitCount = 0;
  let activeAlerts = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === todayStr) { // Date col
      visitCount++;
      if (data[i][10].includes('EMERGENCY') || data[i][10].includes('OVERDUE')) activeAlerts++;
    }
  }
  
  if (visitCount > 0) {
     const recipient = Session.getEffectiveUser().getEmail();
     MailApp.sendEmail({
       to: recipient,
       subject: `Daily Safety Summary - ${CONFIG.ORG_NAME}`,
       htmlBody: `<h2>Daily Report: ${todayStr}</h2><ul><li>Total Visits: ${visitCount}</li><li>Alerts Triggered: ${activeAlerts}</li></ul><p>Check spreadsheet for details.</p>`
     });
  }
}

// 3. PDF Generation (Advanced Stub)
function generateVisitPdf(rowIndex) {
    // This function would interact with a Google Doc template.
    // Requires CONFIG.DOC_TEMPLATE_ID and CONFIG.PDF_FOLDER_ID to be set.
    if (!CONFIG.DOC_TEMPLATE_ID || !CONFIG.PDF_FOLDER_ID) return;
    
    // Logic to copy template, replace placeholders ({Worker Name}, etc), 
    // save as PDF, and email link would go here.
    // Kept lightweight for this version to prevent timeout errors during setup.
}
