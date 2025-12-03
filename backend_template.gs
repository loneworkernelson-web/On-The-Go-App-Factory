// ==========================================
// OTG APPSUITE - MASTER BACKEND v5.0 (Smart Update)
// ==========================================
const CONFIG = {
  SECRET_KEY: "%%SECRET_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%",
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%"
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  // Wait up to 30 seconds for other processes to finish
  lock.tryLock(30000); 
  
  try {
    const p = e.parameter;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // 1. SETUP HEADERS
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp","Date","Worker Name","Worker Phone Number","Emergency Contact Name","Emergency Contact Number","Emergency Contact Email","Escalation Contact Name","Escalation Contact Number","Escalation Contact Email","Alarm Status","Notes","Location Name","Location Address","Last Known GPS","GPS Timestamp","Battery Level","Photo 1","Distance (km)"];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
    
    // 2. PROCESS PHOTO
    let photoUrl = "";
    if(p['Photo 1'] && p['Photo 1'].includes('base64')) {
      try {
        const data = Utilities.base64Decode(p['Photo 1'].split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', p['Worker Name'] + '_' + Date.now() + '.jpg');
        
        let file;
        if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
           file = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID).createFile(blob);
        } else {
           file = DriveApp.createFile(blob);
        }
        
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch(err) { photoUrl = "Err: "+err; }
    }

    // 3. SMART ROW UPDATE LOGIC
    const status = p['Alarm Status'];
    const worker = p['Worker Name'];
    let rowUpdated = false;

    // If this is NOT a new start (e.g. it's Depart, Update, or Check-in), try to find the previous row
    if (status !== 'ON SITE') {
      const lastRow = sheet.getLastRow();
      // Only search last 50 rows for performance
      const searchDepth = Math.min(lastRow, 50); 
      
      if (lastRow > 1) {
        // Get column C (Worker Name) and K (Alarm Status)
        // Indices are 0-based in the array, so C=2, K=10
        const data = sheet.getRange(lastRow - searchDepth + 1, 1, searchDepth, 19).getValues();
        
        // Loop backwards from bottom
        for (let i = data.length - 1; i >= 0; i--) {
          const rowWorker = data[i][2]; // Col C
          const rowStatus = data[i][10]; // Col K
          
          // Find latest row for this worker that is NOT "DEPARTED" (i.e. it's active)
          if (rowWorker === worker && rowStatus !== 'DEPARTED' && rowStatus !== 'COMPLETED') {
             const realRowIndex = lastRow - searchDepth + 1 + i;
             
             // UPDATE THE ROW
             // Update Status (Col 11/K)
             sheet.getRange(realRowIndex, 11).setValue(status);
             
             // Append Notes (Col 12/L) - Don't overwrite, append
             if (p['Notes']) {
                const oldNotes = data[i][11];
                const newNotes = oldNotes + " | " + p['Notes'];
                sheet.getRange(realRowIndex, 12).setValue(newNotes);
             }
             
             // Update GPS (Col 15/O)
             if (p['Last Known GPS']) sheet.getRange(realRowIndex, 15).setValue(p['Last Known GPS']);
             
             // Update Battery (Col 17/Q)
             if (p['Battery Level']) sheet.getRange(realRowIndex, 17).setValue(p['Battery Level']);
             
             // Update Photo (Col 18/R) - Only if new photo provided
             if (photoUrl) sheet.getRange(realRowIndex, 18).setValue(photoUrl);
             
             // Update Distance (Col 19/S)
             if (p['Distance']) sheet.getRange(realRowIndex, 19).setValue(p['Distance']);

             rowUpdated = true;
             break; // Stop searching
          }
        }
      }
    }

    // 4. IF NO UPDATE HAPPENED (New Visit, or Fallback), APPEND NEW ROW
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
          p['Alarm Status'],
          p['Notes'],
          p['Location Name'] || '',
          p['Location Address'] || '',
          p['Last Known GPS'],
          p['Timestamp'] || new Date().toISOString(),
          p['Battery Level'] || '',
          photoUrl,
          p['Distance'] || ''
        ];
        sheet.appendRow(row);
    }

    // 5. SEND ALERTS (SMS/Email)
    if(p['Alarm Status'].match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) {
       sendAlert(p);
    }

    return ContentService.createTextOutput("OK");

  } catch(e) {
    return ContentService.createTextOutput("Error: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

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
  
  // EMAIL
  emailRecipients = [...new Set(emailRecipients)].filter(e => e && e.includes('@'));
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p>Worker: ${data['Worker Name']}</p><p>Location: ${data['Location Name']}</p><p>Battery: ${data['Battery Level']}</p><p>Map: <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`;
  if(emailRecipients.length > 0) MailApp.sendEmail({to: emailRecipients.join(','), subject: subject, htmlBody: body});

  // SMS (Textbelt)
  smsNumbers = [...new Set(smsNumbers)].filter(n => n && n.length > 5);
  const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}. Bat: ${data['Battery Level']}. Map: http://maps.google.com/?q=${data['Last Known GPS']}`;
  
  smsNumbers.forEach(phone => sendSms(phone, smsMsg));
}

function sendSms(phone, msg) {
  const cleanPhone = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); 
  const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt';
  
  try {
    UrlFetchApp.fetch('https://textbelt.com/text', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ phone: cleanPhone, message: msg, key: key }),
      muteHttpExceptions: true
    });
  } catch(e) { console.log("SMS Fail", e); }
}

function doGet(e) {
  if(e.parameter.test && e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
  
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if(e.parameter.action === 'getForms') {
     const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklists');
     if(!sh) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const data = sh.getDataRange().getValues();
     let foundRow = data.find(r => r[0] === e.parameter.companyName);
     if(!foundRow) foundRow = data.find(r => r[0] === '(Standard)');
     if(!foundRow) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
     
     const questions = [];
     for(let i=2; i<data[0].length; i++) {
         const val = foundRow[i];
         if(val && val !== "") {
             let type = 'check', text = val;
             if(val.includes('[TEXT]')) { type='text'; text=val.replace('[TEXT]','').trim(); }
             else if(val.includes('[PHOTO]')) { type='photo'; text=val.replace('[PHOTO]','').trim(); }
             else if(val.includes('[YESNO]')) { type='yesno'; text=val.replace('[YESNO]','').trim(); }
             else if(val.includes('[NUMBER]')) { type='number'; text=val.replace('[NUMBER]','').trim(); }
             else if(val.includes('$')) { type='number'; text=val.replace('$','').trim(); } // Mileage handler
             else if(val.includes('[GPS]')) { type='gps'; text=val.replace('[GPS]','').trim(); }
             else if(val.includes('[HEADING]')) { type='header'; text=val.replace('[HEADING]','').trim(); }
             else if(val.includes('[NOTE]')) { type='note'; text=val.replace('[NOTE]','').trim(); }
             else if(val.includes('[SIGN]')) { type='signature'; text=val.replace('[SIGN]','').trim(); }
             questions.push({type, text});
         }
     }
     return ContentService.createTextOutput(JSON.stringify(questions)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("OTG Online");
}
