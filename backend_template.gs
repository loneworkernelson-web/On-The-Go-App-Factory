// ==========================================
// OTG APPSUITE - MASTER BACKEND v2.1
// ==========================================
const CONFIG = {
  SECRET_KEY: "%%SECRET_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", // Optional
  ORS_API_KEY: "%%ORS_API_KEY%%", // Optional
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%" // Optional
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    const p = e.parameter;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // 1. SETUP HEADERS (Auto-detect)
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp","Date","Worker Name","Worker Phone Number","Emergency Contact Name","Emergency Contact Number","Emergency Contact Email","Escalation Contact Name","Escalation Contact Number","Escalation Contact Email","Alarm Status","Notes","Location Name","Location Address","Last Known GPS","GPS Timestamp","Battery Level","Photo 1","Distance (km)"];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    }
    
    // 2. PROCESS PHOTO
    let photoUrl = "";
    if(p['Photo 1'] && p['Photo 1'].includes('base64')) {
      try {
        const data = Utilities.base64Decode(p['Photo 1'].split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', p['Worker Name'] + '_' + Date.now() + '.jpg');
        const file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch(err) { photoUrl = "Err: "+err; }
    }

    // 3. APPEND ROW
    const row = [
      new Date(),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      p['Worker Name'],
      p['Worker Phone Number'],
      p['Emergency Contact Name'] || '',
      p['Emergency Contact Number'] || '',
      p['Emergency Contact Email'] || '',
      p['Escalation Contact Name'] || '',
      p['Escalation Contact Number'] || '',
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

    // 4. SEND ALERTS (To contacts + Owner)
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
  // Logic: Send to Emergency Contact + Script Owner
  // If Escalation, send to Escalation Contact instead
  
  let recipients = [Session.getEffectiveUser().getEmail()]; // Always notify admin
  
  if (data['Alarm Status'] === 'ESCALATION_SENT') {
     if(data['Escalation Contact Email']) recipients.push(data['Escalation Contact Email']);
  } else {
     if(data['Emergency Contact Email']) recipients.push(data['Emergency Contact Email']);
  }
  
  // Remove duplicates and empty strings
  recipients = [...new Set(recipients)].filter(e => e && e.includes('@'));
  
  const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `
    <h1 style="color:red;font-weight:bold;">${data['Alarm Status']}</h1>
    <p><strong>Worker:</strong> ${data['Worker Name']}</p>
    <p><strong>Phone:</strong> <a href="tel:${data['Worker Phone Number']}">${data['Worker Phone Number']}</a></p>
    <p><strong>Location:</strong> ${data['Location Name'] || 'Unknown'}</p>
    <p><strong>Notes:</strong> ${data['Notes']}</p>
    <p><strong>GPS Map:</strong> <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>
    <hr>
    <p><em>This is an automated safety alert. Please verify worker safety immediately.</em></p>
  `;
  
  if(recipients.length > 0) {
    MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body});
  }
}

function doGet(e) {
  if(e.parameter.test && e.parameter.key === CONFIG.SECRET_KEY) {
    return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
  }
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify(rows)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput("OTG Online");
}
