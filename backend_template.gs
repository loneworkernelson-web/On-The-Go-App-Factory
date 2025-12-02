// ==========================================
// OTG APPSUITE - MASTER BACKEND v2.0
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
    
    // 1. SETUP HEADERS
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp","Date","Worker Name","Worker Phone Number","Alarm Status","Notes","Location Name","Last Known GPS","Photo 1","Distance (km)"];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    }
    
    // 2. PROCESS PHOTO (If present)
    let photoUrl = "";
    if(p['Photo 1'] && p['Photo 1'].includes('base64')) {
      if(CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
        try {
          const data = Utilities.base64Decode(p['Photo 1'].split(',')[1]);
          const blob = Utilities.newBlob(data, 'image/jpeg', p['Worker Name'] + '_' + Date.now() + '.jpg');
          const file = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID).createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoUrl = file.getUrl();
        } catch(err) {
          photoUrl = "Error saving: " + err.toString();
        }
      } else {
        photoUrl = "Photo skipped (No Folder ID configured)";
      }
    }

    // 3. APPEND ROW
    sheet.appendRow([
      new Date(),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      p['Worker Name'],
      p['Worker Phone Number'],
      p['Alarm Status'],
      p['Notes'],
      p['Location Name'] || '',
      p['Last Known GPS'],
      photoUrl,
      p['Distance'] || ''
    ]);

    // 4. SEND ALERTS (Emails)
    if(p['Alarm Status'].includes('EMERGENCY') || p['Alarm Status'].includes('DURESS') || p['Alarm Status'].includes('MISSED')) {
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
  const email = Session.getEffectiveUser().getEmail(); // Sends to owner for now
  const subject = "🚨 OTG ALERT: " + data['Worker Name'] + " - " + data['Alarm Status'];
  const body = `
    <h1>SAFETY ALERT</h1>
    <p><strong>Worker:</strong> ${data['Worker Name']}</p>
    <p><strong>Status:</strong> <span style="color:red">${data['Alarm Status']}</span></p>
    <p><strong>Phone:</strong> <a href="tel:${data['Worker Phone Number']}">${data['Worker Phone Number']}</a></p>
    <p><strong>GPS:</strong> <a href="https://maps.google.com/?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>
    <hr>
    <p>Time: ${new Date().toLocaleString()}</p>
  `;
  MailApp.sendEmail({to: email, subject: subject, htmlBody: body});
}

function doGet(e) {
  return ContentService.createTextOutput("OTG Backend Online");
}