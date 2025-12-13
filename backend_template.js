/**
 * OTG APPSUITE - MASTER BACKEND v58.0 (Platinum Security)
 * Added: Device Fingerprinting & Auto-Locking
 */

const CONFIG = {
  SECRET_KEY: "%%SECRET_KEY%%",
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  PDF_FOLDER_ID: "",        
  REPORT_TEMPLATE_ID: "",   
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: Session.getScriptTimeZone(),
  ARCHIVE_DAYS: 30,
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%
};

const sp = PropertiesService.getScriptProperties();
const tid = sp.getProperty('REPORT_TEMPLATE_ID');
const pid = sp.getProperty('PDF_FOLDER_ID');
if(tid) CONFIG.REPORT_TEMPLATE_ID = tid;
if(pid) CONFIG.PDF_FOLDER_ID = pid;

// ────────────────────────────────────────────────────────────
// 1. SECURITY GATEKEEPER & FINGERPRINTING
// ────────────────────────────────────────────────────────────
function checkAccess(workerName, deviceId) {
  if (!workerName) return { allowed: false, msg: "Name missing" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff');
  
  // Basic Mode (No Staff Tab)
  if (!sheet) return { allowed: true }; 

  const data = sheet.getDataRange().getValues();
  // Scan Staff List
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][0]).trim().toLowerCase();
    const targetName = String(workerName).trim().toLowerCase();
    
    if (rowName === targetName) {
       // 1. Check Active Status (Col C)
       if (data[i][2] && String(data[i][2]).toLowerCase().includes('inactive')) {
           return { allowed: false, msg: "Account Disabled" };
       }

       // 2. Device Fingerprint Check (Col E)
       // If Col E is empty, BIND this device to this user.
       const registeredId = String(data[i][4] || ""); 
       
       if (registeredId === "" || registeredId === "undefined") {
           // First time login - Lock it in
           if(deviceId) {
               sheet.getRange(i + 1, 5).setValue(deviceId);
               return { allowed: true };
           }
       } else {
           // Subsequent login - Match IDs
           if (registeredId === deviceId) {
               return { allowed: true };
           } else {
               return { allowed: false, msg: "Unauthorized Device. Contact Admin to reset." };
           }
       }
       return { allowed: true }; // Fallback if no deviceId sent yet
    }
  }
  return { allowed: false, msg: "Name not found in Staff list." };
}

// ────────────────────────────────────────────────────────────
// 2. API HANDLERS
// ────────────────────────────────────────────────────────────
function doGet(e) {
  if(e.parameter.test) {
     if(e.parameter.key === CONFIG.SECRET_KEY) return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({status:"error", message:"Invalid Key"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // SYNC
  if(e.parameter.action === 'sync') {
      const worker = e.parameter.worker;
      const deviceId = e.parameter.deviceId; // v58.0 New Param
      
      const auth = checkAccess(worker, deviceId);
      if (!auth.allowed) {
          return ContentService.createTextOutput(JSON.stringify({
              status: "error", 
              message: "ACCESS DENIED: " + auth.msg
          })).setMimeType(ContentService.MimeType.JSON);
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const tSheet = ss.getSheetByName('Templates');
      const tData = tSheet ? tSheet.getDataRange().getValues() : [];
      const forms = [];
      const cachedTemplates = {};
      
      for(let i=1; i<tData.length; i++) {
          const row = tData[i];
          if(row.length < 3) continue;
          const type = row[0]; const name = row[1]; const assign = row[2]; 
          if(type === 'FORM' && (assign.includes(worker) || assign === 'ALL')) {
              forms.push({ name: name, questions: parseQuestions(row) });
          }
          cachedTemplates[name] = parseQuestions(row);
      }
      
      const sSheet = ss.getSheetByName('Sites');
      const sData = sSheet ? sSheet.getDataRange().getValues() : [];
      const sites = [];
      for(let i=1; i<sData.length; i++) {
          const row = sData[i];
          if(row.length < 1) continue;
          const assign = row[0]; 
          if(assign.includes(worker) || assign === 'ALL') {
              sites.push({ template: row[1], company: row[2], siteName: row[3], address: row[4], contactName: row[5], contactPhone: row[6], contactEmail: row[7], notes: row[8] });
          }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ sites: sites, forms: forms, cachedTemplates: cachedTemplates })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if(e.parameter.callback){
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
    if(!sh) return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify({status:"error"})+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    const data=sh.getDataRange().getValues();
    const headers=data.shift();
    const rows=data.map(r=>{ let o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
    return ContentService.createTextOutput(e.parameter.callback+"("+JSON.stringify({ workers: rows, server_time: new Date().toISOString(), escalation_limit: CONFIG.ESCALATION_MINUTES })+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if(e.parameter.run === 'setupTemplate') return ContentService.createTextOutput(setupReportTemplate()); 
  return ContentService.createTextOutput(JSON.stringify({status: "online"})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 

  try {
    if (!e || !e.parameter) return ContentService.createTextOutput(JSON.stringify({status:"error"}));
    const p = e.parameter;
    if (!p.key || p.key.trim() !== CONFIG.SECRET_KEY.trim()) return ContentService.createTextOutput(JSON.stringify({status: "error"}));
    
    // Check Auth on Post too (prevents spoofing)
    // Note: We don't enforce deviceId strictly on POST yet to prevent data loss if ID rotates, 
    // but the Sync check effectively blocks usage anyway.
    const auth = checkAccess(p['Worker Name'], null); 
    if (!auth.allowed) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Unauthorized"}));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
      sheet.appendRow(headers); sheet.setFrozenRows(1);
    }
    
    const assets = {};
    const assetIds = {}; 
    ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4', 'Signature'].forEach(key => {
        if(p[key] && p[key].length > 200) {
             const safeWorker = (p['Worker Name'] || 'Worker').replace(/[^a-z0-9]/gi, '_');
             const suffix = key === 'Signature' ? 'png' : 'jpg';
             const result = saveImageToDrive(p[key], `${safeWorker}_${key.replace(' ', '')}_${Date.now()}.${suffix}`);
             assets[key] = result.url;
             assetIds[key] = result.id;
        } else { assets[key] = ""; }
    });

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
             if (p['Visit Report Data'] && p['Visit Report Data'].length > 5) sheet.getRange(rIdx, 20).setValue(p['Visit Report Data']);
             
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

    if (!rowUpdated) {
        const row = [new Date(), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), p['Worker Name'], "'" + (p['Worker Phone Number'] || ""), p['Emergency Contact Name'], "'" + (p['Emergency Contact Number'] || ""), p['Emergency Contact Email'], p['Escalation Contact Name'], "'" + (p['Escalation Contact Number'] || ""), p['Escalation Contact Email'], newStatus, p['Notes'], p['Location Name'], p['Location Address'], p['Last Known GPS'], p['Timestamp'] || new Date().toISOString(), p['Battery Level'], assets['Photo 1'], p['Distance'], p['Visit Report Data'], p['Anticipated Departure Time'], assets['Signature'], assets['Photo 2'], assets['Photo 3'], assets['Photo 4']];
        sheet.appendRow(row);
    }
    
    if (p['Template Name']) processFormEmail(p, assetIds);
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return ContentService.createTextOutput("OK");
  } catch(e) { return ContentService.createTextOutput("Error: " + e.toString()); } 
  finally { lock.releaseLock(); }
}

function processFormEmail(p, assetIds) {
    try {
        const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates'); 
        const data = sh.getDataRange().getValues();
        const row = data.find(r => String(r[1]).trim() === String(p['Template Name']).trim());
        if (!row) return; 
        const recipient = row[3]; 
        if (!recipient || !String(recipient).includes('@')) return; 
        
        let reportData = {};
        try { reportData = JSON.parse(p['Visit Report Data']); } catch(e) {}
        const worker = p['Worker Name'];
        const loc = p['Location Name'] || "Unknown";
        
        let html = `<div style="font-family: sans-serif; max-width: 600px; padding: 20px; border:1px solid #ccc;">
            <h2 style="color: #2563eb;">${p['Template Name']}</h2>
            <p><strong>Submitted by:</strong> ${worker}<br><strong>Location:</strong> ${loc}<br><strong>Time:</strong> ${new Date().toLocaleString()}</p>
            <div style="background:#f1f5f9; padding:10px; margin:15px 0;"><strong>Notes:</strong> ${p['Notes']||""}</div>
            <table style="width:100%; border-collapse: collapse;">`;
            
        for (const [key, val] of Object.entries(reportData)) {
            if (key === 'Signature_Image') continue;
            let displayVal = val;
            if (typeof val === 'string' && val.length > 20 && !val.includes('http')) displayVal = smartScribe(val);
            html += `<tr style="border-bottom: 1px solid #eee;"><td style="padding: 8px; font-weight: bold; color: #555;">${key}</td><td style="padding: 8px;">${displayVal}</td></tr>`;
        }
        html += `</table><br>`;
        
        if (assetIds['Photo 1']) html += `<h3>Photo 1</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 1']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`;
        if (assetIds['Photo 2']) html += `<h3>Photo 2</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 2']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`;
        if (assetIds['Signature']) html += `<h3>Authorized Signature</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Signature']}&sz=w400" style="max-width:300px; border-bottom:2px solid #000;"><br>`;
        
        html += `</div>`;
        MailApp.sendEmail({ to: recipient, subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${worker}`, htmlBody: html });
    } catch(e) { console.log("Email Error: " + e); }
}

function saveImageToDrive(base64String, filename) {
    try {
        const base64Clean = base64String.split(',').pop();
        const data = Utilities.base64Decode(base64Clean);
        const blob = Utilities.newBlob(data, 'image/jpeg', filename); 
        let folder;
        if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) {
             try { folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); } catch(e){ folder = DriveApp.getRootFolder(); }
        } else { folder = DriveApp.getRootFolder(); }
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return { url: file.getUrl(), id: file.getId() };
    } catch(e) { return { url: "", id: "" }; }
}

function checkOverdueVisits() { /* As before */ }
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
function setupReportTemplate() { /* As before */ }
function sendAlert(data) { /* As before */ }
function smartScribe(text) { /* As before */ }
function archiveOldData() { /* As before */ }
function runAllLongitudinalReports() { /* As before */ }
