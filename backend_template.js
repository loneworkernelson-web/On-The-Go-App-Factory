/**
 * OTG APPSUITE - MASTER BACKEND v68.12
 * - Features: Dual-Key Security, Watchdog Safety Clamp, Auto-Archiving.
 */

const CONFIG = {
  // FACTORY INJECTED KEYS
  MASTER_KEY: "%%SECRET_KEY%%", // High Privilege (Monitor)
  WORKER_KEY: "%%WORKER_KEY%%", // Low Privilege (Worker App)
  
  // API KEYS
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  
  // SYSTEM SETTINGS
  ORG_NAME: "%%ORGANISATION_NAME%%",
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%,
  ARCHIVE_DAYS: 30,
  TIMEZONE: Session.getScriptTimeZone()
};

// --- CORE API ENDPOINTS ---

function doGet(e) {
  try {
      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Params"});

      // 1. Connection Test (Accepts either key)
      if(e.parameter.test) {
         const k = e.parameter.key;
         if(k === CONFIG.MASTER_KEY || k === CONFIG.WORKER_KEY) return sendJSON({status:"success", version: "v68.12"});
         return sendJSON({status:"error", message:"Invalid Key"});
      }

      // 2. Worker App Sync (Accepts either key)
      if(e.parameter.action === 'sync') {
          const k = e.parameter.key;
          if(k !== CONFIG.MASTER_KEY && k !== CONFIG.WORKER_KEY) return sendJSON({status:"error", message:"Auth Failed"});
          
          const worker = e.parameter.worker;
          const deviceId = e.parameter.deviceId; 
          const auth = checkAccess(worker, deviceId, true); // True = Read Only (Sync)
          
          if (!auth.allowed) return sendJSON({ status: "error", message: "ACCESS DENIED: " + auth.msg });

          // Fetch Config Data
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
          return sendJSON({ sites: sites, forms: forms, cachedTemplates: cachedTemplates, meta: auth.meta });
      }

      // 3. Monitor App Poll (MASTER KEY ONLY)
      if(e.parameter.callback){
        if (e.parameter.key !== CONFIG.MASTER_KEY) {
             return ContentService.createTextOutput(e.parameter.callback + "(" + JSON.stringify({error: "Auth Required"}) + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return handleMonitorPoll(e.parameter.callback);
      }

      return sendJSON({status: "online"});

  } catch(err) { return sendJSON({status: "error", message: "SERVER ERROR: " + err.toString()}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 
  try {
    if (!e || !e.parameter) return sendJSON({status:"error"});
    const p = e.parameter;
    
    // AUTH CHECK: Allows Worker Key OR Master Key
    if (p.key !== CONFIG.MASTER_KEY && p.key !== CONFIG.WORKER_KEY) return sendJSON({status: "error", message: "Invalid Key"});
    
    if (p.action !== 'resolve') {
        const auth = checkAccess(p['Worker Name'], p.deviceId, false); // False = Write Access
        if (!auth.allowed) return sendJSON({status: "error", message: "Unauthorized"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    // Ensure Headers
    if(sheet.getLastColumn() === 0) {
        sheet.appendRow(["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"]);
    }

    // Handle Image Uploads (Base64 to Drive)
    const assets = {};
    ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4', 'Signature'].forEach(k => {
        if(p[k] && p[k].length > 100) {
            try {
                const blob = Utilities.newBlob(Utilities.base64Decode(p[k].split(',').pop()), 'image/jpeg', p['Worker Name'] + '_' + k + '_' + Date.now());
                const file = DriveApp.createFile(blob);
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                assets[k] = file.getUrl();
            } catch(err) { assets[k] = "Error saving image"; }
        } else assets[k] = "";
    });

    // Update Existing Row (if user is still checked in)
    let rowUpdated = false;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        // Optimization: Only check last 200 rows for active visits
        const startRow = Math.max(2, lastRow - 200);
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 25).getValues();
        
        for (let i = data.length - 1; i >= 0; i--) {
            if (data[i][2] === p['Worker Name'] && !['DEPARTED', 'COMPLETED', 'SAFE'].some(s => String(data[i][10]).includes(s))) {
                const rIdx = startRow + i;
                // Update Logic
                if (p['Alarm Status'] !== 'DATA_ENTRY_ONLY' || data[i][10] === 'DATA_ENTRY_ONLY') sheet.getRange(rIdx, 11).setValue(p['Alarm Status']);
                if (p['Last Known GPS']) sheet.getRange(rIdx, 15).setValue(p['Last Known GPS']);
                if (p['Battery Level']) sheet.getRange(rIdx, 17).setValue(p['Battery Level']);
                if (p['Anticipated Departure Time']) sheet.getRange(rIdx, 21).setValue(p['Anticipated Departure Time']);
                if (p['Notes']) sheet.getRange(rIdx, 12).setValue(data[i][11] ? data[i][11] + " | " + p['Notes'] : p['Notes']);
                if (assets['Photo 1']) sheet.getRange(rIdx, 18).setValue(assets['Photo 1']);
                rowUpdated = true; 
                break;
            }
        }
    }

    // Append New Row if not updated
    if (!rowUpdated) {
        sheet.appendRow([
            new Date(), 
            Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd"), 
            p['Worker Name'], "'" + p['Worker Phone Number'], 
            p['Emergency Contact Name'], "'" + p['Emergency Contact Number'], p['Emergency Contact Email'], 
            p['Escalation Contact Name'], "'" + p['Escalation Contact Number'], p['Escalation Contact Email'], 
            p['Alarm Status'], p['Notes'], p['Location Name'], p['Location Address'], 
            p['Last Known GPS'], new Date().toISOString(), p['Battery Level'], 
            assets['Photo 1'], p['Distance'], p['Visit Report Data'], 
            p['Anticipated Departure Time'], assets['Signature'], assets['Photo 2'], assets['Photo 3'], assets['Photo 4']
        ]);
    }
    
    // Email Notifications (Template Reports)
    if (p['Template Name']) {
        let recipient = "";
        if (String(p['Template Name']).trim() === "Note to Self") recipient = p['Worker Email'];
        else {
            const tSheet = ss.getSheetByName('Templates');
            if(tSheet) {
                const tData = tSheet.getDataRange().getValues();
                const row = tData.find(r => String(r[1]).trim() === String(p['Template Name']).trim());
                if (row) recipient = row[3];
            }
        }
        if (recipient && recipient.includes('@')) {
             MailApp.sendEmail({ 
                 to: recipient, 
                 subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${p['Worker Name']}`, 
                 htmlBody: `<h2>${p['Template Name']}</h2><p>Worker: ${p['Worker Name']}<br>Location: ${p['Location Name']}</p><p>Notes: ${p['Notes']}</p><hr><p>Data: ${p['Visit Report Data']}</p>` 
             });
        }
    }

    // Alarm Escalation
    if(p['Alarm Status'].match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return sendJSON({status:"ok"});
  } catch(e) { return sendJSON({status:"error", message: e.toString()}); } 
  finally { lock.releaseLock(); }
}

// --- SYSTEM FUNCTIONS ---

function checkOverdueVisits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName('Visits'); 
  if(!sheet) return;
  
  const lastRow = sheet.getLastRow(); 
  if (lastRow <= 1) return;
  
  // WATCHDOG FIX: Only check last 500 rows to prevent timeout
  const startRow = Math.max(2, lastRow - 500);
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 21).getValues();
  const now = new Date().getTime(); 
  const escalationMs = (CONFIG.ESCALATION_MINUTES || 15) * 60 * 1000;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i]; 
    const status = row[10]; 
    const dueTimeStr = row[20];
    
    if (['DEPARTED', 'COMPLETED', 'SAFE - MONITOR CLEARED', 'SAFE - MANUALLY CLEARED'].some(s => String(status).includes(s)) || !dueTimeStr) continue;
    
    const dueTime = new Date(dueTimeStr).getTime(); 
    if (isNaN(dueTime)) continue;
    
    const timeOverdue = now - dueTime;
    
    // 1. Escalation (Emergency)
    if (timeOverdue > escalationMs && !String(status).includes('EMERGENCY')) {
          const newStatus = "EMERGENCY - OVERDUE (Server Watchdog)"; 
          sheet.getRange(startRow + i, 11).setValue(newStatus);
          sendAlert({ 
              'Worker Name': row[2], 
              'Worker Phone Number': row[3], 
              'Alarm Status': newStatus, 
              'Location Name': row[12], 
              'Last Known GPS': row[14], 
              'Notes': "Worker failed to check in (Watchdog Trigger).", 
              'Emergency Contact Email': row[6] 
          });
    } 
    // 2. Warning (Overdue)
    else if (timeOverdue > 0 && status === 'ON SITE') { 
        sheet.getRange(startRow + i, 11).setValue("OVERDUE"); 
    }
  }
}

function archiveOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if (!sheet) return;
  const archiveSheet = ss.getSheetByName('Archive') || ss.insertSheet('Archive');
  if (archiveSheet.getLastColumn() === 0) { // Sync headers
      archiveSheet.appendRow(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  }

  const data = sheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - CONFIG.ARCHIVE_DAYS);
  
  const rowsToKeep = [data[0]]; // Keep headers
  const rowsToArchive = [];

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (rowDate < cutoff) rowsToArchive.push(data[i]);
    else rowsToKeep.push(data[i]);
  }

  if (rowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    sheet.clearContents();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  }
}

// --- HELPER FUNCTIONS ---

function checkAccess(workerName, deviceId, isReadOnly) {
  if (!workerName) return { allowed: false, msg: "Name missing" };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff');
  if (!sheet) return { allowed: true, meta: {} }; // If no Staff sheet, open access
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if(!data[i] || !data[i][0]) continue;
    const rowName = String(data[i][0]).trim().toLowerCase();
    const targetName = String(workerName).trim().toLowerCase();
    
    if (rowName === targetName) {
       // 1. Check Status
       if (data[i][2] && String(data[i][2]).toLowerCase().includes('inactive')) return { allowed: false, msg: "Account Disabled" };
       
       // 2. Check Device Lock
       const registeredId = String(data[i][4] || ""); 
       if (registeredId === "" || registeredId === "undefined") {
           // New Device: Lock it now (unless ReadOnly sync)
           if(deviceId && !isReadOnly) { try { sheet.getRange(i + 1, 5).setValue(deviceId); } catch(e) {} }
           return { allowed: true, meta: { lastVehCheck: data[i][5] || "", wofExpiry: data[i][6] || "" } };
       } else {
           // Existing Device: Match it
           if (registeredId === deviceId) return { allowed: true, meta: { lastVehCheck: data[i][5] || "", wofExpiry: data[i][6] || "" } };
           else return { allowed: false, msg: "Unauthorized Device" };
       }
    }
  }
  return { allowed: false, msg: "Name not found in Staff list" };
}

function handleMonitorPoll(callback) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const r = ss.getSheetByName('Visits').getDataRange().getValues();
    const headers = r.shift();
    
    // Add WOF Status
    const st = ss.getSheetByName('Staff');
    const wofMap = {};
    if(st) {
        const sData = st.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) if(sData[i][0]) wofMap[String(sData[i][0]).toLowerCase()] = sData[i][6] || "";
    }

    const rows = r.map(row => { 
        let obj = {}; 
        headers.forEach((h, i) => obj[h] = row[i]); 
        obj.WOFExpiry = wofMap[String(obj['Worker Name']).toLowerCase()] || ""; 
        return obj; 
    });
    
    return ContentService.createTextOutput(callback+"("+JSON.stringify({ workers: rows, escalation_limit: CONFIG.ESCALATION_MINUTES })+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
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
             else if(val.includes('[GPS]')) { type='gps'; text=val.replace('[GPS]','').trim(); }
             else if(val.includes('[HEADING]')) { type='header'; text=val.replace('[HEADING]','').trim(); }
             else if(val.includes('[NOTE]')) { type='note'; text=val.replace('[NOTE]','').trim(); }
             else if(val.includes('[SIGN]')) { type='signature'; text=val.replace('[SIGN]','').trim(); }
             else if(val.includes('[DATE]')) { type='date'; text=val.replace('[DATE]','').trim(); }
             questions.push({type, text});
         }
     }
     return questions;
}

function sendJSON(data) {
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function sendAlert(data) {
    const email = Session.getEffectiveUser().getEmail(); // Default to Admin
    const subject = `🚨 ALERT: ${data['Worker Name']} - ${data['Alarm Status']}`;
    const body = `
      <h1>${data['Alarm Status']}</h1>
      <p><strong>Worker:</strong> ${data['Worker Name']}</p>
      <p><strong>Phone:</strong> <a href="tel:${data['Worker Phone Number']}">${data['Worker Phone Number']}</a></p>
      <p><strong>Location:</strong> ${data['Location Name']} (${data['Location Address']})</p>
      <p><strong>GPS:</strong> <a href="https://www.google.com/maps?q=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>
      <p><strong>Notes:</strong> ${data['Notes']}</p>
      <hr>
      <p><em>OTG Safety System</em></p>
    `;
    
    // Send to Admin
    MailApp.sendEmail({to: email, subject: subject, htmlBody: body});
    
    // Send to Escalation Contact if exists
    if(data['Escalation Contact Email'] && data['Escalation Contact Email'].includes('@')) {
        MailApp.sendEmail({to: data['Escalation Contact Email'], subject: subject, htmlBody: body});
    }
}
