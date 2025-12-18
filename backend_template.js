/**
 * OTG APPSUITE - MASTER BACKEND v68.19
 * - Feature: Auto-Magic GPS Linking in Email Reports.
 * - Includes: Two-Key Security, Photo Fixes, Watchdog.
 */

const CONFIG = {
  // SECURITY KEYS (Replace if pasting manually)
  MASTER_KEY: "%%SECRET_KEY%%", 
  WORKER_KEY: "%%WORKER_KEY%%", 

  // API KEYS
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  
  // CONFIG
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

function doGet(e) {
  try {
      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Params"});

      // 1. Connection Test (Accepts EITHER Key)
      if(e.parameter.test) {
         const k = e.parameter.key;
         if(k === CONFIG.MASTER_KEY || k === CONFIG.WORKER_KEY) return sendJSON({status:"success", version: "v68.19"});
         return sendJSON({status:"error", message:"Invalid Key"});
      }

      // 2. Geocode Proxy
      if(e.parameter.action === 'geocode') {
          if(!e.parameter.lat || !e.parameter.lon) return sendJSON({address: "Invalid Coords"});
          try {
              const geo = Maps.newGeocoder().reverseGeocode(e.parameter.lat, e.parameter.lon);
              if (geo.status === 'OK' && geo.results.length > 0) {
                  return sendJSON({address: geo.results[0].formatted_address});
              }
              return sendJSON({address: "Unknown Location"});
          } catch(err) { return sendJSON({address: "Geocode Failed"}); }
      }

      // 3. Worker App Sync (Accepts EITHER Key)
      if(e.parameter.action === 'sync') {
          const k = e.parameter.key;
          if(k !== CONFIG.MASTER_KEY && k !== CONFIG.WORKER_KEY) return sendJSON({status:"error", message:"Auth Failed"});
          
          const worker = e.parameter.worker;
          const deviceId = e.parameter.deviceId; 
          const auth = checkAccess(worker, deviceId, true);
          if (!auth.allowed) return sendJSON({ status: "error", message: "ACCESS DENIED: " + auth.msg });

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

      // 4. Monitor App Poll (MASTER KEY ONLY)
      if(e.parameter.callback){
        if (e.parameter.key !== CONFIG.MASTER_KEY) {
             return ContentService.createTextOutput(e.parameter.callback + "(" + JSON.stringify({error: "Auth Required"}) + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return handleMonitorPoll(e.parameter.callback);
      }

      if(e.parameter.run === 'setupTemplate') return ContentService.createTextOutput(setupReportTemplate()); 
      return sendJSON({status: "online"});

  } catch(err) { return sendJSON({status: "error", message: "SERVER ERROR: " + err.toString()}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); 
  try {
    if (!e || !e.parameter) return sendJSON({status:"error"});
    const p = e.parameter;
    
    // Check keys
    if (p.key !== CONFIG.MASTER_KEY && p.key !== CONFIG.WORKER_KEY) return sendJSON({status: "error", message: "Invalid Key"});
    
    if (p.action !== 'resolve') {
        const auth = checkAccess(p['Worker Name'], p.deviceId, false); 
        if (!auth.allowed) return sendJSON({status: "error", message: "Unauthorized"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    
    if(sheet.getLastColumn() === 0) {
      const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
      sheet.appendRow(headers); sheet.setFrozenRows(1);
    }
    
    const assets = {};
    const assetIds = {}; 
    ['Photo 1', 'Photo 2', 'Photo 3', 'Photo 4', 'Signature'].forEach(key => {
        if(p[key] && p[key].length > 100) { 
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
      const searchDepth = Math.min(lastRow - 1, 200);
      const startRow = lastRow - searchDepth + 1;
      const data = sheet.getRange(startRow, 1, searchDepth, 25).getValues(); 
      for (let i = data.length - 1; i >= 0; i--) {
        const rowWorker = data[i][2];
        const rowStatus = data[i][10];
        
        if (rowWorker === worker && (!['DEPARTED', 'COMPLETED'].includes(rowStatus))) {
             const rIdx = startRow + i;
             
             if (newStatus !== 'DATA_ENTRY_ONLY' || rowStatus === 'DATA_ENTRY_ONLY') sheet.getRange(rIdx, 11).setValue(newStatus);
             if (p['Last Known GPS']) sheet.getRange(rIdx, 15).setValue(p['Last Known GPS']);
             if (p['Battery Level']) sheet.getRange(rIdx, 17).setValue(p['Battery Level']);
             if (p['Anticipated Departure Time']) sheet.getRange(rIdx, 21).setValue(p['Anticipated Departure Time']);
             if (p['Notes'] && !p['Notes'].includes("Locating")) {
                const oldNotes = data[i][11];
                if (!oldNotes.includes(p['Notes'])) sheet.getRange(rIdx, 12).setValue(oldNotes ? oldNotes + " | " + p['Notes'] : p['Notes']);
             }
             if (p['Distance']) sheet.getRange(rIdx, 19).setValue(p['Distance']);
             if (p['Visit Report Data'] && p['Visit Report Data'].length > 5) {
                 const oldData = data[i][19];
                 sheet.getRange(rIdx, 20).setValue(oldData ? oldData + " | " + p['Visit Report Data'] : p['Visit Report Data']);
             }
             
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
    
    if (p['Template Name'] === 'Vehicle Safety Check') { updateStaffVehCheck(worker, p['Visit Report Data']); }
    if (p['Template Name']) processFormEmail(p, assetIds);
    if(newStatus.match(/EMERGENCY|DURESS|MISSED|ESCALATION/)) sendAlert(p);

    return sendJSON({status:"ok"});
  } catch(e) { return sendJSON({status:"error", message: e.toString()}); } 
  finally { lock.releaseLock(); }
}

// --- HELPER FUNCTIONS ---
function checkAccess(workerName, deviceId, isReadOnly) {
  if (!workerName) return { allowed: false, msg: "Name missing" };
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Staff');
  if (!sheet) return { allowed: true, meta: {} }; 
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if(!data[i] || !data[i][0]) continue;
    const rowName = String(data[i][0]).trim().toLowerCase();
    if (rowName === String(workerName).trim().toLowerCase()) {
       if (data[i][2] && String(data[i][2]).toLowerCase().includes('inactive')) return { allowed: false, msg: "Account Disabled" };
       const registeredId = String(data[i][4] || ""); 
       if (registeredId === "" || registeredId === "undefined") {
           if(deviceId && !isReadOnly) { try { sheet.getRange(i + 1, 5).setValue(deviceId); } catch(e) {} }
           return { allowed: true, meta: getRowMeta(data[i]) };
       } else {
           if (registeredId === deviceId) return { allowed: true, meta: getRowMeta(data[i]) };
           else return { allowed: false, msg: "Unauthorized Device. Contact Admin to reset." };
       }
    }
  }
  return { allowed: false, msg: "Name not found in Staff list." };
}
function sendJSON(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function handleMonitorPoll(callback) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const t = ss.getSheetByName('Visits');
    if(!t) return ContentService.createTextOutput(callback+"("+JSON.stringify({status:"error"})+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    const r = t.getDataRange().getValues(); const headers = r.shift();
    const st = ss.getSheetByName('Staff'); const stD = st ? st.getDataRange().getValues() : [];
    const wofMap = {}; if(stD.length > 1) { for(let i=1; i<stD.length; i++) { if(stD[i] && stD[i][0]) wofMap[String(stD[i][0]).toLowerCase()] = stD[i][6] || ""; } }
    const rows = r.map(e => { let obj = {}; headers.forEach((h, idx) => obj[h] = e[idx]); const wName = obj['Worker Name'] ? String(obj['Worker Name']).toLowerCase() : ""; obj.WOFExpiry = wofMap[wName] || ""; return obj; });
    return ContentService.createTextOutput(callback+"("+JSON.stringify({ workers: rows, server_time: new Date().toISOString(), escalation_limit: CONFIG.ESCALATION_MINUTES })+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function getRowMeta(row) { return { lastVehCheck: row[5] || "", wofExpiry: row[6] || "" }; }
function updateStaffVehCheck(worker, jsonString) { try { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Staff'); if(!sheet) return; const data = sheet.getDataRange().getValues(); let wofDate = ""; const now = new Date().toISOString(); try { const j = JSON.parse(jsonString); for (const key in j) { if (key.includes("Expiry") || key.includes("Due")) { wofDate = j[key]; break; } } } catch(e) {} for (let i = 1; i < data.length; i++) { if (String(data[i][0]).toLowerCase() === String(worker).toLowerCase()) { sheet.getRange(i + 1, 6).setValue(now); if (wofDate) sheet.getRange(i + 1, 7).setValue(wofDate); break; } } } catch(e) {} }

// UPGRADED EMAIL PROCESSOR
function processFormEmail(p, assetIds) { 
    try { 
        let recipient = ""; 
        if (String(p['Template Name']).trim() === "Note to Self") { 
            recipient = p['Worker Email']; 
            if (!recipient || !recipient.includes('@')) return; 
        } else { 
            const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Templates'); 
            const data = sh.getDataRange().getValues(); 
            const row = data.find(r => String(r[1]).trim() === String(p['Template Name']).trim()); 
            if (!row) return; 
            recipient = row[3]; 
        } 
        if (!recipient || !String(recipient).includes('@')) return; 
        
        let reportData = {}; 
        try { reportData = JSON.parse(p['Visit Report Data']); } catch(e) {} 
        
        const worker = p['Worker Name']; 
        const loc = p['Location Name'] || "Unknown"; 
        
        let html = `<div style="font-family: sans-serif; max-width: 600px; padding: 20px; border:1px solid #ccc;"><h2 style="color: #2563eb;">${p['Template Name']}</h2><p><strong>Submitted by:</strong> ${worker}<br><strong>Location:</strong> ${loc}<br><strong>Time:</strong> ${new Date().toLocaleString()}</p><div style="background:#f1f5f9; padding:10px; margin:15px 0;"><strong>Notes:</strong> ${p['Notes']||""}</div><table style="width:100%; border-collapse: collapse;">`; 
        
        for (const [key, val] of Object.entries(reportData)) { 
            if (key === 'Signature_Image') continue; 
            let displayVal = val; 
            
            // --- SMART LINKING START ---
            if (typeof val === 'string') {
                // 1. Detect GPS (e.g. -41.2, 174.7)
                if (val.match(/^-?\d+(\.\d+)?,\s*-?\d+(\.\d+)?$/)) {
                    displayVal = `<a href="https://www.google.com/maps/search/?api=1&query=${val}" target="_blank" style="color:#2563eb; text-decoration:underline;">📍 ${val}</a>`;
                }
                // 2. Smart Scribe (if Long Text)
                else if (val.length > 20 && !val.includes('http') && !val.includes('data:image')) {
                    displayVal = smartScribe(val);
                }
            }
            // --- SMART LINKING END ---

            html += `<tr style="border-bottom: 1px solid #eee;"><td style="padding: 8px; font-weight: bold; color: #555;">${key}</td><td style="padding: 8px;">${displayVal}</td></tr>`; 
        } 
        
        html += `</table><br>`; 
        if (assetIds['Photo 1']) html += `<h3>Photo 1</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 1']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`; 
        if (assetIds['Photo 2']) html += `<h3>Photo 2</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 2']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`; 
        if (assetIds['Photo 3']) html += `<h3>Photo 3</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 3']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`; 
        if (assetIds['Photo 4']) html += `<h3>Photo 4</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Photo 4']}&sz=w600" style="max-width:100%; border:1px solid #ccc; border-radius:8px; margin-bottom:10px;"><br>`; 
        if (assetIds['Signature']) html += `<h3>Authorized Signature</h3><img src="https://drive.google.com/thumbnail?id=${assetIds['Signature']}&sz=w400" style="max-width:300px; border-bottom:2px solid #000;"><br>`; 
        html += `</div>`; 
        
        MailApp.sendEmail({ to: recipient, subject: `[${CONFIG.ORG_NAME}] ${p['Template Name']} - ${worker}`, htmlBody: html }); 
    } catch(e) { console.log("Email Error: " + e); } 
}

function saveImageToDrive(base64String, filename) { try { const base64Clean = base64String.split(',').pop(); const data = Utilities.base64Decode(base64Clean); const blob = Utilities.newBlob(data, 'image/jpeg', filename); let folder; if (CONFIG.PHOTOS_FOLDER_ID && CONFIG.PHOTOS_FOLDER_ID.length > 5) { try { folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); } catch(e){ folder = DriveApp.getRootFolder(); } } else { folder = DriveApp.getRootFolder(); } const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return { url: file.getUrl(), id: file.getId() }; } catch(e) { return { url: "", id: "" }; } }
function checkOverdueVisits() { PropertiesService.getScriptProperties().setProperty('LAST_WATCHDOG_RUN', new Date().toISOString()); const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Visits'); if(!sheet) return; const lastRow = sheet.getLastRow(); if (lastRow <= 1) return; const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues(); const now = new Date().getTime(); const escalationMs = (CONFIG.ESCALATION_MINUTES || 15) * 60 * 1000; for (let i = 0; i < data.length; i++) { const row = data[i]; const status = row[10]; const dueTimeStr = row[20]; if (['DEPARTED', 'COMPLETED', 'SAFE - MONITOR CLEARED', 'SAFE - MANUALLY CLEARED'].includes(status) || !dueTimeStr) continue; const dueTime = new Date(dueTimeStr).getTime(); if (isNaN(dueTime)) continue; const timeOverdue = now - dueTime; if (timeOverdue > escalationMs && !status.includes('EMERGENCY')) { const newStatus = "EMERGENCY - OVERDUE (Server Watchdog)"; sheet.getRange(i + 2, 11).setValue(newStatus); sendAlert({ 'Worker Name': row[2], 'Worker Phone Number': row[3], 'Alarm Status': newStatus, 'Location Name': row[12], 'Last Known GPS': row[14], 'Notes': "Worker failed to check in.", 'Emergency Contact Email': row[6], 'Emergency Contact Number': row[5], 'Escalation Contact Email': row[9], 'Escalation Contact Number': row[8], 'Battery Level': row[16] }); } else if (timeOverdue > 0 && status === 'ON SITE') { sheet.getRange(i + 2, 11).setValue("OVERDUE"); } } }
function parseQuestions(row) { const questions = []; for(let i=4; i<row.length; i++) { const val = row[i]; if(val && val !== "") { let type='check', text=val; if(val.includes('[TEXT]')) { type='text'; text=val.replace('[TEXT]','').trim(); } else if(val.includes('[PHOTO]')) { type='photo'; text=val.replace('[PHOTO]','').trim(); } else if(val.includes('[YESNO]')) { type='yesno'; text=val.replace('[YESNO]','').trim(); } else if(val.includes('[NUMBER]')) { type='number'; text=val.replace('[NUMBER]','').trim(); } else if(val.includes('$')) { type='number'; text=val.replace('$','').trim(); } else if(val.includes('[GPS]')) { type='gps'; text=val.replace('[GPS]','').trim(); } else if(val.includes('[HEADING]')) { type='header'; text=val.replace('[HEADING]','').trim(); } else if(val.includes('[NOTE]')) { type='note'; text=val.replace('[NOTE]','').trim(); } else if(val.includes('[SIGN]')) { type='signature'; text=val.replace('[SIGN]','').trim(); } else if(val.includes('[DATE]')) { type='date'; text=val.replace('[DATE]','').trim(); } questions.push({type, text}); } } return questions; }
function setupReportTemplate() { try { const doc = DocumentApp.create(`${CONFIG.ORG_NAME} Master Report Template`); const body = doc.getBody(); body.appendParagraph(CONFIG.ORG_NAME).setHeading(DocumentApp.ParagraphHeading.HEADING1); body.appendParagraph("VISIT REPORT").setHeading(DocumentApp.ParagraphHeading.HEADING2); const cells = [["Worker:", "{{WorkerName}}"], ["Location:", "{{LocationName}}"], ["Date:", "{{Date}}"], ["Status:", "{{AlarmStatus}}"]]; cells.forEach(r => body.appendParagraph(`${r[0]} ${r[1]}`)); body.appendHorizontalRule(); body.appendParagraph("NOTES").setHeading(DocumentApp.ParagraphHeading.HEADING3); body.appendParagraph("{{Notes}}"); body.appendHorizontalRule(); body.appendParagraph("FORM DATA").setHeading(DocumentApp.ParagraphHeading.HEADING3); body.appendParagraph("{{VisitReportData}}"); body.appendHorizontalRule(); body.appendParagraph("Authorized Signature:").setHeading(DocumentApp.ParagraphHeading.HEADING4); body.appendParagraph("{{Signature}}"); doc.saveAndClose(); PropertiesService.getScriptProperties().setProperty('REPORT_TEMPLATE_ID', doc.getId()); return "SUCCESS: Template Created. ID: " + doc.getId(); } catch(e) { return "ERROR: " + e.toString(); } }
function sendAlert(data) { let recipients = [Session.getEffectiveUser().getEmail()]; let smsNumbers = []; if (data['Alarm Status'] === 'ESCALATION_SENT') { if(data['Escalation Contact Email']) recipients.push(data['Escalation Contact Email']); if(data['Escalation Contact Number']) smsNumbers.push(data['Escalation Contact Number']); } else { if(data['Emergency Contact Email']) recipients.push(data['Emergency Contact Email']); if(data['Emergency Contact Number']) smsNumbers.push(data['Emergency Contact Number']); } recipients = [...new Set(recipients)].filter(e => e && e.includes('@')); const subject = "🚨 SAFETY ALERT: " + data['Worker Name'] + " - " + data['Alarm Status']; const body = `<h1 style="color:red;">${data['Alarm Status']}</h1><p><strong>Worker:</strong> ${data['Worker Name']}</p><p><strong>Location:</strong> ${data['Location Name'] || 'Unknown'}</p><p><strong>Battery:</strong> ${data['Battery Level'] || 'Unknown'}</p><p><strong>Map:</strong> <a href="https://www.google.com/maps/search/?api=1&query=${data['Last Known GPS']}">${data['Last Known GPS']}</a></p>`; if(recipients.length > 0) MailApp.sendEmail({to: recipients.join(','), subject: subject, htmlBody: body}); const key = CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5 ? CONFIG.TEXTBELT_API_KEY : 'textbelt'; const smsMsg = `SOS: ${data['Worker Name']} - ${data['Alarm Status']} at ${data['Location Name']}`; smsNumbers.forEach(phone => { const clean = phone.replace(/^'/, '').replace(/[^0-9+]/g, ''); try { UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: clean, message: smsMsg, key: key }), muteHttpExceptions: true }); } catch(e) {} }); }
function smartScribe(text) { if (!CONFIG.GEMINI_API_KEY || !text || text.length < 5) return text; try { const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`; const payload = { "contents": [{ "parts": [{ "text": "Correct grammar to NZ English: " + text }] }] }; const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }); return JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim(); } catch (e) { return text; } }
function archiveOldData() { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Visits'); let archive = ss.getSheetByName('Archive'); if (!archive) archive = ss.insertSheet('Archive'); const data = sheet.getDataRange().getValues(); if (data.length <= 1) return; const today = new Date(); const rowsToKeep = [data[0]]; const rowsToArchive = []; for (let i = 1; i < data.length; i++) { const date = new Date(data[i][0]); const diff = (today - date) / (1000 * 60 * 60 * 24); if (diff > CONFIG.ARCHIVE_DAYS && (data[i][10] === 'DEPARTED' || data[i][10] === 'COMPLETED')) { rowsToArchive.push(data[i]); } else { rowsToKeep.push(data[i]); } } if (rowsToArchive.length > 0) { if (archive.getLastRow() === 0) archive.appendRow(data[0]); archive.getRange(archive.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive); sheet.clearContents(); sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep); } }
function runAllLongitudinalReports() { 
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    let allData = [];
    const sheet = ss.getSheetByName('Visits'); 
    if (sheet) {
        const vData = sheet.getDataRange().getValues();
        if(vData.length > 1) allData = allData.concat(vData.slice(1));
    }
    const archive = ss.getSheetByName('Archive');
    if (archive) {
        const aData = archive.getDataRange().getValues();
        if(aData.length > 1) allData = allData.concat(aData.slice(1));
    }
    if (allData.length === 0) return; 

    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM"); 
    const name = `Longitudinal Report - ${dateStr} - ${CONFIG.ORG_NAME}`; 
    let reportFile; 
    const files = DriveApp.getFilesByName(name); 
    if (files.hasNext()) reportFile = files.next(); else reportFile = DriveApp.getFileById(SpreadsheetApp.create(name).getId()); 
    const reportSS = SpreadsheetApp.open(reportFile); 
    
    let sheetAct = reportSS.getSheetByName('Worker Activity'); 
    if (sheetAct) sheetAct.clear(); else sheetAct = reportSS.insertSheet('Worker Activity'); 
    sheetAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered"]); 
    sheetAct.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dbeafe"); 
    
    const stats = {}; 
    for (let i = 0; i < allData.length; i++) { 
        const worker = allData[i][2]; 
        const status = allData[i][10]; 
        if (!stats[worker]) stats[worker] = { visits: 0, alerts: 0 }; 
        stats[worker].visits++; 
        if (status.includes("EMERGENCY") || status.includes("OVERDUE")) stats[worker].alerts++; 
    } 
    const actRows = Object.keys(stats).map(w => [w, stats[w].visits, stats[w].alerts]); 
    if (actRows.length > 0) sheetAct.getRange(2, 1, actRows.length, 3).setValues(actRows); 

    let sheetTrav = reportSS.getSheetByName('Travel Stats'); 
    if (sheetTrav) sheetTrav.clear(); else sheetTrav = reportSS.insertSheet('Travel Stats'); 
    sheetTrav.appendRow(["Worker Name", "Total Distance (km)", "Trips"]); 
    sheetTrav.getRange(1,1,1,3).setFontWeight("bold").setBackground("#dcfce7"); 
    
    const tStats = {}; 
    for (let i = 0; i < allData.length; i++) { 
        const worker = allData[i][2]; 
        const dist = parseFloat(allData[i][18]) || 0; 
        if (!tStats[worker]) tStats[worker] = { km: 0, trips: 0 }; 
        if (dist > 0) { tStats[worker].km += dist; tStats[worker].trips++; } 
    } 
    const travRows = Object.keys(tStats).map(w => [w, tStats[w].km.toFixed(2), tStats[w].trips]); 
    if (travRows.length > 0) sheetTrav.getRange(2, 1, travRows.length, 3).setValues(travRows); 
    
    MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: `Report: ${name}`, htmlBody: `<a href="${reportSS.getUrl()}">View Report</a>` }); 
}
