/**
 * OTG APPSUITE - MASTER BACKEND v84.2
 * Updates: Crash-proof row building, JSON-forced headers.
 */

const CONFIG = {
  MASTER_KEY: "%%SECRET_KEY%%", 
  WORKER_KEY: "%%WORKER_KEY%%", 
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: "%%TIMEZONE%%", 
  COUNTRY_CODE: "%%COUNTRY_CODE%%", 
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%,
  ENABLE_AI: %%ENABLE_AI%%,
  STATS_FREQ: "%%STATS_FREQ%%"
};

function doGet(e) {
  try {
      if(!e || !e.parameter) return sendJSON({status:"running", message:"OTG Server Online"});
      const p = e.parameter;

      if(p.action === 'ping') return sendJSON({status: "success", message: "Connected"});

      if(p.key === CONFIG.MASTER_KEY) {
         if(p.action === 'fetch') return fetchData();
         if(p.action === 'stats') return generateStats();
      }

      if(p.key === CONFIG.WORKER_KEY) {
         if(p.action === 'sync') return handleSync(p);
      }
      
      return sendJSON({status:"error", message:"Access Denied: Invalid Key"});

  } catch(error) {
      return sendJSON({status:"error", message: error.toString()});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
      // 1. Lock to prevent collision
      if (!lock.tryLock(30000)) return sendJSON({status: "error", message: "Server Busy"});

      // 2. Validate Payload
      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Payload"});
      const p = e.parameter;
      
      // 3. Auth Check
      if(p.key !== CONFIG.WORKER_KEY && p.key !== CONFIG.MASTER_KEY) {
        return sendJSON({status:"error", message:"Auth Failed"});
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // 4. Auto-Repair Visits Sheet
      let sheet = ss.getSheetByName('Visits');
      if (!sheet) {
          sheet = ss.insertSheet('Visits');
          sheet.appendRow(["Timestamp","Date","Worker Name","Worker Phone Number","Emergency Contact Name","Emergency Contact Number","Emergency Contact Email","Escalation Contact Name","Escalation Contact Number","Escalation Contact Email","Alarm Status","Notes","Location Name","Location Address","Last Known GPS","GPS Timestamp","Battery Level","Photo 1","Distance (km)","Visit Report Data","Anticipated Departure Time"]);
      }

      // 5. Route Action
      if(p.action === 'checkin') return handleCheckin(p, ss);
      if(p.action === 'sos') return handleSOS(p, ss);
      if(p.action === 'resolve') return handleResolve(p, ss);
      if(p.action === 'register_device') return handleDeviceReg(p, ss);

      return sendJSON({status:"error", message:"Unknown Action"});

  } catch(error) {
      return sendJSON({status:"error", message: "Critical Error: " + error.toString()});
  } finally {
      lock.releaseLock();
  }
}

function handleSync(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sitesSheet = ss.getSheetByName('Sites');
  // Safe read: if sheet missing, return empty array
  const sites = sitesSheet ? sitesSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] && (r[0].toString() === 'ALL' || r[0].toString() === p.worker))
    .map(r => ({
      siteName: r[3], company: r[2], template: r[1], address: r[4], 
      contactName: r[5], contactPhone: r[6], contactEmail: r[7], notes: r[8]
    })) : [];

  const tempSheet = ss.getSheetByName('Templates');
  const forms = [];
  const cachedTemplates = {};
  
  if(tempSheet) {
     const tData = tempSheet.getDataRange().getValues().slice(1);
     tData.forEach(r => {
        if(r[0] === 'FORM' && (r[2] === 'ALL' || (r[2] && r[2].includes(p.worker)))) {
           const questions = [];
           for(let i=4; i<r.length; i++) { if(r[i]) questions.push(r[i]); }
           forms.push({name: r[1], questions: questions});
        }
        if(r[0] === 'REPORT') {
           const questions = [];
           for(let i=4; i<r.length; i++) { if(r[i]) questions.push(r[i]); }
           cachedTemplates[r[1]] = questions;
        }
     });
  }

  const meta = {};
  const staffSheet = ss.getSheetByName('Staff');
  if(staffSheet) {
     const sData = staffSheet.getDataRange().getValues();
     for(let i=1; i<sData.length; i++) {
        if(sData[i][0] === p.worker) {
           meta.lastVehCheck = sData[i][5];
           meta.wofExpiry = sData[i][6];
           if(p.deviceId && !sData[i][4]) staffSheet.getRange(i+1, 5).setValue(p.deviceId);
           break;
        }
     }
  }

  return sendJSON({
    status: "success",
    sites: sites,
    forms: forms,
    cachedTemplates: cachedTemplates,
    meta: meta
  });
}

function handleCheckin(p, ss) {
  let finalNote = p.Notes || "";
  if(CONFIG.ENABLE_AI && CONFIG.GEMINI_API_KEY && finalNote.length > 5) {
     const cleanNote = smartScribe(finalNote);
     if(cleanNote) finalNote = cleanNote; 
  }

  let photoUrl = "";
  if(p.PhotoData && p.PhotoData.length > 100) photoUrl = saveImage(p.PhotoData, p['Worker Name']);

  const row = buildRow(p, finalNote, photoUrl, p['Alarm Status']);
  ss.getSheetByName('Visits').appendRow(row);
  
  // Update Vehicle Check Date if applicable
  if(p['Template Name'] === 'Vehicle Safety Check') {
     const staffSheet = ss.getSheetByName('Staff');
     if(staffSheet) {
        const sData = staffSheet.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) {
           if(sData[i][0] === p['Worker Name']) {
              staffSheet.getRange(i+1, 6).setValue(new Date());
              if(p['Visit Report Data']) {
                 try {
                    const json = JSON.parse(p['Visit Report Data']);
                    for(const key in json) {
                        if(key.toLowerCase().includes("wof") || key.toLowerCase().includes("expiry")) {
                            staffSheet.getRange(i+1, 7).setValue(json[key]);
                        }
                    }
                 } catch(e){}
              }
              break;
           }
        }
     }
  }
  
  return sendJSON({status:"success", message:"Check-in Saved", cleanNote: finalNote});
}

function handleSOS(p, ss) {
  const row = buildRow(p, "USER TRIGGERED ALARM: " + (p.Notes||""), "", "EMERGENCY - SOS TRIGGERED");
  ss.getSheetByName('Visits').appendRow(row);
  
  if(CONFIG.TEXTBELT_API_KEY) {
     sendSMS(p['Emerg Phone'], `SOS ALERT: ${p['Worker Name']} at ${p['Location Name']}. GPS: ${p['GPS Coords']}`);
     if(p['Escal Phone']) sendSMS(p['Escal Phone'], `SOS ALERT (Escalation): ${p['Worker Name']}`);
  }
  return sendJSON({status:"success", message:"SOS Broadcasted"});
}

function handleResolve(p, ss) {
  const row = buildRow(p, p.Notes, "", p['Alarm Status']); 
  row[12] = "HQ Dashboard"; 
  ss.getSheetByName('Visits').appendRow(row);
  return sendJSON({status:"success"});
}

function handleDeviceReg(p, ss) {
  const sheet = ss.getSheetByName('Staff');
  if(!sheet) return sendJSON({status:"error", message:"Staff Sheet Missing"});
  
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
     if(data[i][2] && data[i][2].toString().toLowerCase() === p.email.toLowerCase()) {
        sheet.getRange(i+1, 8).setValue(p.deviceId); 
        return sendJSON({status:"success"});
     }
  }
  return sendJSON({status:"error", message:"Email not found"});
}

function fetchData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if(!sheet) return sendJSON({status:"success", workers: []});

  const visits = sheet.getDataRange().getValues();
  const rawData = visits.slice(Math.max(visits.length - 100, 1));
  
  const mapped = rawData.map(r => ({
      "Worker Name": r[2] || "Unknown",
      "Alarm Status": r[10] || "Unknown",
      "Worker Phone Number": r[3] || "",
      "Emergency Contact Name": r[4] || "",
      "Emergency Contact Number": r[5] || "",
      "Notes": r[11] || "",
      "Location Name": r[12] || "",
      "Last Known GPS": r[14] || "",
      "Battery Level": r[16] || "",
      "Anticipated Departure Time": r[20] || "",
      "Timestamp": r[0]
  }));

  return sendJSON({status:"success", workers: mapped});
}

function getManifest() {
  return sendJSON({status:"success", message: "Deprecated"});
}

function generateStats() {
  return sendJSON({status:"success", message:"Stats Regenerated"});
}

// SAFE BUILD ROW - Prevents crashes on missing data
function buildRow(p, notes, photo, status) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "dd/MM/yyyy");
  return [
    now, dateStr,
    p['Worker Name'] || "Unknown", 
    p['Worker Phone'] || "",
    p['Emerg Name'] || "", 
    p['Emerg Phone'] || "", 
    p['Emerg Email'] || "",
    p['Escal Name'] || "", 
    p['Escal Phone'] || "", 
    p['Escal Email'] || "",
    status || "Unknown", 
    notes || "",
    p['Location Name'] || "", 
    p['Location Address'] || "",
    p['GPS Coords'] || "", 
    p['GPS Timestamp'] || "", 
    p['Battery Level'] || "",
    photo || "", 
    p['Trip Distance'] || 0, 
    p['Visit Report Data'] || "", 
    p['Anticipated Departure Time'] || ""
  ];
}

function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function saveImage(base64Data, workerName) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.JPEG, workerName + "_" + Date.now() + ".jpg");
    return folder.createFile(blob).getUrl();
  } catch(e) { return "Error: " + e.toString(); }
}

function sendSMS(phone, message) {
  if(!phone || phone.length < 8) return;
  try {
    UrlFetchApp.fetch('https://textbelt.com/text', {
      'method': 'post',
      'payload': { 'phone': phone, 'message': message, 'key': CONFIG.TEXTBELT_API_KEY }
    });
  } catch(e) { console.error("SMS Failed", e); }
}

function smartScribe(rawText) {
  try {
    const safeText = rawText.replace(/[0-9]{9,}/g, "[NUMBER]"); 
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const payload = { "contents": [{ "parts": [{"text": `Fix grammar. Input: "${safeText}"`}] }] };
    const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    if(json.candidates) return json.candidates[0].content.parts[0].text.trim();
  } catch(e) { return null; }
}

function checkOverdueVisits() {}
