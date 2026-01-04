/**
 * OTG APPSUITE - MASTER BACKEND v82.3 (Fixed)
 * Verifies: Key Authentication Only. No Version Checks.
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
         if(p.action === 'sync') return handleSync(p); // Specific Sync Handler
         if(p.action === 'manifest') return getManifest();
      }
      
      return sendJSON({status:"error", message:"Access Denied: Invalid Key"});

  } catch(error) {
      return sendJSON({status:"error", message: error.toString()});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
      if (!lock.tryLock(30000)) return sendJSON({status: "error", message: "Server Busy"});

      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Payload"});
      const p = e.parameter;
      
      if(p.key !== CONFIG.WORKER_KEY && p.key !== CONFIG.MASTER_KEY) {
        return sendJSON({status:"error", message:"Auth Failed"});
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      if(p.action === 'checkin') return handleCheckin(p, ss);
      if(p.action === 'sos') return handleSOS(p, ss);
      if(p.action === 'resolve') return handleResolve(p, ss);
      if(p.action === 'register_device') return handleDeviceReg(p, ss);

      return sendJSON({status:"error", message:"Unknown Action"});

  } catch(error) {
      return sendJSON({status:"error", message: "Critical: " + error.toString()});
  } finally {
      lock.releaseLock();
  }
}

// === NEW SYNC HANDLER ===
function handleSync(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Sites
  const sitesSheet = ss.getSheetByName('Sites');
  if(!sitesSheet) return sendJSON({status:"error", message:"Missing Tab: 'Sites'"});
  
  const siteData = sitesSheet.getDataRange().getValues().slice(1);
  const sites = siteData
    .filter(r => r[0].toString() === 'ALL' || r[0].toString() === p.worker) // Filter by Assigned To
    .map(r => ({
      siteName: r[3],
      company: r[2],
      template: r[1],
      address: r[4],
      contactName: r[5],
      contactPhone: r[6],
      contactEmail: r[7],
      notes: r[8]
    }));

  // 2. Get Templates
  const tempSheet = ss.getSheetByName('Templates');
  const forms = [];
  if(tempSheet) {
     const tData = tempSheet.getDataRange().getValues().slice(1);
     tData.forEach(r => {
        if(r[0] === 'FORM' && (r[2] === 'ALL' || r[2].includes(p.worker))) {
           const questions = [];
           for(let i=4; i<r.length; i++) { if(r[i]) questions.push(r[i]); }
           forms.push({name: r[1], questions: questions});
        }
     });
  }

  // 3. Cache Report Templates
  const cachedTemplates = {};
  if(tempSheet) {
     const tData = tempSheet.getDataRange().getValues().slice(1);
     tData.forEach(r => {
        if(r[0] === 'REPORT') {
           const questions = [];
           for(let i=4; i<r.length; i++) { if(r[i]) questions.push(r[i]); }
           cachedTemplates[r[1]] = questions;
        }
     });
  }

  // 4. Meta Data (Vehicle checks etc)
  const meta = {};
  const staffSheet = ss.getSheetByName('Staff');
  if(staffSheet) {
     const sData = staffSheet.getDataRange().getValues();
     for(let i=1; i<sData.length; i++) {
        if(sData[i][0] === p.worker) {
           meta.lastVehCheck = sData[i][5];
           meta.wofExpiry = sData[i][6];
           // Auto-register device ID if missing
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
  
  // Handle Vehicle Checks specifically
  if(p['Template Name'] === 'Vehicle Safety Check') {
     const staffSheet = ss.getSheetByName('Staff');
     if(staffSheet) {
        const sData = staffSheet.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) {
           if(sData[i][0] === p['Worker Name']) {
              staffSheet.getRange(i+1, 6).setValue(new Date()); // Update Last Check
              // If WOF Expiry was captured in the form, update it
              if(p['Visit Report Data']) {
                 try {
                    const json = JSON.parse(p['Visit Report Data']);
                    // Look for date keys
                    for(const key in json) {
                        if(key.includes("WOF") || key.includes("Expiry")) {
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
  if(!sheet) return sendJSON({status:"error", message:"No Staff Tab"});
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
     if(data[i][2].toString().toLowerCase() === p.email.toLowerCase()) {
        sheet.getRange(i+1, 8).setValue(p.deviceId); 
        return sendJSON({status:"success"});
     }
  }
  return sendJSON({status:"error", message:"Email not found"});
}

function generateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visits = ss.getSheetByName('Visits').getDataRange().getValues();
  const workerStats = {};

  for(let i=1; i<visits.length; i++) {
      const row = visits[i];
      const name = row[2]; 
      const status = row[10]; 
      const dist = parseFloat(row[18]) || 0; 

      if(!workerStats[name]) workerStats[name] = { trips: 0, distance: 0, alerts: 0, visits: 0 };
      
      workerStats[name].visits++;
      if(dist > 0) {
          workerStats[name].distance += dist;
          workerStats[name].trips++;
      }
      if(status.toString().includes('EMERGENCY')) {
          workerStats[name].alerts++;
      }
  }

  let sAct = ss.getSheetByName('Activity Stats');
  if(!sAct) sAct = ss.insertSheet('Activity Stats');
  sAct.clear();
  sAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Last Updated"]);
  sAct.getRange(1,1,1,4).setFontWeight("bold").setBackground("#e0f2fe");
  
  const actRows = Object.keys(workerStats).map(w => [w, workerStats[w].visits, workerStats[w].alerts, new Date()]);
  if(actRows.length > 0) sAct.getRange(2,1,actRows.length, 4).setValues(actRows);

  let sTrav = ss.getSheetByName('Travel Stats');
  if(!sTrav) sTrav = ss.insertSheet('Travel Stats');
  sTrav.clear();
  sTrav.appendRow(["Worker Name", "Trips Recorded", "Total Distance (km)", "Last Updated"]);
  sTrav.getRange(1,1,1,4).setFontWeight("bold").setBackground("#dcfce7");
  
  const travRows = Object.keys(workerStats).map(w => [w, workerStats[w].trips, workerStats[w].distance.toFixed(2), new Date()]);
  if(travRows.length > 0) sTrav.getRange(2,1,travRows.length, 4).setValues(travRows);

  return sendJSON({status:"success", message:"Stats Regenerated"});
}

function fetchData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visits = ss.getSheetByName('Visits').getDataRange().getValues();
  // Get header to map columns correctly
  const header = visits[0];
  // Helper to find index
  const getIdx = (name) => header.indexOf(name);
  
  const rawData = visits.slice(Math.max(visits.length - 100, 1));
  
  const mapped = rawData.map(r => ({
      "Worker Name": r[2],
      "Alarm Status": r[10],
      "Worker Phone Number": r[3],
      "Emergency Contact Name": r[4],
      "Emergency Contact Number": r[5],
      "Notes": r[11],
      "Location Name": r[12],
      "Last Known GPS": r[14],
      "Battery Level": r[16],
      "Anticipated Departure Time": r[20],
      "Timestamp": r[0]
  }));

  return sendJSON({status:"success", workers: mapped});
}

function getManifest() {
  return sendJSON({status:"success", message: "Manifest deprecated. Use sync."});
}

function buildRow(p, notes, photo, status) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "dd/MM/yyyy");
  return [
    now, dateStr,
    p['Worker Name'], p['Worker Phone'],
    p['Emerg Name'], p['Emerg Phone'], p['Emerg Email'],
    p['Escal Name'], p['Escal Phone'], p['Escal Email'],
    status, notes,
    p['Location Name'], p['Location Address'],
    p['GPS Coords'], p['GPS Timestamp'], p['Battery Level'],
    photo, p['Trip Distance'] || 0,
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

function redactPII(text) {
    if (!text) return "";
    let safeText = text.replace(/[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/g, "[EMAIL]");
    const cc = CONFIG.COUNTRY_CODE || "64";
    const phonePattern = new RegExp('(?:\\+?(?:' + cc + ')|0)[0-9]{1,4}[ -]?[0-9]{3,4}[ -]?[0-9]{3,9}', 'g');
    safeText = safeText.replace(phonePattern, "[PHONE]");
    safeText = safeText.replace(/(?<= )\b[A-Z][a-z]+\b/g, (match) => {
        const common = ["The","A","An","Is","In","At","On","To","For","With","High","Low","Safe","Risk","Note","Visit","Check","Site"];
        return common.includes(match) ? match : "[NAME]"; 
    });
    return safeText;
}

function smartScribe(rawText) {
  try {
    const safeText = redactPII(rawText);
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const payload = { "contents": [{ "parts": [{"text": `Fix grammar/spelling (Regional English). concise. Input: "${safeText}"`}] }] };
    const response = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    if(json.candidates) return json.candidates[0].content.parts[0].text.trim();
  } catch(e) { return null; }
}

function checkOverdueVisits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const lastSeen = {};
  
  for(let i=1; i<data.length; i++) {
     const w = data[i][2];
     const t = new Date(data[i][0]);
     if(!lastSeen[w] || t > lastSeen[w].time) {
       lastSeen[w] = { 
         row: i+1, 
         time: t, 
         status: data[i][10].toString(), 
         name: w, 
         emerg: data[i][5] 
       };
     }
  }
  
  const standardLimit = CONFIG.ESCALATION_MINUTES * 60 * 1000;
  
  for(const w in lastSeen) {
     const e = lastSeen[w];
     const isSafe = e.status === "SAFE - MANUALLY CLEARED" || e.status === "DEPARTED";
     const isAlreadyEmergency = e.status.includes("EMERGENCY");
     
     if(!isSafe && !isAlreadyEmergency) {
        let limit = standardLimit;
        if(e.status.includes("HIGH RISK")) {
            limit = 60 * 1000; 
        }

        const diff = now - e.time;
        if(diff > limit) {
           sheet.getRange(e.row, 11).setValue("EMERGENCY - OVERDUE (Watchdog)");
           if(CONFIG.TEXTBELT_API_KEY) {
              const type = e.status.includes("HIGH RISK") ? "ZERO TOLERANCE" : "Standard";
              sendSMS(e.emerg, `URGENT (${type}): ${e.name} is OVERDUE. Last check-in: ${(diff/60000).toFixed(0)} mins ago.`);
           }
        }
     }
  }
  
  if(now.getHours() === 0 && now.getMinutes() < 15) {
      let shouldRun = false;
      const freq = CONFIG.STATS_FREQ || "MONTHLY";
      if(freq === "DAILY") shouldRun = true;
      else if(freq === "WEEKLY" && now.getDay() === 1) shouldRun = true; 
      else if(freq === "MONTHLY" && now.getDate() === 1) shouldRun = true; 
      if(shouldRun) generateStats();
  }
}
