/**
 * OTG APPSUITE - MASTER BACKEND v77.4 (Global + Privacy + Scheduled Stats)
 * Features: Dynamic Region, PII Redaction, Configurable Stats Schedule.
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
  STATS_FREQ: "%%STATS_FREQ%%" // DAILY, WEEKLY, MONTHLY
};

// ==========================================
// 1. GET HANDLER
// ==========================================
function doGet(e) {
  try {
      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Params"});
      const p = e.parameter;

      if(p.action === 'ping') return sendJSON({status: "success", message: "Connected"});

      // AUTH REQUIRED ACTIONS
      if(p.key === CONFIG.MASTER_KEY) {
         if(p.action === 'fetch') return fetchData();
         if(p.action === 'stats') return generateStats(); // Manual Trigger
      }

      if(p.key === CONFIG.WORKER_KEY) {
         if(p.action === 'manifest') return getManifest();
      }
      
      return sendJSON({status:"error", message:"Access Denied"});

  } catch(error) {
      return sendJSON({status:"error", message: error.toString()});
  }
}

// ==========================================
// 2. POST HANDLER
// ==========================================
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
      
      // Routing
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

// ==========================================
// 3. CORE LOGIC
// ==========================================

function handleCheckin(p, ss) {
  // PII & AI Scrubbing
  let finalNote = p.Notes || "";
  if(CONFIG.ENABLE_AI && CONFIG.GEMINI_API_KEY && finalNote.length > 5) {
     const cleanNote = smartScribe(finalNote);
     if(cleanNote) finalNote = cleanNote; 
  }

  // Photo
  let photoUrl = "";
  if(p.PhotoData && p.PhotoData.length > 100) photoUrl = saveImage(p.PhotoData, p['Worker Name']);

  const row = buildRow(p, finalNote, photoUrl, p['Alarm Status']);
  ss.getSheetByName('Visits').appendRow(row);
  
  return sendJSON({status:"success", message:"Check-in Saved", cleanNote: finalNote});
}

function handleSOS(p, ss) {
  const row = buildRow(p, "USER TRIGGERED ALARM: " + (p.Notes||""), "", "EMERGENCY - SOS TRIGGERED");
  ss.getSheetByName('Visits').appendRow(row);
  
  // Escalation
  if(CONFIG.TEXTBELT_API_KEY) {
     sendSMS(p['Emerg Phone'], `SOS ALERT: ${p['Worker Name']} at ${p['Location Name']}. GPS: ${p['GPS Coords']}`);
     if(p['Escal Phone']) sendSMS(p['Escal Phone'], `SOS ALERT (Escalation): ${p['Worker Name']}`);
  }
  return sendJSON({status:"success", message:"SOS Broadcasted"});
}

function handleResolve(p, ss) {
  // Log the resolution as a new line event
  const row = buildRow(p, p.Notes, "", p['Alarm Status']); // Alarm Status should be SAFE
  row[12] = "HQ Dashboard"; // Location Name override
  ss.getSheetByName('Visits').appendRow(row);
  return sendJSON({status:"success"});
}

function handleDeviceReg(p, ss) {
  const sheet = ss.getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
     if(data[i][2].toString().toLowerCase() === p.email.toLowerCase()) {
        sheet.getRange(i+1, 8).setValue(p.deviceId); 
        return sendJSON({status:"success"});
     }
  }
  return sendJSON({status:"error", message:"Email not found"});
}

// ==========================================
// 4. STATS & REPORTING
// ==========================================

function generateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visits = ss.getSheetByName('Visits').getDataRange().getValues();
  
  // Containers
  const workerStats = {}; // { name: { trips: 0, distance: 0, alerts: 0, visits: 0 } }

  // Skip Header (Row 0)
  for(let i=1; i<visits.length; i++) {
      const row = visits[i];
      const name = row[2]; // Worker Name
      const status = row[10]; // Alarm Status
      const dist = parseFloat(row[18]) || 0; // Trip Distance

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

  // 1. Write Activity Stats
  let sAct = ss.getSheetByName('Activity Stats');
  if(!sAct) sAct = ss.insertSheet('Activity Stats');
  sAct.clear();
  sAct.appendRow(["Worker Name", "Total Visits", "Alerts Triggered", "Last Updated"]);
  sAct.getRange(1,1,1,4).setFontWeight("bold").setBackground("#e0f2fe"); // Blue header
  
  const actRows = Object.keys(workerStats).map(w => [
      w, workerStats[w].visits, workerStats[w].alerts, new Date()
  ]);
  if(actRows.length > 0) sAct.getRange(2,1,actRows.length, 4).setValues(actRows);

  // 2. Write Travel Stats
  let sTrav = ss.getSheetByName('Travel Stats');
  if(!sTrav) sTrav = ss.insertSheet('Travel Stats');
  sTrav.clear();
  sTrav.appendRow(["Worker Name", "Trips Recorded", "Total Distance (km)", "Last Updated"]);
  sTrav.getRange(1,1,1,4).setFontWeight("bold").setBackground("#dcfce7"); // Green header
  
  const travRows = Object.keys(workerStats).map(w => [
      w, workerStats[w].trips, workerStats[w].distance.toFixed(2), new Date()
  ]);
  if(travRows.length > 0) sTrav.getRange(2,1,travRows.length, 4).setValues(travRows);

  return sendJSON({status:"success", message:"Stats Regenerated"});
}

// ==========================================
// 5. DATA FETCHERS
// ==========================================

function fetchData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visits = ss.getSheetByName('Visits').getDataRange().getValues();
  const recent = visits.slice(Math.max(visits.length - 100, 1));
  return sendJSON({status:"success", data: recent});
}

function getManifest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staff = ss.getSheetByName('Staff').getDataRange().getValues().slice(1)
    .map(r => ({name: r[0], phone: r[1], email: r[2], pin: r[3], deviceId: r[7]}));
  const sites = ss.getSheetByName('Sites').getDataRange().getValues().slice(1)
    .map(r => ({id: r[0], name: r[1], lat: r[2], lng: r[3], address: r[4]}));
  const templates = ss.getSheetByName('Templates').getDataRange().getValues().slice(1)
    .map(r => ({id: r[0], name: r[1], fields: r[2]})); 
  return sendJSON({status:"success", staff: staff, sites: sites, templates: templates});
}

// ==========================================
// 6. UTILITIES
// ==========================================

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
    photo, p['Trip Distance'] || 0
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

// ==========================================
// 7. PRIVACY & AI 
// ==========================================

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

// ==========================================
// 8. WATCHDOG (Triggers every 5-10 mins)
// ==========================================
function checkOverdueVisits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const lastSeen = {};
  
  // A. Scan for Overdue Workers
  for(let i=1; i<data.length; i++) {
     const w = data[i][2];
     const t = new Date(data[i][0]);
     if(!lastSeen[w] || t > lastSeen[w].time) {
       lastSeen[w] = { row: i+1, time: t, status: data[i][10], name: w, emerg: data[i][5] };
     }
  }
  
  const limit = CONFIG.ESCALATION_MINUTES * 60 * 1000;
  
  for(const w in lastSeen) {
     const e = lastSeen[w];
     if(e.status !== "SAFE - MANUALLY CLEARED" && e.status !== "DEPARTED" && !e.status.includes("EMERGENCY")) {
        const diff = now - e.time;
        if(diff > limit) {
           sheet.getRange(e.row, 11).setValue("EMERGENCY - OVERDUE (Watchdog)");
           if(CONFIG.TEXTBELT_API_KEY) sendSMS(e.emerg, `URGENT: ${e.name} is OVERDUE. Checked in ${(diff/60000).toFixed(0)} mins ago.`);
        }
     }
  }
  
  // B. Scheduled Stats Generation
  // Check if it is Midnight (00:00 - 00:15)
  if(now.getHours() === 0 && now.getMinutes() < 15) {
      let shouldRun = false;
      const freq = CONFIG.STATS_FREQ || "MONTHLY";

      if(freq === "DAILY") shouldRun = true;
      else if(freq === "WEEKLY" && now.getDay() === 1) shouldRun = true; // Monday
      else if(freq === "MONTHLY" && now.getDate() === 1) shouldRun = true; // 1st of Month
      
      if(shouldRun) generateStats();
  }
}
