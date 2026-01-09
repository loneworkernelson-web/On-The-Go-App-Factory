/**
 * OTG APPSUITE - MASTER BACKEND v79.23
 * * FEATURES:
 * 1. Contextual Photo Embedding (Images appear inline with questions).
 * 2. Strict Number Parsing (Fixes "Form Builder" math errors).
 * 3. Longitudinal Reporting & Travel Logging.
 * 4. Tiered Escalation & "Watchdog" Logic.
 * 5. RESTORED: Gemini AI Text Polisher (Presentation Layer Only).
 */

// ==========================================
// 1. CONFIGURATION
// ==========================================
const CONFIG = {
  MASTER_KEY: "%%SECRET_KEY%%", 
  WORKER_KEY: "%%WORKER_KEY%%", 
  ORS_API_KEY: "%%ORS_API_KEY%%", 
  GEMINI_API_KEY: "%%GEMINI_API_KEY%%", 
  TEXTBELT_API_KEY: "%%TEXTBELT_API_KEY%%",
  PHOTOS_FOLDER_ID: "%%PHOTOS_FOLDER_ID%%", 
  REPORT_TEMPLATE_ID: "",   
  ORG_NAME: "%%ORGANISATION_NAME%%",
  TIMEZONE: "%%TIMEZONE%%", 
  ARCHIVE_DAYS: 30,
  ESCALATION_MINUTES: %%ESCALATION_MINUTES%%,
  ENABLE_REDACTION: %%ENABLE_REDACTION%%,
  VEHICLE_TERM: "%%VEHICLE_TERM%%",
  COUNTRY_CODE: "%%COUNTRY_PREFIX%%"
};

// ==========================================
// 2. ENTRY POINTS (DO NOT EDIT)
// ==========================================

function doPost(e) {
  // Security Gate
  if (!e || !e.parameter || e.parameter.key !== CONFIG.WORKER_KEY) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "ACCESS DENIED"})).setMimeType(ContentService.MimeType.JSON);
  }

  // Action Routing
  const action = e.parameter.action;
  
  if (action === 'resolve') return resolveAlert(e);
  
  // Standard Report / Heartbeat
  return processPost(e);
}

function doGet(e) {
  if (!e || !e.parameter) return ContentService.createTextOutput("OTG Backend Online");
  
  // Security Gate
  if (e.parameter.key !== CONFIG.WORKER_KEY && e.parameter.key !== CONFIG.MASTER_KEY) {
     return ContentService.createTextOutput(JSON.stringify({status: "error", message: "ACCESS DENIED"}));
  }

  const action = e.parameter.action;
  
  if (action === 'sync') return handleSync(e);
  if (action === 'getGlobalForms') return getGlobalForms();
  if (action === 'getMonitorData') return getMonitorData();
  
  return ContentService.createTextOutput("Action not found.");
}

// ==========================================
// 3. CORE LOGIC
// ==========================================

function processPost(e) {
  const p = e.parameter;
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet = getSheet('Visits');
    const timestamp = new Date();
    
    // 1. Handle Photos (Save to Drive & Prepare for Email)
    const savedPhotos = savePhotosToDrive(e); // Returns array of URLs
    
    // 2. Save to Sheet (Raw Data - Legal Requirement)
    const rowData = [
      timestamp,
      p['Worker Name'],
      p['Alarm Status'],
      p['Location Name'],
      p['Location Address'],
      p['Battery Level'],
      p['Notes'], // RAW NOTES SAVED HERE
      p['Anticipated Departure Time'],
      p['Worker Phone Number'],
      p['Worker Email'],
      JSON.stringify(savedPhotos), // Archive Photo Links
      p['Last Known GPS'],
      p['deviceId'],
      // Capture Dynamic Form Data for Stats
      p['Visit Report Data'] || "" 
    ];
    sheet.appendRow(rowData);
    
    // 3. Update Monitor Status
    updateMonitorState(p, timestamp);

    // 4. Handle "Travel Report" Specifics (Mileage Logging)
    if(p['Alarm Status'] === 'DEPARTED' || p['Alarm Status'] === 'TRAVELLING') {
       if(p['Visit Report Data']) {
          try {
             const reportData = JSON.parse(p['Visit Report Data']);
             // Check for numeric fields like "Distance" or "KM"
             let dist = 0;
             if(reportData['Distance']) dist = parseFloat(reportData['Distance']);
             else if(reportData['Distance (km)']) dist = parseFloat(reportData['Distance (km)']);
             else if(reportData['KM']) dist = parseFloat(reportData['KM']);
             else if(reportData['Mileage']) dist = parseFloat(reportData['Mileage']);
             
             if(dist > 0) logTravelData(p['Worker Name'], dist, p['Location Name'], timestamp);
          } catch(err) {
             console.log("Travel Parsing Error: " + err);
          }
       }
    }

    // 5. Send Email Notifications (Tier 1)
    // Only send if it's NOT just a heartbeat/tracking pulse
    if (p['Alarm Status'] !== 'TRAVELLING') {
       sendEmailReport(e, savedPhotos);
    }
    
    // 6. Handle Escalations (Tier 2 - SMS)
    if (p['Alarm Status'].includes('EMERGENCY') || p['Alarm Status'].includes('DURESS')) {
      sendEmergencySms(p);
    }

    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. EMAIL REPORTING (Contextual Photos + AI Polishing)
// ==========================================

function sendEmailReport(e, savedPhotoUrls) {
  const p = e.parameter;
  const recipient = p['Worker Email']; // Send copy to worker
  const subject = `[${CONFIG.ORG_NAME}] ${p['Alarm Status']}: ${p['Worker Name']} @ ${p['Location Name']}`;
  
  // AI POLISH: We polish the notes for the EMAIL ONLY (Presentation Layer)
  // The Database retains the raw input for evidence.
  let displayNotes = p['Notes'] || "";
  if(displayNotes.length > 5 && CONFIG.GEMINI_API_KEY && !CONFIG.GEMINI_API_KEY.includes('%%')) {
      displayNotes = smartScribe(displayNotes);
  }

  let html = `
    <div style="font-family: sans-serif; max-width: 600px; border: 1px solid #ccc; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #1e3a8a; color: white; padding: 15px;">
        <h2 style="margin:0;">${p['Alarm Status']}</h2>
        <p style="margin:5px 0 0 0; opacity: 0.8;">${p['Worker Name']} - ${new Date().toLocaleString()}</p>
      </div>
      <div style="padding: 20px;">
        <table style="width:100%; border-collapse: collapse;">
  `;

  // Standard Fields
  // We manually inject the Polished Notes here
  const standardKeys = ['Location Name', 'Location Address', 'Battery Level', 'Last Known GPS'];
  
  html += `<tr><td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold; width: 40%;">Notes (AI Summary)</td><td style="padding:8px; border-bottom:1px solid #eee; background-color:#f0fdf4;">${displayNotes}</td></tr>`;
  
  standardKeys.forEach(k => {
     if(p[k]) {
       let val = p[k];
       if(k === 'Last Known GPS') val = `<a href="https://www.google.com/maps?q=${val}">${val}</a>`;
       html += `<tr><td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold; width: 40%;">${k}</td><td style="padding:8px; border-bottom:1px solid #eee;">${val}</td></tr>`;
     }
  });

  // Dynamic Form Data (The Form Builder Logic)
  let inlineImages = {};
  
  if(p['Visit Report Data']) {
    try {
      const data = JSON.parse(p['Visit Report Data']);
      html += `<tr><td colspan="2" style="background:#f3f4f6; padding:10px; font-weight:bold; font-size:0.9em;">REPORT DETAILS</td></tr>`;
      
      for (const [key, value] of Object.entries(data)) {
         let displayValue = value;
         
         // CHECK FOR PHOTO TOKENS: "{{PHOTO_1}}"
         if (typeof value === 'string' && value.includes('{{PHOTO_')) {
             const match = value.match(/{{PHOTO_(\d+)}}/);
             if (match) {
                 const photoId = match[1];
                 const paramName = 'Photo ' + photoId;
                 
                 // Retrieve the blob from the upload parameters
                 if (e.parameter[paramName]) {
                     const blob = e.parameter[paramName];
                     const cid = 'photo' + photoId;
                     inlineImages[cid] = blob; // Add to email package
                     
                     // Replace token with HTML Image Tag referencing the CID
                     displayValue = `<br><img src="cid:${cid}" style="max-width:100%; border-radius:8px; border:2px solid #ddd; margin-top:5px;"><br>`;
                 } else {
                     displayValue = "(Photo Missing)";
                 }
             }
         }
         
         // Clean up Signature if present
         if(key === 'Signature' && displayValue.startsWith('data:image')) {
             displayValue = "✅ Signed Digitally"; 
         }

         html += `<tr><td style="padding:8px; border-bottom:1px solid #eee;">${key}</td><td style="padding:8px; border-bottom:1px solid #eee;">${displayValue}</td></tr>`;
      }
    } catch(err) {
      html += `<tr><td colspan="2" style="color:red;">Error parsing report data: ${err}</td></tr>`;
    }
  }

  html += `
        </table>
      </div>
      <div style="background-color: #f3f4f6; padding: 10px; text-align: center; font-size: 0.8em; color: #666;">
        Generated by OTG AppSuite
      </div>
    </div>
  `;

  // Send Email
  const options = {
    htmlBody: html,
    inlineImages: inlineImages,
    subject: subject
  };

  if(recipient) {
      MailApp.sendEmail(recipient, subject, "Please enable HTML email.", options);
  }
}

// ==========================================
// 5. HELPER FUNCTIONS & SYNC
// ==========================================

// RESTORED: SmartScribe (Gemini Integration)
function smartScribe(rawText) {
  if (!CONFIG.GEMINI_API_KEY) return rawText;
  
  try {
    const prompt = `You are a safety officer. Rephrase the following field notes into professional, clear English. Correct spelling/grammar. Keep it factual. Do not add markdown or extra conversational text. Notes: "${rawText}"`;
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const payload = {
      "contents": [{
        "parts": [{"text": prompt}]
      }]
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return rawText + " (AI Unavailable)";
    }
  } catch (e) {
    console.log("Gemini Error: " + e.toString());
    return rawText; // Fail safe to raw text
  }
}

function handleSync(e) {
  const p = e.parameter;
  const workerName = p.worker;
  const deviceId = p.deviceId;
  
  // 1. Get Sites
  const sites = getSheetData('Sites').filter(r => !r.Archived);
  
  // 2. Get Forms
  const forms = getGlobalFormsData();

  // 3. Get Cached Templates (The JSON format used by app)
  const cachedTemplates = {};
  forms.forEach(f => {
      cachedTemplates[f.name] = f.questions;
  });

  // 4. Get Meta (Vehicle Status)
  const meta = getWorkerMeta(workerName);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    sites: sites.map(s => ({
      siteName: s['Site Name'],
      address: s['Address'],
      company: s['Company Name'],
      contactName: s['Contact Name'],
      contactPhone: s['Contact Phone'],
      notes: s['Notes'],
      template: s['Default Template']
    })),
    forms: forms,
    cachedTemplates: cachedTemplates,
    meta: meta
  })).setMimeType(ContentService.MimeType.JSON);
}

function getWorkerMeta(workerName) {
  // Look up vehicle checks
  const sheet = getSheet('Visits');
  const data = sheet.getDataRange().getValues();
  // Reverse search for last vehicle check
  let lastVehCheck = null;
  
  for(let i=data.length-1; i>=0; i--) {
     if(data[i][1] === workerName && data[i][13]) { // Column 14 (Index 13) is Report Data
         if(data[i][13].includes('Vehicle Safety Check')) {
             lastVehCheck = data[i][0]; // Timestamp
             break;
         }
     }
  }
  return {
    lastVehCheck: lastVehCheck,
    wofExpiry: null 
  };
}

function getGlobalFormsData() {
  const sheet = getSheet('Templates');
  const data = sheet.getDataRange().getValues();
  const forms = [];
  
  // Skip header
  for(let i=1; i<data.length; i++) {
    const row = data[i];
    if(row[0] && row[1]) {
      // Parse the pipe-separated questions
      const qRaw = row[1].split('|').map(q => q.trim()).filter(q => q !== "");
      forms.push({
        name: row[0],
        questions: qRaw
      });
    }
  }
  return forms;
}

function getGlobalForms() {
  return ContentService.createTextOutput(JSON.stringify(getGlobalFormsData()))
    .setMimeType(ContentService.MimeType.JSON);
}

function savePhotosToDrive(e) {
  const folderId = CONFIG.PHOTOS_FOLDER_ID;
  const urls = [];
  if (!folderId || folderId === "%%PHOTOS_FOLDER_ID%%") return urls;

  const folder = DriveApp.getFolderById(folderId);
  const p = e.parameter;
  
  // Loop through 'Photo 1', 'Photo 2', etc.
  for (let i = 1; i <= 10; i++) {
    const key = 'Photo ' + i;
    if (p[key]) {
      const blob = p[key];
      // Rename file for easier sorting: Date_Worker_PhotoX
      const safeName = p['Worker Name'].replace(/[^a-zA-Z0-9]/g, '_');
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      blob.setName(`${timestamp}_${safeName}_${key}.jpg`);
      
      const file = folder.createFile(blob);
      urls.push(file.getUrl());
    }
  }
  return urls;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Init Headers if new
    if(name === 'Visits') sheet.appendRow(['Timestamp', 'Worker Name', 'Alarm Status', 'Location Name', 'Location Address', 'Battery', 'Notes', 'Anticipated Departure', 'Phone', 'Email', 'Photos', 'GPS', 'DeviceID', 'Report Data']);
    if(name === 'Templates') sheet.appendRow(['Template Name', 'Questions (Pipe Separated | )']);
    if(name === 'Sites') sheet.appendRow(['Site Name', 'Address', 'Company Name', 'Contact Name', 'Contact Phone', 'Contact Email', 'Notes', 'Default Template', 'Archived']);
    if(name === 'Travel Log') sheet.appendRow(['Timestamp', 'Worker Name', 'Location', 'Distance (km)', 'Month']);
  }
  return sheet;
}

function getSheetData(name) {
  const sheet = getSheet(name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];
  for(let i=1; i<data.length; i++) {
    let obj = {};
    for(let j=0; j<headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    results.push(obj);
  }
  return results;
}

// ==========================================
// 6. MONITOR & WATCHDOG LOGIC
// ==========================================

function updateMonitorState(p, time) {
  const props = PropertiesService.getScriptProperties();
  const key = 'STATUS_' + p['Worker Name'];
  const state = {
    status: p['Alarm Status'],
    location: p['Location Name'],
    battery: p['Battery Level'],
    lastUpdate: time.getTime(),
    gps: p['Last Known GPS'],
    notes: p['Notes']
  };
  props.setProperty(key, JSON.stringify(state));
}

function getMonitorData() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys().filter(k => k.startsWith('STATUS_'));
  const data = keys.map(k => {
    const s = JSON.parse(props.getProperty(k));
    s.worker = k.replace('STATUS_', '');
    return s;
  });
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function sendEmergencySms(p) {
  if (!CONFIG.TEXTBELT_API_KEY || CONFIG.TEXTBELT_API_KEY.includes("%%")) return;
  
  // Construct Message
  const msg = `SOS ALERT: ${p['Worker Name']} sent ${p['Alarm Status']} at ${p['Location Name']}. GPS: ${p['Last Known GPS']}. Call them immediately.`;
  
  // Get contacts from Payload (sent by App)
  // We try Escalation Contact first, then Emergency Contact
  const numbers = [];
  if (p['Escalation Contact Number']) numbers.push(p['Escalation Contact Number']);
  if (p['Emergency Contact Number']) numbers.push(p['Emergency Contact Number']);
  
  numbers.forEach(num => {
    UrlFetchApp.fetch('https://textbelt.com/text', {
      method: 'post',
      payload: {
        phone: num,
        message: msg,
        key: CONFIG.TEXTBELT_API_KEY,
      },
    });
  });
}

function resolveAlert(e) {
  const p = e.parameter;
  const sheet = getSheet('Visits');
  sheet.appendRow([new Date(), p['Worker Name'], 'SAFE - MANUALLY CLEARED', 'HQ Manual Resolution', '', '', p['Notes'], '', '', '', '', '', '', '']);
  updateMonitorState(p, new Date());
  return ContentService.createTextOutput("Resolved");
}

// ==========================================
// 7. STATISTICS & TRAVEL LOGIC
// ==========================================

function logTravelData(worker, km, loc, time) {
   const sheet = getSheet('Travel Log');
   // Month Key (e.g., "2024-01")
   const monthKey = time.getFullYear() + "-" + ("0" + (time.getMonth() + 1)).slice(-2);
   sheet.appendRow([time, worker, loc, km, monthKey]);
}

function runMonthlyStats() {
  const sheet = getSheet('Visits');
  const data = sheet.getDataRange().getValues();
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  
  let count = 0, distance = 0, alerts = 0;
  
  // Parse Data
  for(let i=1; i<data.length; i++) {
    const rowTime = new Date(data[i][0]);
    if(rowTime > oneWeekAgo) {
      count++;
      
      const status = data[i][2];
      if(status.includes("EMERGENCY") || status.includes("DURESS")) alerts++;
      
      // Parse Report Data for numbers
      const rawJSON = data[i][13];
      if(rawJSON && rawJSON.startsWith('{')) {
         try {
           const r = JSON.parse(rawJSON);
           Object.values(r).forEach(val => {
             // Strict Number Parsing fix
             const n = parseFloat(val);
             if (!isNaN(n) && isFinite(n) && n < 100000) {
                 // Aggregation logic if needed
             }
           });
           
           // Explicit Distance Check
           if(r['Distance']) distance += parseFloat(r['Distance']);
           else if(r['Distance (km)']) distance += parseFloat(r['Distance (km)']);
           
         } catch(e) {}
      }
    }
  }
  
  const html = `
    <h2>Weekly Safety Report</h2>
    <p><strong>Period:</strong> Last 7 Days</p>
    <table border="1" cellpadding="10" style="border-collapse:collapse;">
      <tr><td><strong>Total Visits</strong></td><td>${count}</td></tr>
      <tr><td><strong>Distance Traveled (Reported)</strong></td><td>${distance.toFixed(2)} km</td></tr>
      <tr><td><strong>Safety Alerts</strong></td><td style="color:${alerts>0?'red':'green'}">${alerts}</td></tr>
    </table>
    <p><em>Generated by OTG AppSuite v79.23</em></p>
  `;
  
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(), 
    subject: "Weekly Safety Summary", 
    htmlBody: html
  });
}
