/**
 * OTG APPSUITE - MASTER BACKEND v77.1 (International)
 * Features: Zero Tolerance Mode, Privacy Redaction, Staged Escalation, i18n Terminology.
 */

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
  VEHICLE_TERM: "%%VEHICLE_TERM%%"
};

// Initialize Dynamic Properties
const sp = PropertiesService.getScriptProperties();
const tid = sp.getProperty('REPORT_TEMPLATE_ID');
if(tid) CONFIG.REPORT_TEMPLATE_ID = tid;

// ==========================================
// 1. GET HANDLER (Read/Sync Operations)
// ==========================================
function doGet(e) {
  try {
      if(!e || !e.parameter) return sendJSON({status:"error", message:"No Params"});
      const p = e.parameter;

      // A. Connection Test
      if(p.test) {
          if(p.key === CONFIG.MASTER_KEY) return sendJSON({status:"success", message:"OTG Online"});
          return sendJSON({status:"error", message:"Auth Fail"});
      }

      // B. Monitor Dashboard Polling (Protected by Master Key)
      if(p.key === CONFIG.MASTER_KEY && !p.action) {
          return sendJSON(getDashboardData());
      }

      // C. Worker App Sync (Protected by Shared Worker Key)
      if(p.action === 'sync') {
          // Worker-level auth: Can use Master Key OR Worker Key
          if(p.key !== CONFIG.MASTER_KEY && p.key !== CONFIG.WORKER_KEY) return sendJSON({status:"error", message:"ACCESS DENIED"});
          return sendJSON(getSyncData(p.worker, p.deviceId));
      }
      
      // D. Get Global Forms
      if(p.action === 'getGlobalForms') {
          return sendJSON(getGlobalForms());
      }

      return sendJSON({status:"error", message:"Invalid Request"});

  } catch(err) {
      return sendJSON({status:"error", message: err.toString()});
  }
}

// ==========================================
// 2. POST HANDLER (Write Operations)
// ==========================================
function doPost(e) {
  if(!e || !e.parameter) return sendJSON({status:"error", message:"No Data"});
  
  // Auth Check
  if(e.parameter.key !== CONFIG.MASTER_KEY && e.parameter.key !== CONFIG.WORKER_KEY) {
      return sendJSON({status:"error", message:"Auth Failed"});
  }

  const p = e.parameter;
  
  // A. Monitor Resolving an Alert
  if(p.action === 'resolve') {
      return resolveAlert(p);
  }

  // B. Worker Posting Data
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
      try {
          handleWorkerPost(p, e);
          return sendJSON({status:"success"});
      } catch(err) {
          return sendJSON({status:"error", message: err.toString()});
      } finally {
          lock.releaseLock();
      }
  } else {
      return sendJSON({status:"error", message:"Server Busy"});
  }
}

// ==========================================
// 3. CORE LOGIC
// ==========================================

function handleWorkerPost(p, e) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Visits');
    if(!sheet) {
        // Auto-create visits tab
        sheet = ss.insertSheet('Visits');
        sheet.appendRow(["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"]);
    }

    // Parse Photos
    let p1="", p2="", p3="", p4="", sig="";
    if(p['Photo 1']) p1 = saveImage(p['Photo 1']);
    if(p['Photo 2']) p2 = saveImage(p['Photo 2']);
    if(p['Photo 3']) p3 = saveImage(p['Photo 3']);
    if(p['Photo 4']) p4 = saveImage(p['Photo 4']);
    if(p['Signature']) sig = saveImage(p['Signature']); // Signature is just an image

    // Handle "Visit Report Data" - if it exists, it might trigger an AI summary
    let reportSummary = "";
    if(p['Visit Report Data']) {
       try {
           const reportObj = JSON.parse(p['Visit Report Data']);
           if(CONFIG.GEMINI_API_KEY && CONFIG.GEMINI_API_KEY.length > 10) {
               reportSummary = smartScribe(reportObj, p['Template Name'] || "Report", p['Notes']);
               if(reportSummary) p['Notes'] = (p['Notes'] + "\n[AI]: " + reportSummary).trim();
           }
       } catch(e) {}
    }

    // Append Row
    const ts = new Date();
    const dateStr = Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd");
    const row = [
        ts.toISOString(),
        dateStr,
        p['Worker Name'],
        p['Worker Phone Number'],
        p['Emergency Contact Name'],
        p['Emergency Contact Number'],
        p['Emergency Contact Email'],
        p['Escalation Contact Name'],
        p['Escalation Contact Number'],
        p['Escalation Contact Email'],
        p['Alarm Status'],
        p['Notes'],
        p['Location Name'],
        p['Location Address'],
        p['Last Known GPS'],
        p['Timestamp'], // GPS Timestamp from phone
        p['Battery Level'],
        p1,
        p['Distance'] || "",
        p['Visit Report Data'],
        p['Anticipated Departure Time'],
        sig, p2, p3, p4
    ];
    sheet.appendRow(row);

    // Update Staff Status (DeviceID binding)
    updateStaffStatus(p);

    // Check Immediate Escalation
    if(p['Alarm Status'].includes("EMERGENCY") || p['Alarm Status'].includes("PANIC") || p['Alarm Status'].includes("DURESS")) {
        triggerAlerts(p, "IMMEDIATE");
    }
}

function updateStaffStatus(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Staff');
    if(!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
        if(data[i][0] === p['Worker Name']) {
            sheet.getRange(i+1, 5).setValue(p['deviceId']); // Col E is DeviceID
            
            // Check for Vehicle Check Updates
            if(p['Template Name'] && p['Template Name'].includes('Vehicle')) {
                sheet.getRange(i+1, 6).setValue(new Date()); // LastVehCheck
                try {
                    const rData = JSON.parse(p['Visit Report Data']);
                    // Look for Expiry Key using Configured Term or standard keywords
                    const term = CONFIG.VEHICLE_TERM || "WOF";
                    const expKey = Object.keys(rData).find(k => k.includes('Expiry') || k.includes(term) || k.includes('Rego'));
                    if(expKey && rData[expKey]) {
                        sheet.getRange(i+1, 7).setValue(rData[expKey]); // WOFExpiry
                    }
                } catch(e){}
            }
            break;
        }
    }
}

function triggerAlerts(p, type) {
    const subject = `🚨 ${type}: ${p['Worker Name']} - ${p['Alarm Status']}`;
    const body = `SAFETY ALERT\n\nWorker: ${p['Worker Name']}\nStatus: ${p['Alarm Status']}\nLocation: ${p['Location Name']}\nNotes: ${p['Notes']}\nGPS: https://maps.google.com/?q=${p['Last Known GPS']}\nBattery: ${p['Battery Level']}`;
    
    // 1. Email
    const emails = [p['Emergency Contact Email'], p['Escalation Contact Email']].filter(e => e && e.includes('@'));
    if(emails.length > 0) {
        MailApp.sendEmail({to: emails.join(','), subject: subject, body: body});
    }

    // 2. SMS (Textbelt)
    if(CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5) {
        const numbers = [p['Emergency Contact Number'], p['Escalation Contact Number']].filter(n => n && n.length > 5);
        numbers.forEach(num => {
            UrlFetchApp.fetch('https://textbelt.com/text', {
                method: 'post',
                payload: { phone: num, message: subject + " " + p['Last Known GPS'], key: CONFIG.TEXTBELT_API_KEY }
            });
        });
    }
}

function resolveAlert(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const ts = new Date();
    
    // Append resolution row
    sheet.appendRow([
        ts.toISOString(),
        Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd"),
        p['Worker Name'],
        "", "", "", "", "", "", "", // Skip contacts
        p['Alarm Status'], // "SAFE - MANUALLY CLEARED"
        p['Notes'], // "[HQ RESOLVED]: ..."
        p['Location Name'],
        "", "", "", p['Battery Level'],
        "", "", "", "", "", "", "", ""
    ]);
    
    return sendJSON({status:"success"});
}

// ==========================================
// 4. WATCHDOG (Time-Driven Trigger)
// ==========================================
function checkOverdueVisits() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    if(!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    // Get latest status per worker
    const latest = {};
    for(let i=1; i<data.length; i++) {
        const row = data[i];
        const name = row[2]; // Worker Name
        if(!latest[name]) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
        else if(new Date(row[0]) > latest[name].time) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
    }

    Object.keys(latest).forEach(worker => {
        const entry = latest[worker].rowData;
        const status = entry[10]; // Alarm Status
        const dueTimeStr = entry[20]; // Anticipated Departure
        
        // Skip if safe
        if(status.includes("DEPARTED") || status.includes("SAFE") || status.includes("COMPLETED")) return;
        
        if(dueTimeStr) {
            const due = new Date(dueTimeStr);
            const diffMins = (now - due) / 60000;
            
            // Logic: Zero Tolerance check
            const isZeroTolerance = (entry[11] && entry[11].includes("[ZERO_TOLERANCE]"));
            const threshold = isZeroTolerance ? 0 : CONFIG.ESCALATION_MINUTES;

            if(diffMins > threshold && !status.includes("EMERGENCY")) {
                // ESCALATE
                const newStatus = isZeroTolerance ? "EMERGENCY - ZERO TOLERANCE OVERDUE" : "EMERGENCY - OVERDUE";
                
                // Add new row to log the escalation
                const newRow = [...entry];
                newRow[0] = new Date().toISOString(); // New timestamp
                newRow[10] = newStatus; // New Status
                newRow[11] = entry[11] + " [SYSTEM AUTO-ESCALATION]"; // Append note
                
                sheet.appendRow(newRow);
                
                // Fire Alerts
                triggerAlerts({
                    'Worker Name': worker,
                    'Alarm Status': newStatus,
                    'Location Name': entry[12],
                    'Notes': "Worker is overdue and has not checked out.",
                    'Last Known GPS': entry[14],
                    'Battery Level': entry[16],
                    'Emergency Contact Email': entry[6],
                    'Escalation Contact Email': entry[9],
                    'Emergency Contact Number': entry[5],
                    'Escalation Contact Number': entry[8]
                }, "OVERDUE");
            }
        }
    });
}

// ==========================================
// 5. DATA FETCHERS
// ==========================================
function getDashboardData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const staffSheet = ss.getSheetByName('Staff');
    if(!sheet) return {workers: []};
    
    // Get last 500 rows for performance
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(1, lastRow - 500);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 25).getValues();
    const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
    
    // Map to JSON
    const workers = data.map(r => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = r[i]);
        return obj;
    });

    // Inject Staff Metadata (WOF Expiry)
    if(staffSheet) {
        const sData = staffSheet.getDataRange().getValues();
        workers.forEach(w => {
            for(let i=1; i<sData.length; i++) {
                if(sData[i][0] === w['Worker Name']) {
                    w['WOFExpiry'] = sData[i][6]; // Col G
                }
            }
        });
    }

    return {workers: workers, escalation_limit: CONFIG.ESCALATION_MINUTES};
}

function getSyncData(workerName, deviceId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get Sites
    const siteSheet = ss.getSheetByName('Sites');
    const sites = [];
    if(siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) {
            // Check assignment (Col A)
            const assigned = sData[i][0];
            if(assigned === "ALL" || assigned.includes(workerName)) {
                sites.push({
                    template: sData[i][1],
                    company: sData[i][2],
                    siteName: sData[i][3],
                    address: sData[i][4],
                    contactName: sData[i][5],
                    contactPhone: sData[i][6],
                    contactEmail: sData[i][7],
                    notes: sData[i][8]
                });
            }
        }
    }

    // 2. Get Forms (Templates)
    const tSheet = ss.getSheetByName('Templates');
    const forms = [];
    const cachedTemplates = {};
    if(tSheet) {
        const tData = tSheet.getDataRange().getValues();
        for(let i=1; i<tData.length; i++) {
            const row = tData[i];
            if(row[2] === "ALL" || row[2].includes(workerName)) {
                const questions = [];
                for(let q=4; q<9; q++) { if(row[q]) questions.push(row[q]); }
                forms.push({name: row[1], type: row[0], questions: questions});
                cachedTemplates[row[1]] = questions;
            }
        }
    }
    
    // 3. Get Metadata (WOF Status)
    const meta = {};
    const stSheet = ss.getSheetByName('Staff');
    if(stSheet) {
        const stData = stSheet.getDataRange().getValues();
        for(let i=1; i<stData.length; i++) {
            if(stData[i][0] === workerName) {
                // Security Check: Bind DeviceID if empty
                if(!stData[i][4]) stSheet.getRange(i+1, 5).setValue(deviceId);
                else if(stData[i][4] !== deviceId) return {status:"error", message:"DEVICE MISMATCH. Contact Admin."};
                
                meta.lastVehCheck = stData[i][5];
                meta.wofExpiry = stData[i][6];
            }
        }
    }

    return {sites, forms, cachedTemplates, meta};
}

function getGlobalForms() {
    // Public endpoint for "Quick Notes" without full sync
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName('Templates');
    if(!tSheet) return [];
    const tData = tSheet.getDataRange().getValues();
    const forms = [];
    for(let i=1; i<tData.length; i++) {
        const row = tData[i];
        if(row[2] === "ALL") {
            const questions = [];
            for(let q=4; q<9; q++) { if(row[q]) questions.push(row[q]); }
            forms.push({name: row[1], questions: questions});
        }
    }
    return forms;
}

// ==========================================
// 6. UTILITIES
// ==========================================

function saveImage(b64) {
    if(!b64 || !CONFIG.PHOTOS_FOLDER_ID) return "";
    try {
        const data = Utilities.base64Decode(b64.split(',')[1]);
        const blob = Utilities.newBlob(data, 'image/jpeg', 'photo_' + Date.now() + '.jpg');
        const folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return "Error saving photo"; }
}

function smartScribe(data, type, notes) {
    if(!CONFIG.GEMINI_API_KEY) return "";
    
    // 1. Data Preparation & Redaction (v77.0 + v77.1)
    let safeNotes = notes || "";
    let safeData = JSON.stringify(data || {});
    
    if(CONFIG.ENABLE_REDACTION) {
        // Redact Emails
        const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
        safeNotes = safeNotes.replace(emailRegex, "[EMAIL_REDACTED]");
        safeData = safeData.replace(emailRegex, "[EMAIL_REDACTED]");
        
        // Redact Phones (Global approximate matches)
        const phoneRegex = /\b(\+?\d{1,3}[- ]?)?\(?\d{3}\)?[- ]?\d{3}[- ]?\d{4}\b/g;
        safeNotes = safeNotes.replace(phoneRegex, "[PHONE_REDACTED]");
        safeData = safeData.replace(phoneRegex, "[PHONE_REDACTED]");
    }

    // 2. Terminology Injection (v77.1)
    const term = CONFIG.VEHICLE_TERM || "Vehicle Inspection";
    
    // 3. Prompt Construction
    const prompt = `Analyze this ${type} report using terminology relevant to "${term}".
    User Notes: ${safeNotes}
    Form Data: ${safeData}
    
    Output a single sentence summary of the key issue or confirmation of safety. If it is a Vehicle Check, explicitly mention the ${term} status.`;

    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) };
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());
        return json.candidates[0].content.parts[0].text.trim();
    } catch (e) { return ""; }
}

function sendJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function archiveOldData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const archive = ss.getSheetByName('Archive') || ss.insertSheet('Archive');
    
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return;

    const today = new Date();
    const cutoff = new Date(today.setDate(today.getDate() - CONFIG.ARCHIVE_DAYS));
    
    const keep = [data[0]]; // Headers
    const move = [];
    
    for(let i=1; i<data.length; i++) {
        // Archive based on Date and 'Closed' status
        if(new Date(data[i][0]) < cutoff && (data[i][10].includes('DEPARTED') || data[i][10].includes('SAFE') || data[i][10].includes('COMPLETED'))) {
            move.push(data[i]);
        } else {
            keep.push(data[i]);
        }
    }
    
    if(move.length > 0) {
        archive.getRange(archive.getLastRow()+1, 1, move.length, move[0].length).setValues(move);
        sheet.clearContents();
        sheet.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
    }
}
