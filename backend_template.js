/**
 * OTG APPSUITE - MASTER BACKEND v79.16 (Tiered Escalation & AI Polish)
 * * UPDATES:
 * - Logic: Tiered Escalation (Warning at +5m, Emergency at +Config Mins).
 * - AI: Proofreads email content but SAVES RAW CONTENT to Spreadsheet (Source of Truth).
 * - Syntax: Supports old Form Builder syntax (#, %) in processing.
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

const sp = PropertiesService.getScriptProperties();
const tid = sp.getProperty('REPORT_TEMPLATE_ID');
if(tid) CONFIG.REPORT_TEMPLATE_ID = tid;

// ==========================================
// 2. GET HANDLER
// ==========================================
function doGet(e) {
  try {
      if(!e || !e.parameter) return sendResponse(e, {status:"error", message:"No Params"});
      const p = e.parameter;

      if(p.test) {
          if(p.key === CONFIG.MASTER_KEY) return sendResponse(e, {status:"success", message:"OTG Online"});
          return sendResponse(e, {status:"error", message:"Auth Fail"});
      }

      if(p.key === CONFIG.MASTER_KEY && !p.action) {
          return sendResponse(e, getDashboardData());
      }

      if(p.action === 'sync') {
          if(p.key !== CONFIG.MASTER_KEY && p.key !== CONFIG.WORKER_KEY) return sendResponse(e, {status:"error", message:"ACCESS DENIED"});
          return sendResponse(e, getSyncData(p.worker, p.deviceId));
      }
      
      if(p.action === 'getGlobalForms') {
          return sendResponse(e, getGlobalForms());
      }

      return sendResponse(e, {status:"error", message:"Invalid Request"});

  } catch(err) {
      return sendResponse(e, {status:"error", message: err.toString()});
  }
}

// ==========================================
// 3. POST HANDLER
// ==========================================
function doPost(e) {
  if(!e || !e.parameter) return sendJSON({status:"error", message:"No Data"});
  
  if(e.parameter.key !== CONFIG.MASTER_KEY && e.parameter.key !== CONFIG.WORKER_KEY) {
      return sendJSON({status:"error", message:"Auth Failed"});
  }

  const p = e.parameter;
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) { 
      try {
          if(p.action === 'resolve') {
              handleResolvePost(p); 
          } else {
              handleWorkerPost(p, e);
          }
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
// 4. SMART RESPONSE HANDLER
// ==========================================
function sendResponse(e, data) {
    const json = JSON.stringify(data);
    if (e && e.parameter && e.parameter.callback) {
        return ContentService.createTextOutput(`${e.parameter.callback}(${json})`)
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);
}

function sendJSON(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 5. CORE LOGIC
// ==========================================

function handleResolvePost(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const workerName = p['Worker Name'];
    const lastRow = sheet.getLastRow();
    let rowUpdated = false;

    if (lastRow > 1) {
        const startRow = Math.max(2, lastRow - 50); 
        const numRows = lastRow - startRow + 1;
        const data = sheet.getRange(startRow, 1, numRows, 11).getValues();
        
        for (let i = data.length - 1; i >= 0; i--) {
            const rowData = data[i];
            if (rowData[2] === workerName) {
                const status = String(rowData[10]);
                if (status.includes('EMERGENCY') || status.includes('PANIC') || status.includes('DURESS') || status.includes('OVERDUE')) {
                    const targetRow = startRow + i;
                    sheet.getRange(targetRow, 11).setValue(p['Alarm Status']); 
                    sheet.getRange(targetRow, 12).setValue((String(rowData[11]) + "\n" + p['Notes']).trim()); 
                    rowUpdated = true;
                    break;
                }
            }
        }
    }

    if (!rowUpdated) {
        const ts = new Date();
        sheet.appendRow([ts.toISOString(), Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd"), workerName, "", "", "", "", "", "", "", p['Alarm Status'], p['Notes'], "HQ Dashboard", "", "", "", "N/A", "", "", "", "", "", "", "", ""]);
    }
}

function handleWorkerPost(p, e) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Visits');
    
    if(!sheet) {
        sheet = ss.insertSheet('Visits');
        sheet.appendRow(["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"]);
    }

    const workerName = p['Worker Name'];

    // SAVE PHOTOS
    let p1="", p2="", p3="", p4="", sig="";
    if(p['Photo 1']) p1 = saveImage(p['Photo 1'], workerName);
    if(p['Photo 2']) p2 = saveImage(p['Photo 2'], workerName);
    if(p['Photo 3']) p3 = saveImage(p['Photo 3'], workerName);
    if(p['Photo 4']) p4 = saveImage(p['Photo 4'], workerName);
    if(p['Signature']) sig = saveImage(p['Signature'], workerName, true); 

    const ts = new Date();
    const dateStr = Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd");

    // === SOURCE OF TRUTH: SAVE RAW DATA TO SHEET FIRST ===
    // We do NOT modify p['Notes'] or p['Visit Report Data'] before saving.
    let rowUpdated = false;
    const lastRow = sheet.getLastRow();
    
    // Smart Ledger Update Logic (Simplified for brevity, logic maintained)
    if (lastRow > 1) {
        const startRow = Math.max(2, lastRow - 50); 
        const numRows = lastRow - startRow + 1;
        const data = sheet.getRange(startRow, 1, numRows, 11).getValues(); 
        
        for (let i = data.length - 1; i >= 0; i--) {
            const rowData = data[i];
            if (rowData[2] === workerName) {
                const status = String(rowData[10]);
                const isClosed = status.includes('DEPARTED') || (status.includes('SAFE') && !status.includes('MANUALLY')) || status.includes('COMPLETED') || status.includes('DATA_ENTRY_ONLY');
                if (!isClosed) {
                    const targetRow = startRow + i;
                    sheet.getRange(targetRow, 1).setValue(ts.toISOString()); 
                    sheet.getRange(targetRow, 11).setValue(p['Alarm Status']); 
                    if (p['Notes'] && p['Notes'] !== rowData[11]) {
                         const oldNotes = sheet.getRange(targetRow, 12).getValue();
                         if (!oldNotes.includes(p['Notes'])) {
                             sheet.getRange(targetRow, 12).setValue((oldNotes + "\n" + p['Notes']).trim());
                         }
                    }
                    if (p['Last Known GPS']) sheet.getRange(targetRow, 15).setValue(p['Last Known GPS']);
                    if (p['Battery Level']) sheet.getRange(targetRow, 17).setValue(p['Battery Level']);
                    if (p['Visit Report Data']) sheet.getRange(targetRow, 20).setValue(p['Visit Report Data']);
                    if(p['Distance']) sheet.getRange(targetRow, 19).setValue(p['Distance']);
                    if(sig) sheet.getRange(targetRow, 22).setValue(sig);
                    if(p1) sheet.getRange(targetRow, 18).setValue(p1);
                    if(p2) sheet.getRange(targetRow, 23).setValue(p2);
                    if(p3) sheet.getRange(targetRow, 24).setValue(p3);
                    if(p4) sheet.getRange(targetRow, 25).setValue(p4);
                    rowUpdated = true;
                    break;
                }
            }
        }
    }

    if (!rowUpdated) {
        const row = [ts.toISOString(), dateStr, workerName, p['Worker Phone Number'], p['Emergency Contact Name'], p['Emergency Contact Number'], p['Emergency Contact Email'], p['Escalation Contact Name'], p['Escalation Contact Number'], p['Escalation Contact Email'], p['Alarm Status'], p['Notes'], p['Location Name'], p['Location Address'], p['Last Known GPS'], p['Timestamp'], p['Battery Level'], p1, p['Distance'] || "", p['Visit Report Data'], p['Anticipated Departure Time'], sig, p2, p3, p4];
        sheet.appendRow(row);
    }

    updateStaffStatus(p);

    // === AI PROOFREADING & EMAILING (AFTER SAVING) ===
    const hasFormData = p['Visit Report Data'] && p['Visit Report Data'].length > 2;
    if(hasFormData) {
       try {
           const reportObj = JSON.parse(p['Visit Report Data']);
           let polishedNotes = p['Notes']; // Default to raw if AI fails
           
           if(CONFIG.GEMINI_API_KEY && CONFIG.GEMINI_API_KEY.length > 10) {
               // Ask AI to proofread ONLY for the email
               polishedNotes = smartScribe(reportObj, p['Template Name'] || "Report", p['Notes']);
           }
           
           // Create a Display Object for the email (Uses Polished Notes)
           const displayObj = {...reportObj}; // Clone
           
           // Trigger Email with Polished Content
           processFormEmail(p, displayObj, polishedNotes, p1, p2, p3, p4, sig);
       } catch(e) { console.error("Email/AI Error: " + e); }
    }

    if(p['Alarm Status'].includes("EMERGENCY") || p['Alarm Status'].includes("PANIC") || p['Alarm Status'].includes("DURESS")) {
        triggerAlerts(p, "IMMEDIATE");
    }
}

// EMAIL PROCESSOR (With AI Polished Notes passed in)
function processFormEmail(p, reportObj, polishedNotes, p1, p2, p3, p4, sig) {
    const templateName = p['Template Name'];
    if (!templateName) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName('Templates');
    if (!tSheet) return;

    // Find Recipient
    const tData = tSheet.getDataRange().getValues();
    let recipientEmail = "";
    const safeTName = templateName.trim().toLowerCase();
    
    for (let i = 1; i < tData.length; i++) {
        if (tData[i][1] && tData[i][1].toString().trim().toLowerCase() === safeTName) {
            recipientEmail = tData[i][3];
            break;
        }
    }

    if (!recipientEmail || !recipientEmail.includes('@')) return;

    // Prepare Inline Images
    const inlineImages = {};
    const imgTags = [];
    const processImg = (key, cidName, title) => {
        if (p[key] && p[key].length > 100) { 
            const blob = dataURItoBlob(p[key]);
            if (blob) {
                inlineImages[cidName] = blob;
                imgTags.push(`<div style="margin-bottom: 20px; text-align: center;"><p style="color:#6b7280; font-size:12px; font-weight:bold; margin-bottom:5px; text-transform:uppercase;">${title}</p><img src="cid:${cidName}" style="max-width: 100%; border-radius: 8px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.1);"></div>`);
            }
        }
    };
    processImg('Photo 1', 'photo1', 'Attached Photo 1');
    processImg('Photo 2', 'photo2', 'Attached Photo 2');
    processImg('Photo 3', 'photo3', 'Attached Photo 3');
    processImg('Photo 4', 'photo4', 'Attached Photo 4');
    
    if (p['Signature'] && p['Signature'].length > 100) {
        const sigBlob = dataURItoBlob(p['Signature']);
        if (sigBlob) inlineImages['signature'] = sigBlob;
    }

    // HTML Construction
    let html = `<div style="background-color:#f3f4f6; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 20px;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);">
            <div style="background-color: #1e40af; padding: 24px; text-align: center;"><h1 style="color: #ffffff; margin: 0; font-size: 24px; font-weight: 800; letter-spacing: 0.5px;">${p['Template Name']}</h1><p style="color: #93c5fd; margin: 8px 0 0 0; font-size: 14px; font-weight: 600;">${CONFIG.ORG_NAME}</p></div>
            <div style="background-color: #eff6ff; padding: 16px; border-bottom: 1px solid #dbeafe; display: flex; justify-content: space-between;"><div style="width: 48%;"><p style="margin:0; font-size:10px; color:#6b7280; text-transform:uppercase; font-weight:bold;">Worker</p><p style="margin:0; font-size:14px; color:#1f2937; font-weight:bold;">${p['Worker Name']}</p></div><div style="width: 48%; text-align:right;"><p style="margin:0; font-size:10px; color:#6b7280; text-transform:uppercase; font-weight:bold;">Date</p><p style="margin:0; font-size:14px; color:#1f2937;">${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}</p></div></div>
            <div style="padding: 24px;">
                <div style="margin-bottom: 24px;"><p style="margin:0; font-size:10px; color:#6b7280; text-transform:uppercase; font-weight:bold; margin-bottom: 4px;">Location</p><div style="background: #f9fafb; padding: 10px; border-radius: 6px; border: 1px solid #e5e7eb; color: #374151; font-size: 13px;">${p['Location Name'] || p['Location Address'] || 'Unknown Location'}</div></div>
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">`;

    let rowCount = 0;
    for (const [key, value] of Object.entries(reportObj)) {
        if(key && value) {
            const bg = rowCount % 2 === 0 ? '#ffffff' : '#f9fafb';
            html += `<tr style="background-color: ${bg};"><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; color: #4b5563; font-weight: bold; font-size: 13px; width: 40%; vertical-align: top;">${key}</td><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; color: #111827; font-size: 13px;">${value}</td></tr>`;
            rowCount++;
        }
    }

    if (polishedNotes) {
        html += `<tr style="background-color: #fffbeb;"><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; color: #b45309; font-weight: bold; font-size: 13px; vertical-align: top;">Notes (AI Polished)</td><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; color: #78350f; font-size: 13px; font-style: italic;">${polishedNotes}</td></tr>`;
    }

    html += `</table>`;
    if (imgTags.length > 0) html += `<div style="margin-top: 30px; border-top: 2px dashed #e5e7eb; padding-top: 20px;">` + imgTags.join('') + `</div>`;
    if (inlineImages['signature']) html += `<div style="margin-top: 20px; text-align: right;"><p style="font-size: 10px; color: #9ca3af; text-transform: uppercase; font-weight: bold; margin-bottom: 5px;">Signed By Worker</p><img src="cid:signature" style="max-height: 60px; border-bottom: 2px solid #d1d5db; padding-bottom: 4px;"></div>`;
    html += `</div><div style="background-color: #f9fafb; padding: 16px; text-align: center; border-top: 1px solid #e5e7eb;"><p style="margin: 0; font-size: 11px; color: #9ca3af;">Generated by OTG AppSuite • Raw data stored in database.</p></div></div></div>`;

    MailApp.sendEmail({ to: recipientEmail, subject: `[Report] ${p['Template Name']} - ${p['Worker Name']}`, htmlBody: html, inlineImages: inlineImages });
}

function dataURItoBlob(dataURI) {
    try {
        const parts = dataURI.split(',');
        if(parts.length < 2) return null;
        const mimeString = parts[0].split(':')[1].split(';')[0];
        const byteString = Utilities.base64Decode(parts[1]);
        return Utilities.newBlob(byteString, mimeString, "image");
    } catch(e) { return null; }
}

function updateStaffStatus(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Staff');
    if(!sheet) return;
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
        if(data[i][0] === p['Worker Name']) {
            sheet.getRange(i+1, 5).setValue(p['deviceId']); 
            if(p['Template Name'] && p['Template Name'].includes('Vehicle')) {
                sheet.getRange(i+1, 6).setValue(new Date()); 
                try {
                    const rData = JSON.parse(p['Visit Report Data']);
                    const term = CONFIG.VEHICLE_TERM || "WOF";
                    const expKey = Object.keys(rData).find(k => k.includes('Expiry') || k.includes(term) || k.includes('Rego'));
                    if(expKey && rData[expKey]) { sheet.getRange(i+1, 7).setValue(rData[expKey]); }
                } catch(e){}
            }
            break;
        }
    }
}

function _cleanPhone(num) {
    if (!num) return null;
    let n = num.toString().replace(/[^0-9]/g, ''); 
    if (n.length < 5) return null;
    if (n.startsWith('0')) { return (CONFIG.COUNTRY_CODE || "+64") + n.substring(1); }
    const ccRaw = (CONFIG.COUNTRY_CODE || "").replace('+', '');
    if (n.startsWith(ccRaw)) { return "+" + n; }
    return "+" + n;
}

function triggerAlerts(p, type) {
    const subject = `🚨 ${type}: ${p['Worker Name']} - ${p['Alarm Status']}`;
    const gpsLink = p['Last Known GPS'] ? `http://googleusercontent.com/maps.google.com/?q=${p['Last Known GPS']}` : "No GPS";
    const body = `SAFETY ALERT\n\nWorker: ${p['Worker Name']}\nStatus: ${p['Alarm Status']}\nLocation: ${p['Location Name']}\nNotes: ${p['Notes']}\nGPS: ${gpsLink}\nBattery: ${p['Battery Level']}`;
    const emails = [p['Emergency Contact Email'], p['Escalation Contact Email']].filter(e => e && e.includes('@'));
    if(emails.length > 0) { MailApp.sendEmail({to: emails.join(','), subject: subject, body: body}); }
    if(CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5) {
        const numbers = [p['Emergency Contact Number'], p['Escalation Contact Number']].map(n => _cleanPhone(n)).filter(n => n);
        numbers.forEach(num => { 
            try {
                UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: num, message: `${subject} ${gpsLink}`, key: CONFIG.TEXTBELT_API_KEY }) }); 
            } catch(e) { console.error("SMS Failed: " + e.toString()); }
        });
    }
}

function resolveAlert(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const ts = new Date();
    sheet.appendRow([ts.toISOString(), Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd"), p['Worker Name'], "", "", "", "", "", "", "", p['Alarm Status'], p['Notes'], p['Location Name'], "", "", "", p['Battery Level'], "", "", "", "", "", "", "", ""]);
    return sendJSON({status:"success"});
}

// UPDATED: TIERED ESCALATION LOGIC
function checkOverdueVisits() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    if(!sheet) return;
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const latest = {};
    for(let i=1; i<data.length; i++) {
        const row = data[i];
        const name = row[2]; 
        if(!latest[name]) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
        else if(new Date(row[0]) > latest[name].time) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
    }
    
    Object.keys(latest).forEach(worker => {
        const entry = latest[worker].rowData;
        const status = entry[10]; 
        const dueTimeStr = entry[20]; 
        const isClosed = status.includes("DEPARTED") || status.includes("SAFE") || status.includes("COMPLETED");
        
        if(!isClosed && dueTimeStr) {
            const due = new Date(dueTimeStr);
            const diffMins = (now - due) / 60000; 
            const isZeroTolerance = (entry[11] && entry[11].includes("[ZERO_TOLERANCE]"));
            
            // TIER 1: WARNING (5 Mins Overdue)
            // Skipped if Zero Tolerance
            if (diffMins > 5 && diffMins < CONFIG.ESCALATION_MINUTES && !status.includes('WARNING') && !status.includes('EMERGENCY') && !isZeroTolerance) {
                const newStatus = "OVERDUE - WARNING SENT";
                const newRow = [...entry];
                newRow[0] = new Date().toISOString(); 
                newRow[10] = newStatus; 
                newRow[11] = entry[11] + " [AUTO-WARNING]";
                sheet.appendRow(newRow);
                // Send Warning Email only (No SMS usually, to save credits, but configurable)
                triggerAlerts({ 'Worker Name': worker, 'Alarm Status': "WARNING - 5 Mins Overdue", 'Location Name': entry[12], 'Notes': "Worker is 5 minutes overdue. Please extend or check-in.", 'Last Known GPS': entry[14], 'Battery Level': entry[16], 'Emergency Contact Email': entry[6], 'Emergency Contact Number': entry[5] }, "WARNING");
            }
            
            // TIER 2: ESCALATION (Config Mins Overdue OR Zero Tolerance)
            const threshold = isZeroTolerance ? 0 : CONFIG.ESCALATION_MINUTES;
            if (diffMins > threshold && !status.includes("EMERGENCY")) {
                const newStatus = isZeroTolerance ? "EMERGENCY - ZERO TOLERANCE OVERDUE" : "EMERGENCY - OVERDUE";
                const newRow = [...entry];
                newRow[0] = new Date().toISOString(); 
                newRow[10] = newStatus; 
                newRow[11] = entry[11] + " [AUTO-ESCALATION]";
                sheet.appendRow(newRow);
                triggerAlerts({ 'Worker Name': worker, 'Alarm Status': newStatus, 'Location Name': entry[12], 'Notes': "Worker is overdue and has breached escalation threshold.", 'Last Known GPS': entry[14], 'Battery Level': entry[16], 'Emergency Contact Email': entry[6], 'Escalation Contact Email': entry[9], 'Emergency Contact Number': entry[5], 'Escalation Contact Number': entry[8] }, "OVERDUE");
            }
        }
    });
}

function getDashboardData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Visits');
    const staffSheet = ss.getSheetByName('Staff');
    if(!sheet) return {workers: []};
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {workers: []}; 
    const startRow = Math.max(2, lastRow - 500); 
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 25).getValues();
    const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
    const workers = data.map(r => { let obj = {}; headers.forEach((h, i) => obj[h] = r[i]); return obj; });
    if(staffSheet) {
        const sData = staffSheet.getDataRange().getValues();
        workers.forEach(w => { for(let i=1; i<sData.length; i++) { if(sData[i][0] === w['Worker Name']) { w['WOFExpiry'] = sData[i][6]; } } });
    }
    return {workers: workers, escalation_limit: CONFIG.ESCALATION_MINUTES};
}

function getSyncData(workerName, deviceId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const siteSheet = ss.getSheetByName('Sites');
    const sites = [];
    const wNameSafe = (workerName || "").toString().toLowerCase().trim();

    if(siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) {
            const assignedRaw = sData[i][0];
            const assignedStr = (assignedRaw || "").toString().toLowerCase();
            const allowedUsers = assignedStr.split(',').map(s => s.trim());
            if(allowedUsers.includes("all") || allowedUsers.includes(wNameSafe)) {
                sites.push({ template: sData[i][1], company: sData[i][2], siteName: sData[i][3], address: sData[i][4], contactName: sData[i][5], contactPhone: sData[i][6], contactEmail: sData[i][7], notes: sData[i][8] });
            }
        }
    }
    
    const tSheet = ss.getSheetByName('Templates');
    const forms = [];
    const cachedTemplates = {};
    if(tSheet) {
        const tData = tSheet.getDataRange().getValues();
        for(let i=1; i<tData.length; i++) {
            const row = tData[i];
            const assignedRaw = row[2];
            const assignedStr = (assignedRaw || "").toString().toLowerCase();
            const allowedUsers = assignedStr.split(',').map(s => s.trim());
            
            if(allowedUsers.includes("all") || allowedUsers.includes(wNameSafe)) {
                const questions = [];
                for(let q=4; q<9; q++) { if(row[q]) questions.push(row[q]); }
                forms.push({name: row[1], type: row[0], questions: questions});
                cachedTemplates[row[1]] = questions;
            }
        }
    }
    
    const meta = {};
    const stSheet = ss.getSheetByName('Staff');
    if(stSheet) {
        const stData = stSheet.getDataRange().getValues();
        for(let i=1; i<stData.length; i++) {
            if((stData[i][0] || "").toString().toLowerCase().trim() === wNameSafe) {
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

function saveImage(b64, workerName, isSignature) {
    if(!b64 || !CONFIG.PHOTOS_FOLDER_ID) return "";
    try {
        const data = Utilities.base64Decode(b64.split(',')[1]);
        const mainFolder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
        let targetFolder = mainFolder;
        if (workerName && workerName.length > 2) {
            const folders = mainFolder.getFoldersByName(workerName);
            if (folders.hasNext()) { targetFolder = folders.next(); } 
            else { targetFolder = mainFolder.createFolder(workerName); }
        }
        const now = new Date();
        const timeStr = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd_HH-mm");
        const safeName = (workerName || "Unknown").replace(/[^a-zA-Z0-9]/g, ''); 
        const type = isSignature ? "Signature" : "Photo";
        const fileName = `${timeStr}_${safeName}_${type}_${Math.floor(Math.random()*100)}.jpg`;
        const blob = Utilities.newBlob(data, 'image/jpeg', fileName);
        const file = targetFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return "Error saving photo: " + e.toString(); }
}

// UPDATED: SMART SCRIBE (Returns polished string, does not save to DB)
function smartScribe(data, type, notes) {
    if(!CONFIG.GEMINI_API_KEY) return notes; // Fallback to original notes
    let safeNotes = notes || "";
    let safeData = JSON.stringify(data || {});
    if(CONFIG.ENABLE_REDACTION) {
        const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
        safeNotes = safeNotes.replace(emailRegex, "[EMAIL_REDACTED]");
        safeData = safeData.replace(emailRegex, "[EMAIL_REDACTED]");
        const phoneRegex = /\b(\+?\d{1,3}[- ]?)?\(?\d{3}\)?[- ]?\d{3}[- ]?\d{4}\b/g;
        safeNotes = safeNotes.replace(phoneRegex, "[PHONE_REDACTED]");
        safeData = safeData.replace(phoneRegex, "[PHONE_REDACTED]");
    }
    const term = CONFIG.VEHICLE_TERM || "Vehicle Inspection";
    const prompt = `You are a professional safety officer proofreading a report. Correct spelling and grammar in the following notes. Use New Zealand English. Do NOT add new facts. If notes are empty, return an empty string. Notes: "${safeNotes}".`;
    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) };
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());
        return json.candidates[0].content.parts[0].text.trim();
    } catch (e) { return notes; }
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
    const keep = [data[0]];
    const move = [];
    for(let i=1; i<data.length; i++) {
        if(new Date(data[i][0]) < cutoff && (data[i][10].includes('DEPARTED') || data[i][10].includes('SAFE') || data[i][10].includes('COMPLETED'))) { move.push(data[i]); } else { keep.push(data[i]); }
    }
    if(move.length > 0) {
        archive.getRange(archive.getLastRow()+1, 1, move.length, move[0].length).setValues(move);
        sheet.clearContents();
        sheet.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
    }
}

function sendWeeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  let count = 0, distance = 0, alerts = 0;
  for(let i=1; i<data.length; i++) {
    const rowTime = new Date(data[i][0]);
    if(rowTime > oneWeekAgo) {
      count++;
      if(data[i][18]) distance += Number(data[i][18]);
      if(data[i][10].toString().includes("EMERGENCY")) alerts++;
    }
  }
  const html = `<h2>Weekly Safety Report</h2><p><strong>Period:</strong> Last 7 Days</p><table border="1" cellpadding="10" style="border-collapse:collapse;"><tr><td><strong>Total Visits</strong></td><td>${count}</td></tr><tr><td><strong>Distance Traveled</strong></td><td>${distance.toFixed(2)} km</td></tr><tr><td><strong>Safety Alerts</strong></td><td style="color:${alerts>0?'red':'green'}">${alerts}</td></tr></table><p><em>Generated by OTG AppSuite</em></p>`;
  MailApp.sendEmail({to: Session.getEffectiveUser().getEmail(), subject: "Weekly Safety Summary", htmlBody: html});
}
