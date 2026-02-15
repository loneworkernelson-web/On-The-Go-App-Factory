/**
 * OTG APPSUITE - MASTER BACKEND v79.35
 * FIXED: SOS Map URLs, SMS Payloads, and GAS Environment Stability
 */

const CONFIG = {
  VERSION: "v79.35", // New diagnostic property
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
  COUNTRY_CODE: "%%COUNTRY_PREFIX%%", 
  LOCALE: "%%LOCALE%%"
};

const sp = PropertiesService.getScriptProperties();
const tid = sp.getProperty('REPORT_TEMPLATE_ID');
if(tid) CONFIG.REPORT_TEMPLATE_ID = tid;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ OTG Admin')
      .addItem('1. Setup Client Reporting', 'setupClientReporting')
      .addItem('2. Run Monthly Stats', 'runMonthlyStats')
      .addItem('3. Run Travel Report', 'generateWorkerTravelReport')
      .addSeparator()
      .addItem('Force Sync Forms', 'getGlobalForms')
      .addToUi();
}

// ==========================================
// 3. WEB HANDLERS (GET/POST)
// ==========================================
function doGet(e) {
  try {
      if(!e || !e.parameter) return sendResponse(e, {status:"error", message:"No Params"});
      const p = e.parameter;
   // NEW: Version Ping for System Info
      if (p.action === 'ping') {
          return sendResponse(e, { status: "success", version: CONFIG.VERSION });
      }
      if (p.action === 'getDistance' && p.start && p.end) {
          const dist = getRouteDistance(p.start, p.end);
          return sendResponse(e, { status: "success", km: dist });
      } 
      if(p.test) return (p.key === CONFIG.MASTER_KEY) ? sendResponse(e, {status:"success"}) : sendResponse(e, {status:"error"});
      if(p.key === CONFIG.MASTER_KEY && !p.action) return sendResponse(e, getDashboardData());
      if(p.action === 'sync') return (p.key === CONFIG.MASTER_KEY || p.key === CONFIG.WORKER_KEY) ? sendResponse(e, getSyncData(p.worker, p.deviceId)) : sendResponse(e, {status:"error"});
      if(p.action === 'getGlobalForms') return sendResponse(e, getGlobalForms());
      return sendResponse(e, {status:"error"});
  } catch(err) { return sendResponse(e, {status:"error", message: err.toString()}); }
}

/**
 * PATCHED: Master Entry Point
 * Integrated routing for Site Procedures and Notice Acknowledgments.
 */
function doPost(e) {
  if(!e || !e.parameter) return sendJSON({status:"error"});
  if(e.parameter.key !== CONFIG.MASTER_KEY && e.parameter.key !== CONFIG.WORKER_KEY) return sendJSON({status:"error"});
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) { 
      try {
          const p = e.parameter;
          
          if(p.action === 'resolve') {
              handleResolvePost(p); 
          }
          else if(p.action === 'registerDevice') {
              return sendJSON(handleRegisterDevice(p));
          }
          // NEW 3: Handle Notice Acknowledgments
          else if(p.action === 'acknowledgeNotice') {
              return sendJSON(handleNoticeAck(p));
          }
          else if(p.action === 'uploadEmergencyProcedures') {
              updateSiteEmergencyProcedures(p);
              handleWorkerPost(p);
          }
          else if (p.action === 'notifySafety') {
            return sendJSON(handleSafetyResolution(p));
          }
          else {
              handleWorkerPost(p);
          }
          
          return sendJSON({status:"success"});
          
      } catch(err) { 
          return sendJSON({status:"error", message: err.toString()}); 
      } 
      finally { 
          lock.releaseLock(); 
      }
  } else { 
      return sendJSON({status:"error", message:"Busy"}); 
  }
}

// ==========================================
// 4. REPORTING ENGINE (BI LAYER)
// ==========================================

function setupClientReporting() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Setup Client Reporting", "Enter exact Client Company Name (as it appears in 'Sites' tab):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const clientName = resp.getResponseText().trim();
  if (!clientName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let indexSheet = ss.getSheetByName('Reporting');
  if (!indexSheet) {
      indexSheet = ss.insertSheet('Reporting');
      indexSheet.appendRow(["Client Name", "Report Sheet ID", "Last Updated"]);
      indexSheet.getRange(1,1,1,3).setFontWeight("bold").setBackground("#e2e8f0");
  }

  const newSheetName = `Stats - ${clientName}`;
  let reportSheet = ss.getSheetByName(newSheetName);
  if (reportSheet) { ui.alert("Sheet already exists!"); return; }
  
  reportSheet = ss.insertSheet(newSheetName);
  reportSheet.appendRow(["Month", "Total Visits", "Total Hours", "Avg Duration", "Safety Checks %", "Numeric Sums (Mileage/etc)"]);
  reportSheet.setFrozenRows(1);
  reportSheet.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1e40af").setFontColor("white");

  indexSheet.appendRow([clientName, reportSheet.getSheetId().toString(), new Date()]);
  ui.alert(`✅ Reporting setup for ${clientName}. \n\nYou can now run 'Monthly Stats' to populate this sheet.`);
}

function runMonthlyStats() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Run Monthly Stats", "Enter Month (YYYY-MM):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const monthStr = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(monthStr)) { ui.alert("Invalid format. Use YYYY-MM."); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitsSheet = ss.getSheetByName('Visits');
  const indexSheet = ss.getSheetByName('Reporting');
  
  if (!visitsSheet || !indexSheet) { ui.alert("Missing 'Visits' or 'Reporting' tabs."); return; }

  const data = visitsSheet.getDataRange().getValues();
  const headers = data.shift();
  
  const dateIdx = headers.indexOf("Timestamp");
  const compIdx = headers.indexOf("Location Name"); 
  const reportIdx = headers.indexOf("Visit Report Data");
  
  const start = new Date(monthStr + "-01");
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);

  const stats = {}; 

  data.forEach(row => {
      const d = new Date(row[dateIdx]);
      if (d >= start && d <= end) {
          let client = "Unknown";
          const clientList = indexSheet.getDataRange().getValues().map(r => r[0]);
          const locName = row[compIdx].toString();
          
          const matchedClient = clientList.find(c => locName.includes(c));
          if (matchedClient) client = matchedClient;
          else return; 

          if (!stats[client]) stats[client] = { visits: 0, duration: 0, sums: {} };
          stats[client].visits++;
          
          const jsonStr = row[reportIdx];
          if (jsonStr && jsonStr.startsWith("{")) {
              try {
                  const report = JSON.parse(jsonStr);
                  for (const [k, v] of Object.entries(report)) {
                      const num = parseFloat(v);
                      if (!isNaN(num)) {
                          if (!stats[client].sums[k]) stats[client].sums[k] = 0;
                          stats[client].sums[k] += num;
                      }
                  }
              } catch(e) {}
          }
      }
  });

  const clients = indexSheet.getDataRange().getValues();
  let updatedCount = 0;

  clients.forEach(row => {
      const clientName = row[0];
      const sheetId = row[1];
      if (stats[clientName]) {
          const allSheets = ss.getSheets();
          const targetSheet = allSheets.find(s => s.getSheetId().toString() === sheetId.toString());
          
          if (targetSheet) {
              const s = stats[clientName];
              const sumStr = Object.entries(s.sums).map(([k,v]) => `${k}: ${v}`).join(", ");
              
              targetSheet.appendRow([
                  monthStr,
                  s.visits,
                  (s.visits * 0.5).toFixed(1), 
                  "N/A",
                  "100%",
                  sumStr
              ]);
              updatedCount++;
          }
      }
  });

  ui.alert(`Stats Run Complete. Updated ${updatedCount} client sheets.`);
}

function generateWorkerTravelReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitsSheet = ss.getSheetByName('Visits');
  if(!visitsSheet) { SpreadsheetApp.getUi().alert("Error: 'Visits' sheet not found."); return; }

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Run Travel Report", "Enter Month (YYYY-MM):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const monthStr = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(monthStr)) { ui.alert("Invalid format. Use YYYY-MM."); return; }

  const reportSheetName = "Travel Report - " + monthStr;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) ss.deleteSheet(reportSheet);
  reportSheet = ss.insertSheet(reportSheetName);

  const data = visitsSheet.getDataRange().getValues();
  const headers = data.shift();
  
  const col = {
    worker: headers.indexOf("Worker Name"),
    arrival: headers.indexOf("Timestamp"), 
    depart: headers.indexOf("Anticipated Departure Time"),
    report: headers.indexOf("Visit Report Data"),
    location: headers.indexOf("Location Name")
  };

  const start = new Date(monthStr + "-01");
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
  
  const workerStats = {};

  data.forEach(row => {
    const d = new Date(row[col.arrival]);
    if (d >= start && d <= end) {
        const worker = row[col.worker];
        if (!worker) return;
        if (!workerStats[worker]) workerStats[worker] = { trips: [], totalDist: 0, totalDurMs: 0 };

        let distance = 0;
        let reportJson = row[col.report];
        
        if (reportJson && reportJson.startsWith("{")) {
            try {
                const r = JSON.parse(reportJson);
                for (let key in r) {
                    if (/km|mil|dist/i.test(key)) { 
                        let val = parseFloat(r[key]);
                        if (!isNaN(val)) distance += val;
                    }
                }
            } catch(e){}
        }
        
        const distCol = headers.indexOf("Distance (km)");
        if(distCol > -1 && row[distCol]) {
             let val = parseFloat(row[distCol]);
             if(!isNaN(val)) distance = val; 
        }

        workerStats[worker].trips.push({
            date: d,
            location: row[col.location],
            distance: distance
        });
        
        workerStats[worker].totalDist += distance;
    }
  });

  let rowIdx = 1;
  reportSheet.getRange(rowIdx, 1).setValue("Travel Report: " + monthStr).setFontWeight("bold").setFontSize(14);
  rowIdx += 2;

  const sortedWorkers = Object.keys(workerStats).sort();

  sortedWorkers.forEach(worker => {
      const data = workerStats[worker];
      
      reportSheet.getRange(rowIdx, 1).setValue(worker).setFontWeight("bold").setBackground("#e2e8f0");
      reportSheet.getRange(rowIdx, 1, 1, 4).merge();
      rowIdx++;
      
      const headerRange = reportSheet.getRange(rowIdx, 1, 1, 4);
      headerRange.setValues([["Date", "Location", "Distance (km)", "Notes"]]);
      headerRange.setFontWeight("bold").setBorder(false, false, true, false, false, false);
      rowIdx++;

      data.trips.sort((a,b) => a.date - b.date).forEach(trip => {
          reportSheet.getRange(rowIdx, 1, 1, 4).setValues([[
              trip.date.toLocaleDateString() + " " + trip.date.toLocaleTimeString(),
              trip.location,
              trip.distance > 0 ? trip.distance : "-",
              ""
          ]]);
          rowIdx++;
      });

      const subTotalRow = reportSheet.getRange(rowIdx, 1, 1, 4);
      subTotalRow.setValues([["TOTALS:", "", data.totalDist.toFixed(1), ""]]);
      subTotalRow.setFontWeight("bold").setBorder(true, false, false, false, false, false);
      rowIdx += 2; 
  });

  reportSheet.autoResizeColumns(1, 4);
  ui.alert("Travel Report Generated!");
}

// ==========================================
// 5. CORE LOGIC (WORKER/MONITOR)
// ==========================================

/**
 * RE-ENGINEERED: handleResolvePost
 * Logic: Updates the Visit record AND triggers "All Clear" alerts to contacts.
 */
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
                // Targets active safety alerts
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
    
    // Fallback: If no active visit is found, log the resolution as a new entry
    if (!rowUpdated) {
        const ts = new Date();
        const dateStr = Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd");
        const row = [
            ts.toISOString(), dateStr, workerName, p['Worker Phone Number'], 
            p['Emergency Contact Name'], p['Emergency Contact Number'], p['Emergency Contact Email'], 
            p['Escalation Contact Name'], p['Escalation Contact Number'], p['Escalation Contact Email'], 
            p['Alarm Status'], p['Notes'], p['Location Name'], p['Location Address'], 
            p['Last Known GPS'], p['Timestamp'], p['Battery Level'], "", "", "", "", "", "", "", ""
        ];
        sheet.appendRow(row);
    }

    // NEW: TRIGGER "ALL CLEAR" NOTIFICATIONS
    // This sends the Email and SMS to both emergency contacts immediately.
    handleSafetyResolution(p); 
}
function handleWorkerPost(p, e) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Visits');
    const workerName = p['Worker Name'];
    const templateName = p['Template Name'] || "";
    const isNoteToSelf = (templateName.trim().toLowerCase() === 'note to self');

    let p1="", p2="", p3="", p4="", sig="";
    if(p['Photo 1']) p1 = saveImage(p['Photo 1'], workerName);
    if(p['Photo 2']) p2 = saveImage(p['Photo 2'], workerName);
    if(p['Photo 3']) p3 = saveImage(p['Photo 3'], workerName);
    if(p['Photo 4']) p4 = saveImage(p['Photo 4'], workerName);
    if(p['Signature']) sig = saveImage(p['Signature'], workerName, true); 

    const ts = new Date();
    const dateStr = Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd");
    let polishedNotes = p['Notes'] || "";
    const hasFormData = p['Visit Report Data'] && p['Visit Report Data'].length > 2;

    let distanceValue = p['Distance'] || ""; 

    if (hasFormData) {
        try {
            const reportObj = JSON.parse(p['Visit Report Data']);
            if (CONFIG.GEMINI_API_KEY && CONFIG.GEMINI_API_KEY.length > 10) {
                polishedNotes = smartScribe(reportObj, templateName, p['Notes']);
            }
            for (let key in reportObj) {
                if (/km|mil|dist|odo/i.test(key)) { 
                    let val = parseFloat(reportObj[key]);
                    if (!isNaN(val)) { distanceValue = val; break; }
                }
            }
        } catch(e) { console.error("Data Parsing Error: " + e); }
    }
    
    if (!isNoteToSelf) {
        if(!sheet) {
            sheet = ss.insertSheet('Visits');
            sheet.appendRow(["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"]);
        }
        
        let rowUpdated = false;
        const lastRow = sheet.getLastRow();
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const distColIdx = headers.indexOf("Distance (km)");
        
        if (lastRow > 1) {
            const startRow = Math.max(2, lastRow - 50); 
            const numRows = lastRow - startRow + 1;
            const data = sheet.getRange(startRow, 1, numRows, 11).getValues(); 
            for (let i = data.length - 1; i >= 0; i--) {
                const rowData = data[i];
                if (rowData[2] === workerName) {
                    const status = String(rowData[10]);
                    const isClosed = status.includes('DEPARTED') || status.includes('COMPLETED') || status.includes('DATA_ENTRY_ONLY');
                    
                    if (!isClosed) {
                        const targetRow = startRow + i;
                        sheet.getRange(targetRow, 1).setValue(ts.toISOString()); 
                        sheet.getRange(targetRow, 11).setValue(p['Alarm Status']); 
                        if (distanceValue && distColIdx > -1) sheet.getRange(targetRow, distColIdx + 1).setValue(distanceValue);
                        if (polishedNotes && polishedNotes !== rowData[11]) {
                             const oldNotes = sheet.getRange(targetRow, 12).getValue();
                             if (!oldNotes.includes(polishedNotes)) sheet.getRange(targetRow, 12).setValue((oldNotes + "\n" + polishedNotes).trim());
                        }
                        if (p['Last Known GPS']) sheet.getRange(targetRow, 15).setValue(p['Last Known GPS']);
                        if (p['Visit Report Data']) sheet.getRange(targetRow, headers.indexOf("Visit Report Data") + 1).setValue(p['Visit Report Data']);
                        rowUpdated = true;
                        break;
                    }
                }
            }
        }

// Ensure these fallbacks are in your backend script to catch the frontend keys
const emgPhone = p['Emergency Contact Number'] || p['Emergency Contact Phone'] || "";
const escPhone = p['Escalation Contact Number'] || p['Escalation Contact Phone'] || "";

if (!rowUpdated) {
    const row = [
        ts.toISOString(), 
        dateStr, 
        workerName, 
        p['Worker Phone Number'], 
        p['Emergency Contact Name'], 
        emgPhone, // FIXED: Maps frontend 'Phone' to backend 'Number'
        p['Emergency Contact Email'], 
        p['Escalation Contact Name'], 
        escPhone, // FIXED: Maps frontend 'Phone' to backend 'Number'
        p['Escalation Contact Email'], 
        p['Alarm Status'], 
        polishedNotes, 
        p['Location Name'], 
        p['Location Address'], 
        p['Last Known GPS'], 
        p['Timestamp'], 
        p['Battery Level'], 
        p1, 
        distanceValue, 
        p['Visit Report Data'], 
        p['Anticipated Departure Time'], 
        sig, 
        p2, 
        p3, 
        p4
    ];
    sheet.appendRow(row);
}
    }

    updateStaffStatus(p);
    if(hasFormData) {
        try {
            const reportObj = JSON.parse(p['Visit Report Data']);
            processFormEmail(p, reportObj, polishedNotes, p1, p2, p3, p4, sig);
        } catch(e) { console.error("Email Error: " + e); }
    }

    if(p['Alarm Status'].includes("EMERGENCY") || p['Alarm Status'].includes("PANIC") || p['Alarm Status'].includes("DURESS")) {
        triggerAlerts(p, "IMMEDIATE");
    }
}
function processFormEmail(p, reportObj, polishedNotes, p1, p2, p3, p4, sig) {
    const templateName = p['Template Name'] || "";
    if (!templateName) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName('Templates');
    const safeTName = templateName.trim().toLowerCase();
    
    let recipientEmail = "";
    
    // Audit Fix: Mandatory Worker Routing for Private Notes
    if (safeTName === 'note to self') {
        recipientEmail = p['Worker Email']; 
    } else if (tSheet) {
        const tData = tSheet.getDataRange().getValues();
        for (let i = 1; i < tData.length; i++) {
            if (tData[i][1] && tData[i][1].toString().trim().toLowerCase() === safeTName) {
                recipientEmail = tData[i][3]; 
                break;
            }
        }
    }
    
    if (!recipientEmail || !recipientEmail.includes('@')) {
        console.warn("No valid recipient found for email routing.");
        return;
    }

    const inlineImages = {};
    const imgTags = [];
    const processImg = (key, cidName, title) => {
        if (p[key] && p[key].length > 100) { 
            const blob = dataURItoBlob(p[key]);
            if (blob) {
                inlineImages[cidName] = blob;
                imgTags.push(`<div style="margin-bottom: 20px;"><p style="font-size:12px;font-weight:bold;">${title}</p><img src="cid:${cidName}" style="max-width:100%;border-radius:8px;"></div>`);
            }
        }
    };

    processImg('Photo 1', 'photo1', 'Attachment 1');
    processImg('Photo 2', 'photo2', 'Attachment 2');
    processImg('Photo 3', 'photo3', 'Attachment 3');
    processImg('Photo 4', 'photo4', 'Attachment 4');
    
    if (p['Signature']) {
        const sigBlob = dataURItoBlob(p['Signature']);
        if (sigBlob) inlineImages['signature'] = sigBlob;
    }

    // NEW: Visit Location Intelligence Block
let mapHtml = "";
    if (p['Last Known GPS']) {
        const gps = p['Last Known GPS'];
        // FIXED: Added missing $ and corrected template literal syntax
        const mapUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(gps)}`;
        mapHtml = `
        <div style="margin-top:20px; padding:15px; background:#f0f7ff; border-radius:8px; border:1px solid #cfe2ff; text-align:center;">
            <p style="margin:0 0 10px 0; font-size:11px; font-weight:800; color:#1e40af; text-transform:uppercase;">📍 Visit Location Intelligence</p>
            <a href="${mapUrl}" style="display:inline-block; padding:12px 24px; background:#1e40af; color:#ffffff; text-decoration:none; border-radius:6px; font-weight:bold;">View Location on Google Maps</a>
        </div>`;
    }

    let subject = (safeTName === 'note to self') ? `[PRIVATE] Note to Self` : `[${templateName}] - ${p['Worker Name']}`;
    let html = `<div style="font-family:Arial,sans-serif;padding:20px;max-width:600px;border:1px solid #eee;border-radius:12px;background-color:#ffffff;color:#333;">
        <h2 style="color:#1e40af;margin-top:0;">${templateName}</h2>
        <p style="color:#666;font-size:12px;">Worker: ${p['Worker Name']} | Sent: ${new Date().toLocaleString()}</p>
        <hr style="border:0;border-top:1px solid #eee;margin:20px 0;">
        
        <div style="background:#f9fafb;padding:15px;border-radius:8px;margin-bottom:20px;border-left:4px solid #1e40af;">
            <p style="white-space:pre-wrap;margin:0;font-size:14px;line-height:1.6;">${polishedNotes}</p>
        </div>
        
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
            <tbody>`;
        
    for (const [key, value] of Object.entries(reportObj)) {
        html += `<tr style="border-bottom:1px solid #f3f4f6;"><td style="padding:8px 0;font-size:13px;color:#6b7280;width:40%;">${key}</td><td style="padding:8px 0;font-size:13px;font-weight:600;color:#111827;">${value}</td></tr>`;
    }
    
    html += `</tbody></table>
        ${mapHtml} 
        <div style="margin-top:25px;">${imgTags.join('')}</div>
        ${p['Signature'] ? '<div style="margin-top:20px;padding-top:20px;border-top:1px solid #eee;"><p style="font-size:11px;color:#999;text-transform:uppercase;">Digital Signature</p><img src="cid:signature" style="max-height:80px;"></div>' : ''}
    </div>`;

    MailApp.sendEmail({ to: recipientEmail, subject: subject, htmlBody: html, inlineImages: inlineImages });

    // Privacy Purge for Private Notes
    if (p['autoDelete'] === 'true' && safeTName === 'note to self') {
        const fileUrls = [p1, p2, p3, p4, sig];
        fileUrls.forEach(url => {
            if (url && url.includes('id=')) {
                try { DriveApp.getFileById(url.split('id=')[1]).setTrashed(true); } catch(e) {}
            }
        });
    }
}

function dataURItoBlob(dataURI) {
    try {
        if (!dataURI) return null;
        let contentType = 'image/jpeg';
        let base64Data = dataURI;

        if (dataURI.includes('base64,')) {
            const parts = dataURI.split(',');
            if (parts.length < 2) return null;
            contentType = parts[0].split(':')[1].split(';')[0];
            base64Data = parts[1];
        }

        const byteString = Utilities.base64Decode(base64Data);
        return Utilities.newBlob(byteString, contentType, "image");
    } catch(e) { 
        console.error("Error decoding base64: " + e.toString());
        return null; 
    }
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
    // Strip non-numeric characters
    let n = num.toString().replace(/[^0-9]/g, ''); 
    if (n.length < 5) return null;
    
    // Handle local '0' prefix (e.g., 021 becomes +6421)
    if (n.startsWith('0')) { 
        return (CONFIG.COUNTRY_CODE || "+64") + n.substring(1); 
    }
    
    // Ensure the '+' prefix is present for Textbelt
    return n.startsWith('+') ? n : "+" + n;
}

/**
 * RE-ENGINEERED: High-Urgency Alert Router
 * Fixes: GPS Variable injection and Dual-Contact SMS Routing.
 */
function triggerAlerts(p, type) {
    // FIXED: Added missing $ for template literal variable injection
    const gpsLink = p['Last Known GPS'] ? `https://www.google.com/maps/search/?api=1&query=$${encodeURIComponent(p['Last Known GPS'])}` : "No GPS Available";
    
    // DEFAULT CONTENT
    let subject = `🚨 ${type}: ${p['Worker Name']} - ${p['Alarm Status']}`;
    let body = `SAFETY ALERT\n\nWorker: ${p['Worker Name']}\nStatus: ${p['Alarm Status']}\nLocation: ${p['Location Name']}\nNotes: ${p['Notes']}\nGPS: ${gpsLink}\nBattery: ${p['Battery Level']}`;

    // SPECIAL CASE: CRITICAL TIMING MODE
    if (p['Alarm Status'].includes("CRITICAL TIMING")) {
        subject = `🚨 URGENT: CRITICAL SAFETY BREACH - ${p['Worker Name']}`;
        body = `⚠️ URGENT SAFETY ALERT (CRITICAL TIMING MODE)\n\n` +
               `The worker, ${p['Worker Name']}, is now OVERDUE from a visit they self-identified as high-risk.\n\n` +
               `In this mode, 15 minutes is considered mission-critical. Please attempt to make contact with the worker IMMEDIATELY.\n\n` +
               `Location: ${p['Location Name']}\n` +
               `Last Known GPS: ${gpsLink}\n` +
               `Device Battery: ${p['Battery Level']}`;
    }
    
    // EMAIL ROUTING: Filter valid addresses and join for multi-recipient delivery
    const emails = [p['Emergency Contact Email'], p['Escalation Contact Email']].filter(e => e && e.includes('@'));
    if(emails.length > 0) { 
        MailApp.sendEmail({to: emails.join(','), subject: subject, body: body}); 
    }
    
    // SMS ROUTING: Sent immediately to all active contacts in the payload
    if(CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5) {
        // Ensure we capture both phone keys used in the worker app
        const numbers = [
            p['Emergency Contact Number'] || p['Emergency Contact Phone'], 
            p['Escalation Contact Number'] || p['Escalation Contact Phone']
        ].map(n => _cleanPhone(n)).filter(n => n);
        
        numbers.forEach(num => { 
            try {
                UrlFetchApp.fetch('https://textbelt.com/text', {
                    'method': 'post',
                    'payload': { 'phone': num, 'message': `${subject}\nGPS: ${gpsLink}`, 'key': CONFIG.TEXTBELT_API_KEY }
                }); 
            } catch(e) { console.error("SMS Failed: " + e.toString()); }
        });
    }
}

/**
 * RE-ENGINEERED: Multi-Stage Escalation Engine
 * Intervals: 15, 30, 45 (Primary) | 60 (Dual) | [CRITICAL_TIMING] (Immediate Dual)
 */
/**
 * RE-ENGINEERED: Multi-Stage Escalation Engine
 * Logic: Handles 15/30/45 alerts and immediate [CRITICAL_TIMING] dual-alerts.
 */
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
        if(!latest[name] || new Date(row[0]) > latest[name].time) {
            latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
        }
    }
    
    Object.keys(latest).forEach(worker => {
        try {
            const entry = latest[worker].rowData;
            const status = String(entry[10]); 
            const dueTimeStr = entry[20]; 
            const isClosed = status.includes("DEPARTED") || status.includes("COMPLETED") || status.includes("DATA_ENTRY_ONLY");
            
            if(!isClosed && dueTimeStr) {
                const due = new Date(dueTimeStr);
                const diffMins = (now - due) / 60000; 
                const isCritical = (entry[11] && entry[11].includes("[CRITICAL_TIMING]"));

                // 1. CRITICAL TIMING: Immediate Dual Alert at 0 mins
                if (isCritical && diffMins >= 0 && !status.includes("EMERGENCY")) {
                    triggerEscalation(sheet, entry, "EMERGENCY - CRITICAL TIMING BREACH", true);
                    return; 
                }

                // 2. STANDARD: 15/30/45/60 min escalations
                if (!isCritical && diffMins >= 15 && diffMins < 30 && !status.includes('15MIN')) {
                    triggerEscalation(sheet, entry, "OVERDUE - 15MIN ALERT", false);
                }
                else if (diffMins >= 30 && diffMins < 45 && !status.includes('30MIN')) {
                    triggerEscalation(sheet, entry, "OVERDUE - 30MIN ALERT", false);
                }
                else if (diffMins >= 45 && diffMins < 60 && !status.includes('45MIN')) {
                    triggerEscalation(sheet, entry, "OVERDUE - 45MIN ALERT", false);
                }
                else if (diffMins >= 60 && !status.includes("EMERGENCY")) {
                    triggerEscalation(sheet, entry, "EMERGENCY - 60MIN BREACH", true);
                }
            }
        } catch (err) { console.error(`Escalation Error: ${err.toString()}`); }
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
            for(let q=4; q<34; q++) { if(row[q]) questions.push(row[q]); }
            forms.push({name: row[1], questions: questions});
        }
    }
    return forms;
}

function saveImage(b64, workerName, isSignature) {
    if(!b64 || !CONFIG.PHOTOS_FOLDER_ID) return "";
    try {
        const blob = dataURItoBlob(b64);
        if (!blob) return "";

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
        blob.setName(fileName); 
        
        const file = targetFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch(e) { return "Error saving photo: " + e.toString(); }
}

function smartScribe(data, type, notes) {
    if(!CONFIG.GEMINI_API_KEY) return notes;
    let safeNotes = notes || "";
    let safeData = JSON.stringify(data || {});
    
    if(CONFIG.ENABLE_REDACTION) {
        const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
        safeNotes = safeNotes.replace(emailRegex, "[EMAIL_REDACTED]");
        const phoneRegex = /\b(\+?6[14][\s-]?|0)[289][0-9][\s-]?[0-9]{3}[\s-]?[0-9]{3,4}\b/g;
        safeNotes = safeNotes.replace(phoneRegex, "[PHONE_REDACTED]");
    }

    // THE MASTER EDITOR PROMPT (Universal for all Work Documents)
    const prompt = `You are the Lead Administrator for ${CONFIG.ORG_NAME}. 
    Task: Convert the provided raw field data and informal notes into a formal, structured professional report.
    Format: Professional work documentation.
    Language: Use formal ${CONFIG.LOCALE} English (e.g., if en-NZ, use 'authorised' instead of 'authorized').
    Context: This is a "${type}" report.
    Style: Clear, objective, and professional. 
    Constraint: Correct all grammar/spelling. Do NOT invent new facts. Maintain technical specificities.
    
    RAW DATA: ${safeData}
    FIELD NOTES: "${safeNotes}"
    
    Output only the polished, professional report text.`;
    
    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());
        
        if (json.candidates && json.candidates.length > 0) {
            const aiText = json.candidates[0].content.parts[0].text.trim();
            if (aiText.length < 5 || aiText.includes("I cannot")) return notes;
            return aiText;
        } else { return notes; }
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

function sendResponse(e, data) {
    const json = JSON.stringify(data);
    if (e && e.parameter && e.parameter.callback) {
        return ContentService.createTextOutput(`${e.parameter.callback}(${json})`)
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);
}

// SECURE ORS PROXY (Fixes API Key Leakage)
function getRouteDistance(start, end) {
  if (!CONFIG.ORS_API_KEY || CONFIG.ORS_API_KEY.length < 5) return null;
  
  try {
    // Reverse coordinates for ORS requirements (lon,lat)
    const p1 = start.split(',').reverse().join(',');
    const p2 = end.split(',').reverse().join(',');
    
    const url = `https://api.openrouteservice.org/v2/directions/driving-car?api_key=${CONFIG.ORS_API_KEY}&start=${p1}&end=${p2}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const meters = json.features[0].properties.segments[0].distance;
      return (meters / 1000).toFixed(2); // Return km
    }
  } catch (e) {
    console.warn("ORS Proxy Error: " + e.toString());
  }
  return null;
}
/**
 * PRIVACY SWEEP: Automatically moves private 'Note to Self' sent emails to the trash.
 * This should be set to run on a time-based trigger (e.g., every hour).
 */
function cleanupPrivateSentNotes() {
  try {
    // Search only in the Sent folder for the specific private subject line
    const threads = GmailApp.search('label:sent subject:"[PRIVATE] Note to Self"');
    
    if (threads.length > 0) {
      for (let i = 0; i < threads.length; i++) {
        threads[i].moveToTrash();
      }
      console.log(`Privacy Sweep: Moved ${threads.length} private threads to trash.`);
    }
  } catch (e) {
    console.warn("Privacy Sweep Error: " + e.toString());
  }
}

/**
 * REFINED: getSyncData with Unified Targeting
 * Logic: Pulls worker groups and applies a single Targeting Engine to filter all data.
 */
function getSyncData(workerName, deviceId) {
    // 1. THE TARGETING ENGINE (Defined once at the top)
    const isAuthorised = (targetStr, name, groups) => {
        const allowed = (targetStr || "").toString().toLowerCase().split(',').map(s => s.trim());
        if (allowed.includes("all")) return true;
        if (allowed.includes(name)) return true;
        
        const myGroups = groups.split(',').map(s => s.trim()).filter(g => g !== "");
        return myGroups.some(g => allowed.includes(g));
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stSheet = ss.getSheetByName('Staff');
    const wNameSafe = (workerName || "").toString().toLowerCase().trim();
    
    if (!stSheet) return {status: "error", message: "Staff sheet missing."};
    
    const stData = stSheet.getDataRange().getValues();
    let workerFound = false;
    let workerGroups = ""; 
    let meta = {};

    // 2. Identify Worker & Their Groups
    for (let i = 1; i < stData.length; i++) {
        if ((stData[i][0] || "").toString().toLowerCase().trim() === wNameSafe) {
            workerFound = true;
            // Column D (Index 3) is 'Group Membership'
            workerGroups = (stData[i][3] || "").toString().toLowerCase(); 
            meta.lastVehCheck = stData[i][5];
            meta.wofExpiry = stData[i][6];
            break; 
        }
    }

    if (!workerFound) return {status: "error", message: "Access Denied."};

    // 3. Filter Sites
    const sites = [];
    const siteSheet = ss.getSheetByName('Sites');
    if (siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
            if (isAuthorised(sData[i][0], wNameSafe, workerGroups)) {
                sites.push({ 
                    template: sData[i][1], company: sData[i][2], siteName: sData[i][3], 
                    address: sData[i][4], contactName: sData[i][5], 
                    contactPhone: sData[i][6], contactEmail: sData[i][7], 
                    notes: sData[i][8], emergencyProcedures: sData[i][9] 
                });
            }
        }
    }
    
    // 4. Filter Templates (Forms)
    const forms = [];
    const cachedTemplates = {};
    const tSheet = ss.getSheetByName('Templates');
    if (tSheet) {
        const tData = tSheet.getDataRange().getValues();
        for (let i = 1; i < tData.length; i++) {
            if (isAuthorised(tData[i][2], wNameSafe, workerGroups)) {
                const questions = [];
                for (let q = 4; q < 34; q++) { if (tData[i][q]) questions.push(tData[i][q]); }
                forms.push({name: tData[i][1], type: tData[i][0], questions: questions});
                cachedTemplates[tData[i][1]] = questions;
            }
        }
    }

// 5. Filter Notices (History)
    const noticeHistory = [];
    const noticeSheet = ss.getSheetByName('Notices');
    if (noticeSheet) {
        const nData = noticeSheet.getDataRange().getValues();
        for (let i = nData.length - 1; i > 0 && noticeHistory.length < 10; i--) {
            if (nData[i][6] === 'Active' && isAuthorised(nData[i][7], wNameSafe, workerGroups)) {
                noticeHistory.push({
                    id: nData[i][1], priority: nData[i][2], title: nData[i][3], 
                    content: nData[i][4], date: nData[i][0]
                });
            }
        }
        meta.noticeHistory = noticeHistory; 
        if (noticeHistory.length > 0) meta.activeNotice = noticeHistory[0];
    }
  
// 6. Filter Resources
    const resources = [];
    const resSheet = ss.getSheetByName('Resources');
    if (resSheet) {
        const rData = resSheet.getDataRange().getValues();
        for (let i = 1; i < rData.length; i++) {
            if (isAuthorised(rData[i][4], wNameSafe, workerGroups)) {
                resources.push({
                    category: rData[i][0], title: rData[i][1], 
                    type: rData[i][2], url: rData[i][3]
                });
            }
        }
        meta.resources = resources;
    }
    
    return {sites, forms, cachedTemplates, meta, version: CONFIG.VERSION};
}
/**
 * BACKEND logic: Specifically updates the 'Sites' tab
 */
function updateSiteEmergencyProcedures(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName("Sites");
  if (!siteSheet) return { status: 'error', message: 'Sites tab not found' };

  const data = siteSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 1. Identify "Emergency Procedures" column
  let colIdx = headers.indexOf("Emergency Procedures");
  if (colIdx === -1) {
    colIdx = headers.length;
    siteSheet.getRange(1, colIdx + 1).setValue("Emergency Procedures");
  }

  // 2. Locate the specific site row
  let targetRow = -1;
  const siteCol = headers.indexOf("Site Name");
  const compCol = headers.indexOf("Company Name");

  for (let i = 1; i < data.length; i++) {
    if (data[i][siteCol] === payload.siteName && data[i][compCol] === payload.companyName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) return { status: 'error', message: 'Site match failed' };

  // 3. Process Photos & Generate Links
  const photoUrls = [];
  const folder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID);
  
  (payload.photos || []).forEach((base64, idx) => {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(",")[1]), "image/jpeg", `EP_${payload.siteName}_${idx}.jpg`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    photoUrls.push(file.getUrl());
  });

  // 4. Update the Sites cell
  siteSheet.getRange(targetRow, colIdx + 1).setValue(photoUrls.join(", "));
  return { status: 'success', links: photoUrls };
}

/**
 * MISSION-CRITICAL: Notice Acknowledgment Logger
 * Appends worker name to Column I of the Notices tab.
 */
function handleNoticeAck(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Notices');
    const noticeId = p.noticeId;
    const worker = p['Worker Name'];

    if (!sheet) return { status: "error", message: "Notices tab missing" };
    
    const data = sheet.getDataRange().getValues();
    // Logic: Find the row by ID and update the 'Acknowledged By' column (Index 8 / Column I)
    for (let i = 1; i < data.length; i++) {
        if (data[i][1] === noticeId) {
            let currentAcks = data[i][8] ? data[i][8].toString().split(',').map(s => s.trim()) : [];
            if (!currentAcks.includes(worker)) {
                currentAcks.push(worker);
                sheet.getRange(i + 1, 9).setValue(currentAcks.join(', '));
            }
            break;
        }
    }
    // Record in the Visits tab for audit history
    handleWorkerPost(p); 
    return { status: "success" };
}

/**
 * HELPER: Unified Escalation Handler
 * Logic: Appends row and routes to Primary (isDual=false) or Both (isDual=true).
 */
function triggerEscalation(sheet, entry, newStatus, isDual) {
    const newRow = [...entry];
    newRow[0] = new Date().toISOString(); 
    newRow[10] = newStatus; 
    newRow[11] = entry[11] + ` [AUTO-${newStatus}]`;
    sheet.appendRow(newRow);

    const payload = {
        'Worker Name': entry[2],
        'Alarm Status': newStatus,
        'Location Name': entry[12],
        'Notes': `Alert: Worker is ${newStatus}.`,
        'Last Known GPS': entry[14],
        'Battery Level': entry[16],
        'Emergency Contact Email': entry[6],
        'Emergency Contact Number': entry[5], // Map frontend 'Phone' to backend 'Number'
        'Escalation Contact Email': isDual ? entry[9] : "",
        'Escalation Contact Number': isDual ? entry[8] : ""
    };
    triggerAlerts(payload, isDual ? "CRITICAL ESCALATION" : "OVERDUE WARNING");
}

/**
 * NEW: handleSafetyResolution
 * Logic: Notifies both contacts that the emergency has ended.
 */
function handleSafetyResolution(p) {
    // 1. Update the Visit Record for the audit trail
    handleWorkerPost(p); 

    // 2. Draft the Resolution Messages
    const subject = `✅ ALL CLEAR: ${p['Worker Name']} is Safe`;
    const body = `SAFETY RESOLUTION\n\n` +
                 `The worker, ${p['Worker Name']}, has checked in and confirmed they are SAFE.\n\n` +
                 `The previous safety alert is now resolved. No further action is required.\n\n` +
                 `Timestamp: ${new Date().toLocaleString()}\n` +
                 `Location: ${p['Location Name']}`;

    // 3. Dual-Contact Email
    const emails = [p['Emergency Contact Email'], p['Escalation Contact Email']].filter(e => e && e.includes('@'));
    if(emails.length > 0) { 
        MailApp.sendEmail({to: emails.join(','), subject: subject, body: body}); 
    }
    
    // 4. Dual-Contact SMS
    if(CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5) {
        const numbers = [p['Emergency Contact Number'], p['Escalation Contact Number']].map(n => _cleanPhone(n)).filter(n => n);
        numbers.forEach(num => { 
            try {
                UrlFetchApp.fetch('https://textbelt.com/text', {
                    'method': 'post',
                    'payload': { 'phone': num, 'message': `${subject}. Alert resolved.`, 'key': CONFIG.TEXTBELT_API_KEY }
                }); 
            } catch(e) {}
        });
    }
    return { status: "success" };
}


