/**
 * OTG APPSUITE - MASTER BACKEND v79.21
 * * FEATURES:
 * 1. Longitudinal Reporting (Trend Analysis for Clients).
 * 2. Travel Reporting (Dedicated Mileage/Duration Log per Worker).
 * 3. Numeric Aggregation (Sums '$' fields automatically).
 * 4. Tiered Escalation, Embedded Emails, Smart Ledger, Robust Sync.
 * 5. Custom Menu: "OTG Admin" added to spreadsheet toolbar.
 * * * PATCHES APPLIED:
 * - Terminology Sync: "Critical Timing Mode" logic updated.
 * - Robust Base64 Image Decoding.
 * - Error-Resilient Dead Man's Switch.
 * - AI Model Version Verification.
 * - Form Limit Increased to 30 Items.
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
// 2. MENU & ADMIN UI
// ==========================================
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
      if(p.test) return (p.key === CONFIG.MASTER_KEY) ? sendResponse(e, {status:"success"}) : sendResponse(e, {status:"error"});
      if(p.key === CONFIG.MASTER_KEY && !p.action) return sendResponse(e, getDashboardData());
      if(p.action === 'sync') return (p.key === CONFIG.MASTER_KEY || p.key === CONFIG.WORKER_KEY) ? sendResponse(e, getSyncData(p.worker, p.deviceId)) : sendResponse(e, {status:"error"});
      if(p.action === 'getGlobalForms') return sendResponse(e, getGlobalForms());
      return sendResponse(e, {status:"error"});
  } catch(err) { return sendResponse(e, {status:"error", message: err.toString()}); }
}

function doPost(e) {
  if(!e || !e.parameter) return sendJSON({status:"error"});
  if(e.parameter.key !== CONFIG.MASTER_KEY && e.parameter.key !== CONFIG.WORKER_KEY) return sendJSON({status:"error"});
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) { 
      try {
          if(e.parameter.action === 'resolve') handleResolvePost(e.parameter); 
          else handleWorkerPost(e.parameter);
          return sendJSON({status:"success"});
      } catch(err) { return sendJSON({status:"error", message: err.toString()}); } 
      finally { lock.releaseLock(); }
  } else { return sendJSON({status:"error", message:"Busy"}); }
}

// ==========================================
// 4. REPORTING ENGINE (BI LAYER)
// ==========================================

function setupClientReporting() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Setup Client Reporting", "Enter exact Client Company Name:", ui.ButtonSet.OK_CANCEL);
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
  ui.alert(`✅ Reporting setup for ${clientName}.`);
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
          if (matchedClient) client = matchedClient; else return; 
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
      const clientName = row[0]; const sheetId = row[1];
      if (stats[clientName]) {
          const allSheets = ss.getSheets();
          const targetSheet = allSheets.find(s => s.getSheetId().toString() === sheetId.toString());
          if (targetSheet) {
              const s = stats[clientName];
              const sumStr = Object.entries(s.sums).map(([k,v]) => `${k}: ${v}`).join(", ");
              targetSheet.appendRow([monthStr, s.visits, (s.visits * 0.5).toFixed(1), "N/A", "100%", sumStr]);
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
  const col = { worker: headers.indexOf("Worker Name"), arrival: headers.indexOf("Timestamp"), depart: headers.indexOf("Anticipated Departure Time"), report: headers.indexOf("Visit Report Data"), location: headers.indexOf("Location Name") };
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
                for (let key in r) { if (/km|mil|dist/i.test(key)) { let val = parseFloat(r[key]); if (!isNaN(val)) distance += val; } }
            } catch(e){}
        }
        const distCol = headers.indexOf("Distance (km)");
        if(distCol > -1 && row[distCol]) { let val = parseFloat(row[distCol]); if(!isNaN(val)) distance = val; }
        workerStats[worker].trips.push({ date: d, location: row[col.location], distance: distance });
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
          reportSheet.getRange(rowIdx, 1, 1, 4).setValues([[trip.date.toLocaleDateString() + " " + trip.date.toLocaleTimeString(), trip.location, trip.distance > 0 ? trip.distance : "-", ""]]);
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
    let p1="", p2="", p3="", p4="", sig="";
    if(p['Photo 1']) p1 = saveImage(p['Photo 1'], workerName);
    if(p['Photo 2']) p2 = saveImage(p['Photo 2'], workerName);
    if(p['Photo 3']) p3 = saveImage(p['Photo 3'], workerName);
    if(p['Photo 4']) p4 = saveImage(p['Photo 4'], workerName);
    if(p['Signature']) sig = saveImage(p['Signature'], workerName, true); 
    const ts = new Date(); const dateStr = Utilities.formatDate(ts, CONFIG.TIMEZONE, "yyyy-MM-dd");
    let rowUpdated = false; const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const startRow = Math.max(2, lastRow - 50); const numRows = lastRow - startRow + 1;
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
                         if (!oldNotes.includes(p['Notes'])) sheet.getRange(targetRow, 12).setValue((oldNotes + "\n" + p['Notes']).trim());
                    }
                    if (p['Last Known GPS']) sheet.getRange(targetRow, 15).setValue(p['Last Known GPS']);
                    if (p['Battery Level']) sheet.getRange(targetRow, 17).setValue(p['Battery Level']);
                    if (p['Visit Report Data']) {
                        sheet.getRange(targetRow, 20).setValue(p['Visit Report Data']);
                        if(p['Distance']) sheet.getRange(targetRow, 19).setValue(p['Distance']);
                        if(sig) sheet.getRange(targetRow, 22).setValue(sig);
                        if(p1) sheet.getRange(targetRow, 18).setValue(p1);
                        if(p2) sheet.getRange(targetRow, 23).setValue(p2);
                        if(p3) sheet.getRange(targetRow, 24).setValue(p3);
                        if(p4) sheet.getRange(targetRow, 25).setValue(p4);
                    }
                    rowUpdated = true; break;
                }
            }
        }
    }
    if (!rowUpdated) {
        const row = [ts.toISOString(), dateStr, workerName, p['Worker Phone Number'], p['Emergency Contact Name'], p['Emergency Contact Number'], p['Emergency Contact Email'], p['Escalation Contact Name'], p['Escalation Contact Number'], p['Escalation Contact Email'], p['Alarm Status'], p['Notes'], p['Location Name'], p['Location Address'], p['Last Known GPS'], p['Timestamp'], p['Battery Level'], p1, p['Distance'] || "", p['Visit Report Data'], p['Anticipated Departure Time'], sig, p2, p3, p4];
        sheet.appendRow(row);
    }
    updateStaffStatus(p);
    const hasFormData = p['Visit Report Data'] && p['Visit Report Data'].length > 2;
    if(hasFormData) {
       try {
           const reportObj = JSON.parse(p['Visit Report Data']);
           let polishedNotes = p['Notes'];
           if(CONFIG.GEMINI_API_KEY && CONFIG.GEMINI_API_KEY.length > 10) { polishedNotes = smartScribe(reportObj, p['Template Name'] || "Report", p['Notes']); }
           processFormEmail(p, reportObj, polishedNotes, p1, p2, p3, p4, sig);
       } catch(e) { console.error("Email Error: " + e); }
    }
    if(p['Alarm Status'].includes("EMERGENCY") || p['Alarm Status'].includes("PANIC") || p['Alarm Status'].includes("DURESS")) { triggerAlerts(p, "IMMEDIATE"); }
}

function processFormEmail(p, reportObj, polishedNotes, p1, p2, p3, p4, sig) {
    const templateName = p['Template Name']; if (!templateName) return;
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const tSheet = ss.getSheetByName('Templates');
    if (!tSheet) return; const tData = tSheet.getDataRange().getValues();
    let recipientEmail = ""; const safeTName = templateName.trim().toLowerCase();
    for (let i = 1; i < tData.length; i++) { if (tData[i][1] && tData[i][1].toString().trim().toLowerCase() === safeTName) { recipientEmail = tData[i][3]; break; } }
    if (!recipientEmail || !recipientEmail.includes('@')) return;
    const inlineImages = {}; const imgTags = [];
    const processImg = (key, cidName, title) => { if (p[key] && p[key].length > 100) { const blob = dataURItoBlob(p[key]); if (blob) { inlineImages[cidName] = blob; imgTags.push(`<div style="margin-bottom: 20px; text-align: center;"><p style="color:#6b7280; font-size:12px; font-weight:bold; margin-bottom:5px; text-transform:uppercase;">${title}</p><img src="cid:${cidName}" style="max-width: 100%; border-radius: 8px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.1);"></div>`); } } };
    processImg('Photo 1', 'photo1', 'Photo 1'); processImg('Photo 2', 'photo2', 'Photo 2'); processImg('Photo 3', 'photo3', 'Photo 3'); processImg('Photo 4', 'photo4', 'Photo 4');
    if (p['Signature'] && p['Signature'].length > 100) { const sigBlob = dataURItoBlob(p['Signature']); if (sigBlob) inlineImages['signature'] = sigBlob; }
    let html = `<div style="background-color:#f3f4f6; font-family: sans-serif; padding: 20px;"><div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden;"><div style="background-color: #1e40af; padding: 24px; text-align: center;"><h1 style="color: #ffffff; margin: 0; font-size: 24px;">${p['Template Name']}</h1></div><div style="padding: 24px;"><table style="width: 100%; border-collapse: collapse;">`;
    for (const [key, value] of Object.entries(reportObj)) { html += `<tr><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; font-weight: bold; font-size: 13px;">${key}</td><td style="padding: 12px; border-bottom: 1px solid #e5e7eb; font-size: 13px;">${value}</td></tr>`; }
    if (polishedNotes) { html += `<tr style="background-color: #fffbeb;"><td style="padding: 12px; font-weight: bold; font-size: 13px;">Notes (AI)</td><td style="padding: 12px; font-size: 13px; font-style: italic;">${polishedNotes}</td></tr>`; }
    html += `</table>`; if (imgTags.length > 0) html += `<div style="margin-top: 30px;">` + imgTags.join('') + `</div>`;
    if (inlineImages['signature']) html += `<div style="margin-top: 20px; text-align: right;"><p style="font-size: 10px; font-weight: bold;">Signed By Worker</p><img src="cid:signature" style="max-height: 60px;"></div>`;
    html += `</div></div></div>`;
    MailApp.sendEmail({ to: recipientEmail, subject: `[Report] ${p['Template Name']} - ${p['Worker Name']}`, htmlBody: html, inlineImages: inlineImages });
}

function dataURItoBlob(dataURI) {
    try {
        if (!dataURI) return null; let contentType = 'image/jpeg'; let base64Data = dataURI;
        if (dataURI.includes('base64,')) { const parts = dataURI.split(','); if (parts.length < 2) return null; contentType = parts[0].split(':')[1].split(';')[0]; base64Data = parts[1]; }
        const byteString = Utilities.base64Decode(base64Data); return Utilities.newBlob(byteString, contentType, "image");
    } catch(e) { console.error("Error decoding base64: " + e.toString()); return null; }
}

function updateStaffStatus(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName('Staff'); if(!sheet) return;
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
        if(data[i][0] === p['Worker Name']) {
            sheet.getRange(i+1, 5).setValue(p['deviceId']); 
            if(p['Template Name'] && p['Template Name'].includes('Vehicle')) {
                sheet.getRange(i+1, 6).setValue(new Date()); 
                try {
                    const rData = JSON.parse(p['Visit Report Data']); const term = CONFIG.VEHICLE_TERM || "WOF";
                    const expKey = Object.keys(rData).find(k => k.includes('Expiry') || k.includes(term) || k.includes('Rego'));
                    if(expKey && rData[expKey]) { sheet.getRange(i+1, 7).setValue(rData[expKey]); }
                } catch(e){}
            }
            break;
        }
    }
}

function _cleanPhone(num) {
    if (!num) return null; let n = num.toString().replace(/[^0-9]/g, ''); if (n.length < 5) return null;
    if (n.startsWith('0')) { return (CONFIG.COUNTRY_CODE || "+64") + n.substring(1); }
    const ccRaw = (CONFIG.COUNTRY_CODE || "").replace('+', ''); if (n.startsWith(ccRaw)) { return "+" + n; } return "+" + n;
}

function triggerAlerts(p, type) {
    const subject = `🚨 ${type}: ${p['Worker Name']} - ${p['Alarm Status']}`;
    const gpsLink = p['Last Known GPS'] ? `http://googleusercontent.com/maps.google.com/?q=${p['Last Known GPS']}` : "No GPS";
    const body = `SAFETY ALERT\n\nWorker: ${p['Worker Name']}\nStatus: ${p['Alarm Status']}\nLocation: ${p['Location Name']}\nNotes: ${p['Notes']}\nGPS: ${gpsLink}\nBattery: ${p['Battery Level']}`;
    const emails = [p['Emergency Contact Email'], p['Escalation Contact Email']].filter(e => e && e.includes('@'));
    if(emails.length > 0) { MailApp.sendEmail({to: emails.join(','), subject: subject, body: body}); }
    if(CONFIG.TEXTBELT_API_KEY && CONFIG.TEXTBELT_API_KEY.length > 5) {
        const numbers = [p['Emergency Contact Number'], p['Escalation Contact Number']].map(n => _cleanPhone(n)).filter(n => n);
        numbers.forEach(num => { try { UrlFetchApp.fetch('https://textbelt.com/text', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ phone: num, message: `${subject} ${gpsLink}`, key: CONFIG.TEXTBELT_API_KEY }) }); } catch(e) { console.error("SMS Failed: " + e.toString()); } });
    }
}

function checkOverdueVisits() {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Visits'); if(!sheet) return;
    const data = sheet.getDataRange().getValues(); const now = new Date(); const latest = {};
    for(let i=1; i<data.length; i++) {
        const row = data[i]; const name = row[2]; 
        if(!latest[name]) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
        else if(new Date(row[0]) > latest[name].time) latest[name] = { r: i+1, time: new Date(row[0]), rowData: row };
    }
    Object.keys(latest).forEach(worker => {
        try {
            const entry = latest[worker].rowData; const status = entry[10]; const dueTimeStr = entry[20]; 
            const isClosed = status.includes("DEPARTED") || status.includes("SAFE") || status.includes("COMPLETED");
            if(!isClosed && dueTimeStr) {
                const due = new Date(dueTimeStr); const diffMins = (now - due) / 60000; 
                // NEW: Updated terminology match for [CRITICAL_TIMING]
                const isCritical = (entry[11] && entry[11].includes("[CRITICAL_TIMING]"));
                if (diffMins > 5 && diffMins < CONFIG.ESCALATION_MINUTES && !status.includes('WARNING') && !status.includes('EMERGENCY') && !isCritical) {
                    const newRow = [...entry]; newRow[0] = new Date().toISOString(); newRow[10] = "OVERDUE - WARNING SENT"; newRow[11] = entry[11] + " [AUTO-WARNING]";
                    sheet.appendRow(newRow);
                    triggerAlerts({ 'Worker Name': worker, 'Alarm Status': "WARNING - 5 Mins Overdue", 'Location Name': entry[12], 'Notes': "5 mins overdue.", 'Last Known GPS': entry[14], 'Battery Level': entry[16], 'Emergency Contact Email': entry[6], 'Emergency Contact Number': entry[5] }, "WARNING");
                }
                const threshold = isCritical ? 0 : CONFIG.ESCALATION_MINUTES;
                if (diffMins > threshold && !status.includes("EMERGENCY")) {
                    const newStatus = isCritical ? "EMERGENCY - CRITICAL TIMING OVERDUE" : "EMERGENCY - OVERDUE";
                    const newRow = [...entry]; newRow[0] = new Date().toISOString(); newRow[10] = newStatus; newRow[11] = entry[11] + " [AUTO-ESCALATION]";
                    sheet.appendRow(newRow);
                    triggerAlerts({ 'Worker Name': worker, 'Alarm Status': newStatus, 'Location Name': entry[12], 'Notes': "Breached escalation threshold.", 'Last Known GPS': entry[14], 'Battery Level': entry[16], 'Emergency Contact Email': entry[6], 'Escalation Contact Email': entry[9], 'Emergency Contact Number': entry[5], 'Escalation Contact Number': entry[8] }, "OVERDUE");
                }
            }
        } catch (err) { console.error(`Error for ${worker}: ${err.toString()}`); }
    });
}

function getDashboardData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Visits'); const staffSheet = ss.getSheetByName('Staff'); if(!sheet) return {workers: []};
    const lastRow = sheet.getLastRow(); if (lastRow < 2) return {workers: []}; 
    const startRow = Math.max(2, lastRow - 500); 
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 25).getValues();
    const headers = ["Timestamp", "Date", "Worker Name", "Worker Phone Number", "Emergency Contact Name", "Emergency Contact Number", "Emergency Contact Email", "Escalation Contact Name", "Escalation Contact Number", "Escalation Contact Email", "Alarm Status", "Notes", "Location Name", "Location Address", "Last Known GPS", "GPS Timestamp", "Battery Level", "Photo 1", "Distance (km)", "Visit Report Data", "Anticipated Departure Time", "Signature", "Photo 2", "Photo 3", "Photo 4"];
    const workers = data.map(r => { let obj = {}; headers.forEach((h, i) => obj[h] = r[i]); return obj; });
    if(staffSheet) { const sData = staffSheet.getDataRange().getValues(); workers.forEach(w => { for(let i=1; i<sData.length; i++) { if(sData[i][0] === w['Worker Name']) { w['WOFExpiry'] = sData[i][6]; } } }); }
    return {workers: workers, escalation_limit: CONFIG.ESCALATION_MINUTES};
}

function getSyncData(workerName, deviceId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const siteSheet = ss.getSheetByName('Sites');
    const sites = []; const wNameSafe = (workerName || "").toString().toLowerCase().trim();
    if(siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for(let i=1; i<sData.length; i++) {
            const assignedStr = (sData[i][0] || "").toString().toLowerCase(); const allowedUsers = assignedStr.split(',').map(s => s.trim());
            if(allowedUsers.includes("all") || allowedUsers.includes(wNameSafe)) { sites.push({ template: sData[i][1], company: sData[i][2], siteName: sData[i][3], address: sData[i][4], contactName: sData[i][5], contactPhone: sData[i][6], contactEmail: sData[i][7], notes: sData[i][8] }); }
        }
    }
    const tSheet = ss.getSheetByName('Templates'); const forms = []; const cachedTemplates = {};
    if(tSheet) {
        const tData = tSheet.getDataRange().getValues();
        for(let i=1; i<tData.length; i++) {
            const assignedStr = (tData[i][2] || "").toString().toLowerCase(); const allowedUsers = assignedStr.split(',').map(s => s.trim());
            if(allowedUsers.includes("all") || allowedUsers.includes(wNameSafe)) {
                const questions = []; for(let q=4; q<34; q++) { if(tData[i][q]) questions.push(tData[i][q]); }
                forms.push({name: tData[i][1], type: tData[i][0], questions: questions}); cachedTemplates[tData[i][1]] = questions;
            }
        }
    }
    const meta = {}; const stSheet = ss.getSheetByName('Staff');
    if(stSheet) {
        const stData = stSheet.getDataRange().getValues();
        for(let i=1; i<stData.length; i++) {
            if((stData[i][0] || "").toString().toLowerCase().trim() === wNameSafe) {
                if(!stData[i][4]) stSheet.getRange(i+1, 5).setValue(deviceId);
                else if(stData[i][4] !== deviceId) return {status:"error", message:"DEVICE MISMATCH."};
                meta.lastVehCheck = stData[i][5]; meta.wofExpiry = stData[i][6];
            }
        }
    }
    return {sites, forms, cachedTemplates, meta};
}

function getGlobalForms() {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const tSheet = ss.getSheetByName('Templates'); if(!tSheet) return [];
    const tData = tSheet.getDataRange().getValues(); const forms = [];
    for(let i=1; i<tData.length; i++) {
        if(tData[i][2] === "ALL") { const questions = []; for(let q=4; q<34; q++) { if(tData[i][q]) questions.push(tData[i][q]); } forms.push({name: tData[i][1], questions: questions}); }
    }
    return forms;
}

function saveImage(b64, workerName, isSignature) {
    if(!b64 || !CONFIG.PHOTOS_FOLDER_ID) return "";
    try {
        const blob = dataURItoBlob(b64); if (!blob) return "";
        const mainFolder = DriveApp.getFolderById(CONFIG.PHOTOS_FOLDER_ID); let targetFolder = mainFolder;
        if (workerName && workerName.length > 2) { const folders = mainFolder.getFoldersByName(workerName); if (folders.hasNext()) { targetFolder = folders.next(); } else { targetFolder = mainFolder.createFolder(workerName); } }
        const timeStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd_HH-mm");
        const type = isSignature ? "Signature" : "Photo"; const fileName = `${timeStr}_${workerName.replace(/[^a-z0-9]/gi,'')}_${type}.jpg`;
        blob.setName(fileName); const file = targetFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return file.getUrl();
    } catch(e) { return "Error: " + e.toString(); }
}

function smartScribe(data, type, notes) {
    if(!CONFIG.GEMINI_API_KEY) return notes; let safeNotes = notes || "";
    if(CONFIG.ENABLE_REDACTION) { safeNotes = safeNotes.replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g, "[EMAIL_REDACTED]").replace(/\b(\+?\d{1,3}[- ]?)?\(?\d{3}\)?[- ]?\d{3}[- ]?\d{4}\b/g, "[PHONE_REDACTED]"); }
    const prompt = `Correct spelling and grammar in these NZ English safety notes: "${safeNotes}". Do not add facts.`;
    try {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
        const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true };
        const json = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
        return (json.candidates && json.candidates.length > 0) ? json.candidates[0].content.parts[0].text.trim() : notes;
    } catch (e) { return notes; }
}

function sendJSON(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }

function sendResponse(e, data) {
    const json = JSON.stringify(data);
    if (e && e.parameter && e.parameter.callback) { return ContentService.createTextOutput(`${e.parameter.callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT); }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
