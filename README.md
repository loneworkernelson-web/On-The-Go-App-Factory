# 💎 OTG AppSuite v76.0: Technical Reconstruction Blueprint

**Version:** 76.0 (Global Edition)
**Date:** December 22, 2025
**System Architecture:** Distributed Serverless PWA Ecosystem.
**Core Stack:** HTML5, TailwindCSS, Vanilla JS, Google Apps Script, Google Sheets.
**Offline Strategy:** Service Worker Caching + LocalStorage Transaction Queue.

---

## 1. Data Schemas & Constants

### A. Spreadsheet Database Structure
The backend relies on a Google Sheet with specific tab names and column orders.

#### 1. Tab: `Visits` (Main Ledger)
* **A:** `Timestamp` (ISO String of entry creation)
* **B:** `Date` (YYYY-MM-DD for filtering)
* **C:** `Worker Name` (String)
* **D:** `Worker Phone Number` (Global format: +64...)
* **E-G:** `Emergency Contact` (Name, Number, Email)
* **H-J:** `Escalation Contact` (Name, Number, Email)
* **K:** `Alarm Status` (e.g., `ON SITE`, `EMERGENCY - OVERDUE (Stage 1)`, `DEPARTED`, `SAFE - MANUALLY CLEARED`)
* **L:** `Notes` (Append-only log)
* **M:** `Location Name`
* **N:** `Location Address`
* **O:** `Last Known GPS` (Lat,Lon)
* **P:** `GPS Timestamp` (Last heartbeat time)
* **Q:** `Battery Level` (e.g., "85%")
* **R:** `Photo 1` (Google Drive URL)
* **S:** `Distance` (Float, km)
* **T:** `Visit Report Data` (Stringified JSON of form answers)
* **U:** `Anticipated Departure Time` (ISO String)
* **V:** `Signature` (Base64/URL)
* **W-Y:** `Photo 2-4` (URLs)

#### 2. Tab: `Staff` (Auth & Meta)
* **A:** `Name` (Primary Key for Auth)
* **B:** `Role` (Display only)
* **C:** `Status` (If contains 'Inactive', block access)
* **D:** `Token` (Unused legacy)
* **E:** `DeviceID` (UUID string. If empty, auto-binds on next sync. If mismatch, block access)
* **F:** `LastVehCheck` (ISO Timestamp)
* **G:** `WOFExpiry` (YYYY-MM-DD)

#### 3. Tab: `Templates` (Form Builder)
* **Columns:** `Type`, `Name`, `Assigned To`, `Email Recipient`, `Question 1`, `Question 2`...
* **Parser Rules:**
    * `[TEXT]`: Textarea.
    * `[YESNO]`: Radio buttons.
    * `[PHOTO]`: File input + Canvas compression.
    * `[GPS]`: Button to capture current coords.
    * `[SIGN]`: Canvas signature pad.
    * `[HEADING]`: H3 divider.

### B. Worker App State Object
The client-side `localStorage` key `loneWorkerState` persists this JSON structure:
```json
{
  "settings": {
    "workerName": "John Doe",
    "workerPhone": "+6421...",
    "pinCode": "1234",
    "duressPin": "9999",
    "googleSheetUrl": "https://..."
  },
  "locations": [
    { "id": "loc_123", "name": "Office", "address": "...", "noReport": false },
    { "id": "travel", "name": "Travelling", "noReport": true } // User toggle pref
  ],
  "activeVisit": {
    "locationId": "loc_123",
    "startTime": "ISO_STRING",
    "anticipatedDepartureTime": "ISO_STRING",
    "startGPS": "-41.2,174.7",
    "fiveMinWarned": false,
    "criticalSent": false,
    "isPanic": false
  },
  "pendingUploads": [ { "payload": "..." } ],
  "meta": { "wofExpiry": "2025-12-01", "lastVehCheck": "..." }
}
