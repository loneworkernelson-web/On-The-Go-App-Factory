# 💎 OTG AppSuite v76.5: Technical Reconstruction Blueprint

**Version:** 76.5 (Global Edition)
**Date:** December 23, 2025
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
* **D:** `Worker Phone Number` (Global format: `+64...`, `+61...`)
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
The client-side `localStorage` key `loneWorkerState` persists a JSON structure containing:
* `settings`: User profile (Name, Phone, PINs).
* `locations`: Array of site objects (ID, Name, Address, Report settings).
* `activeVisit`: Current timer state (Start Time, Due Time, GPS).
* `pendingUploads`: Queue of offline payloads waiting for sync.
* `meta`: Vehicle WOF expiry and last check dates.

---

## 2. Component Specifications

### A. Factory App (`index.html`)
**Role:** Configuration Wizard & Build Engine.
**Libraries:** `JSZip` (for zip generation).

**Detailed Logic:**
1.  **Inputs:** Captures 10+ config variables (Org Name, Keys, Timezone, Toggles).
2.  **Region Logic:**
    * Dropdown selects Country (e.g., `+61`).
    * Browser API detects Timezone (e.g., `Australia/Sydney`).
    * Both are injected into the final build.
3.  **Template Injection:**
    * **User Guide:** Loads `guide_template.html`. Performs Regex replacement on conditional blocks based on user checkboxes.
    * **Ops Manual:** Loads `ops_manual_template.html`. Injects Keys and URLs.
4.  **Zip Construction:**
    * Bundles `WorkerApp` (index.html, sw.js, manifest.json, USER_GUIDE.html).
    * Bundles `MonitorApp` (index.html).
    * Bundles `Code.gs` and `OPERATIONS_MANUAL.html`.
    * Naming Convention: `OrgName_AppSuite.zip`.

### B. Worker App (`worker_template.html`)
**Role:** Primary User Interface.
**UI Framework:** Single HTML file, Tab-based navigation.

**Key Logic Modules:**

1.  **GPS Pre-warming:**
    * Action: When user taps a Location Tile.
    * Logic: Call `navigator.geolocation.getCurrentPosition` immediately (fire-and-forget).
    * Result: Wakes up GPS radio so the "Hold to Start" lock acquires a fix almost instantly.

2.  **Gatekeeper (Phone Formatting):**
    * Function: `cleanGlobalPhone(number)`
    * Logic: Strips non-integers. Checks `CONFIG.countryPrefix`.
    * If input starts with `0`, replace `0` with Prefix.
    * Ensures consistent E.164 format for SMS.

3.  **Smart Timer (The "Overdue" State):**
    * `tick()` function runs every 1000ms.
    * **Phase 1 (Normal):** Timer > 0. UI Green/Grey.
    * **Phase 2 (Grace Period):** Timer < 0 AND Time < EscalationLimit.
        * UI: Flashing Red "OVERDUE".
        * Button: "Hold to End" (Red) **remains visible**. User can end without PIN.
    * **Phase 3 (Alarm):** Time < EscalationLimit (e.g., -15 mins).
        * Action: Trigger `triggerAutoAlert('EMERGENCY - OVERDUE')`.
        * UI: Hides "Hold to End". Shows "I AM SAFE" (Green).
        * Requirement: User must enter PIN to clear.

4.  **Travel Logic:**
    * **Toggle:** "Report Required?" (Yes/No) is editable via the `i` button.
    * **End Visit Flow:**
        * If `noReport == true`: Stop timer, submit `DEPARTED`, skip form/mileage.
        * If `noReport == false`: Get GPS (End point), calc distance using `haversine` or `ORS API`, open Report Modal.

5.  **Edit Location Logic:**
    * Manual Locations (`loc_...`) and "Travelling" (`travel`) can be edited.
    * Fields: Business Name, Site Name, Address (GPS Button), Report Required (Bool), Template Name.
    * Persistence: Saves to `state.locations` in `localStorage`.

### C. Monitor App (`monitor_template.html`)
**Role:** Operations Dashboard.
**Libraries:** `Leaflet.js` (Maps), `Tone.js` (Audio).

**Key Logic Modules:**

1.  **Robust GPS Parsing:**
    * Raw data from sheet may be messy (e.g., `" -41.2, 174.0 "`).
    * Logic: `str.replace(/[^0-9.,-]/g, '')`. Split by comma. Parse Float. Check `isNaN`.
    * Only plot on map if valid.

2.  **Map Safety:**
    * `handleData` calls `renderMap`.
    * **Safety Check:** `if(typeof renderMap === 'function')`. Prevents race conditions on load.

3.  **Alert State Machine:**
    * Track `acknowledgedAlerts` (Set of names).
    * If `status.includes('EMERGENCY')` AND name NOT in `acknowledgedAlerts`:
        * Trigger Full Screen Red Alert.
        * Play Siren Loop (`Tone.js`).

### D. Backend (`backend_template.js`)
**Role:** API Gateway & Watchdog.

**Key Logic Modules:**

1.  **Staged Escalation (Watchdog):**
    * Trigger: Time-driven (Every 10 mins).
    * **Check 1:** Is `DueTime + 15mins < Now`?
        * Set Status: `EMERGENCY - OVERDUE (Stage 1)`.
        * Notify: Emergency Contact ONLY.
    * **Check 2:** Is Status == `Stage 1` AND `DueTime + 25mins < Now`?
        * Set Status: `EMERGENCY - OVERDUE (Stage 2)`.
        * Notify: Escalation Contact + Emergency Contact (Second Notice).

2.  **Timezone Handling:**
    * Fix: `Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd")`.
    * Ensures "Today's Visits" align with the user's actual day.

3.  **Resolution Logic (`action=resolve`):**
    * Input: `Worker Name`, `Notes`.
    * Action: Find last row for Worker.
    * Write: New row with status `SAFE - MANUALLY CLEARED`.
    * Notify: Send "Green Alert" email to Emergency Contacts saying "All Clear".

### E. Documentation Templates

1.  **`guide_template.html` (User Guide):**
    * **Role:** Instructions for the end-user (Worker).
    * **Features:** GPS Signal explanation, Adding Locations, Emergency Disclaimer ("Does not call Police").
    * **Dynamic:** Factory removes "Vehicle Safety" section if disabled.

2.  **`ops_manual_template.html` (Admin Manual):**
    * **Role:** Instructions for the Administrator.
    * **Features:**
        * Display Master Keys.
        * Deployment Guide (Google Apps Script + Netlify).
        * Monitor Dashboard Guide (Status Codes, Resolution).
        * AI Co-Pilot Guide (How to prompt for customizations).

---

## 3. Implementation Steps (The "Recipe")

To rebuild this system from scratch:

1.  **Setup Google Environment:**
    * Create a Google Sheet. Setup tabs: `Visits`, `Staff`, `Templates`, `Sites`.
    * Create a Google Apps Script project linked to the Sheet.
    * Paste `backend_template.js` logic.
    * **Critical:** Deploy as Web App -> Execute as Me -> Access: Anyone.

2.  **Build the Factory:**
    * Create `index.html`.
    * Create `guide_template.html` and `ops_manual_template.html`.
    * Create `worker_template.html` and `monitor_template.html`.
    * Implement the `JSZip` logic to bundle the configured files.

3.  **Deploy Client Apps:**
    * Use the Factory to generate the ZIP.
    * Host `WorkerApp` on a static HTTPS host (Netlify).
    * Host `MonitorApp` similarly.

4.  **Initialize:**
    * Open Worker App -> Run Wizard (Generates UUID).
    * Open Sheet -> `Staff` tab -> Clear `DeviceID` for the worker to Authorize the new UUID.
    * Test Sync.
