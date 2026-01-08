# Meeting the OTG AppSuite: A "BYO" Safety System for Lean Organizations

I designed the **On-The-Go (OTG) AppSuite** to solve a specific problem: professional Lone Worker safety systems are often too expensive or complex for small charities and community organizations.

The OTG AppSuite is not a standard SaaS product. You don't sign up for an account, you don't pay a monthly subscription, and your data doesn't live on our servers. Instead, it is a **self-hosted system engine** that you deploy into your own Google Cloud environment.

Here is a realistic breakdown of what it does and who it is for.

---

### **How It Works**
The system is "Serverless." It uses **Google Sheets** as its database and **Google Apps Script** as its brain.

1.  **The Worker App:** Staff install a web app (PWA) on their phones. It handles GPS tracking, check-in timers, and form reporting. It works offline and syncs when connection is restored.
2.  **The Backend:** A script running in your Google Drive receives data, saves photos to your Drive folders, and acts as a "Watchdog." If a worker misses a check-in, the script triggers an escalation (Email/SMS).
3.  **The Dashboard:** Managers view a live status board on their office PC, showing active workers, battery levels, and locations.

### **Key Capabilities**
* **Tiered Escalation:** Sends a gentle "nudge" email if a worker is 5 minutes overdue, followed by a full "Emergency" SMS/Email alert to managers if they remain unresponsive.
* **Zero Tolerance:** Workers entering high-risk situations can toggle a mode that skips the warning and alerts HQ immediately if the timer expires.
* **Business Intelligence:** It doesn't just log safety; it tracks work. The system generates monthly PDF reports showing visit trends, total hours on-site, and aggregated numeric data (e.g., total mileage).
* **AI Integration:** Raw data is stored exactly as typed for legal accuracy, but email notifications use AI to "polish" hasty notes into professional English for management updates.

---

### **The "Litmus Test": Is this right for you?**

This solution is **NOT** a fit for everyone. Use this checklist to assess suitability:

**✅ You are the ideal user if:**
* **Budget is a primary constraint.** You are a charity, non-profit, or small business that cannot justify $20/user/month fees.
* **You value Data Sovereignty.** You want to own your data in your own Google Drive, not trust a third-party vendor.
* **You have "One Tech-Savvy Person."** You don't need a developer, but you need someone comfortable copying and pasting code, generating API keys, and managing a Google Sheet.
* **Your fleet is Small to Medium.** The system works best for teams of 5 to ~50 active workers.

**❌ This is likely NOT for you if:**
* **You need an SLA.** Because you host it, you are the support team. There is no 24/7 helpdesk to call if you break your spreadsheet.
* **You need Enterprise Integration.** It does not plug into Active Directory, SAP, or complex HR systems out of the box.
* **You have 500+ staff.** Google Sheets has processing limits (quotas) that very large organizations might hit.

---

### **Best-Fit Organizations**
Based on its architecture, the OTG AppSuite is tailored for:

1.  **Community Care & Social Services:** Organizations visiting clients in their homes where staff safety is a concern but funding is tight.
2.  **Environmental & Conservation Groups:** Staff working in remote areas who need offline-capable check-in tools and GPS tracking.
3.  **Property Management & Real Estate:** Solo agents performing inspections who need a discreet "Dead Man's Switch."
4.  **Volunteer Networks:** Temporary or casual fleets where installing a heavy, paid corporate app is impractical.

---

### **Summary**
The OTG AppSuite is a **"Build Your Own"** professional safety platform. It trades the convenience of a "Sign Up" button for the power of **ownership, customization, and zero running costs**. If you are willing to spend 20 minutes setting it up, you get a safety system that rivals commercial alternatives for free.# OTG AppSuite v79.17: Comprehensive Technical Reference Manual

**Version:** 79.17 (Global Golden Master)  
**Date:** January 8, 2026  
**Architecture:** Serverless / Distributed PWA  
**Runtime:** Google V8 Engine (Backend) / ECMA6 Browser (Frontend)  
**License:** MIT / Open Source "Forever Free"  

---

## 1. Architectural Philosophy & Topology

The OTG AppSuite rejects the traditional SaaS model (Software as a Service) in favor of a **Factory Pattern** deployment. This ensures data sovereignty, zero ongoing licensing costs, and resilience against vendor lock-in.

### 1.1 The Factory Pattern (Instantiation)
The `index.html` (Factory App) acts as a client-side compiler. It does not communicate with a central OTG server.
1.  **Template Loading:** It holds the raw source code for the Worker and Monitor apps as internal string variables.
2.  **Configuration Injection:** It accepts user inputs (Org Name, API Keys, Timers) and performs a "Find & Replace" operation on the raw source code using global Regex keys (e.g., `%%SECRET_KEY%%`).
3.  **Cryptographic Generation:** It generates a random 9-character alphanumeric `WORKER_KEY` used to secure the handshake between the Worker App and the Google Script.
4.  **Bundling:** It utilizes the `JSZip` library to package the compiled HTML files, manifest, and service workers into a deployable ZIP archive.

### 1.2 Data Topology (The "Thick Client" Model)
To minimize server costs and latency, logic is pushed to the client ("Thick Client"):
* **Worker App:** Handles GPS triangulation, form validation, countdown timers, and retry queues locally.
* **Monitor App:** Handles sorting, filtering, and alert rendering locally.
* **Backend:** Acts primarily as a RESTful API endpoint and database interface, only performing heavy lifting for Reporting and Escalation.



---

## 2. Backend Logic Specification (`Code.gs`)

The backend is hosted on Google Apps Script, exposing a Web App URL (`/exec`).

### 2.1 Concurrency & Locking
Google Sheets is not a transactional database. To prevent race conditions (two workers writing simultaneously causing data overwrites), the system uses `LockService`.
* **Mechanism:** `LockService.getScriptLock()`
* **Timeout:** 10,000ms (10 seconds).
* **Behavior:** If the lock cannot be acquired within 10s, the backend returns a `Server Busy` JSON error. The Worker App detects this and keeps the payload in its retry queue.

### 2.2 The "Smart Ledger" Algorithm (`handleWorkerPost`)
The system does not simply append every request as a new row. It attempts to maintain a coherent "Session" for each visit.

**Logic Flow:**
1.  **Receive Payload:** Worker sends `Worker Name` and `Alarm Status`.
2.  **Scan Context:** The script reads the **last 50 rows** of the `Visits` sheet.
3.  **Match Session:** It looks for a row where:
    * Column C (`Worker Name`) matches the incoming payload.
    * Column K (`Alarm Status`) is **NOT** a "Closed" state (`DEPARTED`, `SAFE`, `COMPLETED`, `DATA_ENTRY_ONLY`).
4.  **Decision:**
    * **Match Found:** The script **UPDATES** the existing row (updating Timestamp, Battery, GPS, Notes). This prevents "row spam" during long visits with multiple updates.
    * **No Match:** The script **APPENDS** a new row to the bottom of the sheet.

### 2.3 Tiered Escalation Watchdog (`checkOverdueVisits`)
This function must be triggered by a Time-Driven Trigger (recommended frequency: 10 minutes).

**State Machine:**
* **Input:** Iterates through all active rows in `Visits`.
* **Calculation:** `Diff = Current_Time - Anticipated_Departure_Time`.
* **Zero Tolerance Check:** If the notes field contains `[ZERO_TOLERANCE]`, the `Grace Period` variable is forced to 0 minutes. Otherwise, it uses `CONFIG.ESCALATION_MINUTES`.

**Trigger Levels:**
1.  **Tier 1 (Nudge):**
    * **Condition:** `Diff > 5 minutes` AND `Status != WARNING`.
    * **Action:** Updates status to `OVERDUE - WARNING SENT`. Sends Email to Worker/Manager.
    * **Constraint:** Skipped if Zero Tolerance is active.
2.  **Tier 2 (Escalation):**
    * **Condition:** `Diff > Grace Period` AND `Status != EMERGENCY`.
    * **Action:** Updates status to `EMERGENCY - OVERDUE`.
    * **Notification:** Sends Email to Escalation Contact + SMS via TextBelt.

### 2.4 Photo Handling & Sub-folders
The system prevents the root "Safety Photos" folder from becoming a dumping ground.
1.  **Decode:** Accepts Base64 string from payload.
2.  **Locate/Create:** Checks if a sub-folder matching `Worker Name` exists inside `PHOTOS_FOLDER_ID`. If not, creates it (`createFolder`).
3.  **Naming Convention:** Saves file as `YYYY-MM-DD_HH-mm_WorkerName_[Type].jpg` to ensure sortability.
4.  **Return:** Returns the `drive.google.com/open?id=...` URL to be written to the spreadsheet.

---

## 3. Worker App Specification (PWA)

### 3.1 Service Worker & Offline Capability
* **File:** `sw.js` (Generated dynamically).
* **Strategy:** Cache-First.
* **Behavior:** On first load, the Service Worker caches `index.html`, `manifest.json`, and the Icon. Subsequent loads (even in Airplane Mode) serve these files from the Cache Storage API.
* **Network Queue:** When offline, POST requests are pushed to a `localStorage` array (`state.pendingUploads`). A recursive function `processUploadQueue()` monitors `navigator.onLine` and flushes this queue when connectivity is restored.

### 3.2 GPS & Battery Watchdogs
* **GPS:** Uses `navigator.geolocation.watchPosition`.
    * **High Accuracy:** Enabled (`enableHighAccuracy: true`).
    * **UI Feedback:**
        * Accuracy < 20m: **3 Green Bars** (Safe).
        * Accuracy < 50m: **2 Amber Bars** (Caution).
        * Accuracy > 50m: **1 Red Bar** (Unsafe/Indoors).
* **Battery:** Uses `navigator.getBattery()`.
    * Event Listener: `levelchange`.
    * Payload: Appends `Battery Level: XX%` to every single heartbeat sent to the server.

### 3.3 Form Builder Syntax (Complete)
The app dynamically builds forms based on headers in the `Templates` sheet using a prefix parser:.

**Structure & Instructions**
* `# Header Name` or `[HEADING] Name` -> Creates a large Section Heading.
* `[NOTE] Text` -> Creates read-only instruction text (not an input).

**Standard Inputs**
* `Standard Text` -> Creates a single-line text input.
* `% Question` -> Creates a Multi-line Text Area.
* `[DATE] Label` -> Creates a Date Picker.

**Choices**
* `[YESNO] Label` -> Creates Yes/No Radio Buttons.
* `[CHECK] Label` -> Creates a simple Checkbox (e.g., "Tick to confirm").

**Smart Inputs**
* `$ Label` or `[NUMBER] Label` -> Creates a Number Input (Automatically summed in Monthly Reports).
* `[PHOTO] Label` -> Creates a Camera/Upload button.
* `[GPS] Label` -> Creates a button to capture current coordinates.
* `[SIGN] Label` -> Creates a Touchscreen Signature Pad.
---

## 4. Monitor App Specification

### 4.1 Communication Protocol (JSONP)
Because Google Apps Script Web Apps do not support CORS (Cross-Origin Resource Sharing) for GET requests from 3rd party domains (like Netlify), the Monitor App uses **JSONP (JSON with Padding)**.
* **Request:** `<script src="SCRIPT_URL?callback=cb_12345">`
* **Response:** The Google Script returns `cb_12345({ ...json_data... })` which executes immediately as JavaScript in the browser, bypassing CORS restrictions.

### 4.2 "Sound of Silence" Watchdog
Safety dashboards are dangerous if they freeze without the user knowing.
* **Logic:** The app records the timestamp of the last successful JSONP packet (`lastHeartbeat`).
* **Check:** A local timer runs every 10 seconds.
* **Trigger:** If `Date.now() - lastHeartbeat > 300,000ms` (5 minutes), the "Connection Lost" overlay covers the screen and an audio warning plays.

### 4.3 Intelligent Sorting
The Dashboard Grid does not sort alphabetically. It sorts by **Urgency Score**:
1.  **Score 2000+:** Duress/Panic (Always Top).
2.  **Score 1000+:** Emergency/Overdue.
3.  **Score 500+:** Warning State (Overdue but within Grace Period).
4.  **Score 100:** Active/Travelling.
5.  **Score 0:** Departed/Safe (Filtered out by default).

---

## 5. Database Schema (Google Sheets)

### Tab 1: `Visits` (The Ledger)
This is the transactional database.
* **Col A (Timestamp):** ISO 8601. System time of entry.
* **Col B (Date):** YYYY-MM-DD. Used for archiving/partitioning.
* **Col C (Worker Name):** The Primary Key for session matching.
* **Col K (Alarm Status):** The State Variable.
    * `ON SITE` / `TRAVELLING`: Active.
    * `DEPARTED` / `SAFE`: Closed.
    * `EMERGENCY`: Triggers Red Alerts.
    * `DURESS_CODE_ACTIVATED`: Triggers Purple Alerts.
* **Col O (Last Known GPS):** Format `lat,lon`. Parsed by Monitor Map.
* **Col T (Visit Report Data):** A JSON string containing all form answers.
* **Col U (Anticipated Departure):** ISO 8601. Used by the Escalation Watchdog.

### Tab 2: `Sites` (Configuration)
Controls the drop-down menu in the Worker App.
* **Col A (Assigned To):** Access Control List.
    * `ALL`: Visible to everyone.
    * `John Doe, Jane Smith`: Visible only to exact matches (case-insensitive).
* **Col B (Template Name):** Links the site to a specific form layout in the `Templates` tab.

### Tab 3: `Templates` (Form Definitions)
Defines the questions asked at Start/End of visit.
* **Col D (Email Recipient):** The specific email address that receives the HTML Report for this specific form type.
* **Col E onwards:** The questions/fields.

### Tab 4: `Reporting` (System Index)
* **Purpose:** Maintains a registry of Client Reporting Sheets.
* **Generated By:** The `setupClientReporting()` admin function.
* **Structure:** `Client Name | Sheet ID | Last Updated`.

---

## 6. Business Intelligence (BI) Engine

The system includes a Longitudinal Reporting module to analyze trends over time.

### 6.1 Logic Flow (`runMonthlyStats`)
1.  **Input:** Administrator inputs a month (e.g., `2025-11`).
2.  **Query:** The script fetches all rows from `Visits` where the timestamp falls within that month.
3.  **Aggregation:**
    * It groups visits by Client (based on Location Name matching).
    * **Summation:** It parses the `Visit Report Data` JSON column. It identifies any field that corresponds to a `$` (Numeric) input type and sums the values (e.g., Total Mileage).
4.  **Output:** It locates the specific `Stats - [ClientName]` sheet via the `Reporting` tab index and appends a new summary row: `Month | Total Visits | Hours | Summed Metrics`.

---

## 7. Security & Privacy Specifications

### 7.1 Data Privacy
* **Legal Source of Truth:** The `Visits` sheet contains raw, unaltered text entered by the worker. This ensures evidentiary integrity for Health & Safety audits.
* **Presentation Layer:** When generating Email Reports, the system sends the notes to Google Gemini with a prompt to *"Correct spelling and grammar"*. This polished text is used **only** in the email HTML, not saved to the database.

### 7.2 API Security
* **TextBelt:** The backend normalizes all phone numbers to **E.164 format** (removing leading zeros, adding country prefix) and sends the payload as `application/json` to ensure compatibility with strict API gateways.
* **Web App:** The Google Script accepts POST requests from "Anyone", but the first line of code checks `if (e.parameter.key !== CONFIG.WORKER_KEY) return 403;`. This prevents unauthorized data injection.

---

## 8. Integration Reference

### External APIs Used
| Service | Purpose | Auth Method | Notes |
| :--- | :--- | :--- | :--- |
| **OpenRouteService** | Calculates driving distance for "Travel" reports. | API Key | Fallback to Crow-flies if key invalid. |
| **Google Gemini** | Proofreads worker notes & summarizes reports. | API Key | Non-destructive (Sheet keeps raw data). |
| **TextBelt** | Sends SMS for Tier 2 Escalations. | API Key | Free tier: 1 SMS/day/IP. |
| **Leaflet.js** | Renders Maps in Monitor App. | Open Source | Uses OpenStreetMap tiles (Free). |
