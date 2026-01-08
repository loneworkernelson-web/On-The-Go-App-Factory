# OTG AppSuite v79.17: Master System Architecture & Logic Specification

**Version:** 79.17 (Global Golden Master)  
**Date:** January 8, 2026  
**Lead Architect:** Russell Nimmo (Assisted by Google Gemini)  
**System Type:** Distributed Serverless Progressive Web App (PWA) + Business Intelligence Engine  
**Target:** Small organisations, charities, and lone worker fleets running on lean budgets.

---

## 1. System Topology & Data Flow

The OTG AppSuite operates on a **Factory Pattern**. It is not a monolithic SaaS; instead, it generates distinct, air-gapped ecosystems for each client.

### A. The Factory Pattern
* **Input:** The Administrator visits `index.html` (The Factory).
* **Configuration:** Administrator inputs Org Name, Secret Keys, and Regional Settings (NZ/AU/UK/US).
* **Process:** The Factory injects these constants into raw HTML/JS templates using Regex replacement.
* **Output:** A ZIP file containing a self-contained **Worker App**, **Monitor App**, **Operations Manual**, and **Backend Code**.

### B. Data Lifecycle
1.  **Creation:** Worker App captures data (GPS, Timestamps, Forms, Battery). Data is queued in `localStorage` if offline.
2.  **Transmission:** Worker App POSTs JSON data to the Google Apps Script Web App URL.
3.  **Ingestion (Smart Ledger):** The Backend determines whether to **Append** a new visit row or **Update** an existing active session.
4.  **Monitoring:** Monitor App polls the Web App every 10 seconds via JSONP.
5.  **Alerting:** The "Watchdog" (Time-Driven Trigger) scans for overdue visits and triggers Tiered Escalation (Email/SMS).
6.  **Intelligence:** The Reporting Engine aggregates data into longitudinal trend sheets.

---

## 2. The Backend (Code.gs)

The Backend acts as the API, Database Controller, and Notification Engine.

### A. Configuration Object (CONFIG)
Hardcoded during Factory generation:
* `MASTER_KEY`: Admin password for Monitor access.
* `WORKER_KEY`: Shared secret for App authentication.
* `PHOTOS_FOLDER_ID`: Google Drive folder for evidence.
* `ESCALATION_MINUTES`: The "Grace Period" buffer (default 15m).
* `COUNTRY_CODE`: E.164 prefix (e.g., +64) for SMS normalization.

### B. API Endpoints
* **`doGet(e)`**: Handles Read requests.
    * **Monitor Poll:** Returns the last 500 rows of `Visits` + Staff metadata.
    * **Worker Sync:** Returns `Sites`, `Templates` (Forms), and `Staff` data using **Robust Matching** (case-insensitive, ignores trailing spaces).
* **`doPost(e)`**: Handles Write requests.
    * **Smart Ledger Logic:** Before writing, the script scans the last 50 rows. If it finds an active session for the worker (Status != DEPARTED), it **updates** that row (e.g., adds departure time/notes) instead of creating a duplicate.
    * **Photo Processing:** Decodes Base64 strings, saves them to **Sub-folders by Worker Name** (e.g., `/Safety Photos/John Doe/`), and writes the Drive Link to the sheet.

### C. The Watchdog (Tiered Escalation)
A critical safety function triggered by a Google Clock Trigger (every 10 mins).
1.  **Tier 1 (Warning):** If a worker is **5 minutes overdue**, a "Warning" email is sent to the worker/supervisor.
2.  **Tier 2 (Emergency):** If a worker exceeds the `ESCALATION_MINUTES` threshold, an **EMERGENCY** alert is sent via Email and SMS (TextBelt).
3.  **Zero Tolerance Mode:** If the worker flagged "High Risk", Tier 1 is skipped, and Tier 2 fires immediately upon expiry.

### D. Reporting & BI Engine
Accessible via a custom **"🛡️ OTG Admin"** menu in the Spreadsheet.
* **Longitudinal Reporting:** Aggregates data Month-by-Month into separate tabs (e.g., `Stats - Acme Corp`).
* **Numeric Aggregation:** Automatically sums up form fields marked with `$` (e.g., Mileage, Expenses).
* **Metrics:** Tracks Total Visits, Total Hours On-Site, and Compliance percentages.

### E. AI Integration (Non-Destructive)
* **The Source of Truth:** The Google Sheet always stores the **Raw** text entered by the worker.
* **The Presentation:** When sending an email notification, the system asks Google Gemini to "proofread and polish" the notes. This polished text is used in the email body for professionalism, but the legal record remains unaltered.

---

## 3. The Worker App (PWA)

A resilient, offline-first mobile application.

### A. UI & UX
* **Grid Layout:** Locations are displayed in a responsive 2-column grid.
* **Hero Travel Tile:** A distinct, gradient-colored tile for "Travelling" mode with GPS mileage tracking.
* **GPS Watchdog:** Visual signal strength bars (Green/Amber/Red) prevent workers from starting a visit with poor accuracy.

### B. Form Builder Syntax
The app dynamically builds forms based on headers in the `Templates` sheet:
* `# Header Name` -> Creates a Section Heading.
* `% Question` -> Creates a Large Text Area.
* `$ Label` -> Creates a Number Input (Summed in Reports).
* `[PHOTO] Label` -> Camera Button.
* `[YESNO] Label` -> Radio Buttons.
* `[GPS] Label` -> Capture Coordinates button.
* `Standard Text` -> Single-line Input.

### C. Safety Logic
* **Battery Watchdog:** Captures `navigator.getBattery()` level and sends it with every ping.
* **Dead Man's Switch:** Local countdown timer.
    * **T-5 Mins:** Visual/Audio/Haptic Warning.
    * **T-0:** Counter turns Red/Negative.
    * **Escalation:** Triggers automatic SOS payload to backend.

---

## 4. The Monitor App (Dashboard)

A situational awareness dashboard for HQ.

* **Protocol:** Uses JSONP to bypass CORS.
* **Connection Watchdog:** "Sound of Silence" feature alerts HQ (visual + audio) if the dashboard loses internet connection or stops receiving data from Google.
* **Map View:** Integrated Leaflet.js map showing the last known location of all active workers.
* **Resolution Logic:** Allows managers to mark alerts as "Resolved" with a note. Uses robust string matching to ensure the correct worker is targeted.

---

## 5. Database Schema (Google Sheets)

### Tab 1: `Visits` (The Ledger)
* **Col A-B:** Timestamps.
* **Col K (Alarm Status):** Controls Monitor color (Green/Amber/Red/Purple).
* **Col O (GPS):** Lat/Lon used by Monitor Map.
* **Col R, W, X, Y:** Photo URLs (from Drive).
* **Col V:** Signature URL.
* **Col T (Report Data):** Raw JSON of form submissions.

### Tab 2: `Sites` (Configuration)
* **Col A (Assigned To):** Comma-separated list of names (e.g., "John, Jane") or "ALL".
* **Col B (Template):** Name of the form to load (matches `Templates` tab).

### Tab 3: `Templates` (Form Definitions)
* **Col A (Type):** REPORT or FORM.
* **Col D (Recipient):** Email address to receive the PDF/HTML report.
* **Col E+:** Questions using the Form Builder Syntax.

### Tab 4: `Reporting` (System Index)
* Created automatically by the "Setup Client Reporting" script. Maps Client Names to their specific Stats Sheet IDs.

---

## 6. Security & Integrations

* **Email Engine:** Sends professional HTML emails with **Inline Embedded Images (CID)**. Photos are visible instantly without clicking links.
* **TextBelt (SMS):** Uses `application/json` payload and enforces E.164 phone number formatting (`+64...`) to ensure delivery.
* **Authentication:**
    * **Worker:** Authenticates via `WORKER_KEY` injection.
    * **Monitor:** Authenticates via `MASTER_KEY` (localStorage).
    * **Device Locking:** The backend binds a worker's name to a specific `deviceId` (UUID) upon first sync to prevent spoofing.
