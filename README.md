# 💎 OTG AppSuite: Technical Specifications

**Version:** 74.0 (Golden Master)
**Date:** December 21, 2025
**Philosophy:** "High Safety, High Trust" — Serverless, Offline-First, Tactical UX.

---

## 1. System Architecture

The OTG AppSuite is a distributed, serverless safety system designed for small non-profits. It relies on **Google Sheets** as the database and **Google Apps Script** as the API Gateway/Logic Engine. The frontend components are decoupled Progressive Web Apps (PWAs).

### Core Components

1.  **Factory App (The Builder):** * *Role:* Client-side HTML tool.
    * *Output:* Generates `Diamond.zip` containing the Worker and Monitor apps with injected cryptographic keys.
    * *Features:* Validates config, generates `sw.js` for offline support, injects logos.

2.  **Worker App (The Field Tool):** * *Role:* Mobile PWA for field staff.
    * *Key Features:* GPS tracking, "Dead Man's Switch" timers, "Panic" broadcasting, Vehicle Checks (WOF), Offline Queue.
    * *UX:* "Tactical" buttons, First-Run Safety Wizard.

3.  **Monitor App (The Dashboard):** * *Role:* Secure HQ dashboard for real-time awareness.
    * *Key Features:* Audio sirens (Tone.js), Traffic Light status, Breadcrumb mapping, Remote Resolution.

4.  **Backend (The Brain):** * *Role:* Google Apps Script Web App.
    * *Key Features:* API Gateway (`doGet`/`doPost`), Database CRUD, Watchdog Timers, Email/SMS Dispatch.

---

## 2. Security Model: Dual-Key Authentication

* **Master Key (High Privilege):** * *Used By:* Monitor App, Backend Admin.
    * *Capabilities:* Read all data, Resolve alerts, Poll status.
* **Worker Key (Low Privilege):** * *Used By:* Worker App.
    * *Capabilities:* Write data (Check-ins), Read own configuration.
* **Device Fingerprinting:** * *Logic:* Worker accounts are locked to a specific browser/device UUID upon first sync.
    * *Enforcement:* Backend rejects payloads if `DeviceID` does not match the 'Staff' sheet.

---

## 3. Data Schema (Google Sheet)

The system requires a Google Sheet with the following tabs:

### Tab 1: `Visits` (The Ledger)
* **Columns A-Z:** Stores every "Heartbeat" or "Action".
* **Key Columns:**
    * `K` (Alarm Status): 'ON SITE', 'EMERGENCY', 'SAFE - MANUALLY CLEARED'.
    * `P` (Device Timestamp): The actual time on the phone.
    * `T` (Visit Report Data): JSON string containing form answers.

### Tab 2: `Staff` (User Management)
* **Column A:** Worker Name (Must match App input exactly).
* **Column E:** DeviceID (UUID). *Clear this cell to reset a user's phone access.*
* **Column F:** Last Vehicle Check Date.
* **Column G:** WOF Expiry Date (YYYY-MM-DD). Controlled by Vehicle Check reports.

### Tab 3: `Templates` (Form Builder)
* **Structure:** `Type | Name | Assigned To | Email Recipient | Question 1...`
* **Tags:** `[TEXT]`, `[NUMBER]`, `[YESNO]`, `[PHOTO]`, `[GPS]`, `[SIGN]`, `[DATE]`.

---

## 4. Critical Workflows

### 1. Onboarding (The Safety Wizard)
* *Trigger:* First run or missing configuration.
* *Step 1 (Identity):* Name, Phone, Email.
* *Step 2 (Security):* User sets local **PIN** and **Duress Code**.
* *Step 3 (Contacts):* User defines **Emergency** and **Escalation** contacts (Email/Phone).

### 2. The "WOF Interceptor"
* *Logic:* If `WOFExpiry < Today`, the App blocks the "Travel" button.
* *Override:* User can acknowledge a liability waiver to proceed with non-driving work.

### 3. The "Panic" Sequence
1.  Worker taps "SOS" 3x.
2.  **Immediate UI:** Screen turns Red, "I AM SAFE" button appears.
3.  **Background:** App sends `Status: EMERGENCY - PANIC BUTTON`.
4.  **Monitor:** Detects status -> Plays Siren -> Flashes Red.
5.  **Backend:** Emails/SMS contacts with Map Link.

### 4. Alert Resolution (Green Alert)
* *Scenario:* HQ confirms worker is safe.
* *Action:* Monitor operator clicks "Resolve Alert".
* *Outcome:* * Backend logs `SAFE - MANUALLY CLEARED`.
    * System sends "ALL CLEAR" email to Emergency Contacts.
    * Dashboard returns to Green.

---

## 5. Deployment Rules

### File Structure (Diamond.zip)
* `/WorkerApp/` -> `index.html`, `manifest.json`, `sw.js` (Host on Netlify/HTTPS).
* `/MonitorApp/` -> `index.html` (Host privately or local).
* `Code.gs` -> Paste into Google Apps Script.

### Service Worker (`sw.js`)
* **Purpose:** Caches the Worker App for offline use.
* **Requirement:** Must be served via **HTTPS**.
* **Behavior:** The app will load instantly, even in "Airplane Mode". Queue data is stored in `localStorage` until connection returns.

### Audio Policy
* **Browser Restriction:** Audio cannot play automatically.
* **Solution:** Monitor App uses a "Click to Start" overlay to unlock the `AudioContext`. Worker App unlocks audio on the first user interaction (PIN entry).
