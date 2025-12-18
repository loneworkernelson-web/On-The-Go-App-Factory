
# 💎 OTG AppSuite: Technical Specifications

**Version:** 70.1 (Tactical Edition)
**Date:** December 19, 2025
**Philosophy:** "High Safety, High Trust" — Serverless, Offline-First, Tactical UX.

---

## 1. System Architecture

The OTG AppSuite is a distributed, serverless safety system. It relies on **Google Sheets** as the database and **Google Apps Script** as the API Gateway/Logic Engine. The frontend components are decoupled Progressive Web Apps (PWAs) and Single Page Applications (SPAs).

### Core Components

1. **Factory App (The Builder):** A local HTML tool that generates the deployed apps with injected cryptographic keys.
2. **Worker App (The Field Tool):** A PWA for staff. Handles GPS tracking, safety timers, reporting, and "Panic" broadcasting.
3. **Monitor App (The Dashboard):** A secure HQ dashboard for real-time situational awareness, audio alerts, and breadcrumb tracking.
4. **Backend (The Brain):** A Google Apps Script deployment that validates requests, manages the database, performs geocoding, and dispatches email/SMS alerts.

### Security Model: Dual-Key Authentication

* **Master Key (High Privilege):** Injected into the Monitor App and Backend. Allows reading all data, resolving alerts, and polling status.
* **Worker Key (Low Privilege):** Injected into the Worker App and Backend. Allows writing data (check-ins) and reading *only* the worker's specific configuration.
* **Device Fingerprinting:** Worker accounts are locked to a specific browser/device UUID upon first login.

---

## 2. Component Specifications

### A. The Factory App (`index.html`)

* **Version:** v68.9.3
* **Role:** Configuration & Build Engine.
* **Function:**
* Accepts Org Name, Logo, and Master Key.
* Generates a random `WORKER_KEY` during build.
* Uses `fetch` to load templates dynamically.
* Injects keys into specific placeholders (`%%SECRET_KEY%%`, `%%WORKER_KEY%%`).
* Outputs a `Diamond.zip` file.



### B. The Worker App (`WorkerApp/index.html`)

* **Version:** v70.1
* **UX Philosophy:** "Tactical/Gloves-On" — Large buttons, clear visuals, high contrast.
* **Core Features:**
* **GPS Traffic Light:** Real-time signal confidence indicator:
* 🟩 (3 Bars): <20m accuracy.
* 🟨 (2 Bars): 20m-100m accuracy.
* 🟥 (1 Bar): >100m accuracy.


* **Quick Time Chips:** One-tap duration setting (15m, 30m, 1h, 2h).
* **WOF Interceptor:** Blocks "Travelling" mode if the Warrant of Fitness is expired, forcing an acknowledgment checkbox.
* **Smart Sync:** Merges server site lists with local "Manual Entries" (does not wipe manual data).
* **Icon Grid Menu:** Replaces text lists with large 2-column emoji tiles for Report Selection.
* **Tactical Overdue State:**
* Timer turns **RED** and counts negative (e.g., `-05:00`).
* Prominent "OVERDUE" label appears.
* **T-5 Mins:** Warning tone + Vibrate.
* **T-2 Mins:** Voice Alert (NZ Accent prioritized) + Strong Vibrate.


* **Offline Mode:** Queues requests in `localStorage` and auto-retries on connection.



### C. The Monitor App (`MonitorApp/index.html`)

* **Version:** v70.0
* **UX Philosophy:** "Situational Awareness" — Instant visual status, breadcrumb trails.
* **Core Features:**
* **Audio Policy Handler:** "Click to Start" overlay ensures browser allows siren sounds.
* **Connection Watchdog:** If heartbeat > 45s, a full-screen "Connection Lost" overlay appears.
* **Breadcrumb Trails:** Draws a fading polyline tail behind worker markers (last 10 points) to show direction of travel.
* **Status Logic:**
* **Red:** Safety Emergency (Panic/Overdue).
* **Purple:** Duress (Silent Alarm).
* **Amber:** Administrative Issue (WOF Expired).
* **Green:** On Time / Safe.


* **Manual Resolution:** Allows HQ to force-clear an alarm with a logged note.



### D. The Backend (`Backend/Code.gs`)

* **Version:** v68.20
* **Core Capabilities:**
* **Dual-Key Auth:** Validates `MASTER_KEY` or `WORKER_KEY` depending on the action.
* **Smart Routing:**
* `action=geocode`: Server-side Proxy to Google Maps API (bypasses Client CORS).
* `action=sync`: Returns config based on Worker Name.


* **Database Logic:**
* **Append:** Creates new rows for new visits.
* **Update:** Smartly updates existing rows for check-ins/reports. Explicitly preserves Photos/Signatures during updates.


* **Notification Engine:**
* **Red Alerts:** Emails Manager + Emergency Contacts.
* **Green Alerts (Resolution):** Notifies Emergency Contacts when an alert is resolved.
* **Smart Linking:** Scans report text for GPS coordinates (e.g., `-41.2, 174.7`) and auto-converts them to clickable Google Maps links in emails.


* **Server Watchdog:** Time-based trigger (every 10-15m) that auto-escalates overdue visits to "EMERGENCY" status.



---

## 3. Data Dictionary (Google Sheets)

### Tab 1: `Visits` (Transactional Data)

| Col | Header | Description |
| --- | --- | --- |
| A | Timestamp | Server time of entry. |
| B | Date | Format: YYYY-MM-DD. |
| C | Worker Name | Unique identifier. |
| D-J | Contacts | Phone/Email for Worker, Emergency, Escalation. |
| K | Alarm Status | `ON SITE`, `DEPARTED`, `OVERDUE`, `EMERGENCY`, `PANIC`, `SAFE`, `TRAVELLING`. |
| L | Notes | Running log of updates (appended via pipe ` |
| O | Last Known GPS | Lat,Lon. **Must** be populated on initial check-in. |
| Q | Battery Level | Phone battery percentage. |
| T | Visit Report Data | JSON string containing form answers. |
| U | Anticipated Departure | ISO Timestamp. Used for Watchdog calculations. |
| V-Y | Photos/Signature | Drive Links. |

### Tab 2: `Staff` (Access Control)

| Col | Header | Description |
| --- | --- | --- |
| A | Name | Must match Worker App input exactly (Case Insensitive). |
| C | Status | "Active" or "Inactive". |
| E | DeviceID | Auto-filled on first sync. Locks account to device. Clear to reset. |
| G | WOFExpiry | Date (YYYY-MM-DD). Controlled by Vehicle Check reports. |

### Tab 3: `Templates` (Form Builder)

* **Columns E-Z:** Define questions using tags:
* `[TEXT]`, `[NUMBER]`, `[YESNO]`, `[PHOTO]`, `[GPS]`, `[SIGN]`, `[DATE]`.



---

## 4. Workflows

### 1. Start Visit (with WOF Check)

1. Worker selects site/activity.
2. **WOF Check:** If `WOFExpiry < Today`, App blocks "Travelling" and demands acknowledgment.
3. **GPS Check:** App displays Signal Strength.
4. Worker sets duration (Slider or Chip).
5. App captures GPS (via Backend Proxy for address).
6. App sends payload (`Status: ON SITE`, `GPS: <coords>`).

### 2. The "Panic" Sequence

1. Worker taps "SOS" 3x.
2. **Immediate UI:** Screen turns Red, "I AM SAFE" button appears.
3. **Background:** App sends `Status: EMERGENCY - PANIC BUTTON`.
4. **Monitor:** Detects status change -> Plays Siren -> Shows Red Overlay.
5. **Backend:** Emails/SMS Contacts with clickable Map Link.

### 3. Overdue Escalation

1. **T-5m:** App warns user (Sound/Vibrate).
2. **T-2m:** App speaks warning (NZ Voice).
3. **T-0m:** App Timer turns Red ("OVERDUE").
4. **T+Escalation Time:** Server Watchdog sees `Now > Due`.
* Updates Sheet to `EMERGENCY - OVERDUE`.
* Triggers Alerts.



### 4. Resolution (All Clear)

1. Worker taps "I AM SAFE" (requires PIN) OR Monitor clicks "Resolve".
2. Status updates to `SAFE - MANUALLY CLEARED`.
3. Backend detects `SAFE` status.
4. Backend sends "✅ ALL CLEAR" email to the Emergency Contacts who received the original alarm.
