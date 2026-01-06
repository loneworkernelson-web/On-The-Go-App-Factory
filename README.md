# 📘 OTG AppSuite v78.0: Master System Architecture & Logic Specification

**Version:** 78.0 (Global Gold Master)
**Date:** January 6, 2026
**Lead Architect:** Russell Nimmo - Assisted by Google Gemini
**System Type:** Distributed Serverless Progressive Web App (PWA)
**License / Philosophy:** "Forever Free" (Zero-Dependency, Client-Side Logic)

---

## 1. System Topology & Data Flow

The OTG AppSuite is not a single SaaS application. It is a **Factory Pattern** system that generates standalone, air-gapped instances for individual organizations.

### A. The Factory Pattern
* **Input:** The Administrator visits `index.html` (The Factory).
* **Process:** The Factory takes configuration data (Org Name, Secrets, Region), loads raw template strings (HTML/JS), performs Regex injection (`.replace(/%%KEY%%/g, value)`), and bundles the result into a ZIP file using `JSZip`.
* **Output:** A completely self-contained `WorkerApp`, `MonitorApp`, and `BackendScript` that operate independently of the Factory.

### B. Data Lifecycle
1.  **Creation:** Worker App captures data (GPS, Timestamps, Forms). Data is queued in `localStorage` if offline.
2.  **Transmission:** Worker App `POST`s JSON data to the Google Apps Script Web App URL.
3.  **Storage:** Google Apps Script appends data to a specific Google Sheet (`Visits` tab).
4.  **Monitoring:** Monitor App `GET`s data from the Web App (polling every 10s).
5.  **Alerting:** Google Apps Script "Watchdog" (Time-Driven Trigger) scans the Sheet for overdue timestamps and triggers external APIs (Gmail, Textbelt).

---

## 2. The Backend (`Code.gs`)

The Backend is the "Brain" of the system. It runs on Google's V8 Engine within the Apps Script environment.

### A. Configuration Object (`CONFIG`)
Hardcoded at the top of the file during Factory generation:
* `MASTER_KEY`: Admin password for Monitor/Factory access.
* `WORKER_KEY`: Shared secret for Worker App authentication.
* `TEXTBELT_API_KEY`: Key for SMS gateway (Optional).
* `GEMINI_API_KEY`: Key for AI summarization (Optional).
* `ESCALATION_MINUTES`: The "Grace Period" buffer (default 15m).
* `ENABLE_REDACTION`: Boolean flag for PII scrubbing.
* `VEHICLE_TERM`: Localized string (e.g., "WOF", "Rego").

### B. API Endpoints
The script is deployed as a Web App (`Execute as: Me`, `Access: Anyone`).

#### 1. `doGet(e)` - Read Operations
* **Connection Test (`?test=1`):** Returns HTTP 200 if Key is valid. Used by Factory.
* **Monitor Poll (`?callback=...`):**
    * **Auth:** Requires `MASTER_KEY`.
    * **Logic:** Reads last 500 rows of `Visits` + `Staff` metadata.
    * **Return:** JSONP payload containing worker status, battery, GPS, and WOF expiry.
* **Worker Sync (`?action=sync`):**
    * **Auth:** Requires `WORKER_KEY` or `MASTER_KEY`.
    * **Logic:** Reads `Sites`, `Templates`, and `Staff` tabs. Returns JSON config for the app.

#### 2. `doPost(e)` - Write Operations
* **Auth:** Strict check of `key` parameter. Returns 403 error if invalid.
* **Concurrency:** Uses `LockService.getScriptLock()` (30s wait) to prevent race conditions during simultaneous writes.
* **Image Handling:** Decodes Base64 strings from POST payload -> Creates Blobs -> Saves to Google Drive -> Returns Drive URL.
* **Row Logic:**
    * **Update:** Searches last 200 rows for an active session (same Worker Name, Status != DEPARTED/SAFE). If found, updates columns (e.g., appends Notes).
    * **Append:** If no active session, appends a new row.

### C. The Watchdog (`checkOverdueVisits`)
**CRITICAL SAFETY MECHANISM.** This function must be triggered by a Google Clock Trigger (recommended: every 10 minutes).
1.  **Scan:** Reads active visits from the Sheet.
2.  **Calculate:** `TimeOverdue = Now - Anticipated_Departure_Time`.
3.  **Zero Tolerance Logic:**
    * If row notes contain `[ZERO_TOLERANCE]`, the `Grace Period` is effectively **0 minutes**.
    * Otherwise, `Grace Period` = `CONFIG.ESCALATION_MINUTES`.
4.  **Escalation Logic:**
    * **Condition:** `TimeOverdue > Threshold` AND Status is not yet `EMERGENCY`.
    * **Action:** Updates Status to `EMERGENCY - OVERDUE`.
    * **Notification:** Calls `sendAlert()` targeting **Emergency Contacts**.
    * **Stage 2:** If already in `Stage 1` and +10 mins have passed, triggers `Stage 2` (escalates to Manager/HQ).

### D. Alerting Systems (`sendAlert`)
* **Email:** Uses native `MailApp.sendEmail`. Costs: 0 (Quotas apply).
* **SMS (Textbelt):**
    * **Service:** Sends HTTP POST to `https://textbelt.com/text`.
    * **Free Tier:** If no key is provided, it attempts to use the 'textbelt' free key (1 SMS per IP per day).
    * **Paid Tier:** If `CONFIG.TEXTBELT_API_KEY` is present, it uses that for reliable delivery.
    * **Content:** "SOS: [Worker] - [Status] at [Location]".

### E. AI Integration (`smartScribe`)
* **Purpose:** Cleans up and summarizes messy voice-to-text notes using Google Gemini.
* **Privacy Layer (Redaction):**
    * Before sending to AI, regex replaces emails (`/...@.../`) with `[EMAIL_REDACTED]`.
    * Regex replaces phone numbers (International `+64...` and Local `021...`) with `[PHONE_REDACTED]`.
* **Context:** Injects `VEHICLE_TERM` into the system prompt so the AI knows to use "Rego" vs "WOF".

---

## 3. The Worker App (`worker_template.html`)

A single-file HTML5 application designed for resilience in low-connectivity environments.

### A. Offline Engine
* **Service Worker:** A generated `sw.js` file caches the HTML, Manifest, and Icon. The app loads instantly in Airplane Mode.
* **Transaction Queue:**
    * All outbound requests (Start, Heartbeat, Stop) are pushed to `state.pendingUploads` array in `localStorage`.
    * **Sync Loop:** A timer runs every few seconds. If `navigator.onLine` is true, it attempts to send the oldest request.
    * **Retry Logic:** If a request fails, it stays in the queue.

### B. "Dead Man's Switch" Timer
The app maintains a local countdown timer (`setInterval`).
* **State 1: Active:** Timer counts down from Duration (e.g., 60 mins).
* **State 2: Warning (T-5 Mins):**
    * **Visual:** "5 MINUTES REMAINING" Banner.
    * **Audio:** Plays "Pre-Alert" tone (High C6 beep).
    * **Haptic:** Vibrate pattern.
* **State 3: Overdue (T < 0):**
    * **Visual:** Counter turns RED and counts UP (negative time).
    * **Audio:** "Warning" tone repeats.
* **State 4: Imminent Alarm (T + Grace Period - 2 Mins):**
    * **Visual:** Flashing warning.
    * **Audio:** **Spoken Warning** via `SpeechSynthesis`.
    * **Voice:** *"Warning. Safety alarm will activate in two minutes."* (Uses localized accent).
* **State 5: Alarm Sent:**
    * **Trigger:** `TimeOverdue > EscalationMinutes`.
    * **Action:** App pushes `EMERGENCY - OVERDUE` payload to backend.

### C. Installation Logic (Smart Install)
* **Android/Chrome:** Listens for `beforeinstallprompt`. Prevents default. Shows a custom "Install Banner" in the app UI. Clicking it triggers the native install prompt.
* **iOS/Safari:** Detects User Agent. If not in `standalone` mode, shows a CSS overlay pointing to the bottom "Share" button with instructions.

### D. Zero Tolerance Mode
* **UI:** A toggle switch on the main screen.
* **Logic:** Sets `state.activeVisit.highRisk = true`.
* **Impact:** Adds `[ZERO_TOLERANCE]` tag to the "Started" payload notes. The Backend Watchdog sees this tag and sets the Grace Period to 0 minutes.
* **Client Side:** If timer hits 00:00, the client *immediately* fires the `triggerAutoAlert('EMERGENCY')` function, skipping the local grace period visual states.

---

## 4. The Monitor App (`monitor_template.html`)

A read-only situational awareness dashboard.

### A. Polling Architecture
* **Mechanism:** Uses `JSONP` (script injection) to bypass CORS restrictions when fetching data from Google Apps Script.
* **Frequency:** Refreshes data every 10 seconds.
* **Heartbeat:** Shows a "Connection Lost" overlay if data is stale (>45 seconds).

### B. Status Visualization Logic
* **Green (Active):** Worker is checked in, time remaining > 0.
* **Amber (Overdue):** `Due Time < Now`, but within Grace Period.
* **Red (Emergency):** Status contains `EMERGENCY` or `PANIC`. Triggers "Siren" sound (Tone.js).
* **Purple (Duress):** Status contains `DURESS`. Triggers "Stealth" visual alert (no sound, or distinct sound).

---

## 5. Database Schema (Google Sheets)

### Tab 1: `Visits` (The Ledger)
| Col | Field | Usage |
| :-- | :--- | :--- |
| **A** | `Timestamp` | System time of entry. Used for sorting Monitor feed. |
| **B** | `Date` | Used for archiving logic. |
| **C** | `Worker Name` | Unique ID for the worker. |
| **D** | `Worker Phone` | Contact info (Normalized E.164). |
| **K** | `Alarm Status` | **Key State.** Controls Monitor color and Backend alerts. |
| **L** | `Notes` | Contains user notes, `[ZERO_TOLERANCE]` tags, and `[AI]` summaries. |
| **O** | `Last Known GPS` | Format: `lat,lon`. Parsed by Monitor for Map View. |
| **U** | `Anticipated Departure` | ISO String. The target time for the Dead Man's Switch. |

### Tab 2: `Staff` (Auth)
* **Col A (Name):** Must match Worker Name in app settings exactly.
* **Col E (DeviceID):** Locks a user to a specific browser instance. Backend rejects posts if DeviceID mismatches.

### Tab 3: `Templates` (Forms)
* **Structure:** `Type | Name | Assigned To | Recipient | Q1 | Q2 ...`
* **Logic:** If `Assigned To` is "ALL", every worker sees it. If "Name", only that worker sees it.

---

## 6. Security & Limitations

### Threat Model
1.  **Public Endpoint:** The Web App URL is technically public.
    * *Mitigation:* App logic requires a valid `key` parameter (Secret Key) to read or write.
2.  **API Key Exposure:** Keys are visible in client source.
    * *Mitigation:* Use "Free Tier" keys. Worst case scenario is quota exhaustion.
3.  **Spoofing:** A worker could spoof a "SAFE" signal.
    * *Mitigation:* The "Duress PIN" feature allows a worker to simulate a safe clearance while silently triggering a Purple Alert.

### Limitations
1.  **Textbelt Free Tier:** 1 SMS per day. Highly recommended to buy a $5 key for production use.
2.  **Netlify "Drop":** Sites deleted after 1 hour unless claimed (Free Account required).
3.  **Browser Sleep:** Mobile browsers aggressively throttle JavaScript timers when the screen is off. The Server-Side Watchdog (`checkOverdueVisits`) is the ultimate fail-safe.
