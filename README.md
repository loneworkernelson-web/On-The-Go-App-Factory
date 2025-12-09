This is the **Master Specification & Deployment Guide** for the **OTG AppSuite (v34.0)**, formatted as a comprehensive `README.md` for GitHub.

It includes the rigorous technical constraints for an AI developer alongside a "Plain English" guide for non-technical administrators.

***

# OTG AppSuite: Serverless Lone Worker Safety System
### Version 34.0 (Smart Scribe & Smart Mileage Edition)

**OTG AppSuite** is a professional-grade, privacy-focused safety system designed for Non-Profits, NGOs, and community organizations. It provides robust lone worker monitoring, automated safety escalation, and digital reporting without monthly subscription fees.

---

## 🚀 Key Features

### For the Worker (App)
* **Simple "Two-Tap" Start:** Select a location, set a duration, and go.
* **Smart Safety:** Triple-tap Panic button, "Smart PIN" (Duress vs. Safe), and battery level monitoring.
* **Offline-First:** Works without signal. Queues reports and syncs automatically when connection returns.
* **Smart Mileage:** Automatic road distance calculation (via OpenRouteService) with "Crow Flies" fallback.
* **Quick Extensions:** Extend visit time by +10/15/30m with a single menu.

### For the Office (Monitor Dashboard)
* **Command Center:** Live dashboard with "Grid View" and "Map View" (Leaflet.js).
* **Server-Side Watchdog:** If a worker's phone is destroyed or loses battery, the *server* triggers the alarm when they are overdue.
* **Visual & Audio Alarms:** Flashing red tiles and siren audio for emergencies.
* **Audio Heartbeat:** Detects if the browser has "slept" the speakers and warns the operator immediately.

### For Administration (Backend)
* **Zero Cost:** Runs entirely on Google Sheets & Google Apps Script (Free Tier).
* **Smart Scribe:** AI-powered grammar and spelling correction for reports (using Gemini), keeping the audit trail professional.
* **Data Sovereignty:** You own the data. No third-party vendor access.

---

## 🛠️ Deployment Guide (No Coding Required)

**Prerequisites:**
* A Google Account (Gmail or Workspace).
* A computer (PC or Mac).
* ~15 minutes.

### Phase 1: Use the "Factory"
We do not ask you to write code. We provide a **Factory Tool** (`index.html`) that writes the code for you.

1.  **Open the Factory:** Double-click the `index.html` file provided in this repository. It will open in your web browser.
2.  **Step 1 - Configuration:**
    * **Org Name:** Enter your organization's name.
    * **Secret Key:** Create a password (e.g., `SafeWork2025!`). **Write this down.** It protects your database.
    * **API Keys (Optional but Recommended):**
        * *OpenRouteService Key:* For accurate road mileage.
        * *Gemini Key:* For AI grammar fixing.
        * *Textbelt Key:* For SMS alerts (leave blank to use the free tier).
3.  **Step 2 - The Spreadsheet:**
    * Click the button to create a **New Google Sheet**.
    * Rename the sheet to **"Safety System"**.
    * Rename the bottom tab to **"Visits"** (Case sensitive!).
    * **Copy Headers:** Click the button in the Factory, then paste (Ctrl+V) into Cell A1 of your sheet.
    * **Add "Checklists" Tab:** Create a new tab, name it "Checklists", and use the Factory button to copy/paste the template structure.
4.  **Step 3 - The Backend:**
    * In your Google Sheet, click **Extensions > Apps Script**.
    * **Paste Code:** The Factory will generate a large block of code. Copy it and paste it into the script editor (delete any default text there first).
    * **Deploy:**
        * Click the blue **Deploy** button (top right) -> **New Deployment**.
        * Click the "Gear" icon ⚙️ -> Select **Web App**.
        * **Description:** "v1".
        * **Execute as:** "Me" (your email).
        * **Who has access:** **Anyone** (This is crucial for the app to work without workers needing Google accounts).
        * Click **Deploy** -> **Authorize Access** -> **Allow**.
    * **Copy URL:** Copy the "Web App URL" (ends in `/exec`) and paste it back into the Factory tool.
5.  **Step 4 - Automation (The Watchdog):**
    * Back in the Apps Script window, click the **Alarm Clock icon** (Triggers) on the left.
    * Add a Trigger: `checkOverdueVisits` -> `Time-driven` -> `Minutes timer` -> `Every 10 minutes`.
    * Add a Trigger: `archiveOldData` -> `Time-driven` -> `Week timer`.
6.  **Step 5 - Download:**
    * Click **Download AppSuite.zip** in the Factory. This contains your custom Worker App and Monitor Dashboard.

### Phase 2: Going Live
1.  **Hosting:** Go to [Netlify Drop](https://app.netlify.com/drop). Drag and drop the `WorkerApp` folder from your zip file. It will give you a link (e.g., `safety-app.netlify.app`). Share this with your workers.
2.  **Monitor:** Repeat the process for the `MonitorApp` folder. Open this link on your office PC.

---

## 🤖 Master Technical Specification
*For AI Agents or Developers recreating/modifying the system.*

### 1. Architecture
* **Pattern:** Serverless PWA with Google Apps Script Backend.
* **Write Protocol:** `POST` (mode: `no-cors`). Frontend "fire-and-forgets" data to Google Sheets.
* **Read Protocol:** `JSONP`. Monitor polls backend via callback function to bypass CORS on GET requests.
* **State Management:** `localStorage` for Offline Queue and App State.

### 2. Backend Logic (`backend_template.gs`)
* **Smart De-duplication:** On `doPost`, scan last 50 rows. If `Worker Name` matches an active row (Status != DEPARTED), **UPDATE** that row. Else, **APPEND**.
* **Watchdog:** Runs every 10 mins. Checks `Col U` (Anticipated Time) vs `Now`. If `Overdue > EscalationMinutes`, update status to `EMERGENCY - OVERDUE` and trigger alerts.
* **Smart Scribe:** Use `UrlFetchApp` to call Gemini API. Prompt: *"Correct grammar/spelling to NZ English."* Apply only to Report Output (PDF/Email), never to raw DB data.

### 3. Worker App Logic (`worker_template.html`)
* **Mileage:** On `preDepart`, call OpenRouteService API. If fail/offline, fallback to Haversine. If fallback used, inject note: *"⚠️ As-the-crow-flies calculation only."*
* **Extension:** "Extend Time" button requires **Long Press (1.5s)** -> Opens Modal (+10m, +15m, +30m). PIN required only if status is `OVERDUE` or `PANIC`.
* **Battery:** Listen to `navigator.getBattery()`. Store level in global var. Send with every payload.

### 4. Monitor App Logic (`monitor_template.html`)
* **Audio Heartbeat:** Check `Tone.context.state` every 2s. If `suspended`, show Red Banner overlay.
* **Acknowledge:** Alarm trigger shows Full Screen Overlay. "HOLD TO ACKNOWLEDGE" button silences audio but keeps Tile/Pin Red & Flashing.
* **Resolution:** Operator must type Worker Name exactly to send `SAFE - MONITOR CLEARED` status.

### 5. Data Dictionary (Google Sheet Schema)
The system relies on a specific 25-column structure in the 'Visits' tab.

| Col | Header | Description |
| :-- | :--- | :--- |
| **A** | Timestamp | Server-side receipt time (Start). |
| **B** | Date | YYYY-MM-DD helper. |
| **C** | Worker Name | **Primary Key** for de-duplication. |
| **D** | Worker Phone | International format. |
| **E-J** | Contacts | Emergency & Escalation details. |
| **K** | **Alarm Status** | The State Machine (ON SITE, DEPARTED, EMERGENCY). |
| **L** | Notes | Appended log. Raw text. |
| **M** | Location Name | Site visited. |
| **N** | Location Address | Human readable address. |
| **O** | Last Known GPS | `Lat,Lon`. |
| **P** | GPS Timestamp | Client-side time. |
| **Q** | **Battery Level** | e.g. "84%". |
| **R** | Photo 1 | Drive URL. |
| **S** | Distance (km) | ORS or Haversine result. |
| **T** | Visit Report Data | JSON string of form answers. |
| **U** | **Anticipated Departure**| ISO String. Watchdog target. |
| **V** | Signature | Drive URL. |
| **W-Y**| Photos 2-4 | Drive URLs. |

---

## 📝 Reporting (AI Prompt Recipes)
The system includes an "AI Reporting Guide". To generate custom reports (e.g., Mileage, Timesheets), create a new tab in your sheet named **"AI Reporting"** and paste the guide generated by the Factory tool.

**Crucial Note for Long-Term Reports:**
Data older than 30 days is moved to the **'Archive'** tab. When asking an AI for annual reports, you must explicitly instruct it to: *"Read data from BOTH the 'Visits' tab and the 'Archive' tab."*

---

## ⚠️ Security Warning
The **Secret Key** set in the Factory is the only barrier between the internet and your database.
* **Do not share it** outside of the Factory generation process.
* **If compromised:** Change the key in the Factory, generate a new Backend Script, update the deployment, and re-distribute the Worker App.
