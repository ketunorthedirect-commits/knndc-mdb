# KNNDCmdb — Complete Setup & Deployment Guide

**Ketu North NDC Members Database**  
Version 1.0.0 | Elections & IT Directorate, 2026

---

## 📋 Overview

KNNDCmdb is a web-based membership registration portal for the NDC Ketu North Constituency. It runs on GitHub Pages (free, no server required) and syncs data to Google Sheets.

**Features:**
- 4 user roles: Data Entry Officer, Ward Coordinator, Constituency Executive, System Administrator
- Offline mode with auto-sync when back online
- Live dashboard with charts
- Audit log of all activities
- Export to CSV/Excel and PDF
- Role-based access control
- Google Sheets as backend database

---

## 🚀 STEP-BY-STEP DEPLOYMENT

---

### PHASE 1: Set Up Google Sheets (Database)

#### Step 1.1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **"+ Blank"** to create a new spreadsheet
3. Name it: **"KNNDCmdb Members Database"**
4. **Copy the Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/[THIS_IS_YOUR_SHEET_ID]/edit
   ```
   Example: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`

#### Step 1.2 — Set Up the Apps Script (Backend)

1. In your Google Sheet, click **Extensions → Apps Script**
2. A new tab opens with a script editor
3. **Delete** all existing code
4. **Copy and paste** the entire contents of `google-apps-script.js` (included in this project)
5. Click **💾 Save** (Ctrl+S)
6. Name the project: **KNNDCmdb Backend**

#### Step 1.3 — Deploy the Apps Script as a Web App

1. In the Apps Script editor, click **Deploy → New deployment**
2. Click the **gear icon ⚙️** next to "Select type" → choose **"Web App"**
3. Fill in:
   - Description: `KNNDCmdb API v1`
   - Execute as: **"Me"**
   - Who has access: **"Anyone"** *(required for the app to write data)*
4. Click **"Deploy"**
5. Click **"Authorize access"** and follow the prompts (click "Allow")
6. **Copy the Web App URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycbx.../exec
   ```
   **Save this URL — you'll need it in Step 3.**

#### Step 1.4 — Get a Google Sheets API Key (for reading data)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (name it "KNNDCmdb")
3. Go to **APIs & Services → Library**
4. Search for **"Google Sheets API"** and click **Enable**
5. Go to **APIs & Services → Credentials**
6. Click **"+ Create Credentials" → "API Key"**
7. **Copy the API Key** (starts with `AIzaSy...`)
8. Click **"Restrict Key"**:
   - Under "API restrictions" select "Restrict key" → tick "Google Sheets API"
   - Under "Application restrictions" → "HTTP referrers" → add your GitHub Pages URL later

---

### PHASE 2: Create the GitHub Repository

#### Step 2.1 — Create a GitHub Account (if you don't have one)

1. Go to [github.com](https://github.com) and sign up (free)

#### Step 2.2 — Create a New Repository

1. Click the **"+"** icon → **"New repository"**
2. Fill in:
   - Repository name: `knndc-mdb` *(or your preferred name)*
   - Description: `Ketu North NDC Members Database`
   - Visibility: **Public** *(required for free GitHub Pages)*
   - ✅ Add a README file
3. Click **"Create repository"**

#### Step 2.3 — Upload Your Files

**Option A: Using GitHub Web Interface (Easiest)**

1. Open your repository on GitHub
2. Click **"Add file" → "Upload files"**
3. Drag and drop ALL the project files/folders:
   ```
   index.html
   css/
     styles.css
   js/
     app.js
     pages.js
   assets/
     ndc-logo.png
   ```
4. Scroll down, write commit message: `Initial deployment`
5. Click **"Commit changes"**

**Option B: Using Git (For developers)**

```bash
# Clone your repository
git clone https://github.com/YOUR_USERNAME/knndc-mdb.git

# Copy all project files into the cloned folder
# Then:
cd knndc-mdb
git add .
git commit -m "Initial deployment"
git push origin main
```

---

### PHASE 3: Enable GitHub Pages

1. In your GitHub repository, click **"Settings"** (top menu)
2. Scroll down to **"Pages"** (left sidebar)
3. Under "Source", select:
   - Branch: **main**
   - Folder: **/ (root)**
4. Click **"Save"**
5. Wait 1-2 minutes, then your site is live at:
   ```
   https://YOUR_USERNAME.github.io/knndc-mdb/
   ```

---

### PHASE 4: Configure the App (Google Sheets Connection)

1. Open your live app: `https://YOUR_USERNAME.github.io/knndc-mdb/`
2. Login as admin: **admin / admin123**
3. Click **"⚙️ Settings"** in the navigation
4. Click **"📊 Google Sheets"** in the settings sidebar
5. Fill in:
   - **Google Sheet ID**: paste from Step 1.1
   - **API Key**: paste from Step 1.4
   - **Apps Script Web App URL**: paste from Step 1.3
6. Click **"💾 Save Connection"**

---

### PHASE 5: First-Time Configuration

#### 5.1 — Change the Admin Password

1. Settings → Users (in the nav)
2. Find "admin" → click **✏️ Edit**
3. Change the password to something secure
4. Click **Save User**

#### 5.2 — Add Polling Stations / Branches

1. Settings → **🏢 Polling Stations**
2. Add each polling station with:
   - Station Name (e.g. "Aflao Community A")
   - Station Code (e.g. "PS-001")
   - Branch Name (e.g. "Aflao Branch")
   - Branch Code (e.g. "BR-001")

#### 5.3 — Add Data Entry Officers

1. Click **👥 User Mgmt** in the nav
2. Click **➕ Add New User**
3. Fill in name, username, password
4. Select Role: **Data Entry Officer**
5. Assign their Polling Station Code and Branch
6. Click **Save User**

#### 5.4 — Configure the App Name (Optional)

1. Settings → **⚙️ General**
2. Change "Application Name" if needed
3. Click **Save Settings**

---

## 🔑 DEFAULT LOGIN CREDENTIALS

| Username  | Password  | Role                    |
|-----------|-----------|-------------------------|
| admin     | admin123  | System Administrator    |
| exec      | exec123   | Constituency Executive  |
| ward1     | ward123   | Ward Coordinator        |
| officer1  | off123    | Data Entry Officer      |
| officer2  | off456    | Data Entry Officer      |

⚠️ **Change all passwords immediately after deployment!**

---

## 👥 USER ROLES & ACCESS

| Feature              | Officer | Ward | Exec | Admin |
|----------------------|---------|------|------|-------|
| Data Entry Form      | ✅      | ❌   | ❌   | ✅    |
| My Records           | ✅      | ❌   | ❌   | ✅    |
| All Records          | ❌      | ✅   | ✅   | ✅    |
| Reports              | ❌      | ✅   | ✅   | ✅    |
| Analytics Dashboard  | ❌      | ❌   | ✅   | ✅    |
| Audit Log            | ❌      | ❌   | ❌   | ✅    |
| User Management      | ❌      | ❌   | ❌   | ✅    |
| Settings             | ❌      | ❌   | ❌   | ✅    |

---

## 📱 HOW TO USE

### For Data Entry Officers:
1. Login with your credentials
2. The form automatically shows your assigned Polling Station, Branch, Station Code, and Branch Code
3. Fill in: Last Name, First Name, Other Names, Party ID, Voter ID, Phone
4. Click **✅ Save Member Record**
5. Works offline — records sync when internet is restored

### For Ward Coordinators:
1. Login → view your Dashboard
2. Click **🗃️ All Records** to view and search members in your ward
3. Click **📈 Reports** for station-level statistics
4. Use Edit button to correct errors (a reason is required and logged)

### For Constituency Executives:
- Same as Ward Coordinators + access to **🔬 Analytics** dashboard
- Full view across all stations

### For Administrators:
- Full access to everything
- Manage users, stations, app settings
- View complete **🛡️ Audit Log** of all activities
- Export data to CSV/Excel or print PDF reports

---

## 🔄 UPDATING THE APP

When you make changes to the code:

**Via GitHub Web:**
1. Open the file in GitHub
2. Click the ✏️ pencil edit icon
3. Make changes
4. Click "Commit changes"
5. GitHub Pages updates automatically in 1-2 minutes

**Via Git:**
```bash
git add .
git commit -m "Description of changes"
git push origin main
```

---

## 📊 GOOGLE SHEETS STRUCTURE

After the first record is saved, your sheet will have these columns:

| Column | Field |
|--------|-------|
| A | First Name |
| B | Surname |
| C | Party ID Number |
| D | Voter ID Number |
| E | Telephone Number |
| F | Polling Station |
| G | Station Code |
| H | Branch Name |
| I | Branch Code |
| J | Other Names |
| K | Officer ID |
| L | Officer Name |
| M | Date/Time Added |
| N | Record ID |

---

## 🌐 OFFLINE MODE

The app works without internet:
- Records are saved to the browser's local storage
- A yellow banner appears when offline
- When internet is restored, all queued records sync automatically
- The offline queue count shows on the Dashboard

---

## 🛡️ SECURITY NOTES

1. **Passwords** are stored in localStorage for demo purposes. For production, use a proper auth system (Firebase Auth recommended).
2. **Google Sheets API Key** — restrict it to your GitHub Pages domain in Google Cloud Console.
3. **Apps Script** — currently open to "Anyone". For extra security, add a secret token check.
4. **Change default passwords** before sharing with users.

---

## ❓ TROUBLESHOOTING

| Problem | Solution |
|---------|----------|
| Page shows 404 | Wait 2-3 min after enabling GitHub Pages |
| Data not syncing to Sheets | Check Script URL in Settings → Google Sheets |
| Login not working | Clear browser cache, check credentials |
| Offline queue not syncing | Click Settings → Google Sheets → Sync button |
| Logo not showing | Ensure `assets/ndc-logo.png` was uploaded |

---

## 📞 SUPPORT

Powered by the **Elections and IT Directorate**  
Ketu North Constituency, NDC, 2026
