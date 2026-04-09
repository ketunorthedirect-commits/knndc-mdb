/* ============================================================
   KNNDCmdb – Core Application Logic  v1.3
   ============================================================ */
'use strict';

const CONFIG = {
  SHEET_ID:      '',
  API_KEY:       '',
  SCRIPT_URL:    '',
  APP_NAME:      'Ketu North NDC Members Database',
  CONSTITUENCY:  'Ketu North',
  VERSION:       '1.3.0',
  INACTIVITY_MS: 10 * 60 * 1000,
  DEFAULT_PASSWORD: 'Ketu@2026',   // reset-to default
};

const LS = {
  SESSION:        'knndc_session',
  SETTINGS:       'knndc_settings',
  OFFLINE_Q:      'knndc_offline_queue',
  MEMBERS:        'knndc_members',
  USERS:          'knndc_users',
  AUDIT:          'knndc_audit',
  LOCKOUT:        'knndc_lockout',
  ATTEMPTS:       'knndc_attempts',
  DEMO_CLEARED:   'knndc_demo_cleared',  // flag: admin has cleared demo
};

// ─── SYSTEM USERS (non-deletable admin + demo accounts) ──────
const SYSTEM_USERS = [
  { id:'u001', username:'admin',    password:CONFIG.DEFAULT_PASSWORD, name:'System Administrator',    role:'admin',   ward:'', station:'', branch:'', assignedStations:[], active:true, mustChangePassword:false, isSystem:true },
  { id:'u002', username:'exec',     password:CONFIG.DEFAULT_PASSWORD, name:'Constituency Executive',  role:'exec',    ward:'', station:'', branch:'', assignedStations:[], active:true, mustChangePassword:true,  isSystem:true },
  { id:'u003', username:'ward1',    password:CONFIG.DEFAULT_PASSWORD, name:'Ward Coordinator (Aflao)',role:'ward',    ward:'Aflao Ward', station:'PS-001', branch:'Aflao Branch', assignedStations:['PS-001','PS-002'], active:true, mustChangePassword:true, isSystem:false },
  { id:'u004', username:'officer1', password:CONFIG.DEFAULT_PASSWORD, name:'Data Entry Officer 1',    role:'officer', ward:'Aflao Ward', station:'PS-001', branch:'Aflao Branch', assignedStations:['PS-001'], active:true, mustChangePassword:true, isSystem:false },
  { id:'u005', username:'officer2', password:CONFIG.DEFAULT_PASSWORD, name:'Data Entry Officer 2',    role:'officer', ward:'Denu Ward',  station:'PS-003', branch:'Denu Branch',  assignedStations:['PS-003'], active:true, mustChangePassword:true, isSystem:false },
];

const DEMO_POLLING_STATIONS = [
  { zone:'Zone A', ward:'Aflao Ward',    code:'PS-001', name:'Aflao A Polling Station',   branch:'Aflao Branch',    branchCode:'BR-001' },
  { zone:'Zone A', ward:'Aflao Ward',    code:'PS-002', name:'Aflao B Polling Station',   branch:'Aflao Branch',    branchCode:'BR-001' },
  { zone:'Zone B', ward:'Denu Ward',     code:'PS-003', name:'Denu Polling Station',      branch:'Denu Branch',     branchCode:'BR-002' },
  { zone:'Zone B', ward:'Agbozume Ward', code:'PS-004', name:'Agbozume Polling Station',  branch:'Agbozume Branch', branchCode:'BR-003' },
  { zone:'Zone C', ward:'Klikor Ward',   code:'PS-005', name:'Klikor Polling Station',    branch:'Klikor Branch',   branchCode:'BR-004' },
  { zone:'Zone C', ward:'Adafienu Ward', code:'PS-006', name:'Adafienu Polling Station',  branch:'Adafienu Branch', branchCode:'BR-005' },
];

// Demo members — only present before admin clears demo data
const DEMO_MEMBERS = [
  { id:'m001', firstName:'Kofi',   lastName:'Mensah',  otherNames:'Agyei',  gender:'Male',   zone:'Zone A', partyId:'NDC-2024-001', voterId:'GH-V-001', phone:'0244001122', ward:'Aflao Ward',    station:'Aflao A Polling Station',  stationCode:'PS-001', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'15/01/2024, 09:30:00', _demo:true },
  { id:'m002', firstName:'Abena',  lastName:'Korkor',  otherNames:'',       gender:'Female', zone:'Zone A', partyId:'NDC-2024-002', voterId:'GH-V-002', phone:'0244002233', ward:'Aflao Ward',    station:'Aflao A Polling Station',  stationCode:'PS-001', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'15/01/2024, 10:15:00', _demo:true },
  { id:'m003', firstName:'Yaw',    lastName:'Tetteh',  otherNames:'Kwame',  gender:'Male',   zone:'Zone B', partyId:'NDC-2024-003', voterId:'GH-V-003', phone:'0554003344', ward:'Denu Ward',     station:'Denu Polling Station',     stationCode:'PS-003', branch:'Denu Branch',     branchCode:'BR-002', officer:'officer2', officerName:'Data Entry Officer 2', timestamp:'16/01/2024, 08:45:00', _demo:true },
  { id:'m004', firstName:'Akosua', lastName:'Kporku',  otherNames:'',       gender:'Female', zone:'Zone A', partyId:'NDC-2024-004', voterId:'GH-V-004', phone:'0244004455', ward:'Aflao Ward',    station:'Aflao B Polling Station',  stationCode:'PS-002', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'16/01/2024, 11:20:00', _demo:true },
  { id:'m005', firstName:'Efo',    lastName:'Dordor',  otherNames:'Selorm', gender:'Male',   zone:'Zone B', partyId:'NDC-2024-005', voterId:'GH-V-005', phone:'0504005566', ward:'Agbozume Ward', station:'Agbozume Polling Station', stationCode:'PS-004', branch:'Agbozume Branch', branchCode:'BR-003', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'17/01/2024, 09:00:00', _demo:true },
];

// ─── APP ─────────────────────────────────────────────────────
const App = {
  currentUser:      null,
  currentPage:      'login',
  settings:         {},
  members:          [],
  users:            [],
  auditLog:         [],
  pollingStations:  [],
  offlineQueue:     [],
  isOnline:         navigator.onLine,
  _inactivityTimer: null,
  _syncInterval:    null,

  init() {
    this.loadSettings();
    this.loadData();
    this.applyAppName();
    this.setupNetworkListeners();
    this.checkSession();
  },

  // ── SETTINGS ──────────────────────────────────────────────
  loadSettings() {
    const saved = JSON.parse(localStorage.getItem(LS.SETTINGS) || '{}');
    this.settings = {
      sheetId:         saved.sheetId         || CONFIG.SHEET_ID,
      apiKey:          saved.apiKey          || CONFIG.API_KEY,
      scriptUrl:       saved.scriptUrl       || CONFIG.SCRIPT_URL,
      appName:         saved.appName         || CONFIG.APP_NAME,
      constituency:    saved.constituency    || CONFIG.CONSTITUENCY,
      pollingStations: saved.pollingStations || DEMO_POLLING_STATIONS,
    };
    this.pollingStations = this.settings.pollingStations;
  },
  saveSettings() {
    localStorage.setItem(LS.SETTINGS, JSON.stringify(this.settings));
    // Defer settings push — never block UI
    if (this.settings.scriptUrl) setTimeout(() => this._pushSettingsToSheet(), 0);
  },

  // Push settings to the Sheet's App Settings tab
  _pushSettingsToSheet() {
    if (!this.settings.scriptUrl) return;
    this._xhrPost(this.settings.scriptUrl, {
      action:       'saveSettings',
      scriptUrl:    this.settings.scriptUrl,
      appName:      this.settings.appName,
      constituency: this.settings.constituency,
      sheetId:      this.settings.sheetId,
      demoCleared:  localStorage.getItem(LS.DEMO_CLEARED) || '',
      updatedBy:    this.currentUser?.username || 'system',
    });
  },

  // Fetch remote settings from the Sheet and apply them to this device
  // Called on every login; safe to call without a scriptUrl (returns early)
  async _fetchAndApplyRemoteSettings() {
    if (!this.isOnline || !this.settings.scriptUrl) return false;
    try {
      const res  = await fetch(this.settings.scriptUrl + '?action=getSettings&t=' + Date.now());
      const data = await res.json();
      if (!data?.settings || !data.exists) return false;

      const remote = data.settings;
      let changed = false;

      // Apply remote values — remote wins for everything except apiKey (never stored server-side)
      const apply = (key, remoteKey) => {
        const rv = remote[remoteKey || key];
        if (rv && rv !== this.settings[key]) {
          this.settings[key] = rv;
          changed = true;
        }
      };

      apply('scriptUrl');
      apply('appName');
      apply('constituency');
      apply('sheetId');

      // Sync demoCleared flag across devices
      if (remote.demoCleared === '1' && !localStorage.getItem(LS.DEMO_CLEARED)) {
        localStorage.setItem(LS.DEMO_CLEARED, '1');
        // Remove demo members on this device too
        const local = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
        const clean = local.filter(m => !m._demo);
        if (clean.length !== local.length) {
          localStorage.setItem(LS.MEMBERS, JSON.stringify(clean));
          this.members = clean;
        }
        changed = true;
      }

      if (changed) {
        localStorage.setItem(LS.SETTINGS, JSON.stringify(this.settings));
        this.pollingStations = this.settings.pollingStations;
        this.applyAppName();
      }
      return changed;
    } catch(_) {
      return false;
    }
  },

  // ── DATA LOADING ──────────────────────────────────────────
  loadData() {
    const su = localStorage.getItem(LS.USERS);
    const parsedUsers = su ? JSON.parse(su) : null;
    if (parsedUsers && parsedUsers.length > 0) {
      this.users = parsedUsers;
    } else {
      // No users in storage — seed with SYSTEM_USERS defaults
      this.users = JSON.parse(JSON.stringify(SYSTEM_USERS));
      localStorage.setItem(LS.USERS, JSON.stringify(this.users)); // write directly, no push to Sheet
    }

    const sm = localStorage.getItem(LS.MEMBERS);
    const demoCleared = localStorage.getItem(LS.DEMO_CLEARED);
    if (sm) {
      this.members = JSON.parse(sm);
    } else if (!demoCleared) {
      this.members = JSON.parse(JSON.stringify(DEMO_MEMBERS));
      this.saveMembers();
    } else {
      this.members = [];
      this.saveMembers();
    }

    this.auditLog     = JSON.parse(localStorage.getItem(LS.AUDIT)    || '[]');
    this.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q)|| '[]');
  },

  saveMembers()  { localStorage.setItem(LS.MEMBERS,   JSON.stringify(this.members)); },

  // saveUsers: write to localStorage instantly, then push to Sheet in background
  saveUsers(usersArray) {
    const toSave = Array.isArray(usersArray) ? usersArray : (this.users || []);
    // Never overwrite real users with an empty array
    if (toSave.length === 0) return;
    this.users = toSave;
    localStorage.setItem(LS.USERS, JSON.stringify(toSave));
    // Defer the network call entirely — UI never waits for it
    const url = (JSON.parse(localStorage.getItem(LS.SETTINGS) || '{}')).scriptUrl || this.settings.scriptUrl;
    if (url) setTimeout(() => this._pushUsersToSheet(toSave, url), 0);
  },

  // Background push — fire-and-forget, no-cors avoids redirect/CORS issues
  _pushUsersToSheet(usersArray, scriptUrl) {
    const url = scriptUrl || this.settings.scriptUrl;
    if (!url) return;
    this._xhrPost(url, { action: 'saveUsers', users: usersArray });
  },

  // Manual trigger — called from the ☁️ Push Users to Sheet button
  async forcePushUsersToSheet() {
    const url = this.settings.scriptUrl || JSON.parse(localStorage.getItem(LS.SETTINGS)||'{}').scriptUrl;
    if (!url) {
      Toast.show('No Script URL', 'Set the Apps Script URL in Settings → Google Sheets first.', 'error');
      return;
    }
    const users = JSON.parse(localStorage.getItem(LS.USERS) || '[]');
    if (!users.length) {
      Toast.show('No Users', 'No user accounts found in local storage.', 'warning');
      return;
    }

    Toast.show('Pushing Users…', `Sending ${users.length} account(s) to Google Sheets…`, 'info', 8000);

    // Use XMLHttpRequest — more reliable than fetch for cross-origin POST to Apps Script.
    // XHR sends the request and ignores the opaque redirect response without throwing.
    const body = JSON.stringify({ action: 'saveUsers', users });
    const success = await new Promise((resolve) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.timeout = 15000; // 15 second timeout
      xhr.onload  = () => resolve(true);   // any response = data was received
      xhr.onerror = () => resolve(false);
      xhr.ontimeout = () => resolve(false);
      xhr.send(body);
    });

    if (success) {
      Toast.show('Users Pushed ✅', `${users.length} account(s) sent to Google Sheets. Open the Users tab to confirm.`, 'success', 7000);
      this.logAudit('SYNC_USERS', `Pushed ${users.length} users to Google Sheets`, this.currentUser?.username || 'admin');
    } else {
      Toast.show('Push Failed', 'Request timed out or network error. Check your internet connection and Script URL, then try again.', 'error', 7000);
    }
  },

  // Pull users from the Sheet and MERGE safely into localStorage.
  // Rules:
  //  - Sheet adds new users that aren't local yet
  //  - Sheet updates non-password fields (name, role, ward, station, assignedStations, active)
  //  - Password in Sheet only wins if it is non-empty AND different from local
  //  - Local admin record is ALWAYS preserved — it is never deleted by a Sheet fetch
  async _fetchUsersFromSheet() {
    if (!this.isOnline || !this.settings.scriptUrl) return false;
    try {
      const res  = await fetch(this.settings.scriptUrl + '?action=getUsers&t=' + Date.now());
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch(_) { return false; }

      if (!data?.users || !Array.isArray(data.users) || data.users.length === 0) {
        return false; // Sheet empty — keep local users as-is
      }

      // Start from current local users as the base
      const local  = JSON.parse(localStorage.getItem(LS.USERS) || '[]');
      const merged = [...local];

      data.users.forEach(sheetUser => {
        if (!sheetUser.username && !sheetUser.id) return; // skip malformed rows
        const idx = merged.findIndex(u =>
          u.id === sheetUser.id || u.username === sheetUser.username
        );
        if (idx >= 0) {
          // Update existing user — but preserve local password if Sheet has empty/same
          const localPwd = merged[idx].password;
          const sheetPwd = sheetUser.password?.trim();
          merged[idx] = {
            ...merged[idx],
            ...sheetUser,
            // Keep local password if sheet password is missing or blank
            password: sheetPwd || localPwd,
          };
        } else {
          // New user from Sheet — add only if they have a password
          if (sheetUser.password?.trim()) {
            merged.push(sheetUser);
          }
        }
      });

      // Safety net: ensure admin account is always present and has a password
      const hasAdmin = merged.some(u => u.role === 'admin' && u.password && u.active !== false);
      if (!hasAdmin) {
        // Don't overwrite — abort this sync to protect login
        return false;
      }

      // Only write back if we ended up with a valid set
      if (!merged.length) return false;

      localStorage.setItem(LS.USERS, JSON.stringify(merged));
      this.users = merged;
      return true;
    } catch(_) {
      return false;
    }
  },

  applyAppName() {
    const n = this.settings.appName || CONFIG.APP_NAME;
    document.querySelectorAll('.app-title-text').forEach(el => el.textContent = n);
    document.title = n;
  },

  // ── NETWORK ───────────────────────────────────────────────
  setupNetworkListeners() {
    window.addEventListener('online',  () => { this.isOnline = true;  this.updateOnlineStatus(); this.flushOfflineQueue(); });
    window.addEventListener('offline', () => { this.isOnline = false; this.updateOnlineStatus(); });
    this.updateOnlineStatus();
  },
  updateOnlineStatus() {
    document.getElementById('offline-banner')?.classList.toggle('show', !this.isOnline);
    const dot = document.getElementById('conn-dot');
    if (dot) dot.className = this.isOnline ? 'online-dot' : 'offline-dot';
  },

  // ── SESSION / INACTIVITY ──────────────────────────────────
  checkSession() {
    const s = sessionStorage.getItem(LS.SESSION);
    if (s) {
      try { this.currentUser = JSON.parse(s); this.showApp(); return; } catch(_) {}
    }
    this.showLogin();
  },

  resetInactivityTimer() {
    if (this._inactivityTimer) clearTimeout(this._inactivityTimer);
    this._inactivityTimer = setTimeout(() => {
      if (!this.currentUser) return;
      this.logAudit('AUTO_LOGOUT','Session expired — 10 min inactivity', this.currentUser.username);
      Toast.show('Session Expired','Logged out after 10 minutes of inactivity.','warning', 5000);
      setTimeout(() => this.logout(), 1800);
    }, CONFIG.INACTIVITY_MS);
  },
  setupInactivityTracking() {
    ['mousemove','keydown','mousedown','touchstart','scroll','click'].forEach(e =>
      document.addEventListener(e, () => this.resetInactivityTimer(), { passive:true })
    );
    this.resetInactivityTimer();
  },
  stopInactivityTracking() {
    if (this._inactivityTimer) clearTimeout(this._inactivityTimer);
    this._inactivityTimer = null;
  },

  // ── AUTH ──────────────────────────────────────────────────
  // NOTE: login() is synchronous so the lockout check stays fast.
  // Users are fetched from the Sheet inside showApp() (after a successful
  // login) and also every 2 minutes via the sync timer. On the very first
  // login on a new device, if the Sheet has no users yet the local SYSTEM_USERS
  // seed is used; the admin then creates real accounts which push to the Sheet,
  // and subsequent logins on all devices pick them up automatically.
  login(username, password) {
    const MAX = 5, LOCK_MS = 2 * 60 * 1000;
    const lockData = JSON.parse(localStorage.getItem(LS.LOCKOUT) || 'null');
    if (lockData) {
      const rem = lockData.until - Date.now();
      if (rem > 0) return { locked:true, seconds:Math.ceil(rem/1000) };
      localStorage.removeItem(LS.LOCKOUT);
      localStorage.removeItem(LS.ATTEMPTS);
    }

    // Read users from localStorage; treat missing OR empty array as "use defaults"
    const storedUsers = JSON.parse(localStorage.getItem(LS.USERS) || 'null');
    this.users = (storedUsers && storedUsers.length > 0) ? storedUsers : JSON.parse(JSON.stringify(SYSTEM_USERS));
    const user = this.users.find(u => u.username === username && u.password === password && u.active);

    if (!user) {
      const n = (parseInt(localStorage.getItem(LS.ATTEMPTS)||'0')) + 1;
      localStorage.setItem(LS.ATTEMPTS, n);
      if (n >= MAX) {
        localStorage.setItem(LS.LOCKOUT, JSON.stringify({ until: Date.now() + LOCK_MS }));
        localStorage.removeItem(LS.ATTEMPTS);
        this.logAudit('LOCKOUT', `Locked after ${MAX} failed attempts for: ${username}`, 'system');
        return { locked:true, seconds: LOCK_MS/1000 };
      }
      return { failed:true, attemptsLeft: MAX - n };
    }

    localStorage.removeItem(LS.ATTEMPTS);
    localStorage.removeItem(LS.LOCKOUT);
    this.currentUser = user;
    sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
    this.logAudit('LOGIN','Logged in successfully', user.username);

    // Force password change on first login
    if (user.mustChangePassword) return { success:true, mustChangePassword:true };
    return { success:true };
  },

  // ── DEMO DATA CLEAR (admin trigger) ───────────────────────
  clearDemoData() {
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]').filter(m => !m._demo);
    this.saveMembers();
    localStorage.setItem(LS.DEMO_CLEARED, '1');
    this.logAudit('CLEAR_DEMO','Administrator cleared all demo data', this.currentUser?.username || 'admin');
    Toast.show('Demo Data Cleared','All sample data has been removed. Only real records remain.','success', 5000);
    // Push the cleared flag to the Sheet so other devices skip demo data too
    if (this.isOnline && this.settings.scriptUrl) {
      this._pushSettingsToSheet().catch(() => {});
    }
  },

  // ── PASSWORD MANAGEMENT ───────────────────────────────────
  changePassword(userId, newPassword) {
    const stored = localStorage.getItem(LS.USERS);
    this.users = (stored ? JSON.parse(stored) : null) || JSON.parse(JSON.stringify(SYSTEM_USERS));
    const u = this.users.find(x => x.id === userId);
    if (!u) return false;
    u.password = newPassword;
    u.mustChangePassword = false;
    if (this.currentUser?.id === userId) {
      this.currentUser = { ...this.currentUser, password:newPassword, mustChangePassword:false };
      sessionStorage.setItem(LS.SESSION, JSON.stringify(this.currentUser));
    }
    this.saveUsers(this.users);
    this.logAudit('PASSWORD_CHANGE', `Password changed for user: ${u.username}`, u.username);
    return true;
  },

  resetPasswordToDefault(userId) {
    const stored = localStorage.getItem(LS.USERS);
    this.users = (stored ? JSON.parse(stored) : null) || JSON.parse(JSON.stringify(SYSTEM_USERS));
    const u = this.users.find(x => x.id === userId);
    if (!u) return false;
    u.password = CONFIG.DEFAULT_PASSWORD;
    u.mustChangePassword = true;
    this.saveUsers(this.users);
    this.logAudit('PASSWORD_RESET', `Password reset to default for: ${u.username}`, this.currentUser.username);
    Toast.show('Password Reset', `${u.name}'s password reset to default. They must change it on next login.`, 'success');
    return true;
  },

  logout() {
    this.logAudit('LOGOUT','User logged out', this.currentUser?.username);
    this.stopInactivityTracking();
    if (this._syncInterval) { clearInterval(this._syncInterval); this._syncInterval = null; }
    this.currentUser = null;
    sessionStorage.removeItem(LS.SESSION);
    this.showLogin();
  },

  showLogin() {
    document.getElementById('app-shell').style.display    = 'none';
    document.getElementById('login-screen').style.display = 'flex';
    this.stopInactivityTracking();
  },

  showApp() {
    // Check if admin needs to see demo-clear prompt
    const demoCleared = localStorage.getItem(LS.DEMO_CLEARED);
    if (this.currentUser?.role === 'admin' && !demoCleared) {
      const hasDemo = (JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]')).some(m => m._demo);
      if (hasDemo) {
        setTimeout(() => Modal.open('modal-demo-clear'), 800);
      }
    }

    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-shell').style.display    = 'flex';
    document.getElementById('app-shell').style.flexDirection = 'column';
    this.renderNav();
    this.renderUserHeader();
    this.setupInactivityTracking();
    this.navigate('dashboard');

    if (this.isOnline && this.settings.scriptUrl) {
      // Fetch settings AND users in parallel on every login
      Promise.all([
        this._fetchAndApplyRemoteSettings(),
        this._fetchUsersFromSheet(),
      ]).then(([settingsChanged]) => {
        if (settingsChanged) {
          this.applyAppName();
          this.renderUserHeader();
          if (this.currentPage === 'settings') PageRenderers.settings();
          Toast.show('Settings Synced','App settings updated from Google Sheets.','info',3000);
        }
        // Re-validate session against freshly fetched users
        this._revalidateSession();
        this._startSyncTimer();
      });
    } else {
      this._startSyncTimer();
    }
  },

  // After fetching users from Sheet, check the current session is still valid.
  _revalidateSession() {
    if (!this.currentUser) return;
    const fresh = this.users.find(u => u.id === this.currentUser.id || u.username === this.currentUser.username);
    if (!fresh) return; // user not in merged set — don't touch session
    // Only force logout if the account is explicitly marked inactive
    if (fresh.active === false) {
      Toast.show('Account Disabled','Your account has been deactivated. Contact your administrator.','error', 8000);
      setTimeout(() => this.logout(), 2500);
      return;
    }
    // Update session with latest non-sensitive fields (role, name, stations)
    // Never overwrite password in session from a Sheet fetch
    this.currentUser = {
      ...fresh,
      password: this.currentUser.password, // keep session password (user's own device knows it)
    };
    sessionStorage.setItem(LS.SESSION, JSON.stringify(this.currentUser));
  },

  _startSyncTimer() {
    if (this._syncInterval) clearInterval(this._syncInterval);
    if (this.settings.scriptUrl) {
      // Every 2 minutes: sync members AND users
      this._syncInterval = setInterval(() => {
        this.fetchFromSheets();
        this._fetchUsersFromSheet();
      }, 2 * 60 * 1000);
      this.fetchFromSheets(); // immediate member sync on login
    }
  },

  // ── FETCH FROM GOOGLE SHEETS (bidirectional sync) ─────────
  async fetchFromSheets() {
    if (!this.isOnline || !this.settings.scriptUrl) return;
    try {
      const url = this.settings.scriptUrl + '?action=getMembers&t=' + Date.now();
      const res = await fetch(url);
      const data = await res.json();
      if (!data?.members?.length) return;

      // Merge sheet records with local records — sheet is source of truth
      const local = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
      const localIds = new Set(local.map(m => m.id || m['Record ID']));
      const demoCleared = localStorage.getItem(LS.DEMO_CLEARED);

      let merged = [...local.filter(m => !m._demo || !demoCleared)];
      let added = 0;

      data.members.forEach(sheetRow => {
        // Normalise sheet columns to app field names
        const norm = this._normaliseSheetRow(sheetRow);
        if (!norm.id) return;
        if (!localIds.has(norm.id)) {
          merged.unshift(norm);
          added++;
        } else {
          // Update existing from sheet (sheet wins)
          const idx = merged.findIndex(m => m.id === norm.id);
          if (idx >= 0) merged[idx] = { ...merged[idx], ...norm };
        }
      });

      if (added > 0) {
        merged.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
        this.members = merged;
        this.saveMembers();
        Toast.show('Synced from Sheets', `${added} new record(s) pulled from Google Sheets.`, 'info', 3000);
        // Refresh current page
        if (['dashboard','records','my-records','reports','analytics'].includes(this.currentPage)) {
          PageRenderers[this.currentPage]?.();
        }
      }

      // Also fetch polling stations from sheet — update existing + add new
      const stationsRes = await fetch(this.settings.scriptUrl + '?action=getStations&t=' + Date.now());
      const stationsData = await stationsRes.json();
      if (stationsData?.stations?.length) {
        const local = [...this.pollingStations];
        let changed = false;
        stationsData.stations.forEach(s => {
          if (!s.code) return;
          const idx = local.findIndex(ps => ps.code === s.code);
          if (idx >= 0) {
            // Update existing — sheet wins for zone/ward/name/branch
            const before = JSON.stringify(local[idx]);
            local[idx] = { ...local[idx], ...s };
            if (JSON.stringify(local[idx]) !== before) changed = true;
          } else {
            local.push(s);
            changed = true;
          }
        });
        if (changed) {
          this.pollingStations = local;
          this.settings.pollingStations = local;
          this.saveSettings();
        }
      }
    } catch(e) {
      // Silent fail — offline will handle it
    }
  },

  _normaliseSheetRow(row) {
    // Handle both named-column objects from Apps Script and raw arrays
    const g = (keys) => {
      for (const k of keys) {
        const val = row[k];
        if (val !== undefined && val !== null && val !== '') return String(val).trim();
      }
      return '';
    };
    return {
      id:          g(['Record ID','id','ID']),
      firstName:   g(['First Name','firstName']),
      lastName:    g(['Surname','lastName','Last Name']),
      otherNames:  g(['Other Names','otherNames']),
      gender:      g(['Gender','gender']),
      zone:        g(['Zone','zone']),
      partyId:     g(['Party ID Number','partyId']),
      voterId:     g(['Voter ID Number','voterId']),
      phone:       g(['Telephone Number','phone']),
      ward:        g(['Ward Name','ward']),
      station:     g(['Polling Station','station']),
      stationCode: g(['Station Code','stationCode']),
      branch:      g(['Branch Name','branch']),
      branchCode:  g(['Branch Code','branchCode']),
      officer:     g(['Officer ID','officer']),
      officerName: g(['Officer Name','officerName']),
      timestamp:   g(['Date/Time Added','timestamp']),
      _fromSheet:  true,
    };
  },

  // ── ROLE ACCESS ───────────────────────────────────────────
  ROLE_PAGES: {
    officer: ['dashboard','entry','my-records'],
    ward:    ['dashboard','records','reports'],
    exec:    ['dashboard','records','reports','analytics'],
    admin:   ['dashboard','entry','records','reports','analytics','audit','users','settings'],
  },
  canAccess(page) { return (this.ROLE_PAGES[this.currentUser?.role]||[]).includes(page); },

  navigate(page) {
    if (!this.canAccess(page)) { Toast.show('Access Denied','You do not have permission.','error'); return; }
    this.currentPage = page;
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const el = document.getElementById('page-' + page);
    if (el) { el.classList.add('active'); PageRenderers[page]?.(); }
    document.querySelectorAll('.nav-link').forEach(a => a.classList.toggle('active', a.dataset.page === page));
    window.scrollTo(0, 0);
  },

  renderNav() {
    const nav = document.getElementById('main-nav');
    if (!nav || !this.currentUser) return;
    const pages = [
      { id:'dashboard', icon:'📊', label:'Dashboard'  },
      { id:'entry',     icon:'✍️',  label:'Data Entry'  },
      { id:'my-records',icon:'📋', label:'My Records'  },
      { id:'records',   icon:'🗃️',  label:'All Records' },
      { id:'reports',   icon:'📈', label:'Reports'     },
      { id:'analytics', icon:'🔬', label:'Analytics'   },
      { id:'audit',     icon:'🛡️',  label:'Audit Log'   },
      { id:'users',     icon:'👥', label:'User Mgmt'   },
      { id:'settings',  icon:'⚙️',  label:'Settings'    },
    ];
    nav.innerHTML = pages.filter(p => this.canAccess(p.id)).map(p =>
      `<a class="nav-link" data-page="${p.id}" onclick="App.navigate('${p.id}')">
         <span class="nav-icon">${p.icon}</span>${p.label}
       </a>`
    ).join('');
  },

  renderUserHeader() {
    const u = this.currentUser;
    if (!u) return;
    const initials = u.name.split(' ').map(n=>n[0]).slice(0,2).join('');
    const rl = {officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Executive',admin:'System Administrator'}[u.role]||u.role;
    document.getElementById('user-avatar').textContent   = initials;
    document.getElementById('user-name-hdr').textContent = u.name;
    document.getElementById('user-role-hdr').textContent = rl;
  },

  // ── MEMBER CRUD ───────────────────────────────────────────
  addMember(data) {
    // ── PRIMARY KEY CHECK: Party ID must be unique ──────────
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    if (data.partyId && data.partyId.trim()) {
      const existing = this.members.find(m =>
        m.partyId?.trim().toLowerCase() === data.partyId.trim().toLowerCase()
        && !m._demo
      );
      if (existing) {
        Toast.show(
          'Duplicate Entry Blocked',
          `Member with Party ID "${data.partyId}" is already enrolled as ${existing.firstName} ${existing.lastName} (${existing.station}).`,
          'error', 7000
        );
        this.logAudit('DUPLICATE_BLOCKED',
          `Blocked duplicate Party ID: ${data.partyId} — existing: ${existing.firstName} ${existing.lastName}`,
          this.currentUser.username
        );
        return null; // signal to caller that save was rejected
      }
    }

    const member = { id:'m'+Date.now(), ...data, officer:this.currentUser.username, officerName:this.currentUser.name, timestamp:new Date().toLocaleString('en-GH') };
    this.members.unshift(member);
    this.saveMembers();
    this.logAudit('ADD_MEMBER',`Added: ${data.firstName} ${data.lastName} (${data.partyId}) — ${data.station}`, this.currentUser.username);
    if (this.isOnline && this.settings.scriptUrl) this.syncToSheets(member);
    else { this.offlineQueue.push({type:'add',data:member}); this.saveOfflineQ(); if(!this.isOnline) Toast.show('Saved Offline','Will sync when online.','warning'); }
    return member;
  },

  // ── PERMISSION CHECK — can this user edit/delete this record? ──
  canModifyMember(m) {
    const u = this.currentUser;
    if (!u) return false;
    if (u.role === 'admin') return true;   // admin can do anything
    if (u.role === 'exec')  return false;  // exec view-only (no edit/delete)
    const codes = (u.assignedStations||[]).length ? u.assignedStations : (u.station ? [u.station] : []);
    if (u.role === 'ward') {
      // Ward coordinator can only modify records in their assigned stations/ward
      return codes.includes(m.stationCode) || m.ward === u.ward || m.branch === u.branch;
    }
    if (u.role === 'officer') {
      // Officer can only modify records they personally entered AND in their assigned stations
      return m.officer === u.username && codes.includes(m.stationCode);
    }
    return false;
  },

  updateMember(id, updates, reason) {
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    const idx = this.members.findIndex(m=>m.id===id);
    if (idx<0) return false;
    // Permission check
    if (!this.canModifyMember(this.members[idx])) {
      Toast.show('Permission Denied','You can only edit records from your assigned stations.','error');
      this.logAudit('EDIT_DENIED',`Blocked edit attempt on record ${id} by ${this.currentUser.username}`,this.currentUser.username);
      return false;
    }
    const before = {...this.members[idx]};
    this.members[idx] = {...this.members[idx],...updates, lastModified:new Date().toLocaleString('en-GH'), modifiedBy:this.currentUser.username};
    this.saveMembers();
    this.logAudit('EDIT_MEMBER',`Edited: ${before.firstName} ${before.lastName}. Reason: ${reason}`, this.currentUser.username, {before, after:this.members[idx], reason});
    if (this.isOnline && this.settings.scriptUrl) {
      this.syncToSheets({...this.members[idx], action:'updateMember'});
    }
    return true;
  },

  deleteMember(id, reason) {
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    const m = this.members.find(m=>m.id===id);
    if (!m) return false;
    // Permission check
    if (!this.canModifyMember(m)) {
      Toast.show('Permission Denied','You can only delete records from your assigned stations.','error');
      this.logAudit('DELETE_DENIED',`Blocked delete attempt on record ${id} by ${this.currentUser.username}`,this.currentUser.username);
      return false;
    }
    this.members = this.members.filter(mx=>mx.id!==id);
    this.saveMembers();
    this.logAudit('DELETE_MEMBER',`Deleted: ${m.firstName} ${m.lastName} (${m.partyId}). Reason: ${reason}`, this.currentUser.username);
    if (this.isOnline && this.settings.scriptUrl) {
      this.syncToSheets({id, action:'deleteMember', reason});
    }
    return true;
  },

  getMembersForUser() {
    const all = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    this.members = all;
    const u = this.currentUser;
    if (!u) return [];
    if (u.role==='admin'||u.role==='exec') return all;
    // Stations this user is authorised to view
    const codes = (u.assignedStations||[]).length ? u.assignedStations : (u.station ? [u.station] : []);
    if (u.role==='ward') {
      // Ward coordinators see all records in their assigned stations OR matching ward
      return all.filter(m => codes.includes(m.stationCode) || m.ward === u.ward);
    }
    if (u.role==='officer') {
      // Officers see only records from their assigned stations (not free-range across the DB)
      return all.filter(m => codes.includes(m.stationCode));
    }
    return [];
  },

  logAudit(action, details, user, extra={}) {
    const entry = {id:'a'+Date.now(),action,details,user:user||'system',timestamp:new Date().toLocaleString('en-GH'),...extra};
    this.auditLog = JSON.parse(localStorage.getItem(LS.AUDIT)||'[]');
    this.auditLog.unshift(entry);
    if (this.auditLog.length>10000) this.auditLog=this.auditLog.slice(0,10000);
    this.saveAudit();
    // Defer audit sheet sync — never block UI
    if (this.settings.scriptUrl) setTimeout(() => this._syncAuditEntry(entry), 0);
  },

  _syncAuditEntry(entry) {
    if (!this.settings.scriptUrl) return;
    this._xhrPost(this.settings.scriptUrl, {
      action:      'logAudit',
      auditAction: entry.action,
      timestamp:   entry.timestamp,
      user:        entry.user,
      details:     entry.details,
      extra:       entry.reason || '',
    });
  },

  syncToSheets(data) {
    if (!this.settings.scriptUrl) return;
    const payload = data.action ? data : {...data, action:'addMember'};
    setTimeout(() => this._xhrPost(this.settings.scriptUrl, payload), 0);
  },

  // Reliable fire-and-forget POST using XHR — handles Apps Script redirects correctly
  _xhrPost(url, data) {
    try {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true); // async
      xhr.timeout = 20000;
      xhr.send(JSON.stringify(data));
      // No response handling — fire and forget
    } catch(_) {}
  },

  getStats() {
    const all   = this.getMembersForUser();
    const today = new Date().toLocaleDateString('en-GH');
    const byStation={}, byDay={}, byGender={Male:0,Female:0,Other:0}, byZone={};
    all.forEach(m => {
      byStation[m.station] = (byStation[m.station]||0)+1;
      const g = m.gender==='Male'?'Male':m.gender==='Female'?'Female':'Other';
      byGender[g]++;
      if (m.zone) byZone[m.zone] = (byZone[m.zone]||0)+1;
    });
    for (let i=6;i>=0;i--) {
      const d=new Date(); d.setDate(d.getDate()-i);
      const key=d.toLocaleDateString('en-GH');
      byDay[key]=all.filter(m=>m.timestamp?.includes(key)).length;
    }
    return { total:all.length, today:all.filter(m=>m.timestamp?.includes(today)).length, byStation, byDay, stations:Object.keys(byStation).length, byGender, byZone };
  },

  async flushOfflineQueue() {
    if (!this.offlineQueue.length || !this.settings.scriptUrl) return;
    Toast.show('Syncing', `Uploading ${this.offlineQueue.length} record(s)…`, 'info');
    const failed = [];
    for (const item of this.offlineQueue) {
      const ok = await new Promise(resolve => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', this.settings.scriptUrl, true);
        xhr.timeout  = 15000;
        xhr.onload   = () => resolve(true);
        xhr.onerror  = () => resolve(false);
        xhr.ontimeout= () => resolve(false);
        xhr.send(JSON.stringify(item.data));
      });
      if (!ok) failed.push(item);
    }
    this.offlineQueue = failed;
    this.saveOfflineQ();
    if (!failed.length) Toast.show('Sync Complete', 'All records uploaded.', 'success');
    else Toast.show('Partial Sync', `${failed.length} record(s) still pending.`, 'warning');
  },

  // Push ALL local members to Sheet (for records entered before Sheets was configured)
  async bulkPushToSheets() {
    if (!this.settings.scriptUrl) {
      Toast.show('No Script URL', 'Configure the Apps Script URL in Settings → Google Sheets.', 'error');
      return;
    }
    const all = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]').filter(m => !m._demo);
    if (!all.length) { Toast.show('No Records', 'There are no real member records to push.', 'warning'); return; }
    Toast.show('Uploading…', `Pushing ${all.length} record(s) to Google Sheets…`, 'info', 6000);
    let ok = 0, fail = 0;
    for (const m of all) {
      const sent = await new Promise(resolve => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', this.settings.scriptUrl, true);
        xhr.timeout  = 15000;
        xhr.onload   = () => resolve(true);
        xhr.onerror  = () => resolve(false);
        xhr.ontimeout= () => resolve(false);
        xhr.send(JSON.stringify({...m, action:'addMember'}));
      });
      if (sent) ok++; else fail++;
    }
    if (fail === 0) Toast.show('Push Complete', `All ${ok} records uploaded to Google Sheets.`, 'success');
    else            Toast.show('Partial Upload', `${ok} uploaded, ${fail} failed. Try again.`, 'warning');
    this.logAudit('BULK_PUSH', `Bulk pushed ${ok}/${all.length} records to Google Sheets`, this.currentUser?.username || 'admin');
  },
};

// ─── TOAST ───────────────────────────────────────────────────
const Toast = {
  show(title, msg='', type='success', duration=4000) {
    const c = document.getElementById('toast-container');
    if (!c) return;
    const icons={success:'✅',error:'❌',warning:'⚠️',info:'ℹ️'};
    const t=document.createElement('div');
    t.className=`toast ${type}`;
    t.innerHTML=`<span class="toast-icon">${icons[type]||'✅'}</span>
      <div class="toast-content"><div class="toast-title">${title}</div>${msg?`<div class="toast-msg">${msg}</div>`:''}</div>
      <button class="toast-close" onclick="this.parentElement.remove()">✕</button>`;
    c.appendChild(t);
    setTimeout(()=>{t.classList.add('hiding');setTimeout(()=>t.remove(),300);},duration);
  }
};

const Modal = {
  open(id)   { document.getElementById(id)?.classList.add('open'); },
  close(id)  { document.getElementById(id)?.classList.remove('open'); },
  closeAll() { document.querySelectorAll('.modal-overlay').forEach(m=>m.classList.remove('open')); },
};
document.addEventListener('click',e=>{ if(e.target.classList.contains('modal-overlay')) Modal.closeAll(); });
