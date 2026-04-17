/* ============================================================
   KNNDCmdb – Core Application Logic  v2.9.8.1
   ============================================================ */
'use strict';

const CONFIG = {
  SHEET_ID:      '',
  API_KEY:       '',
  SCRIPT_URL:    '',          // ← Fill this once before deploying: paste your Apps Script Web App URL here.
                              //   Every new device will automatically inherit all settings from the Sheet.
  APP_NAME:      'Ketu North NDC Members Database',
  CONSTITUENCY:  'Ketu North',
  VERSION:       '2.9.8',
  INACTIVITY_MS: 10 * 60 * 1000,
  DEFAULT_PASSWORD: 'Ketu@2026',   // reset-to default for non-admin accounts
  ADMIN_PASSWORD:   'admin123',    // default admin password
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
  { id:'u001', username:'admin',    password:CONFIG.ADMIN_PASSWORD,   name:'System Administrator',    role:'admin',   ward:'', station:'', branch:'', assignedStations:[], active:true, mustChangePassword:false, isSystem:true },
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
    App.loadSettings();
    App.loadData();
    App.applyAppName();
    App.setupNetworkListeners();
    App.checkSession();
  },

  // ── SETTINGS ──────────────────────────────────────────────
  loadSettings() {
    const saved = JSON.parse(localStorage.getItem(LS.SETTINGS) || '{}');
    App.settings = {
      sheetId:         saved.sheetId         || CONFIG.SHEET_ID,
      apiKey:          saved.apiKey          || CONFIG.API_KEY,
      // CONFIG.SCRIPT_URL is the baked-in fallback — lets fresh devices bootstrap without manual entry
      scriptUrl:       saved.scriptUrl       || CONFIG.SCRIPT_URL,
      appName:         saved.appName         || CONFIG.APP_NAME,
      constituency:    saved.constituency    || CONFIG.CONSTITUENCY,
      pollingStations: saved.pollingStations || DEMO_POLLING_STATIONS,
    };
    App.pollingStations = App.settings.pollingStations;
  },
  saveSettings() {
    localStorage.setItem(LS.SETTINGS, JSON.stringify(App.settings));
    // Defer settings push — never block UI
    if (App.settings.scriptUrl) setTimeout(() => App._pushSettingsToSheet(), 0);
  },

  // Push settings to the Sheet's App Settings tab
  _pushSettingsToSheet() {
    if (!App.settings.scriptUrl) return;
    App._xhrPost(App.settings.scriptUrl, {
      action:          'saveSettings',
      scriptUrl:       App.settings.scriptUrl,
      appName:         App.settings.appName,
      constituency:    App.settings.constituency,
      sheetId:         App.settings.sheetId,
      demoCleared:     localStorage.getItem(LS.DEMO_CLEARED) || '',
      pollingStations: JSON.stringify(App.pollingStations || []),
      updatedBy:       App.currentUser?.username || 'system',
    });
  },

  // Fetch remote settings from the Sheet and apply them to App device
  // Called on every login; safe to call without a scriptUrl (returns early)
  async _fetchAndApplyRemoteSettings() {
    if (!App.isOnline || !App.settings.scriptUrl) return false;
    try {
      const res  = await fetch(App.settings.scriptUrl + '?action=getSettings&t=' + Date.now(), { cache: 'no-store' });
      const data = await res.json();
      if (!data?.settings || !data.exists) return false;

      const remote = data.settings;
      let changed = false;

      // Apply remote values — remote wins for everything except apiKey (never stored server-side)
      const apply = (key, remoteKey) => {
        const rv = remote[remoteKey || key];
        if (rv && rv !== App.settings[key]) {
          App.settings[key] = rv;
          changed = true;
        }
      };

      apply('scriptUrl');
      apply('appName');
      apply('constituency');
      apply('sheetId');

      // Sync pollingStations from remote settings if present and non-empty
      if (Array.isArray(remote.pollingStations) && remote.pollingStations.length) {
        const localStations  = App.settings.pollingStations || [];
        const merged         = [...localStations];
        let   stationsChanged = false;
        remote.pollingStations.forEach(s => {
          if (!s.code) return;
          const idx = merged.findIndex(ps => ps.code === s.code);
          if (idx >= 0) {
            const before = JSON.stringify(merged[idx]);
            merged[idx] = { ...merged[idx], ...s };
            if (JSON.stringify(merged[idx]) !== before) stationsChanged = true;
          } else {
            merged.push(s);
            stationsChanged = true;
          }
        });
        if (stationsChanged || !localStations.length) {
          App.settings.pollingStations = merged;
          App.pollingStations = merged;
          changed = true;
        }
      }

      // Sync demoCleared flag across devices
      if (remote.demoCleared === '1' && !localStorage.getItem(LS.DEMO_CLEARED)) {
        localStorage.setItem(LS.DEMO_CLEARED, '1');
        // Remove demo members on App device too
        const local = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
        const clean = local.filter(m => !m._demo);
        if (clean.length !== local.length) {
          localStorage.setItem(LS.MEMBERS, JSON.stringify(clean));
          App.members = clean;
        }
        changed = true;
      }

      if (changed) {
        localStorage.setItem(LS.SETTINGS, JSON.stringify(App.settings));
        App.pollingStations = App.settings.pollingStations;
        App.applyAppName();
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
      App.users = parsedUsers;
    } else {
      // No users in storage — seed with SYSTEM_USERS defaults
      App.users = JSON.parse(JSON.stringify(SYSTEM_USERS));
      localStorage.setItem(LS.USERS, JSON.stringify(App.users)); // write directly, no push to Sheet
    }

    const sm = localStorage.getItem(LS.MEMBERS);
    const demoCleared = localStorage.getItem(LS.DEMO_CLEARED);
    if (sm) {
      App.members = JSON.parse(sm);
    } else if (!demoCleared) {
      App.members = JSON.parse(JSON.stringify(DEMO_MEMBERS));
      App.saveMembers();
    } else {
      App.members = [];
      App.saveMembers();
    }

    App.auditLog     = JSON.parse(localStorage.getItem(LS.AUDIT)    || '[]');
    App.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q)|| '[]');
  },

  saveMembers() {
    try {
      localStorage.setItem(LS.MEMBERS, JSON.stringify(App.members));
    } catch(e) {
      if (e.name !== 'QuotaExceededError' && e.code !== 22) return; // unexpected error, ignore

      const role = App.currentUser?.role;
      const isFullViewRole = role === 'admin' || role === 'exec';

      if (isFullViewRole) {
        // Admin/exec: NEVER trim member records.
        // Instead, free space by clearing non-critical keys that can be rebuilt.
        try { localStorage.removeItem(LS.AUDIT);   } catch(_) {}
        try { localStorage.removeItem(LS.OFFLINE_Q);} catch(_) {}
        // Try again after freeing audit/queue space
        try {
          localStorage.setItem(LS.MEMBERS, JSON.stringify(App.members));
          return; // success — no toast needed
        } catch(_) {}
        // Still failing: the members data itself is too large for this device's quota.
        // Tell admin to use Pull All from Sheet on the next load instead of trimming.
        Toast.show(
          '⚠️ Storage Full',
          'This device\'s storage is full. Use 🔄 Pull All from Sheet on the Records page to reload data directly from Google Sheets.',
          'warning', 12000
        );
        return;
      }

      // Officer / ward coordinator: trim oldest Sheet-confirmed records
      const synced   = App.members.filter(m => m._fromSheet);
      const unsynced = App.members.filter(m => !m._fromSheet);
      let trimmed = [...unsynced, ...synced];
      let saved = false;
      while (trimmed.length > 100 && !saved) {
        trimmed = trimmed.slice(0, Math.floor(trimmed.length * 0.8));
        try { localStorage.setItem(LS.MEMBERS, JSON.stringify(trimmed)); saved = true; } catch(_) {}
      }
      if (!saved) {
        try {
          const emergency = App.members.slice(0, 500);
          localStorage.setItem(LS.MEMBERS, JSON.stringify(emergency));
          saved = true;
        } catch(_) {}
      }
      App.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
      Toast.show(
        '⚠️ Storage Full',
        'This device is running low on storage. Push your records to Google Sheets now, then the oldest synced records will be cleared automatically.',
        'warning', 10000
      );
      if (App.isOnline && App.settings.scriptUrl && App.currentUser) {
        setTimeout(() => App.pushMyRecordsToSheet(), 1000);
      }
    }
  },
  saveAudit()     { localStorage.setItem(LS.AUDIT,     JSON.stringify(App.auditLog)); },
  saveOfflineQ()  { localStorage.setItem(LS.OFFLINE_Q, JSON.stringify(App.offlineQueue)); },

  // saveUsers: write to localStorage instantly, then push to Sheet in background
  saveUsers(usersArray) {
    const toSave = Array.isArray(usersArray) ? usersArray : (App.users || []);
    // Never overwrite real users with an empty array
    if (toSave.length === 0) return;
    App.users = toSave;
    localStorage.setItem(LS.USERS, JSON.stringify(toSave));
    // Defer the network call entirely — UI never waits for it
    const url = (JSON.parse(localStorage.getItem(LS.SETTINGS) || '{}')).scriptUrl || App.settings.scriptUrl;
    if (url) setTimeout(() => App._pushUsersToSheet(toSave, url), 0);
  },

  // Background push — fire-and-forget, no-cors avoids redirect/CORS issues
  _pushUsersToSheet(usersArray, scriptUrl) {
    const url = scriptUrl || App.settings.scriptUrl;
    if (!url) return;
    App._xhrPost(url, { action: 'saveUsers', users: usersArray });
  },

  // Manual trigger — called from the ☁️ Push Users to Sheet button
  async forcePushUsersToSheet() {
    const url = App.settings.scriptUrl || JSON.parse(localStorage.getItem(LS.SETTINGS)||'{}').scriptUrl;
    if (!url) {
      Toast.show('No Script URL', 'Set the Apps Script URL in Settings → Google Sheets first.', 'error');
      return;
    }
    const users = JSON.parse(localStorage.getItem(LS.USERS) || '[]');
    if (!users.length) {
      Toast.show('No Users', 'No user accounts found in local storage.', 'warning');
      return;
    }

    const btns = document.querySelectorAll('[onclick*="forcePushUsersToSheet"]');
    btns.forEach(b => { b.disabled = true; b.dataset._orig = b.textContent; b.textContent = '⏳ Pushing…'; });

    Toast.show('Pushing Users…', `Sending ${users.length} account(s) to Google Sheets…`, 'info', 10000);

    const body = JSON.stringify({ action: 'saveUsers', users });
    const success = await new Promise((resolve) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.timeout = 30000;
      xhr.onload    = () => resolve(true);
      xhr.onerror   = () => resolve(false);
      xhr.ontimeout = () => resolve(false);
      xhr.send(body);
    });

    btns.forEach(b => { b.disabled = false; b.textContent = b.dataset._orig || '☁️ Push Users to Sheet'; });

    if (success) {
      Toast.show('Users Pushed ✅', `${users.length} account(s) sent to Google Sheets. Open the Users tab to confirm.`, 'success', 7000);
      App.logAudit('SYNC_USERS', `Pushed ${users.length} users to Google Sheets`, App.currentUser?.username || 'admin');
    } else {
      Toast.show('Push Failed', 'Apps Script did not respond. Verify the Script URL is correct and the Web App deployment is active, then try again.', 'error', 9000);
    }
  },

  // Manual trigger — called from the ☁️ Push Stations to Sheet button in Settings
  async forcePushStationsToSheet() {
    const url = App.settings.scriptUrl || JSON.parse(localStorage.getItem(LS.SETTINGS)||'{}').scriptUrl;
    if (!url) {
      Toast.show('No Script URL', 'Set the Apps Script URL in Settings → Google Sheets first.', 'error');
      return;
    }
    const stations = App.pollingStations || JSON.parse(localStorage.getItem(LS.SETTINGS)||'{}').pollingStations || [];
    if (!stations.length) {
      Toast.show('No Stations', 'No polling stations are configured locally. Add stations in Settings → Polling Stations first.', 'warning');
      return;
    }

    const btns = document.querySelectorAll('[onclick*="forcePushStationsToSheet"]');
    btns.forEach(b => { b.disabled = true; b.dataset._orig = b.textContent; b.textContent = '⏳ Pushing…'; });

    Toast.show('Pushing Stations…', `Sending ${stations.length} station(s) to Google Sheets…`, 'info', 10000);

    const body = JSON.stringify({ action: 'saveStations', stations });
    const success = await new Promise((resolve) => {
      const xhr = new XMLHttpRequest();
      xhr.open('POST', url, true);
      xhr.timeout   = 30000;
      xhr.onload    = () => resolve(true);
      xhr.onerror   = () => resolve(false);
      xhr.ontimeout = () => resolve(false);
      xhr.send(body);
    });

    btns.forEach(b => { b.disabled = false; b.textContent = b.dataset._orig || '☁️ Push Stations to Sheet'; });

    if (success) {
      Toast.show('Stations Pushed ✅', `${stations.length} station(s) sent to Google Sheets. Open the Polling Stations tab to confirm.`, 'success', 7000);
      App.logAudit('PUSH_STATIONS', `Pushed ${stations.length} configured stations to Google Sheets`, App.currentUser?.username || 'admin');
    } else {
      Toast.show('Push Failed', 'Apps Script did not respond. Verify the Script URL is correct and the Web App deployment is active, then try again.', 'error', 9000);
    }
  },

  // Pull users from the Sheet and MERGE safely into localStorage.
  // Rules:
  //  - Sheet adds new users that aren't local yet
  //  - Sheet updates non-password fields (name, role, ward, station, assignedStations, active)
  //  - Password in Sheet only wins if it is non-empty AND different from local
  //  - Local admin record is ALWAYS preserved — it is never deleted by a Sheet fetch
  async _fetchUsersFromSheet() {
    if (!App.isOnline || !App.settings.scriptUrl) return false;
    try {
      const res  = await fetch(App.settings.scriptUrl + '?action=getUsers&t=' + Date.now(), { cache: 'no-store' });
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
          // For the currently-logged-in user: keep local password (they know their own password)
          // For all other users: sheet password wins (propagates password changes across devices)
          const isCurrentUser = App.currentUser &&
            (merged[idx].id === App.currentUser.id || merged[idx].username === App.currentUser.username);
          const localPwd  = merged[idx].password;
          const sheetPwd  = sheetUser.password?.trim();
          merged[idx] = {
            ...merged[idx],
            ...sheetUser,
            password: isCurrentUser ? (localPwd || sheetPwd) : (sheetPwd || localPwd),
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
        // Don't overwrite — abort App sync to protect login
        return false;
      }

      // Only write back if we ended up with a valid set
      if (!merged.length) return false;

      localStorage.setItem(LS.USERS, JSON.stringify(merged));
      App.users = merged;
      return true;
    } catch(_) {
      return false;
    }
  },

  applyAppName() {
    const n = App.settings.appName || CONFIG.APP_NAME;
    document.querySelectorAll('.app-title-text').forEach(el => el.textContent = n);
    document.title = n;
  },

  // ── NETWORK ───────────────────────────────────────────────
  setupNetworkListeners() {
    window.addEventListener('online', () => {
      App.isOnline = true;
      App.updateOnlineStatus();
      // Reload queue from localStorage — in-memory array may be stale after page reload
      App.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q) || '[]');
      if (App.offlineQueue.length) App.flushOfflineQueue();
    });
    window.addEventListener('offline', () => { App.isOnline = false; App.updateOnlineStatus(); });
    App.updateOnlineStatus();
  },
  updateOnlineStatus() {
    document.getElementById('offline-banner')?.classList.toggle('show', !App.isOnline);
    const dot = document.getElementById('conn-dot');
    if (dot) dot.className = App.isOnline ? 'online-dot' : 'offline-dot';
  },

  // ── SESSION / INACTIVITY ──────────────────────────────────
  checkSession() {
    const s = sessionStorage.getItem(LS.SESSION);
    if (s) {
      try { App.currentUser = JSON.parse(s); App.showApp(); return; } catch(_) {}
    }
    App.showLogin();
  },

  resetInactivityTimer() {
    if (App._inactivityTimer) clearTimeout(App._inactivityTimer);
    App._inactivityTimer = setTimeout(() => {
      if (!App.currentUser) return;
      App.logAudit('AUTO_LOGOUT','Session expired — 10 min inactivity', App.currentUser.username);
      Toast.show('Session Expired','Logged out after 10 minutes of inactivity.','warning', 5000);
      setTimeout(() => App.logout(), 1800);
    }, CONFIG.INACTIVITY_MS);
  },
  setupInactivityTracking() {
    ['mousemove','keydown','mousedown','touchstart','scroll','click'].forEach(e =>
      document.addEventListener(e, () => App.resetInactivityTimer(), { passive:true })
    );
    App.resetInactivityTimer();
  },
  stopInactivityTracking() {
    if (App._inactivityTimer) clearTimeout(App._inactivityTimer);
    App._inactivityTimer = null;
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
    App.users = (storedUsers && storedUsers.length > 0) ? storedUsers : JSON.parse(JSON.stringify(SYSTEM_USERS));
    const user = App.users.find(u => u.username === username && u.password === password && u.active);

    if (!user) {
      const n = (parseInt(localStorage.getItem(LS.ATTEMPTS)||'0')) + 1;
      localStorage.setItem(LS.ATTEMPTS, n);
      if (n >= MAX) {
        localStorage.setItem(LS.LOCKOUT, JSON.stringify({ until: Date.now() + LOCK_MS }));
        localStorage.removeItem(LS.ATTEMPTS);
        App.logAudit('LOCKOUT', `Locked after ${MAX} failed attempts for: ${username}`, 'system');
        return { locked:true, seconds: LOCK_MS/1000 };
      }
      return { failed:true, attemptsLeft: MAX - n };
    }

    localStorage.removeItem(LS.ATTEMPTS);
    localStorage.removeItem(LS.LOCKOUT);
    App.currentUser = user;
    sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
    App.logAudit('LOGIN','Logged in successfully', user.username);

    // Force password change on first login
    if (user.mustChangePassword) return { success:true, mustChangePassword:true };
    return { success:true };
  },

  // ── DEMO DATA CLEAR (admin trigger) ───────────────────────
  clearDemoData() {
    App.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]').filter(m => !m._demo);
    App.saveMembers();
    localStorage.setItem(LS.DEMO_CLEARED, '1');
    App.logAudit('CLEAR_DEMO','Administrator cleared all demo data', App.currentUser?.username || 'admin');
    Toast.show('Demo Data Cleared','All sample data has been removed. Only real records remain.','success', 5000);
    // Push the cleared flag to the Sheet so other devices skip demo data too
    if (App.isOnline && App.settings.scriptUrl) {
      App._pushSettingsToSheet().catch(() => {});
    }
  },

  // ── PASSWORD MANAGEMENT ───────────────────────────────────
  changePassword(userId, newPassword) {
    const stored = localStorage.getItem(LS.USERS);
    App.users = (stored ? JSON.parse(stored) : null) || JSON.parse(JSON.stringify(SYSTEM_USERS));
    const u = App.users.find(x => x.id === userId);
    if (!u) return false;
    u.password = newPassword;
    u.mustChangePassword = false;
    if (App.currentUser?.id === userId) {
      App.currentUser = { ...App.currentUser, password:newPassword, mustChangePassword:false };
      sessionStorage.setItem(LS.SESSION, JSON.stringify(App.currentUser));
    }
    App.saveUsers(App.users);
    App.logAudit('PASSWORD_CHANGE', `Password changed for user: ${u.username}`, u.username);
    return true;
  },

  // resetPasswordToDefault: admin account resets to ADMIN_PASSWORD; others reset to DEFAULT_PASSWORD
  resetPasswordToDefault(userId) {
    const stored = localStorage.getItem(LS.USERS);
    App.users = (stored ? JSON.parse(stored) : null) || JSON.parse(JSON.stringify(SYSTEM_USERS));
    const u = App.users.find(x => x.id === userId);
    if (!u) return false;
    u.password = (u.role === 'admin') ? CONFIG.ADMIN_PASSWORD : CONFIG.DEFAULT_PASSWORD;
    u.mustChangePassword = true;
    App.saveUsers(App.users);
    App.logAudit('PASSWORD_RESET', `Password reset to default for: ${u.username}`, App.currentUser.username);
    Toast.show('Password Reset', `${u.name}'s password reset to default. They must change it on next login.`, 'success');
    return true;
  },

  logout() {
    App.logAudit('LOGOUT','User logged out', App.currentUser?.username);
    App.stopInactivityTracking();
    if (App._syncInterval) { clearInterval(App._syncInterval); App._syncInterval = null; }
    App.currentUser = null;
    sessionStorage.removeItem(LS.SESSION);
    App.showLogin();
  },

  showLogin() {
    document.getElementById('app-shell').style.display    = 'none';
    document.getElementById('login-screen').style.display = 'flex';
    App.stopInactivityTracking();
  },

  showApp() {
    // Check if admin needs to see demo-clear prompt
    const demoCleared = localStorage.getItem(LS.DEMO_CLEARED);
    if (App.currentUser?.role === 'admin' && !demoCleared) {
      const hasDemo = (JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]')).some(m => m._demo);
      if (hasDemo) {
        setTimeout(() => Modal.open('modal-demo-clear'), 800);
      }
    }

    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-shell').style.display    = 'flex';
    document.getElementById('app-shell').style.flexDirection = 'column';
    App.renderNav();
    App.renderUserHeader();
    App.setupInactivityTracking();
    App.navigate('dashboard');

    // Flush any queued offline records now that the user is logged in
    // Reload from localStorage first — in-memory array may be stale after a page reload
    App.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q) || '[]');
    if (App.isOnline && App.settings.scriptUrl && App.offlineQueue.length) {
      setTimeout(() => App.flushOfflineQueue(), 1500); // slight delay so UI settles first
    }

    if (App.isOnline && App.settings.scriptUrl) {
      // Fetch settings AND users in parallel on every login
      Promise.all([
        App._fetchAndApplyRemoteSettings(),
        App._fetchUsersFromSheet(),
      ]).then(([settingsChanged]) => {
        if (settingsChanged) {
          App.applyAppName();
          App.renderUserHeader();
          if (App.currentPage === 'settings') PageRenderers.settings();
          Toast.show('Settings Synced','App settings updated from Google Sheets.','info',3000);
        }
        // Re-validate session against freshly fetched users
        App._revalidateSession();
        App._startSyncTimer();
      });
    } else {
      App._startSyncTimer();
    }
  },

  // After fetching users from Sheet, check the current session is still valid.
  // Propagates any assignment changes (ward, stations, branch, role) to the live session
  // without requiring a re-login.
  _revalidateSession() {
    if (!App.currentUser) return;
    const fresh = App.users.find(u => u.id === App.currentUser.id || u.username === App.currentUser.username);
    if (!fresh) return; // user not in merged set — don't touch session

    // Force logout if the account has been deactivated
    if (fresh.active === false) {
      Toast.show('Account Disabled','Your account has been deactivated. Contact your administrator.','error', 8000);
      setTimeout(() => App.logout(), 2500);
      return;
    }

    // Snapshot fields that affect data access and UI before overwrite
    const prev = {
      stations: JSON.stringify(App.currentUser.assignedStations || []),
      ward:     App.currentUser.ward     || '',
      station:  App.currentUser.station  || '',
      branch:   App.currentUser.branch   || '',
      role:     App.currentUser.role     || '',
    };

    // Update session — never overwrite password from a Sheet fetch
    App.currentUser = {
      ...fresh,
      password: App.currentUser.password,
    };
    sessionStorage.setItem(LS.SESSION, JSON.stringify(App.currentUser));

    // Detect what changed
    const now = {
      stations: JSON.stringify(App.currentUser.assignedStations || []),
      ward:     App.currentUser.ward     || '',
      station:  App.currentUser.station  || '',
      branch:   App.currentUser.branch   || '',
      role:     App.currentUser.role     || '',
    };

    const stationsChanged = prev.stations !== now.stations;
    const wardChanged     = prev.ward     !== now.ward;
    const branchChanged   = prev.branch   !== now.branch;
    const roleChanged     = prev.role     !== now.role;
    const anyChanged      = stationsChanged || wardChanged || branchChanged || roleChanged;

    if (!anyChanged) return;

    // Re-render nav (role may have changed) and header (ward/role display)
    App.renderNav();
    App.renderUserHeader();

    // Re-render the active page if it uses assignment data
    const pagesAffected = ['entry','my-records','records','dashboard','reports','analytics'];
    if (pagesAffected.includes(App.currentPage)) {
      PageRenderers[App.currentPage]?.();
    }

    // Build a meaningful toast message listing what changed
    const parts = [];
    if (stationsChanged) {
      const count = (App.currentUser.assignedStations || []).length;
      parts.push(`${count} polling station${count !== 1 ? 's' : ''}`);
    }
    if (wardChanged && now.ward)   parts.push(`ward: ${now.ward}`);
    if (branchChanged && now.branch) parts.push(`branch: ${now.branch}`);
    if (roleChanged)               parts.push(`role: ${now.role}`);

    if (parts.length) {
      Toast.show(
        'Assignments Updated',
        `Your account has been updated — ${parts.join(', ')}.`,
        'info', 5000
      );
    }
  },

  _startSyncTimer() {
    if (App._syncInterval) clearInterval(App._syncInterval);
    if (App.settings.scriptUrl) {
      // Every 2 minutes: sync members AND users; re-validate session after user fetch
      App._syncInterval = setInterval(async () => {
        App.fetchFromSheets();
        const changed = await App._fetchUsersFromSheet();
        // Always re-validate so live assignment changes (ward, stations, role)
        // are applied to the active session without requiring a re-login
        if (changed !== false) App._revalidateSession();
      }, 2 * 60 * 1000);
      App.fetchFromSheets(); // immediate member sync on login
    }
  },

  // ── FETCH FROM GOOGLE SHEETS (bidirectional sync) ─────────
  async fetchFromSheets() {
    if (!App.isOnline || !App.settings.scriptUrl) return;
    try {
      const url = App.settings.scriptUrl + '?action=getMembers&t=' + Date.now();
      const res = await fetch(url, { cache: 'no-store' });
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
        const norm = App._normaliseSheetRow(sheetRow);
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

      const localCount = local.filter(m => !m._demo || !demoCleared).length;

      if (added > 0 || merged.length > localCount) {
        merged.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
        App.members = merged;
        App.saveMembers();
        if (added > 0) {
          Toast.show('Synced from Sheets', `${added} new record(s) pulled from Google Sheets.`, 'info', 3000);
        }
        // Refresh current page
        if (['dashboard','records','my-records','reports','analytics'].includes(App.currentPage)) {
          PageRenderers[App.currentPage]?.();
        }
      }

      // ── Auto-trim: mark Sheet-confirmed records and free local storage ──────
      // After a successful sync, records confirmed in the Sheet are safe server-side.
      // Mark them as _fromSheet=true. If local storage is >80% full, drop the oldest
      // Sheet-sourced records to free space for new entries — they remain in the Sheet.
      App._autoTrimMembers(data.members);

      // Also fetch polling stations from sheet — update existing + add new
      const stationsRes = await fetch(App.settings.scriptUrl + '?action=getStations&t=' + Date.now(), { cache: 'no-store' });
      const stationsData = await stationsRes.json();
      if (stationsData?.stations?.length) {
        const local = [...App.pollingStations];
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
          App.pollingStations = local;
          App.settings.pollingStations = local;
          App.saveSettings();
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
    officer: ['dashboard','entry','my-records','records'],
    ward:    ['dashboard','records','reports'],
    exec:    ['dashboard','records','reports','analytics'],
    admin:   ['dashboard','entry','records','reports','analytics','audit','users','settings'],
  },
  canAccess(page) { return (App.ROLE_PAGES[App.currentUser?.role]||[]).includes(page); },

  navigate(page) {
    if (!App.canAccess(page)) { Toast.show('Access Denied','You do not have permission.','error'); return; }
    App.currentPage = page;
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const el = document.getElementById('page-' + page);
    if (el) { el.classList.add('active'); PageRenderers[page]?.(); }
    document.querySelectorAll('.nav-link').forEach(a => a.classList.toggle('active', a.dataset.page === page));
    window.scrollTo(0, 0);
  },

  renderNav() {
    const nav = document.getElementById('main-nav');
    if (!nav || !App.currentUser) return;
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
    nav.innerHTML = pages.filter(p => App.canAccess(p.id)).map(p =>
      `<a class="nav-link" data-page="${p.id}" onclick="App.navigate('${p.id}')">
         <span class="nav-icon">${p.icon}</span>${p.label}
       </a>`
    ).join('');
  },

  renderUserHeader() {
    const u = App.currentUser;
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
    App.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    if (data.partyId && data.partyId.trim()) {
      const existing = App.members.find(m =>
        m.partyId?.trim().toLowerCase() === data.partyId.trim().toLowerCase()
        && !m._demo
      );
      if (existing) {
        Toast.show(
          'Duplicate Entry Blocked',
          `Member with Party ID "${data.partyId}" is already enrolled as ${existing.firstName} ${existing.lastName} (${existing.station}).`,
          'error', 7000
        );
        App.logAudit('DUPLICATE_BLOCKED',
          `Blocked duplicate Party ID: ${data.partyId} — existing: ${existing.firstName} ${existing.lastName}`,
          App.currentUser.username
        );
        return null; // signal to caller that save was rejected
      }
    }

    const member = { id:'m'+Date.now(), ...data, officer:App.currentUser.username, officerName:App.currentUser.name, timestamp:new Date().toLocaleString('en-GH'), isoDate:App._todayISO() };
    App.members.unshift(member);
    App.saveMembers();
    App.logAudit('ADD_MEMBER',`Added: ${data.firstName} ${data.lastName} (${data.partyId}) — ${data.station}`, App.currentUser.username);
    if (App.isOnline && App.settings.scriptUrl) App.syncToSheets(member);
    else { App.offlineQueue.push({type:'add',data:member}); App.saveOfflineQ(); if(!App.isOnline) Toast.show('Saved Offline','Will sync when online.','warning'); }
    return member;
  },

  // ── PERMISSION CHECK — can App user edit/delete App record? ──
  // canEditMember — can this user EDIT this record?
  // exec (Constituency Executive) can edit any record; ward/officer are station-scoped.
  canEditMember(m) {
    const u = App.currentUser;
    if (!u) return false;
    if (u.role === 'admin') return true;
    if (u.role === 'exec')  return true;   // exec can edit (but not delete)
    const codes = (u.assignedStations||[]).length ? u.assignedStations : (u.station ? [u.station] : []);
    if (u.role === 'ward') {
      return codes.includes(m.stationCode) || m.ward === u.ward || m.branch === u.branch;
    }
    if (u.role === 'officer') {
      return m.officer === u.username && codes.includes(m.stationCode);
    }
    return false;
  },

  // canDeleteMember — can this user DELETE this record?
  // exec is explicitly blocked from deletion (view + edit only).
  canDeleteMember(m) {
    const u = App.currentUser;
    if (!u) return false;
    if (u.role === 'admin') return true;
    if (u.role === 'exec')  return false;  // exec cannot delete
    const codes = (u.assignedStations||[]).length ? u.assignedStations : (u.station ? [u.station] : []);
    if (u.role === 'ward') {
      return codes.includes(m.stationCode) || m.ward === u.ward || m.branch === u.branch;
    }
    if (u.role === 'officer') {
      return m.officer === u.username && codes.includes(m.stationCode);
    }
    return false;
  },

  // canModifyMember — legacy alias used by deleteMember(); maps to canDeleteMember
  canModifyMember(m) { return App.canDeleteMember(m); },

  updateMember(id, updates, reason) {
    App.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    const idx = App.members.findIndex(m=>m.id===id);
    if (idx<0) return false;
    // Permission check — use canEditMember (exec can edit; officer scoped to own records)
    if (!App.canEditMember(App.members[idx])) {
      Toast.show('Permission Denied','You can only edit records from your assigned stations.','error');
      App.logAudit('EDIT_DENIED',`Blocked edit attempt on record ${id} by ${App.currentUser.username}`,App.currentUser.username);
      return false;
    }
    const before = {...App.members[idx]};
    App.members[idx] = {...App.members[idx],...updates, lastModified:new Date().toLocaleString('en-GH'), modifiedBy:App.currentUser.username};
    App.saveMembers();
    App.logAudit('EDIT_MEMBER',`Edited: ${before.firstName} ${before.lastName}. Reason: ${reason}`, App.currentUser.username, {before, after:App.members[idx], reason});
    if (App.isOnline && App.settings.scriptUrl) {
      App.syncToSheets({...App.members[idx], action:'updateMember'});
    } else if (App.settings.scriptUrl) {
      // Queue for sync when back online
      App.offlineQueue.push({ type:'update', data:{...App.members[idx], action:'updateMember', reason} });
      App.saveOfflineQ();
      if (!App.isOnline) Toast.show('Edit Queued','Changes saved locally and will sync when online.','warning');
    }
    return true;
  },

  deleteMember(id, reason) {
    App.members = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    const m = App.members.find(m=>m.id===id);
    if (!m) return false;
    // Permission check
    if (!App.canModifyMember(m)) {
      Toast.show('Permission Denied','You can only delete records from your assigned stations.','error');
      App.logAudit('DELETE_DENIED',`Blocked delete attempt on record ${id} by ${App.currentUser.username}`,App.currentUser.username);
      return false;
    }
    App.members = App.members.filter(mx=>mx.id!==id);
    App.saveMembers();
    App.logAudit('DELETE_MEMBER',`Deleted: ${m.firstName} ${m.lastName} (${m.partyId}). Reason: ${reason}`, App.currentUser.username);
    if (App.isOnline && App.settings.scriptUrl) {
      App.syncToSheets({id, action:'deleteMember', reason});
    } else if (App.settings.scriptUrl) {
      // Queue for sync when back online
      App.offlineQueue.push({ type:'delete', data:{id, action:'deleteMember', reason} });
      App.saveOfflineQ();
      if (!App.isOnline) Toast.show('Delete Queued','Record removed locally and will sync when online.','warning');
    }
    return true;
  },

  getMembersForUser() {
    const all = JSON.parse(localStorage.getItem(LS.MEMBERS)||'[]');
    App.members = all;
    const u = App.currentUser;
    if (!u) return [];
    if (u.role==='admin'||u.role==='exec') return all;
    // Stations App user is authorised to view
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
    App.auditLog = JSON.parse(localStorage.getItem(LS.AUDIT)||'[]');
    App.auditLog.unshift(entry);
    if (App.auditLog.length>10000) App.auditLog=App.auditLog.slice(0,10000);
    localStorage.setItem(LS.AUDIT, JSON.stringify(App.auditLog));
    // Defer audit sheet sync — never block UI
    if (App.settings.scriptUrl) setTimeout(() => App._syncAuditEntry(entry), 0);
  },

  _syncAuditEntry(entry) {
    if (!App.settings.scriptUrl) return;
    App._xhrPost(App.settings.scriptUrl, {
      action:      'logAudit',
      auditAction: entry.action,
      timestamp:   entry.timestamp,
      user:        entry.user,
      details:     entry.details,
      extra:       entry.reason || '',
    });
  },

  // Pull audit entries from the Sheet and merge with local log.
  // Called when admin navigates to the Audit Log page and Sheet is configured.
  // Returns true if new entries were found, false otherwise.
  async _fetchAuditFromSheet() {
    if (!App.isOnline || !App.settings.scriptUrl) return false;
    try {
      const res  = await fetch(App.settings.scriptUrl + '?action=getAudit&t=' + Date.now(), { cache: 'no-store' });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch(_) { return false; }
      if (!data?.entries?.length) return false;

      // Build a dedup key: timestamp + action + user (Sheet has no id field)
      const local  = JSON.parse(localStorage.getItem(LS.AUDIT) || '[]');
      const seenKeys = new Set(local.map(e => `${e.timestamp}|${e.action}|${e.user}`));

      const newEntries = data.entries
        .filter(e => {
          const k = `${e.timestamp}|${e.action}|${e.user}`;
          return !seenKeys.has(k);
        })
        .map(e => ({
          id:        'a_sheet_' + Math.random().toString(36).slice(2),
          action:    e.action,
          details:   e.details,
          user:      e.user,
          timestamp: e.timestamp,
          reason:    e.extra || '',
          _fromSheet: true,
        }));

      if (!newEntries.length) return false;

      // Merge: combine local + sheet entries, sort newest-first, cap at 10 000
      const merged = [...local, ...newEntries]
        .sort((a, b) => {
          // Sort by id timestamp where available, else lexicographic on timestamp string
          const ta = a.id?.startsWith('a') && !a._fromSheet ? parseInt(a.id.slice(1)) : 0;
          const tb = b.id?.startsWith('a') && !b._fromSheet ? parseInt(b.id.slice(1)) : 0;
          return tb - ta || b.timestamp.localeCompare(a.timestamp);
        })
        .slice(0, 10000);

      App.auditLog = merged;
      localStorage.setItem(LS.AUDIT, JSON.stringify(merged));
      return true;
    } catch(_) {
      return false;
    }
  },

  syncToSheets(data) {
    if (!App.settings.scriptUrl) return;
    // Use upsertMember as default: if record already exists in Sheet it updates in-place,
    // preventing duplicate rows when records sync more than once (e.g. after offline/online cycle).
    const payload = data.action ? data : {...data, action:'upsertMember'};
    setTimeout(() => App._xhrPost(App.settings.scriptUrl, payload), 0);
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

  // Force a complete re-pull from Sheet for admin/exec.
  // Clears the local member cache first so fetchFromSheets re-downloads everything,
  // not just the delta. Used when admin notices local count < Sheet count.
  async _pullAllFromSheet() {
    if (!App.settings.scriptUrl) {
      Toast.show('No Script URL', 'Configure the Apps Script URL in Settings → Google Sheets.', 'error');
      return;
    }
    const btn = document.getElementById('records-pull-btn');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Pulling…'; }
    Toast.show('Pulling from Sheet', 'Downloading all records from Google Sheets…', 'info', 8000);

    try {
      const url  = App.settings.scriptUrl + '?action=getMembers&t=' + Date.now();
      const res  = await fetch(url, { cache: 'no-store' });
      const data = await res.json();

      if (!data?.members?.length) {
        Toast.show('No Records', 'The Google Sheet returned no records. Check the Sheet has data.', 'warning');
        if (btn) { btn.disabled = false; btn.textContent = '🔄 Pull All from Sheet'; }
        return;
      }

      // Keep any locally-entered records that haven't been pushed to Sheet yet
      const local       = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
      const sheetIds    = new Set(data.members.map(m => m.id).filter(Boolean));
      const localOnly   = local.filter(m => m.id && !sheetIds.has(m.id) && !m._demo);

      // Build fresh set: all Sheet records + any local-only unsynced records
      const sheetNormed = data.members.map(r => ({ ...App._normaliseSheetRow(r), _fromSheet: true }));
      const merged      = [...localOnly, ...sheetNormed]
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

      // Free up space from non-critical keys before writing the full member set
      try { localStorage.removeItem(LS.AUDIT);    } catch(_) {}
      try { localStorage.removeItem(LS.OFFLINE_Q); } catch(_) {}

      App.members = merged;
      App.saveMembers(); // uses quota-aware handler

      Toast.show('Pull Complete ✅',
        `${data.members.length} record(s) loaded from Google Sheets${localOnly.length ? ` + ${localOnly.length} local unsynced` : ''}.`,
        'success', 6000);
      App.logAudit('PULL_ALL_SHEET', `Admin force-pulled ${data.members.length} records from Google Sheets`, App.currentUser?.username);

      // Refresh all relevant pages
      if (['records','dashboard','reports','analytics'].includes(App.currentPage)) {
        PageRenderers[App.currentPage]?.();
      }
    } catch(e) {
      Toast.show('Pull Failed', 'Could not reach Google Sheets. Check your connection and Script URL.', 'error');
    }

    if (btn) { btn.disabled = false; btn.textContent = '🔄 Pull All from Sheet'; }
  },
  // Called after every Sheet sync. Marks records confirmed in the Sheet as
  // _fromSheet:true, then checks storage usage on officer/ward devices only.
  // Admin and exec devices are NEVER trimmed — they need the full dataset visible.
  // Only locally-entered unsynced records are always preserved.
  _autoTrimMembers(sheetMembers) {
    if (!sheetMembers?.length) return;

    // Never trim admin or exec — they need full visibility
    const role = App.currentUser?.role;
    const isFullViewRole = role === 'admin' || role === 'exec';

    const sheetIds = new Set(sheetMembers.map(m => m.id).filter(Boolean));
    const local    = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');

    // Mark confirmed records as _fromSheet
    let changed = false;
    local.forEach(m => {
      if (sheetIds.has(m.id) && !m._fromSheet) {
        m._fromSheet = true;
        changed = true;
      }
    });
    if (changed) {
      App.members = local;
      try { localStorage.setItem(LS.MEMBERS, JSON.stringify(local)); } catch(_) {}
    }

    // Admin and exec: stop here — no trimming ever
    if (isFullViewRole) return;

    // Officer/ward: check storage and trim oldest synced records if needed
    const raw = localStorage.getItem(LS.MEMBERS) || '[]';
    const sizeKB = (raw.length * 2) / 1024;
    const THRESHOLD_KB = 3500;

    if (sizeKB < THRESHOLD_KB) return;

    const unsynced = local.filter(m => !m._fromSheet);
    const synced   = local.filter(m =>  m._fromSheet);

    synced.sort((a, b) => {
      const da = a.isoDate || App._isoDate(a.timestamp) || '0000';
      const db = b.isoDate || App._isoDate(b.timestamp) || '0000';
      return da.localeCompare(db);
    });

    let trimmed = [...unsynced, ...synced];
    while (trimmed.length > unsynced.length) {
      trimmed.shift();
      const testSize = (JSON.stringify(trimmed).length * 2) / 1024;
      if (testSize < THRESHOLD_KB) break;
    }

    if (trimmed.length < local.length) {
      const freed = local.length - trimmed.length;
      App.members = trimmed;
      try {
        localStorage.setItem(LS.MEMBERS, JSON.stringify(trimmed));
        App.logAudit('STORAGE_TRIM', `Auto-trimmed ${freed} synced record(s) from local storage to free space. All records remain in Google Sheets.`, 'system');
      } catch(_) {}
    }
  },
  // Returns 'YYYY-MM-DD' from any stored timestamp string, locale-independent.
  // Handles en-GH 'D/M/YYYY, ...' and en-US 'M/D/YYYY, ...' fallback formats.
  _isoDate(timestamp) {
    if (!timestamp) return '';
    // If already ISO format
    if (/^\d{4}-\d{2}-\d{2}/.test(timestamp)) return timestamp.slice(0, 10);
    // Try native Date parse first (works for M/D/YYYY = en-US)
    const d = new Date(timestamp);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
    // Manual parse for D/M/YYYY (en-GH): extract day, month, year
    const m = timestamp.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      const [, day, month, year] = m;
      const forced = new Date(`${year}-${month.padStart(2,'0')}-${day.padStart(2,'0')}`);
      if (!isNaN(forced)) return forced.toISOString().slice(0, 10);
    }
    return '';
  },

  // Returns today's date as 'YYYY-MM-DD' in local time (immune to locale and timezone drift)
  _todayISO() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  },

  // Returns 'YYYY-MM-DD' for a date N days ago
  _daysAgoISO(n) {
    const d = new Date();
    d.setDate(d.getDate() - n);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  },

  getStats() {
    const all   = App.getMembersForUser();
    const todayISO = App._todayISO();
    const byStation={}, byDay={}, byGender={Male:0,Female:0,Other:0}, byZone={};
    all.forEach(m => {
      byStation[m.station] = (byStation[m.station]||0)+1;
      const g = m.gender==='Male'?'Male':m.gender==='Female'?'Female':'Other';
      byGender[g]++;
      if (m.zone) byZone[m.zone] = (byZone[m.zone]||0)+1;
    });
    // Build 7-day trend using ISO date keys — locale-independent
    for (let i=6;i>=0;i--) {
      const key = App._daysAgoISO(i);
      byDay[key] = all.filter(m => (m.isoDate || App._isoDate(m.timestamp)) === key).length;
    }
    const todayCount = all.filter(m => (m.isoDate || App._isoDate(m.timestamp)) === todayISO).length;
    return { total:all.length, today:todayCount, byStation, byDay, stations:Object.keys(byStation).length, byGender, byZone };
  },

  async flushOfflineQueue() {
    // Always reload from localStorage — in-memory array may be stale after a page reload
    App.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q) || '[]');
    if (!App.offlineQueue.length || !App.settings.scriptUrl) return;

    Toast.show('Syncing', `Uploading ${App.offlineQueue.length} queued operation(s)…`, 'info', 6000);
    const failed = [];
    let uploaded = 0;

    for (const item of App.offlineQueue) {
      // Route each item to the correct GAS action based on its type
      let payload;
      if (item.type === 'delete') {
        payload = { action: 'deleteMember', ...item.data };
      } else if (item.type === 'update') {
        payload = { action: 'updateMember', ...item.data };
      } else {
        // 'add' or legacy items without type — use upsert (idempotent, safe to retry)
        payload = { action: 'upsertMember', ...item.data };
      }

      const ok = await new Promise(resolve => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', App.settings.scriptUrl, true);
        xhr.timeout  = 20000;
        xhr.onload   = () => resolve(true);
        xhr.onerror  = () => resolve(false);
        xhr.ontimeout= () => resolve(false);
        xhr.send(JSON.stringify(payload));
      });

      if (ok) { uploaded++; } else { failed.push(item); }
    }

    App.offlineQueue = failed;
    App.saveOfflineQ();

    if (!failed.length) {
      Toast.show('Sync Complete ✅', `${uploaded} queued operation(s) uploaded to Google Sheets.`, 'success', 5000);
    } else {
      Toast.show('Partial Sync', `${uploaded} uploaded, ${failed.length} still pending. Will retry when online.`, 'warning', 6000);
    }

    // Refresh dashboard so the Offline Queue counter updates to 0
    if (App.currentUser && App.currentPage === 'dashboard') PageRenderers.dashboard();
  },

  // Push ALL local members to Sheet (for records entered before Sheets was configured)
  // Uses upsertMember — safe to run multiple times: existing rows are updated, new rows appended.
  async bulkPushToSheets() {
    if (!App.settings.scriptUrl) {
      Toast.show('No Script URL', 'Configure the Apps Script URL in Settings → Google Sheets.', 'error');
      return;
    }
    const all = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]').filter(m => !m._demo);
    if (!all.length) { Toast.show('No Records', 'There are no real member records to push.', 'warning'); return; }

    // Lock button to prevent double-push
    const btn = document.querySelector('[onclick*="bulkPushToSheets"]');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Pushing…'; }

    Toast.show('Uploading…', `Pushing ${all.length} record(s) to Google Sheets…`, 'info', 10000);
    let ok = 0, fail = 0;
    for (const m of all) {
      const sent = await new Promise(resolve => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', App.settings.scriptUrl, true);
        xhr.timeout  = 20000;
        xhr.onload   = () => resolve(true);
        xhr.onerror  = () => resolve(false);
        xhr.ontimeout= () => resolve(false);
        xhr.send(JSON.stringify({ action: 'upsertMember', ...m }));
      });
      if (sent) ok++; else fail++;
    }

    // Restore button
    if (btn) { btn.disabled = false; btn.textContent = '☁️ Push All Records to Sheet'; }

    if (fail === 0) Toast.show('Push Complete ✅', `All ${ok} record(s) uploaded to Google Sheets.`, 'success', 6000);
    else            Toast.show('Partial Upload', `${ok} uploaded, ${fail} failed. Try again.`, 'warning', 6000);
    App.logAudit('BULK_PUSH', `Bulk pushed ${ok}/${all.length} records to Google Sheets`, App.currentUser?.username || 'admin');
  },

  // Push records visible to the current user to Google Sheet.
  // Used from the Data Entry page — scoped by role and assigned stations.
  // Officers push only their station's records; admin/ward push all they can see.
  async pushMyRecordsToSheet() {
    if (!App.settings.scriptUrl) {
      Toast.show('No Script URL', 'Configure the Apps Script URL in Settings → Google Sheets first.', 'error');
      return;
    }
    const records = App.getMembersForUser().filter(m => !m._demo);
    if (!records.length) {
      Toast.show('No Records', 'No records found for your assigned station(s).', 'warning');
      return;
    }

    // Lock all entry-page push buttons during the operation
    const btns = document.querySelectorAll('[onclick*="pushMyRecordsToSheet"]');
    btns.forEach(b => { b.disabled = true; b.dataset._orig = b.textContent; b.textContent = '⏳ Pushing…'; });

    Toast.show('Pushing…', `Sending ${records.length} record(s) to Google Sheets…`, 'info', 10000);
    let ok = 0, fail = 0;
    for (const m of records) {
      const sent = await new Promise(resolve => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', App.settings.scriptUrl, true);
        xhr.timeout  = 20000;
        xhr.onload   = () => resolve(true);
        xhr.onerror  = () => resolve(false);
        xhr.ontimeout= () => resolve(false);
        xhr.send(JSON.stringify({ action: 'upsertMember', ...m }));
      });
      if (sent) ok++; else fail++;
    }

    btns.forEach(b => { b.disabled = false; b.textContent = b.dataset._orig || '☁️ Push Records to Sheet'; });

    if (fail === 0) {
      Toast.show('Push Complete ✅', `All ${ok} record(s) uploaded to Google Sheets.`, 'success', 6000);
      App.logAudit('PUSH_MY_RECORDS', `Pushed ${ok} record(s) to Google Sheets from Data Entry page`, App.currentUser?.username);
    } else {
      Toast.show('Partial Upload', `${ok} uploaded, ${fail} failed. Check your connection and try again.`, 'warning', 6000);
    }
    // Refresh entry page so count badge updates
    if (App.currentPage === 'entry') PageRenderers.entry();
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
      <button class="toast-close" onclick="App.parentElement.remove()">✕</button>`;
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
