/* ============================================================
   KNNDCmdb – Core Application Logic  v1.2
   ============================================================ */

'use strict';

// ─── CONFIG ──────────────────────────────────────────────────
const CONFIG = {
  SHEET_ID:      '',
  API_KEY:       '',
  SCRIPT_URL:    '',
  APP_NAME:      'Ketu North NDC Members Database',
  APP_SHORT:     'KNNDCmdb',
  CONSTITUENCY:  'Ketu North',
  VERSION:       '1.2.0',
  INACTIVITY_MS: 10 * 60 * 1000,   // 10 minutes
};

// ─── LOCAL STORAGE KEYS ──────────────────────────────────────
const LS = {
  SESSION:   'knndc_session',
  SETTINGS:  'knndc_settings',
  OFFLINE_Q: 'knndc_offline_queue',
  MEMBERS:   'knndc_members',        // persistent across all roles
  USERS:     'knndc_users',
  AUDIT:     'knndc_audit',
  LOCKOUT:   'knndc_lockout',
  ATTEMPTS:  'knndc_attempts',
};

// ─── DEMO DATA ───────────────────────────────────────────────
const DEFAULT_USERS = [
  { id:'u001', username:'admin',    password:'Admin@2026',  name:'System Administrator',   role:'admin',   ward:'', station:'', branch:'', assignedStations:[], active:true },
  { id:'u002', username:'exec',     password:'Exec@2026',   name:'Constituency Executive',  role:'exec',    ward:'', station:'', branch:'', assignedStations:[], active:true },
  { id:'u003', username:'ward1',    password:'Ward@2026',   name:'Ward Coordinator (Aflao)',role:'ward',    ward:'Aflao Ward',   station:'PS-001', branch:'Aflao Branch', assignedStations:['PS-001','PS-002'], active:true },
  { id:'u004', username:'officer1', password:'Off1@2026',   name:'Data Entry Officer 1',    role:'officer', ward:'Aflao Ward',   station:'PS-001', branch:'Aflao Branch', assignedStations:['PS-001'], active:true },
  { id:'u005', username:'officer2', password:'Off2@2026',   name:'Data Entry Officer 2',    role:'officer', ward:'Denu Ward',    station:'PS-003', branch:'Denu Branch',  assignedStations:['PS-003'], active:true },
];

const DEMO_POLLING_STATIONS = [
  { ward:'Aflao Ward',    code:'PS-001', name:'Aflao A Polling Station',   branch:'Aflao Branch',    branchCode:'BR-001' },
  { ward:'Aflao Ward',    code:'PS-002', name:'Aflao B Polling Station',   branch:'Aflao Branch',    branchCode:'BR-001' },
  { ward:'Denu Ward',     code:'PS-003', name:'Denu Polling Station',      branch:'Denu Branch',     branchCode:'BR-002' },
  { ward:'Agbozume Ward', code:'PS-004', name:'Agbozume Polling Station',  branch:'Agbozume Branch', branchCode:'BR-003' },
  { ward:'Klikor Ward',   code:'PS-005', name:'Klikor Polling Station',    branch:'Klikor Branch',   branchCode:'BR-004' },
  { ward:'Adafienu Ward', code:'PS-006', name:'Adafienu Polling Station',  branch:'Adafienu Branch', branchCode:'BR-005' },
];

const SEED_MEMBERS = [
  { id:'m001', firstName:'Kofi',   lastName:'Mensah',  otherNames:'Agyei',  partyId:'NDC-2024-001', voterId:'GH-V-001', phone:'0244001122', ward:'Aflao Ward',    station:'Aflao A Polling Station',  stationCode:'PS-001', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'15/01/2024, 09:30:00' },
  { id:'m002', firstName:'Abena',  lastName:'Korkor',  otherNames:'',       partyId:'NDC-2024-002', voterId:'GH-V-002', phone:'0244002233', ward:'Aflao Ward',    station:'Aflao A Polling Station',  stationCode:'PS-001', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'15/01/2024, 10:15:00' },
  { id:'m003', firstName:'Yaw',    lastName:'Tetteh',  otherNames:'Kwame',  partyId:'NDC-2024-003', voterId:'GH-V-003', phone:'0554003344', ward:'Denu Ward',     station:'Denu Polling Station',     stationCode:'PS-003', branch:'Denu Branch',     branchCode:'BR-002', officer:'officer2', officerName:'Data Entry Officer 2', timestamp:'16/01/2024, 08:45:00' },
  { id:'m004', firstName:'Akosua', lastName:'Kporku',  otherNames:'',       partyId:'NDC-2024-004', voterId:'GH-V-004', phone:'0244004455', ward:'Aflao Ward',    station:'Aflao B Polling Station',  stationCode:'PS-002', branch:'Aflao Branch',    branchCode:'BR-001', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'16/01/2024, 11:20:00' },
  { id:'m005', firstName:'Efo',    lastName:'Dordor',  otherNames:'Selorm', partyId:'NDC-2024-005', voterId:'GH-V-005', phone:'0504005566', ward:'Agbozume Ward', station:'Agbozume Polling Station', stationCode:'PS-004', branch:'Agbozume Branch', branchCode:'BR-003', officer:'officer1', officerName:'Data Entry Officer 1', timestamp:'17/01/2024, 09:00:00' },
];

// ─── APP STATE ────────────────────────────────────────────────
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

  init() {
    this.loadSettings();
    this.loadData();
    this.applyAppName();
    this.setupNetworkListeners();
    this.checkSession();
  },

  // ── SETTINGS ──
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
  saveSettings() { localStorage.setItem(LS.SETTINGS, JSON.stringify(this.settings)); },

  // ── DATA LOADING — persistent across logins ──
  loadData() {
    // Users
    const su = localStorage.getItem(LS.USERS);
    this.users = su ? JSON.parse(su) : DEFAULT_USERS;
    if (!su) this.saveUsers();

    // Members — ALWAYS from localStorage; never overwrite if data exists
    const sm = localStorage.getItem(LS.MEMBERS);
    this.members = sm ? JSON.parse(sm) : SEED_MEMBERS;
    if (!sm) this.saveMembers();

    this.auditLog     = JSON.parse(localStorage.getItem(LS.AUDIT)    || '[]');
    this.offlineQueue = JSON.parse(localStorage.getItem(LS.OFFLINE_Q)|| '[]');
  },

  saveMembers()  { localStorage.setItem(LS.MEMBERS,   JSON.stringify(this.members)); },
  saveUsers()    { localStorage.setItem(LS.USERS,     JSON.stringify(this.users)); },
  saveAudit()    { localStorage.setItem(LS.AUDIT,     JSON.stringify(this.auditLog)); },
  saveOfflineQ() { localStorage.setItem(LS.OFFLINE_Q, JSON.stringify(this.offlineQueue)); },

  applyAppName() {
    const n = this.settings.appName || CONFIG.APP_NAME;
    document.querySelectorAll('.app-title-text').forEach(el => el.textContent = n);
    document.title = n;
  },

  // ── NETWORK ──
  setupNetworkListeners() {
    window.addEventListener('online',  () => { this.isOnline = true;  this.updateOnlineStatus(); this.flushOfflineQueue(); });
    window.addEventListener('offline', () => { this.isOnline = false; this.updateOnlineStatus(); });
    this.updateOnlineStatus();
  },
  updateOnlineStatus() {
    const banner = document.getElementById('offline-banner');
    const dot    = document.getElementById('conn-dot');
    if (banner) banner.classList.toggle('show', !this.isOnline);
    if (dot)    dot.className = this.isOnline ? 'online-dot' : 'offline-dot';
  },

  // ── SESSION ──
  checkSession() {
    const s = sessionStorage.getItem(LS.SESSION);
    if (s) {
      try { this.currentUser = JSON.parse(s); this.showApp(); return; } catch(_) {}
    }
    this.showLogin();
  },

  // ── INACTIVITY TIMER ──
  resetInactivityTimer() {
    if (this._inactivityTimer) clearTimeout(this._inactivityTimer);
    this._inactivityTimer = setTimeout(() => {
      if (!this.currentUser) return;
      this.logAudit('AUTO_LOGOUT', 'Session expired — 10 minutes inactivity', this.currentUser.username);
      Toast.show('Session Expired', 'Logged out after 10 minutes of inactivity.', 'warning', 5000);
      setTimeout(() => this.logout(), 1800);
    }, CONFIG.INACTIVITY_MS);
  },

  setupInactivityTracking() {
    ['mousemove','keydown','mousedown','touchstart','scroll','click'].forEach(evt =>
      document.addEventListener(evt, () => this.resetInactivityTimer(), { passive: true })
    );
    this.resetInactivityTimer();
  },

  stopInactivityTracking() {
    if (this._inactivityTimer) clearTimeout(this._inactivityTimer);
    this._inactivityTimer = null;
  },

  // ── AUTH ──
  login(username, password) {
    const MAX = 5, LOCK_MS = 2 * 60 * 1000;

    const lockData = JSON.parse(localStorage.getItem(LS.LOCKOUT) || 'null');
    if (lockData) {
      const rem = lockData.until - Date.now();
      if (rem > 0) return { locked: true, seconds: Math.ceil(rem / 1000) };
      localStorage.removeItem(LS.LOCKOUT);
      localStorage.removeItem(LS.ATTEMPTS);
    }

    // Re-load users from storage in case admin changed passwords
    this.users = JSON.parse(localStorage.getItem(LS.USERS) || 'null') || DEFAULT_USERS;
    const user = this.users.find(u => u.username === username && u.password === password && u.active);

    if (!user) {
      const n = (parseInt(localStorage.getItem(LS.ATTEMPTS) || '0')) + 1;
      localStorage.setItem(LS.ATTEMPTS, n);
      if (n >= MAX) {
        localStorage.setItem(LS.LOCKOUT, JSON.stringify({ until: Date.now() + LOCK_MS }));
        localStorage.removeItem(LS.ATTEMPTS);
        this.logAudit('LOCKOUT', `Locked after ${MAX} failed attempts — username: ${username}`, 'system');
        return { locked: true, seconds: LOCK_MS / 1000 };
      }
      return { failed: true, attemptsLeft: MAX - n };
    }

    localStorage.removeItem(LS.ATTEMPTS);
    localStorage.removeItem(LS.LOCKOUT);
    this.currentUser = user;
    sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
    this.logAudit('LOGIN', `Logged in successfully`, user.username);
    return { success: true };
  },

  logout() {
    this.logAudit('LOGOUT', 'User logged out', this.currentUser?.username);
    this.stopInactivityTracking();
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
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-shell').style.display    = 'flex';
    document.getElementById('app-shell').style.flexDirection = 'column';
    this.renderNav();
    this.renderUserHeader();
    this.setupInactivityTracking();
    this.navigate('dashboard');
  },

  // ── ROLE ACCESS ──
  ROLE_PAGES: {
    officer: ['dashboard','entry','my-records'],
    ward:    ['dashboard','records','reports'],
    exec:    ['dashboard','records','reports','analytics'],
    admin:   ['dashboard','entry','records','reports','analytics','audit','users','settings'],
  },

  canAccess(page) {
    return (this.ROLE_PAGES[this.currentUser?.role] || []).includes(page);
  },

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
      { id:'dashboard', icon:'📊', label:'Dashboard' },
      { id:'entry',     icon:'✍️',  label:'Data Entry' },
      { id:'my-records',icon:'📋', label:'My Records' },
      { id:'records',   icon:'🗃️',  label:'All Records' },
      { id:'reports',   icon:'📈', label:'Reports' },
      { id:'analytics', icon:'🔬', label:'Analytics' },
      { id:'audit',     icon:'🛡️',  label:'Audit Log' },
      { id:'users',     icon:'👥', label:'User Mgmt' },
      { id:'settings',  icon:'⚙️',  label:'Settings' },
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
    const initials  = u.name.split(' ').map(n => n[0]).slice(0, 2).join('');
    const roleLabel = { officer:'Data Entry Officer', ward:'Ward Coordinator', exec:'Constituency Executive', admin:'System Administrator' }[u.role] || u.role;
    document.getElementById('user-avatar').textContent   = initials;
    document.getElementById('user-name-hdr').textContent = u.name;
    document.getElementById('user-role-hdr').textContent = roleLabel;
  },

  // ── MEMBER CRUD ──
  addMember(data) {
    const member = {
      id: 'm' + Date.now(),
      ...data,
      officer:     this.currentUser.username,
      officerName: this.currentUser.name,
      timestamp:   new Date().toLocaleString('en-GH'),
    };
    // Always re-read from storage first to avoid overwriting concurrent records
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
    this.members.unshift(member);
    this.saveMembers();
    this.logAudit('ADD_MEMBER', `Added: ${data.firstName} ${data.lastName} (${data.partyId}) — ${data.station}`, this.currentUser.username);
    if (this.isOnline && this.settings.scriptUrl) {
      this.syncToSheets(member);
    } else {
      this.offlineQueue.push({ type:'add', data: member });
      this.saveOfflineQ();
      if (!this.isOnline) Toast.show('Saved Offline','Will sync when connection is restored.','warning');
    }
    return member;
  },

  updateMember(id, updates, reason) {
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
    const idx = this.members.findIndex(m => m.id === id);
    if (idx < 0) return;
    const before = { ...this.members[idx] };
    this.members[idx] = { ...this.members[idx], ...updates, lastModified: new Date().toLocaleString('en-GH'), modifiedBy: this.currentUser.username };
    this.saveMembers();
    this.logAudit('EDIT_MEMBER', `Edited: ${before.firstName} ${before.lastName}. Reason: ${reason}`, this.currentUser.username, { before, after: this.members[idx], reason });
  },

  deleteMember(id, reason) {
    this.members = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
    const m = this.members.find(m => m.id === id);
    this.members = this.members.filter(m => m.id !== id);
    this.saveMembers();
    this.logAudit('DELETE_MEMBER', `Deleted: ${m?.firstName} ${m?.lastName} (${m?.partyId}). Reason: ${reason}`, this.currentUser.username);
  },

  // ── DATA VISIBILITY — always fresh from storage ──
  getMembersForUser() {
    const all = JSON.parse(localStorage.getItem(LS.MEMBERS) || '[]');
    this.members = all;
    const u = this.currentUser;
    if (!u) return [];
    if (u.role === 'admin' || u.role === 'exec') return all;
    const codes = (u.assignedStations || []).length ? u.assignedStations : (u.station ? [u.station] : []);
    if (u.role === 'ward')    return all.filter(m => codes.includes(m.stationCode) || m.ward === u.ward || m.branch === u.branch);
    if (u.role === 'officer') return all.filter(m => m.officer === u.username || codes.includes(m.stationCode));
    return [];
  },

  // ── AUDIT ──
  logAudit(action, details, user, extra = {}) {
    const entry = { id:'a'+Date.now(), action, details, user: user || 'system', timestamp: new Date().toLocaleString('en-GH'), ...extra };
    this.auditLog = JSON.parse(localStorage.getItem(LS.AUDIT) || '[]');
    this.auditLog.unshift(entry);
    if (this.auditLog.length > 10000) this.auditLog = this.auditLog.slice(0, 10000);
    this.saveAudit();
  },

  // ── STATS ──
  getStats() {
    const all   = this.getMembersForUser();
    const today = new Date().toLocaleDateString('en-GH');
    const byStation = {};
    all.forEach(m => { byStation[m.station] = (byStation[m.station] || 0) + 1; });
    const byDay = {};
    for (let i = 6; i >= 0; i--) {
      const d = new Date(); d.setDate(d.getDate() - i);
      const key = d.toLocaleDateString('en-GH');
      byDay[key] = all.filter(m => m.timestamp?.includes(key)).length;
    }
    return { total: all.length, today: all.filter(m => m.timestamp?.includes(today)).length, byStation, byDay, stations: Object.keys(byStation).length };
  },

  // ── SYNC ──
  async syncToSheets(data) {
    if (!this.settings.scriptUrl) return;
    try { await fetch(this.settings.scriptUrl, { method:'POST', body: JSON.stringify(data) }); }
    catch(e) { this.offlineQueue.push({ type:'add', data }); this.saveOfflineQ(); }
  },

  async flushOfflineQueue() {
    if (!this.offlineQueue.length || !this.settings.scriptUrl) return;
    Toast.show('Syncing', `Uploading ${this.offlineQueue.length} record(s)…`, 'info');
    const failed = [];
    for (const item of this.offlineQueue) {
      try { await this.syncToSheets(item.data); } catch(e) { failed.push(item); }
    }
    this.offlineQueue = failed;
    this.saveOfflineQ();
    if (!failed.length) Toast.show('Sync Complete','All records uploaded.','success');
  },
};

// ─── TOAST ───────────────────────────────────────────────────
const Toast = {
  show(title, msg = '', type = 'success', duration = 4000) {
    const c = document.getElementById('toast-container');
    if (!c) return;
    const icons = { success:'✅', error:'❌', warning:'⚠️', info:'ℹ️' };
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.innerHTML = `<span class="toast-icon">${icons[type]||'✅'}</span>
      <div class="toast-content"><div class="toast-title">${title}</div>${msg?`<div class="toast-msg">${msg}</div>`:''}</div>
      <button class="toast-close" onclick="this.parentElement.remove()">✕</button>`;
    c.appendChild(t);
    setTimeout(() => { t.classList.add('hiding'); setTimeout(() => t.remove(), 300); }, duration);
  }
};

// ─── MODAL ───────────────────────────────────────────────────
const Modal = {
  open(id)   { document.getElementById(id)?.classList.add('open'); },
  close(id)  { document.getElementById(id)?.classList.remove('open'); },
  closeAll() { document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('open')); },
};
document.addEventListener('click', e => { if (e.target.classList.contains('modal-overlay')) Modal.closeAll(); });
