/* ============================================================
   KNNDCmdb – Core Application Logic
   ============================================================ */

'use strict';

// ─── CONFIG (edit these after Google Sheets setup) ───────────
const CONFIG = {
  SHEET_ID:      '',   // Your Google Sheet ID
  API_KEY:       '',   // Your Google Sheets API Key
  SCRIPT_URL:    '',   // Your Google Apps Script Web App URL
  APP_NAME:      'Ketu North NDC Members Database',
  APP_SHORT:     'KNNDCmdb',
  CONSTITUENCY:  'Ketu North',
  VERSION:       '1.0.0',
};

// ─── LOCAL STORAGE KEYS ───────────────────────────────────────
const LS = {
  SESSION:   'knndc_session',
  SETTINGS:  'knndc_settings',
  OFFLINE_Q: 'knndc_offline_queue',
  MEMBERS:   'knndc_members_cache',
  USERS:     'knndc_users',
  AUDIT:     'knndc_audit',
};

// ─── DEFAULT DEMO DATA ────────────────────────────────────────
const DEFAULT_USERS = [
  { id:'u001', username:'admin',    password:'admin123',  name:'System Administrator', role:'admin',        station:'',                branch:'',            active:true },
  { id:'u002', username:'exec',     password:'exec123',   name:'Constituency Exec',    role:'exec',         station:'',                branch:'',            active:true },
  { id:'u003', username:'ward1',    password:'ward123',   name:'Ward Coordinator 1',   role:'ward',         station:'PS-001',           branch:'Aflao Branch',active:true },
  { id:'u004', username:'officer1', password:'off123',    name:'Data Entry Officer 1', role:'officer',      station:'PS-001',           branch:'Aflao Branch',active:true },
  { id:'u005', username:'officer2', password:'off456',    name:'Data Entry Officer 2', role:'officer',      station:'PS-002',           branch:'Denu Branch', active:true },
];

const DEMO_POLLING_STATIONS = [
  { code:'PS-001', name:'Aflao Polling Station',    branch:'Aflao Branch',   branchCode:'BR-001' },
  { code:'PS-002', name:'Denu Polling Station',     branch:'Denu Branch',    branchCode:'BR-002' },
  { code:'PS-003', name:'Agbozume Polling Station', branch:'Agbozume Branch',branchCode:'BR-003' },
  { code:'PS-004', name:'Klikor Polling Station',   branch:'Klikor Branch',  branchCode:'BR-004' },
  { code:'PS-005', name:'Adafienu Polling Station', branch:'Adafienu Branch',branchCode:'BR-005' },
];

const DEMO_MEMBERS = [
  { id:'m001', firstName:'Kofi',    lastName:'Mensah',   otherNames:'Agyei',  partyId:'NDC-2024-001', voterId:'GH-V-001', phone:'0244001122', station:'Aflao Polling Station', stationCode:'PS-001', branch:'Aflao Branch', branchCode:'BR-001', officer:'officer1', timestamp:'2024-01-15 09:30:00' },
  { id:'m002', firstName:'Abena',   lastName:'Korkor',   otherNames:'',       partyId:'NDC-2024-002', voterId:'GH-V-002', phone:'0244002233', station:'Aflao Polling Station', stationCode:'PS-001', branch:'Aflao Branch', branchCode:'BR-001', officer:'officer1', timestamp:'2024-01-15 10:15:00' },
  { id:'m003', firstName:'Yaw',     lastName:'Tetteh',   otherNames:'Kwame',  partyId:'NDC-2024-003', voterId:'GH-V-003', phone:'0554003344', station:'Denu Polling Station',  stationCode:'PS-002', branch:'Denu Branch',  branchCode:'BR-002', officer:'officer2', timestamp:'2024-01-16 08:45:00' },
  { id:'m004', firstName:'Akosua',  lastName:'Kporku',   otherNames:'',       partyId:'NDC-2024-004', voterId:'GH-V-004', phone:'0244004455', station:'Aflao Polling Station', stationCode:'PS-001', branch:'Aflao Branch', branchCode:'BR-001', officer:'officer1', timestamp:'2024-01-16 11:20:00' },
  { id:'m005', firstName:'Efo',     lastName:'Dordor',   otherNames:'Selorm', partyId:'NDC-2024-005', voterId:'GH-V-005', phone:'0504005566', station:'Agbozume Polling Station',stationCode:'PS-003',branch:'Agbozume Branch',branchCode:'BR-003', officer:'officer1', timestamp:'2024-01-17 09:00:00' },
];

// ─── APP STATE ────────────────────────────────────────────────
const App = {
  currentUser:    null,
  currentPage:    'login',
  settings:       {},
  members:        [],
  users:          [],
  auditLog:       [],
  pollingStations:[],
  offlineQueue:   [],
  isOnline:       navigator.onLine,
  charts:         {},

  init() {
    this.loadSettings();
    this.loadData();
    this.applyAppName();
    this.setupNetworkListeners();
    this.checkSession();
    this.renderNav();
  },

  loadSettings() {
    const saved = JSON.parse(localStorage.getItem(LS.SETTINGS) || '{}');
    this.settings = {
      sheetId:     saved.sheetId     || CONFIG.SHEET_ID,
      apiKey:      saved.apiKey      || CONFIG.API_KEY,
      scriptUrl:   saved.scriptUrl   || CONFIG.SCRIPT_URL,
      appName:     saved.appName     || CONFIG.APP_NAME,
      constituency:saved.constituency|| CONFIG.CONSTITUENCY,
      pollingStations: saved.pollingStations || DEMO_POLLING_STATIONS,
    };
    this.pollingStations = this.settings.pollingStations;
  },

  saveSettings() {
    localStorage.setItem(LS.SETTINGS, JSON.stringify(this.settings));
  },

  loadData() {
    this.users       = JSON.parse(localStorage.getItem(LS.USERS)    || 'null') || DEFAULT_USERS;
    this.members     = JSON.parse(localStorage.getItem(LS.MEMBERS)  || '[]');
    this.auditLog    = JSON.parse(localStorage.getItem(LS.AUDIT)    || '[]');
    this.offlineQueue= JSON.parse(localStorage.getItem(LS.OFFLINE_Q)|| '[]');
    if (!this.members.length) { this.members = DEMO_MEMBERS; this.saveMembers(); }
  },

  saveMembers()   { localStorage.setItem(LS.MEMBERS,   JSON.stringify(this.members)); },
  saveUsers()     { localStorage.setItem(LS.USERS,     JSON.stringify(this.users)); },
  saveAudit()     { localStorage.setItem(LS.AUDIT,     JSON.stringify(this.auditLog)); },
  saveOfflineQ()  { localStorage.setItem(LS.OFFLINE_Q, JSON.stringify(this.offlineQueue)); },

  applyAppName() {
    document.querySelectorAll('.app-title-text').forEach(el => el.textContent = this.settings.appName);
    document.title = this.settings.appName;
  },

  setupNetworkListeners() {
    window.addEventListener('online',  () => { this.isOnline = true;  this.updateOnlineStatus(); this.flushOfflineQueue(); });
    window.addEventListener('offline', () => { this.isOnline = false; this.updateOnlineStatus(); });
    this.updateOnlineStatus();
  },

  updateOnlineStatus() {
    const banner = document.getElementById('offline-banner');
    const dot    = document.getElementById('conn-dot');
    if (banner) banner.classList.toggle('show', !this.isOnline);
    if (dot) { dot.className = this.isOnline ? 'online-dot' : 'offline-dot'; }
  },

  checkSession() {
    const s = sessionStorage.getItem(LS.SESSION);
    if (s) {
      try {
        this.currentUser = JSON.parse(s);
        this.showApp();
        return;
      } catch(_) {}
    }
    this.showLogin();
  },

  login(username, password) {
    const user = this.users.find(u => u.username === username && u.password === password && u.active);
    if (!user) return false;
    this.currentUser = user;
    sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
    this.logAudit('LOGIN', `User logged in`, user.username);
    return true;
  },

  logout() {
    this.logAudit('LOGOUT', 'User logged out', this.currentUser?.username);
    this.currentUser = null;
    sessionStorage.removeItem(LS.SESSION);
    this.showLogin();
  },

  showLogin() {
    document.getElementById('app-shell').style.display = 'none';
    document.getElementById('login-screen').style.display = 'flex';
  },

  showApp() {
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-shell').style.display   = 'flex';
    document.getElementById('app-shell').style.flexDirection = 'column';
    this.renderNav();
    this.renderUserHeader();
    this.navigate('dashboard');
  },

  // Role-based page access
  ROLE_PAGES: {
    officer: ['dashboard','entry','my-records'],
    ward:    ['dashboard','records','reports'],
    exec:    ['dashboard','records','reports','analytics'],
    admin:   ['dashboard','entry','records','reports','analytics','audit','settings','users'],
  },

  canAccess(page) {
    const allowed = this.ROLE_PAGES[this.currentUser?.role] || [];
    return allowed.includes(page);
  },

  navigate(page) {
    if (!this.canAccess(page)) { Toast.show('Access Denied', 'You do not have permission to view that page.', 'error'); return; }
    this.currentPage = page;
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const el = document.getElementById('page-' + page);
    if (el) { el.classList.add('active'); PageRenderers[page]?.(); }
    document.querySelectorAll('.nav-link').forEach(a => {
      a.classList.toggle('active', a.dataset.page === page);
    });
    window.scrollTo(0,0);
  },

  renderNav() {
    const nav = document.getElementById('main-nav');
    if (!nav || !this.currentUser) return;
    const pages = [
      { id:'dashboard',  icon:'📊', label:'Dashboard' },
      { id:'entry',      icon:'✍️',  label:'Data Entry', roles:['officer','admin'] },
      { id:'my-records', icon:'📋', label:'My Records',  roles:['officer'] },
      { id:'records',    icon:'🗃️',  label:'All Records', roles:['ward','exec','admin'] },
      { id:'reports',    icon:'📈', label:'Reports',     roles:['ward','exec','admin'] },
      { id:'analytics',  icon:'🔬', label:'Analytics',   roles:['exec','admin'] },
      { id:'audit',      icon:'🛡️',  label:'Audit Log',   roles:['admin'] },
      { id:'users',      icon:'👥', label:'User Mgmt',   roles:['admin'] },
      { id:'settings',   icon:'⚙️',  label:'Settings',    roles:['admin'] },
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
    const roleLabel = {officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Executive',admin:'System Administrator'}[u.role]||u.role;
    document.getElementById('user-avatar').textContent = initials;
    document.getElementById('user-name-hdr').textContent = u.name;
    document.getElementById('user-role-hdr').textContent = roleLabel;
  },

  addMember(data) {
    const member = {
      id: 'm' + Date.now(),
      ...data,
      officer:   this.currentUser.username,
      officerName: this.currentUser.name,
      timestamp: new Date().toLocaleString('en-GH'),
    };
    this.members.unshift(member);
    this.saveMembers();
    this.logAudit('ADD_MEMBER', `Added member: ${data.firstName} ${data.lastName} (${data.partyId})`, this.currentUser.username);
    if (this.isOnline && this.settings.scriptUrl) {
      this.syncToSheets(member);
    } else {
      this.offlineQueue.push({ type:'add', data:member });
      this.saveOfflineQ();
      Toast.show('Saved Offline', 'Record queued for sync when connection is restored.', 'warning');
    }
    return member;
  },

  updateMember(id, updates, reason) {
    const idx = this.members.findIndex(m => m.id === id);
    if (idx < 0) return;
    const old = { ...this.members[idx] };
    this.members[idx] = { ...this.members[idx], ...updates, lastModified: new Date().toLocaleString('en-GH'), modifiedBy: this.currentUser.username };
    this.saveMembers();
    this.logAudit('EDIT_MEMBER', `Updated member ${id}. Reason: ${reason}. Changes: ${JSON.stringify(updates)}`, this.currentUser.username, { before:old, after:this.members[idx], reason });
  },

  deleteMember(id, reason) {
    const m = this.members.find(m=>m.id===id);
    this.members = this.members.filter(m=>m.id!==id);
    this.saveMembers();
    this.logAudit('DELETE_MEMBER', `Deleted member ${id} (${m?.firstName} ${m?.lastName}). Reason: ${reason}`, this.currentUser.username);
  },

  logAudit(action, details, user, extra={}) {
    const entry = { id:'a'+Date.now(), action, details, user:user||'system', timestamp:new Date().toLocaleString('en-GH'), ...extra };
    this.auditLog.unshift(entry);
    if (this.auditLog.length > 5000) this.auditLog = this.auditLog.slice(0, 5000);
    this.saveAudit();
  },

  async syncToSheets(data) {
    if (!this.settings.scriptUrl) return;
    try {
      await fetch(this.settings.scriptUrl, { method:'POST', body:JSON.stringify(data) });
    } catch(e) {
      this.offlineQueue.push({ type:'add', data });
      this.saveOfflineQ();
    }
  },

  async flushOfflineQueue() {
    if (!this.offlineQueue.length || !this.settings.scriptUrl) return;
    Toast.show('Syncing', `Uploading ${this.offlineQueue.length} offline record(s)…`, 'info');
    const failed = [];
    for (const item of this.offlineQueue) {
      try { await this.syncToSheets(item.data); }
      catch(e) { failed.push(item); }
    }
    this.offlineQueue = failed;
    this.saveOfflineQ();
    if (!failed.length) Toast.show('Sync Complete', 'All offline records uploaded.', 'success');
  },

  getMembersForUser() {
    const u = this.currentUser;
    if (!u) return [];
    if (u.role === 'admin' || u.role === 'exec') return this.members;
    if (u.role === 'ward') return this.members.filter(m => m.stationCode === u.station || m.branch === u.branch);
    if (u.role === 'officer') return this.members.filter(m => m.officer === u.username);
    return [];
  },

  getStats() {
    const all = this.getMembersForUser();
    const today = new Date().toLocaleDateString('en-GH');
    const todayRecords = all.filter(m => m.timestamp?.startsWith(today) || false);
    // Group by station
    const byStation = {};
    all.forEach(m => { byStation[m.station] = (byStation[m.station]||0)+1; });
    // Group by day (last 7)
    const byDay = {};
    for (let i=6; i>=0; i--) {
      const d = new Date(); d.setDate(d.getDate()-i);
      const key = d.toLocaleDateString('en-GH');
      byDay[key] = all.filter(m => m.timestamp?.includes(key)).length;
    }
    return { total:all.length, today:todayRecords.length, byStation, byDay, stations:Object.keys(byStation).length };
  },
};

// ─── TOAST ────────────────────────────────────────────────────
const Toast = {
  show(title, msg='', type='success', duration=4000) {
    const container = document.getElementById('toast-container');
    const icons = { success:'✅', error:'❌', warning:'⚠️', info:'ℹ️' };
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.innerHTML = `<span class="toast-icon">${icons[type]||'✅'}</span>
      <div class="toast-content"><div class="toast-title">${title}</div>${msg?`<div class="toast-msg">${msg}</div>`:''}</div>
      <button class="toast-close" onclick="this.parentElement.remove()">✕</button>`;
    container.appendChild(t);
    setTimeout(() => { t.classList.add('hiding'); setTimeout(()=>t.remove(), 300); }, duration);
  }
};

// ─── MODAL HELPER ─────────────────────────────────────────────
const Modal = {
  open(id)  { document.getElementById(id)?.classList.add('open'); },
  close(id) { document.getElementById(id)?.classList.remove('open'); },
  closeAll() { document.querySelectorAll('.modal-overlay').forEach(m=>m.classList.remove('open')); },
};

// Click outside to close
document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay')) Modal.closeAll();
});
