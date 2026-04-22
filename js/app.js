// ============================================================
// KNNDCmdb  app.js  v3.0.5
// Elections & IT Directorate · Ketu North NDC · 2026
//
// Changes from v2.4.0 → v3.0:
//   - All Google Apps Script XHR calls replaced with REST API calls
//   - JWT token stored in App.jwt; added to every request header
//   - Login now calls POST /login and receives a JWT token
//   - Change-password calls POST /change-password
//   - Offline queue bulk-flush calls POST /members/bulk
//   - App.settings.scriptUrl is now the API base URL (e.g. https://knndc-api.up.railway.app)
//   - GAS-specific helpers (xhrGAS, buildGASUrl) removed; replaced with apiGet / apiPost
//   - All localStorage auth logic retained as offline fallback
// ============================================================

var App = (() => {
  'use strict';

  // ── Version ───────────────────────────────────────────────
  const VERSION = '3.0.5';

  // ── localStorage keys ─────────────────────────────────────
  const LS = {
    SESSION:       'knndc_session',
    SETTINGS:      'knndc_settings',
    USERS:         'knndc_users',
    MEMBERS:       'knndc_members',
    AUDIT:         'knndc_audit',
    OFFLINE_QUEUE: 'knndc_offline_queue',
    LOCKOUT:       'knndc_lockout',
    DEMO_CLEARED:  'knndc_demo_cleared',
    JWT:           'knndc_jwt',       // ← new: persisted JWT token
  };

  // ── Role page access ──────────────────────────────────────
  const ROLE_PAGES = {
    admin:   ['dashboard', 'entry', 'records', 'reports', 'analytics', 'users', 'audit', 'settings'],
    exec:    ['dashboard', 'records', 'reports', 'analytics'],
    ward:    ['dashboard', 'entry', 'records', 'reports'],
    officer: ['dashboard', 'entry', 'my-records', 'records'],
  };

  // ── State ─────────────────────────────────────────────────
  let currentUser    = null;
  let settings       = {};
  let pollingStations = [];
  let jwt            = '';          // ← active JWT token
  let _membersCache  = null;        // ← in-memory members (avoids 5MB localStorage limit)

  // ── Getters ───────────────────────────────────────────────
  function getVersion()         { return VERSION; }
  function getCurrentUser()     { return currentUser; }
  function getSettings()        { return settings; }
  function getPollingStations() { return pollingStations; }
  function getJwt()             { return jwt; }

  // ── Persistence helpers ───────────────────────────────────
  function lsGet(key, fallback = null) {
    try { const v = localStorage.getItem(key); return v !== null ? JSON.parse(v) : fallback; }
    catch { return fallback; }
  }
  function lsSet(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch { return false; }
  }
  function lsDel(key) { try { localStorage.removeItem(key); } catch {} }

  // ── Settings ──────────────────────────────────────────────
  function loadSettings() {
    const saved = lsGet(LS.SETTINGS, {});
    settings = Object.assign({
      scriptUrl:       '',
      appName:         'KNNDCmdb',
      constituency:    'Ketu North',
      sheetId:         '',
      pollingStations: [],
    }, saved);
    pollingStations = Array.isArray(settings.pollingStations) ? settings.pollingStations : [];
  }

  function saveSettings(patch) {
    settings = Object.assign(settings, patch);
    lsSet(LS.SETTINGS, settings);
  }

  function getApiBase() {
    return (settings.scriptUrl || '').replace(/\/+$/, '');
  }

  // ── JWT helpers ───────────────────────────────────────────
  function setJwt(token) {
    jwt = token || '';
    if (jwt) lsSet(LS.JWT, jwt);
    else lsDel(LS.JWT);
  }

  function loadJwt() {
    jwt = lsGet(LS.JWT, '') || '';
  }

  function isJwtExpired() {
    if (!jwt) return true;
    try {
      const payload = JSON.parse(atob(jwt.split('.')[1]));
      return Date.now() / 1000 > payload.exp;
    } catch { return true; }
  }

  // ── XHR helpers (JWT-authenticated) ─────────────────────
  // All API calls use these two helpers instead of the old GAS helpers.

  function apiGet(path) {
    return new Promise(resolve => {
      const base = getApiBase();
      if (!base) return resolve({ success: false, error: 'No API URL configured' });

      const xhr = new XMLHttpRequest();
      xhr.open('GET', base + path, true);
      xhr.timeout = 15000;
      xhr.setRequestHeader('Authorization', 'Bearer ' + (jwt || lsGet(LS.JWT,'') || ''));
      xhr.setRequestHeader('Content-Type', 'application/json');
      xhr.onload = () => {
        try { resolve(JSON.parse(xhr.responseText)); }
        catch { resolve({ success: false, error: 'Invalid JSON response' }); }
      };
      xhr.onerror   = () => resolve({ success: false, error: 'Network error' });
      xhr.ontimeout = () => resolve({ success: false, error: 'Request timed out' });
      xhr.send();
    });
  }

  function apiPost(path, data) {
    return new Promise(resolve => {
      const base = getApiBase();
      if (!base) return resolve({ success: false, error: 'No API URL configured' });

      const xhr = new XMLHttpRequest();
      xhr.open('POST', base + path, true);
      xhr.timeout = 20000;
      xhr.setRequestHeader('Authorization', 'Bearer ' + (jwt || lsGet(LS.JWT,'') || ''));
      xhr.setRequestHeader('Content-Type', 'application/json');
      xhr.onload = () => {
        try { resolve(JSON.parse(xhr.responseText)); }
        catch { resolve({ success: false, error: 'Invalid JSON response' }); }
      };
      xhr.onerror   = () => resolve({ success: false, error: 'Network error' });
      xhr.ontimeout = () => resolve({ success: false, error: 'Request timed out' });
      xhr.send(JSON.stringify(data));
    });
  }

  function apiDelete(path) {
    return new Promise(resolve => {
      const base = getApiBase();
      if (!base) return resolve({ success: false, error: 'No API URL configured' });

      const xhr = new XMLHttpRequest();
      xhr.open('DELETE', base + path, true);
      xhr.timeout = 10000;
      xhr.setRequestHeader('Authorization', 'Bearer ' + (jwt || lsGet(LS.JWT,'') || ''));
      xhr.onload = () => {
        try { resolve(JSON.parse(xhr.responseText)); }
        catch { resolve({ success: false, error: 'Invalid JSON response' }); }
      };
      xhr.onerror   = () => resolve({ success: false, error: 'Network error' });
      xhr.ontimeout = () => resolve({ success: false, error: 'Request timed out' });
      xhr.send();
    });
  }

  // ── API connectivity check ────────────────────────────────
  async function pingApi() {
    const base = getApiBase();
    if (!base) return false;
    try {
      const res = await apiGet('/ping');
      return res.status === 'ok';
    } catch { return false; }
  }

  // ── Login ─────────────────────────────────────────────────
  async function login(username, password) {
    const base = getApiBase();

    // ── API login (required when API URL is configured) ────────
    if (base) {
      try {
        const xhr = new XMLHttpRequest();
        const result = await new Promise(resolve => {
          xhr.open('POST', base + '/login', true);
          xhr.timeout = 15000;
          xhr.setRequestHeader('Content-Type', 'application/json');
          xhr.onload = () => {
            try { resolve(JSON.parse(xhr.responseText)); }
            catch { resolve({ success: false }); }
          };
          xhr.onerror   = () => resolve({ success: false, error: 'Network error' });
          xhr.ontimeout = () => resolve({ success: false, error: 'Timeout' });
          xhr.send(JSON.stringify({ username, password }));
        });

        if (result.success && result.token && result.user) {
          setJwt(result.token);
          const user = result.user;

          // Normalise fields to match local user object shape
          if (typeof user.assigned_stations === 'string') {
            try { user.assignedStations = JSON.parse(user.assigned_stations); }
            catch { user.assignedStations = []; }
          } else {
            user.assignedStations = user.assigned_stations || [];
          }
          user.mustChangePassword = !!user.must_change_password;
          user.active = !!user.active;

          // Store in session and update local users cache
          sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
          const users = lsGet(LS.USERS, []);
          const idx = users.findIndex(u => u.id === user.id);
          if (idx >= 0) users[idx] = { ...users[idx], ...user };
          else users.push(user);
          lsSet(LS.USERS, users);

          currentUser = user;
          return { success: true, user };
        }

        // API returned a definitive error — reject immediately, no fallback
        return {
          success: false,
          error: result.error || 'Invalid credentials',
        };

      } catch {
        // Network completely unreachable
        return { success: false, error: 'Cannot reach the server. Check your internet connection.' };
      }
    }

    // ── Offline-only mode (no API URL configured at all) ───────
    // Only reached when the app has never been connected to an API.
    // Once an API URL is set, ALL authentication goes through the API above.
    const users = lsGet(LS.USERS, []);
    const user  = users.find(u => u.username === username && u.active !== false);
    if (!user) return { success: false, error: 'Invalid credentials' };
    if (user.password && user.password !== password)
      return { success: false, error: 'Invalid credentials' };
    setJwt('');
    sessionStorage.setItem(LS.SESSION, JSON.stringify(user));
    currentUser = user;
    return { success: true, user, offline: true };
  }

  // ── Logout ────────────────────────────────────────────────
  function logout() {
    currentUser = null;
    setJwt('');
    sessionStorage.removeItem(LS.SESSION);
  }

  // ── Session restore ───────────────────────────────────────
  function restoreSession() {
    try {
      const raw = sessionStorage.getItem(LS.SESSION);
      if (!raw) return false;
      currentUser = JSON.parse(raw);
      loadJwt();
      // If JWT is expired, clear it (user will need to re-login on next API call)
      if (isJwtExpired()) setJwt('');
      return !!currentUser;
    } catch { return false; }
  }

  // ── Re-validate session (called after sync) ───────────────
  async function _revalidateSession() {
    if (!currentUser) return;
    const users = lsGet(LS.USERS, []);
    const fresh = users.find(u => u.id === currentUser.id);
    if (!fresh) return;

    const changed =
      fresh.role    !== currentUser.role    ||
      fresh.ward    !== currentUser.ward    ||
      fresh.station !== currentUser.station ||
      fresh.branch  !== currentUser.branch;

    currentUser = { ...currentUser, ...fresh };
    sessionStorage.setItem(LS.SESSION, JSON.stringify(currentUser));

    if (changed) {
      _renderNav();
      _renderHeader();
      showToast('Your role or station assignment has been updated.', 'info');
    }
  }

  // ── User helpers ──────────────────────────────────────────
  function getUsers() { return lsGet(LS.USERS, []); }

  function getMembersForUser(allMembers) {
    if (!currentUser) return [];
    const role = currentUser.role;
    if (role === 'admin') return allMembers;
    if (role === 'exec') return allMembers;
    if (role === 'ward') {
      return allMembers.filter(m => m.ward === currentUser.ward);
    }
    if (role === 'officer') {
      const assigned = currentUser.assignedStations || [];
      return allMembers.filter(m => {
        const sc = m.stationCode || m.station_code || '';
        return assigned.includes(sc);
      });
    }
    return [];
  }

  function canModifyMember(member) {
    if (!currentUser) return false;
    const role = currentUser.role;
    if (role === 'admin') return true;
    if (role === 'exec') return false;
    if (role === 'ward') return member.ward === currentUser.ward;
    if (role === 'officer') {
      const sc = member.stationCode || member.station_code || '';
      return (currentUser.assignedStations || []).includes(sc);
    }
    return false;
  }

  function getStationPool() {
    if (!currentUser) return [];
    const role = currentUser.role;
    if (role === 'admin' || role === 'exec') return pollingStations;
    if (role === 'ward') return pollingStations.filter(s => s.ward === currentUser.ward);
    if (role === 'officer') {
      const assigned = currentUser.assignedStations || [];
      return pollingStations.filter(s => assigned.includes(s.code || s.station_code || ''));
    }
    return [];
  }

  // ── Members ───────────────────────────────────────────────
  function getMembers() {
    // Use in-memory cache if available (set by fetchFromApi after login)
    if (_membersCache !== null) return _membersCache;
    // Fall back to localStorage — normalise field names in case snake_case
    // records were stored from a previous API fetch
    const stored = lsGet(LS.MEMBERS, []);
    if (stored.length) {
      _membersCache = stored.map(_normaliseMember);
      return _membersCache;
    }
    return [];
  }

  function saveMembers(members) {
    _membersCache = members;
    // Only write to localStorage if dataset is small enough (< 3MB)
    try {
      const json = JSON.stringify(members);
      if (json.length < 3 * 1024 * 1024) {
        localStorage.setItem(LS.MEMBERS, json);
      } else {
        // Dataset too large for localStorage — memory only
        localStorage.removeItem(LS.MEMBERS);
      }
    } catch {}
  }

  function addMember(member) {
    // Add to memory cache immediately so UI updates without waiting for sync
    if (_membersCache === null) _membersCache = getMembers();
    _membersCache.push(member);
    saveMembers(_membersCache);
    _pushMemberToApi(member, 'add');
    _logAudit('ADD_MEMBER', `Added ${member.firstName} ${member.lastName}`, member.id);
  }

  function updateMember(updated) {
    if (_membersCache === null) _membersCache = getMembers();
    const idx = _membersCache.findIndex(m => m.id === (updated.id || updated));
    const u   = typeof updated === 'string' ? { id: updated } : updated;
    if (idx >= 0) { _membersCache[idx] = { ..._membersCache[idx], ...u }; }
    saveMembers(_membersCache);
    _pushMemberToApi(u, 'update');
    _logAudit('UPDATE_MEMBER', `Updated ${u.firstName||''} ${u.lastName||''}`, u.id);
  }

  function deleteMember(id, reason) {
    if (_membersCache === null) _membersCache = getMembers();
    const found   = _membersCache.find(m => m.id === id);
    _membersCache = _membersCache.filter(m => m.id !== id);
    saveMembers(_membersCache);
    if (found) {
      _deleteMemberFromApi(id);
      _logAudit('DELETE_MEMBER', `Deleted ${found.firstName||found.first_name||''} ${found.lastName||found.last_name||''}`, id, reason);
    }
  }

  function generateMemberId() {
    return 'mbr-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
  }

  // ── API push helpers (fire-and-forget) ───────────────────
  function _pushMemberToApi(member, mode) {
    if (!getApiBase() || isJwtExpired()) {
      _enqueueOffline(member);
      return;
    }
    apiPost('/members/upsert', member).then(res => {
      if (!res.success) _enqueueOffline(member);
    });
  }

  function _deleteMemberFromApi(id) {
    if (!getApiBase() || isJwtExpired()) return;
    apiDelete('/members/' + id);
  }

  // ── Offline queue ─────────────────────────────────────────
  function _enqueueOffline(member) {
    const queue = lsGet(LS.OFFLINE_QUEUE, []);
    const idx   = queue.findIndex(m => m.id === member.id);
    if (idx >= 0) queue[idx] = member; else queue.push(member);
    lsSet(LS.OFFLINE_QUEUE, queue);
  }

  async function flushOfflineQueue() {
    const queue = lsGet(LS.OFFLINE_QUEUE, []);
    if (!queue.length) return { flushed: 0 };
    if (!getApiBase() || isJwtExpired()) return { flushed: 0, reason: 'offline' };

    const res = await apiPost('/members/bulk', { members: queue });
    if (res.success) {
      lsDel(LS.OFFLINE_QUEUE);
      return { flushed: queue.length, errors: res.errors };
    }
    return { flushed: 0, error: res.error };
  }

  function getOfflineQueueCount() {
    return lsGet(LS.OFFLINE_QUEUE, []).length;
  }

  // ── Remote sync ───────────────────────────────────────────

  // ── Normalise API member record (snake_case → camelCase) ──────
  // The MySQL API returns snake_case field names.
  // The frontend everywhere expects camelCase.
  // Run every member through this once on fetch so the rest of the
  // app never needs to handle both forms.
  function _normaliseMember(m) {
    return {
      id:           m.id,
      firstName:    m.firstName    || m.first_name    || '',
      lastName:     m.lastName     || m.last_name     || '',
      otherNames:   m.otherNames   || m.other_names   || '',
      gender:       m.gender       || '',
      zone:         m.zone         || '',
      partyId:      m.partyId      || m.party_id      || '',
      voterId:      m.voterId      || m.voter_id      || '',
      phone:        m.phone        || '',
      ward:         m.ward         || '',
      station:      m.station      || '',
      stationCode:  m.stationCode  || m.station_code  || '',
      branch:       m.branch       || '',
      branchCode:   m.branchCode   || m.branch_code   || '',
      officer:      m.officer      || '',
      officerName:  m.officerName  || m.officer_name  || '',
      timestamp:    m.timestamp    || '',
      isoDate:      m.isoDate      || m.iso_date      || '',
      _demo:        m._demo        || false,
    };
  }

  async function fetchFromApi() {
    if (!getApiBase() || isJwtExpired()) return false;

    try {
      // Fetch members and stations in parallel
      // For large datasets (admin sees all 12k+ records), only cache in memory
      // For scoped users (ward/officer), the dataset is small enough for localStorage
      const [mRes, sRes] = await Promise.all([
        apiGet('/members'),
        apiGet('/stations'),
      ]);

      if (mRes.success && Array.isArray(mRes.members)) {
        const apiIds    = new Set(mRes.members.map(m => m.id));
        const queue     = lsGet(LS.OFFLINE_QUEUE, []);
        const queueIds  = new Set(queue.map(m => m.id));
        const localOnly = (_membersCache || []).filter(m => !apiIds.has(m.id) && queueIds.has(m.id));
        const allMembers = [...mRes.members.map(_normaliseMember), ...localOnly];

        // For non-admin users, also save scoped subset to localStorage
        // so data survives a page refresh without another API call
        if (currentUser && currentUser.role !== 'admin') {
          const scoped = getMembersForUser(allMembers);
          try {
            const json = JSON.stringify(scoped);
            if (json.length < 4 * 1024 * 1024) {
              localStorage.setItem(LS.MEMBERS, json);
            }
          } catch {}
        } else {
          // Admin: memory only (too large for localStorage)
          localStorage.removeItem(LS.MEMBERS);
        }

        _membersCache = allMembers;
      }

      if (sRes.success && Array.isArray(sRes.stations)) {
        pollingStations = sRes.stations.map(s => ({
          code:        s.code,
          zone:        s.zone,
          ward:        s.ward,
          name:        s.name,
          branch:      s.branch,
          branchCode:  s.branch_code || s.branchCode,
        }));
        saveSettings({ pollingStations });
      }

      return true;
    } catch { return false; }
  }

  async function _fetchUsersFromApi() {
    if (!getApiBase() || isJwtExpired()) return false;
    try {
      const res = await apiGet('/users');
      if (!res.success || !Array.isArray(res.users)) return false;

      const remoteUsers = res.users.map(u => ({
        ...u,
        assignedStations: (() => {
          if (Array.isArray(u.assigned_stations)) return u.assigned_stations;
          try { return JSON.parse(u.assigned_stations || '[]'); } catch { return []; }
        })(),
        mustChangePassword: !!u.must_change_password,
        active: u.active !== 0 && u.active !== false,
      }));

      // Safety: never wipe out all admins
      const hasAdmin = remoteUsers.some(u => u.role === 'admin' && u.active);
      if (!hasAdmin) return false;

      // Merge: preserve local user's own local password for offline fallback
      const local = lsGet(LS.USERS, []);
      const localById = Object.fromEntries(local.map(u => [u.id, u]));
      const merged = remoteUsers.map(u =>
        u.id === currentUser?.id
          ? { ...u, password: localById[u.id]?.password }
          : u
      );

      lsSet(LS.USERS, merged);
      await _revalidateSession();
      return true;
    } catch { return false; }
  }

  async function _fetchAndApplyRemoteSettings() {
    if (!getApiBase() || isJwtExpired()) return false;
    try {
      const res = await apiGet('/settings');
      if (!res.success || !res.settings) return false;
      const s = res.settings;
      const patch = {};
      if (s.appName)      patch.appName      = s.appName;
      if (s.constituency) patch.constituency  = s.constituency;
      if (s.pollingStations) {
        try {
          const ps = typeof s.pollingStations === 'string'
            ? JSON.parse(s.pollingStations)
            : s.pollingStations;
          if (Array.isArray(ps) && ps.length) {
            patch.pollingStations = ps;
            pollingStations = ps;
          }
        } catch {}
      }
      if (Object.keys(patch).length) saveSettings(patch);
      return true;
    } catch { return false; }
  }

  // Prefetch before login (settings + stations for Quick Connect)
  async function prefetchOnLoad() {
    if (!getApiBase()) return;
    await Promise.allSettled([
      _fetchAndApplyRemoteSettings(),
      apiGet('/stations').then(res => {
        if (res.success && Array.isArray(res.stations)) {
          pollingStations = res.stations.map(s => ({
            code: s.code, zone: s.zone, ward: s.ward,
            name: s.name, branch: s.branch,
            branchCode: s.branch_code || s.branchCode,
          }));
          saveSettings({ pollingStations });
        }
      }),
    ]);
  }

  // Full sync after login
  async function syncAfterLogin() {
    await Promise.allSettled([
      _fetchAndApplyRemoteSettings(),
      _fetchUsersFromApi(),
    ]);
    await fetchFromApi();
    await flushOfflineQueue();
  }

  // Periodic sync timer (every 2 min)
  let _syncTimer = null;
  function startSyncTimer() {
    if (_syncTimer) clearInterval(_syncTimer);
    _syncTimer = setInterval(async () => {
      if (!currentUser) return;
      await fetchFromApi();
      await _fetchUsersFromApi();
    }, 2 * 60 * 1000);
  }
  function stopSyncTimer() {
    if (_syncTimer) { clearInterval(_syncTimer); _syncTimer = null; }
  }

  // ── Force push helpers (Settings page buttons) ────────────
  async function forcePushStationsToApi() {
    if (!getApiBase()) return { success: false, error: 'Not connected' };
    if (isJwtExpired()) loadJwt(); // reload from localStorage
    return apiPost('/stations/save', { stations: pollingStations });
  }

  async function forcePushUsersToApi() {
    if (!getApiBase()) return { success: false, error: 'Not connected' };
    if (isJwtExpired()) loadJwt();
    const users = getUsers().filter(u => !u.isSystem);
    let ok = 0;
    for (const u of users) {
      const res = await apiPost('/users/upsert', u);
      if (res.success) ok++;
    }
    return { success: true, pushed: ok, total: users.length };
  }

  async function bulkPushToApi() {
    if (!getApiBase()) return { success: false, error: 'Not connected' };
    if (isJwtExpired()) loadJwt();
    return apiPost('/members/bulk', { members: getMembers() });
  }

  // ── Audit log ─────────────────────────────────────────────
  function _logAudit(action, details, ref, extra) {
    const entry = {
      action,
      details,
      user:      currentUser?.username || 'system',
      timestamp: new Date().toISOString(),
      reason:    extra || '',
    };
    const log = lsGet(LS.AUDIT, []);
    log.unshift(entry);
    lsSet(LS.AUDIT, log.slice(0, 500));       // keep last 500 locally
    if (getApiBase() && !isJwtExpired()) {
      apiPost('/audit', entry);               // fire-and-forget
    }
  }

  function getAuditLog() { return lsGet(LS.AUDIT, []); }

  async function fetchAuditFromApi() {
    if (!getApiBase() || isJwtExpired()) return getAuditLog();
    const res = await apiGet('/audit');
    if (res.success && Array.isArray(res.entries)) return res.entries;
    return getAuditLog();
  }

  // ── Change password ───────────────────────────────────────
  async function changePassword(currentPassword, newPassword) {
    const base = getApiBase();
    if (base && !isJwtExpired()) {
      return apiPost('/change-password', { currentPassword, newPassword });
    }
    // Offline fallback
    const users = getUsers();
    const idx = users.findIndex(u => u.id === currentUser?.id);
    if (idx < 0) return { success: false, error: 'User not found' };
    if (users[idx].password !== currentPassword)
      return { success: false, error: 'Current password incorrect' };
    users[idx].password = newPassword;
    users[idx].mustChangePassword = false;
    lsSet(LS.USERS, users);
    currentUser = { ...currentUser, mustChangePassword: false };
    sessionStorage.setItem(LS.SESSION, JSON.stringify(currentUser));
    return { success: true };
  }

  // ── Save user (from Users management page) ────────────────
  async function saveUser(userObj) {
    const users = getUsers();
    const idx = users.findIndex(u => u.id === userObj.id);
    if (idx >= 0) users[idx] = { ...users[idx], ...userObj };
    else users.push(userObj);
    lsSet(LS.USERS, users);

    if (getApiBase() && !isJwtExpired()) {
      await apiPost('/users/upsert', userObj);
    }
  }

  async function deleteUser(id) {
    const users = getUsers().filter(u => u.id !== id);
    lsSet(LS.USERS, users);
    if (getApiBase() && !isJwtExpired()) {
      await apiDelete('/users/' + id);
    }
  }

  // ── Toast / UI helpers (stubs — implemented in index.html) ─
  function showToast(msg, type) {
    if (typeof window._showToast === 'function') window._showToast(msg, type);
  }
  function _renderNav()    { if (typeof window._renderNav    === 'function') window._renderNav(); }
  function _renderHeader() { if (typeof window._renderHeader === 'function') window._renderHeader(); }

  // ── Lockout helpers ───────────────────────────────────────
  const MAX_ATTEMPTS = 5;
  const LOCKOUT_MS   = 15 * 60 * 1000;

  function getLockout() { return lsGet(LS.LOCKOUT, { attempts: 0, until: 0 }); }
  function saveLockout(data) { lsSet(LS.LOCKOUT, data); }

  function isLockedOut() {
    const l = getLockout();
    if (l.until && Date.now() < l.until) return true;
    if (l.until && Date.now() >= l.until) saveLockout({ attempts: 0, until: 0 });
    return false;
  }

  function recordFailedAttempt() {
    const l = getLockout();
    l.attempts = (l.attempts || 0) + 1;
    if (l.attempts >= MAX_ATTEMPTS) l.until = Date.now() + LOCKOUT_MS;
    saveLockout(l);
    return MAX_ATTEMPTS - l.attempts;
  }

  function clearLockout() { lsDel(LS.LOCKOUT); }

  // ── Demo data helpers ─────────────────────────────────────
  function isDemoCleared() { return !!lsGet(LS.DEMO_CLEARED, false); }

  function clearDemoData() {
    const members = getMembers().filter(m => !m._demo);
    saveMembers(members);
    lsSet(LS.DEMO_CLEARED, true);
    _logAudit('CLEAR_DEMO', 'Demo data cleared');
  }

  // ── Initialise ────────────────────────────────────────────
  function init() {
    loadSettings();
    loadJwt();
  }

  // ── Public API ────────────────────────────────────────────
  return {
    // Meta
    VERSION, LS, ROLE_PAGES,
    getVersion, getSettings, getPollingStations, getCurrentUser, getJwt,
    getApiBase, isJwtExpired,

    // Init / session
    init, loadSettings, saveSettings, loadJwt,
    login, logout, restoreSession, changePassword,
    prefetchOnLoad, syncAfterLogin,
    startSyncTimer, stopSyncTimer,

    // Members
    getMembers, saveMembers, addMember, updateMember, deleteMember,
    getMembersForUser, canModifyMember, getStationPool, generateMemberId,

    // Users
    getUsers, saveUser, deleteUser,

    // Stations
    getPollingStations,

    // Sync
    fetchFromApi, flushOfflineQueue, getOfflineQueueCount,
    forcePushStationsToApi, forcePushUsersToApi, bulkPushToApi,

    // Audit
    getAuditLog, fetchAuditFromApi,

    // Lockout
    isLockedOut, recordFailedAttempt, clearLockout,

    // Demo
    isDemoCleared, clearDemoData,

    // UI stubs (called from index.html)
    showToast,

    // Low-level API helpers (used by pages.js for settings page)
    apiGet, apiPost, pingApi,
  };
})();
