/* ============================================================
   KNNDCmdb – Page Renderers  v1.2
   ============================================================ */

const PageRenderers = {

  // ══════════════════════════════════════════════════════════
  // DASHBOARD
  // ══════════════════════════════════════════════════════════
  dashboard() {
    const s = App.getStats();
    const u = App.currentUser;
    const roleLabel = { officer:'Data Entry Officer', ward:'Ward Coordinator', exec:'Constituency Executive', admin:'System Administrator' }[u.role] || u.role;
    document.getElementById('dash-welcome').textContent = `Welcome back, ${u.name.split(' ')[0]}`;
    document.getElementById('dash-role').textContent    = roleLabel;
    document.getElementById('dash-total').textContent   = s.total.toLocaleString();
    document.getElementById('dash-today').textContent   = s.today;
    document.getElementById('dash-stations').textContent= s.stations;
    document.getElementById('dash-offline').textContent = App.offlineQueue.length;

    const recent = App.getMembersForUser().slice(0, 8);
    const tbody  = document.getElementById('dash-recent-tbody');
    if (tbody) tbody.innerHTML = recent.length
      ? recent.map(m => `<tr>
          <td><strong>${m.lastName||''}, ${m.firstName||''}</strong> ${m.otherNames||''}</td>
          <td>${m.partyId||'—'}</td>
          <td>${m.station||'—'}</td>
          <td><span class="badge badge-blue">${m.officer||'—'}</span></td>
          <td>${m.timestamp||'—'}</td>
        </tr>`).join('')
      : `<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">No records yet</div></div></td></tr>`;

    setTimeout(() => this._drawBarChart('dash-chart', s.byDay), 100);
  },

  _drawBarChart(canvasId, byDay) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx    = canvas.getContext('2d');
    const labels = Object.keys(byDay);
    const data   = Object.values(byDay);
    const max    = Math.max(...data, 1);
    const W = canvas.width = canvas.offsetWidth || 400;
    const H = canvas.height = 160;
    ctx.clearRect(0, 0, W, H);
    const barW = (W - 80) / labels.length;
    const gap  = barW * 0.28;

    for (let i = 0; i <= 4; i++) {
      const y = 20 + (H - 50) * i / 4;
      ctx.strokeStyle = '#e5e7eb'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(40, y); ctx.lineTo(W - 10, y); ctx.stroke();
      ctx.fillStyle = '#9ca3af'; ctx.font = '10px Inter'; ctx.textAlign = 'right';
      ctx.fillText(Math.round(max - (max * i / 4)), 36, y + 4);
    }
    data.forEach((v, i) => {
      const x  = 40 + i * barW + gap / 2;
      const bw = barW - gap;
      const bh = ((v / max) * (H - 60)) || 2;
      const y  = H - 30 - bh;
      const g  = ctx.createLinearGradient(0, y, 0, H - 30);
      g.addColorStop(0, '#1a6b3a'); g.addColorStop(1, 'rgba(26,107,58,.25)');
      ctx.fillStyle = g;
      if (ctx.roundRect) { ctx.beginPath(); ctx.roundRect(x, y, bw, bh, 3); ctx.fill(); }
      else { ctx.fillRect(x, y, bw, bh); }
      ctx.fillStyle = '#6b7280'; ctx.font = '9px Inter'; ctx.textAlign = 'center';
      const d = new Date(labels[i]);
      ctx.fillText(isNaN(d) ? labels[i] : d.toLocaleDateString('en', { weekday:'short' }), x + bw / 2, H - 12);
      if (v > 0) { ctx.fillStyle = '#1a6b3a'; ctx.font = 'bold 10px Inter'; ctx.fillText(v, x + bw / 2, y - 4); }
    });
  },

  // ══════════════════════════════════════════════════════════
  // DATA ENTRY
  // ══════════════════════════════════════════════════════════
  entry() {
    const u = App.currentUser;
    // If officer has exactly one assigned station, pre-fill it
    if (u.role === 'officer' && u.assignedStations?.length === 1) {
      const s = App.pollingStations.find(ps => ps.code === u.assignedStations[0]);
      if (s) this._fillStationFields(s);
    }
    this._setupBranchCodeSearch(u.role);
  },

  _fillStationFields(s) {
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    set('f-ward-name',    s.ward);
    set('f-station-name', s.name);
    set('f-branch-name',  s.branch);
    set('f-station-code', s.code);
    set('f-branch-code',  s.branchCode);
  },

  _setupBranchCodeSearch(role) {
    const input    = document.getElementById('f-branch-code');
    const dropdown = document.getElementById('branch-code-dropdown');
    if (!input || !dropdown) return;

    // Officers with multiple stations still need to pick; single-station officers are locked
    const isLocked = role === 'officer' && App.currentUser?.assignedStations?.length === 1;
    if (isLocked) {
      input.setAttribute('readonly', true);
      input.classList.add('auto-filled');
      return;
    }

    input.removeAttribute('readonly');
    input.classList.remove('auto-filled');
    input.placeholder = 'Type branch code, ward or station name…';

    // Scope options to assigned stations for officers
    const u = App.currentUser;
    const getOptions = () => {
      if (u.role === 'officer' && u.assignedStations?.length) {
        return App.pollingStations.filter(s => u.assignedStations.includes(s.code));
      }
      return App.pollingStations;
    };

    const showDropdown = (q) => {
      const pool = getOptions();
      const hits = pool.filter(s =>
        !q ||
        s.branchCode?.toLowerCase().includes(q) ||
        s.branch?.toLowerCase().includes(q) ||
        s.name?.toLowerCase().includes(q) ||
        s.ward?.toLowerCase().includes(q) ||
        s.code?.toLowerCase().includes(q)
      ).slice(0, 30);

      dropdown.innerHTML = hits.length
        ? hits.map((s, i) => `
            <div class="dropdown-item" data-sidx="${App.pollingStations.indexOf(s)}">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
                <div>
                  <strong style="color:var(--ndc-green-dk)">${s.branchCode}</strong>
                  <span style="color:var(--gray-500);font-size:11px;margin-left:6px">${s.name}</span>
                </div>
                <span class="badge badge-green" style="font-size:10px">${s.ward||''}</span>
              </div>
              <div style="font-size:11px;color:var(--gray-400);margin-top:2px">
                Branch: ${s.branch} &nbsp;·&nbsp; Code: ${s.code}
              </div>
            </div>`)
          .join('')
        : '<div class="dropdown-item" style="color:var(--gray-400);text-align:center;padding:16px">No matching stations</div>';

      dropdown.classList.add('open');

      dropdown.querySelectorAll('.dropdown-item[data-sidx]').forEach(item => {
        item.onclick = () => {
          const s = App.pollingStations[parseInt(item.dataset.sidx)];
          if (s) { this._fillStationFields(s); Toast.show('Station Selected', `${s.name} · ${s.branchCode}`, 'success'); }
          dropdown.classList.remove('open');
        };
      });
    };

    // Remove previous listeners by cloning
    const newInput = input.cloneNode(true);
    input.parentNode.replaceChild(newInput, input);

    newInput.addEventListener('focus', () => showDropdown(newInput.value.toLowerCase().trim()));
    newInput.addEventListener('input', () => showDropdown(newInput.value.toLowerCase().trim()));
    document.addEventListener('click', e => {
      if (!newInput.contains(e.target) && !dropdown.contains(e.target)) dropdown.classList.remove('open');
    });
    // Show all options after a moment so admin can browse
    if (role !== 'officer') setTimeout(() => showDropdown(''), 150);
  },

  submitEntry() {
    const required = ['f-branch-code','f-station-code','f-last-name','f-first-name','f-party-id'];
    let ok = true;
    required.forEach(id => {
      const el = document.getElementById(id);
      if (el && !el.value.trim()) { el.style.borderColor = 'var(--ndc-red)'; ok = false; }
      else if (el) el.style.borderColor = '';
    });
    if (!ok) { Toast.show('Validation Error','Please fill all required fields and select a station.','error'); return; }

    const g = id => document.getElementById(id)?.value.trim() || '';
    App.addMember({
      ward:        g('f-ward-name'),
      station:     g('f-station-name'),
      stationCode: g('f-station-code'),
      branch:      g('f-branch-name'),
      branchCode:  g('f-branch-code'),
      lastName:    g('f-last-name'),
      firstName:   g('f-first-name'),
      otherNames:  g('f-other-names'),
      partyId:     g('f-party-id'),
      voterId:     g('f-voter-id'),
      phone:       g('f-phone'),
    });
    Toast.show('Record Saved', `${g('f-first-name')} ${g('f-last-name')} registered.`, 'success');
    ['f-last-name','f-first-name','f-other-names','f-party-id','f-voter-id','f-phone']
      .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
    document.getElementById('f-last-name')?.focus();
  },

  // ══════════════════════════════════════════════════════════
  // MY RECORDS
  // ══════════════════════════════════════════════════════════
  'my-records'() {
    const members = App.getMembersForUser();
    document.getElementById('my-records-count').textContent = `${members.length} record${members.length !== 1 ? 's' : ''}`;
    this._renderMembersTable('my-records-tbody', members, false);
  },

  // ══════════════════════════════════════════════════════════
  // ALL RECORDS
  // ══════════════════════════════════════════════════════════
  records() {
    PageRenderers._allState = PageRenderers._allState || { q:'', page:1 };
    this._renderAllRecords();
  },

  _renderAllRecords() {
    const st = PageRenderers._allState;
    let members = App.getMembersForUser();
    if (st.q) {
      const q = st.q.toLowerCase();
      members = members.filter(m =>
        m.firstName?.toLowerCase().includes(q) || m.lastName?.toLowerCase().includes(q) ||
        m.partyId?.toLowerCase().includes(q)   || m.voterId?.toLowerCase().includes(q)  ||
        m.phone?.includes(q)                   || m.station?.toLowerCase().includes(q)  ||
        m.ward?.toLowerCase().includes(q)
      );
    }
    const per   = 20;
    const total = members.length;
    const pages = Math.ceil(total / per) || 1;
    const page  = Math.min(st.page, pages);
    const slice = members.slice((page - 1) * per, page * per);

    document.getElementById('records-count').textContent = `${total.toLocaleString()} record${total !== 1 ? 's' : ''}`;
    this._renderMembersTable('records-tbody', slice, true);

    const pc = document.getElementById('records-pagination');
    if (!pc) return;
    pc.innerHTML = '';
    if (pages > 1) {
      const mk = (lbl, pg, active, disabled) => {
        const b = document.createElement('button');
        b.className = 'page-btn' + (active ? ' active' : '');
        b.textContent = lbl; b.disabled = disabled;
        b.onclick = () => { PageRenderers._allState.page = pg; PageRenderers._renderAllRecords(); };
        return b;
      };
      pc.appendChild(mk('‹', page - 1, false, page === 1));
      for (let i = 1; i <= pages; i++) {
        if (i === 1 || i === pages || Math.abs(i - page) <= 2) pc.appendChild(mk(i, i, i === page));
        else if (Math.abs(i - page) === 3) { const s = document.createElement('span'); s.className = 'page-btn'; s.textContent = '…'; pc.appendChild(s); }
      }
      pc.appendChild(mk('›', page + 1, false, page === pages));
    }
  },

  _renderMembersTable(tbodyId, members, showActions) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    const canEdit = showActions && ['admin','exec','ward'].includes(App.currentUser?.role);
    const canDel  = showActions && App.currentUser?.role === 'admin';
    tbody.innerHTML = members.length
      ? members.map(m => `<tr>
          <td><strong>${m.lastName||''}</strong>, ${m.firstName||''} ${m.otherNames||''}</td>
          <td>${m.partyId||'—'}</td>
          <td>${m.voterId||'—'}</td>
          <td>${m.phone||'—'}</td>
          <td>${m.ward||'—'}</td>
          <td>${m.station||'—'}</td>
          <td>${m.branch||'—'}</td>
          <td>${m.timestamp||'—'}</td>
          ${showActions ? `<td>
            ${canEdit ? `<button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEdit('${m.id}')">✏️ Edit</button>` : ''}
            ${canDel  ? `<button class="btn btn-sm btn-danger"    onclick="PageRenderers.confirmDelete('${m.id}')">🗑️</button>` : ''}
          </td>` : ''}
        </tr>`)
      .join('')
      : `<tr><td colspan="9"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records found</div></div></td></tr>`;
  },

  openEdit(id) {
    const m = App.members.find(x => x.id === id);
    if (!m) return;
    ['edit-id','edit-first','edit-last','edit-other','edit-party','edit-voter','edit-phone','edit-reason'].forEach(fid => {
      const map = { 'edit-id':id,'edit-first':m.firstName,'edit-last':m.lastName,'edit-other':m.otherNames,'edit-party':m.partyId,'edit-voter':m.voterId,'edit-phone':m.phone,'edit-reason':'' };
      const el = document.getElementById(fid); if (el) el.value = map[fid] || '';
    });
    document.getElementById('edit-current-info').textContent = `Editing: ${m.firstName} ${m.lastName} | Party ID: ${m.partyId} | Added: ${m.timestamp}`;
    Modal.open('modal-edit');
  },

  submitEdit() {
    const id     = document.getElementById('edit-id').value;
    const reason = document.getElementById('edit-reason').value.trim();
    if (!reason) { Toast.show('Reason Required','Please provide a reason for this change.','error'); return; }
    App.updateMember(id, {
      firstName: document.getElementById('edit-first').value.trim(),
      lastName:  document.getElementById('edit-last').value.trim(),
      otherNames:document.getElementById('edit-other').value.trim(),
      partyId:   document.getElementById('edit-party').value.trim(),
      voterId:   document.getElementById('edit-voter').value.trim(),
      phone:     document.getElementById('edit-phone').value.trim(),
    }, reason);
    Modal.close('modal-edit');
    Toast.show('Record Updated','Changes saved.','success');
    PageRenderers.records();
  },

  confirmDelete(id) {
    const m = App.members.find(x => x.id === id);
    document.getElementById('del-id').value  = id;
    document.getElementById('del-name').textContent = m ? `${m.firstName} ${m.lastName} (${m.partyId})` : id;
    document.getElementById('del-reason').value = '';
    Modal.open('modal-delete');
  },

  submitDelete() {
    const id     = document.getElementById('del-id').value;
    const reason = document.getElementById('del-reason').value.trim();
    if (!reason) { Toast.show('Reason Required','Please provide a reason.','error'); return; }
    App.deleteMember(id, reason);
    Modal.close('modal-delete');
    Toast.show('Record Deleted','Member removed.','error');
    PageRenderers.records();
  },

  // ══════════════════════════════════════════════════════════
  // REPORTS
  // ══════════════════════════════════════════════════════════
  reports() {
    const members = App.getMembersForUser();
    const byStation = {};
    members.forEach(m => {
      if (!byStation[m.station]) byStation[m.station] = { station:m.station, branch:m.branch||'', ward:m.ward||'', count:0 };
      byStation[m.station].count++;
    });
    const rows = Object.values(byStation).sort((a,b) => b.count - a.count);
    const tbody = document.getElementById('report-tbody');
    if (tbody) tbody.innerHTML = rows.length
      ? rows.map((r, i) => `<tr>
          <td>${i+1}</td>
          <td>${r.ward||'—'}</td>
          <td><strong>${r.station}</strong></td>
          <td>${r.branch}</td>
          <td>${r.count}</td>
          <td><div class="progress-bar" style="width:100px"><div class="progress-fill" style="width:${members.length?Math.round(r.count/members.length*100):0}%"></div></div></td>
          <td>${members.length ? Math.round(r.count/members.length*100) : 0}%</td>
        </tr>`).join('')
      : `<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--gray-400)">No data available</td></tr>`;

    document.getElementById('report-total').textContent = members.length.toLocaleString();
    const today = new Date().toLocaleDateString('en-GH');
    document.getElementById('daily-count').textContent = members.filter(m => m.timestamp?.includes(today)).length;
    document.getElementById('daily-date').textContent  = today;
    setTimeout(() => this._drawPieChart(byStation, members.length), 100);
  },

  _drawPieChart(byStation, total) {
    const canvas = document.getElementById('report-chart');
    if (!canvas || !total) return;
    const ctx = canvas.getContext('2d');
    const W = canvas.width = canvas.offsetWidth || 280;
    const H = canvas.height = 200;
    ctx.clearRect(0, 0, W, H);
    const cx = W / 2, cy = H / 2, r = Math.min(cx, cy) - 20;
    const colors = ['#1a6b3a','#c8102e','#2563eb','#d97706','#7c3aed','#10b981','#f59e0b','#ef4444','#8b5cf6'];
    let angle = -Math.PI / 2;
    Object.values(byStation).forEach((s, i) => {
      const slice = (s.count / total) * Math.PI * 2;
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, angle, angle + slice); ctx.closePath();
      ctx.fillStyle = colors[i % colors.length]; ctx.fill();
      angle += slice;
    });
    ctx.beginPath(); ctx.arc(cx, cy, r * 0.55, 0, Math.PI * 2); ctx.fillStyle = 'white'; ctx.fill();
    ctx.fillStyle = '#1f2937'; ctx.font = 'bold 20px Outfit'; ctx.textAlign = 'center';
    ctx.fillText(total, cx, cy + 4);
    ctx.fillStyle = '#6b7280'; ctx.font = '10px Inter'; ctx.fillText('Total', cx, cy + 18);
  },

  exportExcel() {
    const members      = App.getMembersForUser();
    const constituency = App.settings.constituency || 'Ketu North';

    if (typeof XLSX === 'undefined') {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      s.onload = () => PageRenderers.exportExcel();
      document.head.appendChild(s);
      Toast.show('Loading','Preparing export…','info'); return;
    }

    const wb  = XLSX.utils.book_new();
    const hdr = ['First Name','Surname','Party ID Number','Voter ID Number','Telephone Number'];

    // ── All Members sheet ──
    const allRows = [
      ['MEMBERSHIP DATABASE','','','',''],
      [`Polling Station / Branch Name: ALL STATIONS`,'','','',`Constituency: ${constituency}`],
      ['','','','',''],
      hdr,
      ...members.map(m => [m.firstName, m.lastName, m.partyId, m.voterId, m.phone])
    ];
    const wsAll = XLSX.utils.aoa_to_sheet(allRows);
    wsAll['!cols'] = [{wch:18},{wch:18},{wch:20},{wch:20},{wch:18}];
    wsAll['!merges'] = [{ s:{r:0,c:0}, e:{r:0,c:3} }, { s:{r:1,c:0}, e:{r:1,c:3} }];
    XLSX.utils.book_append_sheet(wb, wsAll, 'All Members');

    // ── Per-station sheets ──
    const byStation = {};
    members.forEach(m => {
      const key = m.stationCode || 'UNK';
      if (!byStation[key]) byStation[key] = { info:m, members:[] };
      byStation[key].members.push(m);
    });
    Object.values(byStation).forEach(({ info, members: sm }) => {
      const name = (info.station||'Station').replace(/[\\\/\?\*\[\]:]/g,'').substring(0,31);
      const rows = [
        ['MEMBERSHIP DATABASE','','','',''],
        [`Polling Station / Branch Name: ${info.station||''} / ${info.branch||''}`,'','','',`Constituency: ${constituency}`],
        ['','','','',''],
        hdr,
        ...sm.map(m => [m.firstName, m.lastName, m.partyId, m.voterId, m.phone])
      ];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols']   = [{wch:18},{wch:18},{wch:20},{wch:20},{wch:18}];
      ws['!merges'] = [{ s:{r:0,c:0}, e:{r:0,c:3} }, { s:{r:1,c:0}, e:{r:1,c:3} }];
      XLSX.utils.book_append_sheet(wb, ws, name);
    });

    // ── Summary sheet ──
    const sumRows = [
      ['MEMBERSHIP DATABASE — SUMMARY','','',''],
      [`Constituency: ${constituency}`,'','',`Export Date: ${new Date().toLocaleDateString('en-GH')}`],
      ['','','',''],
      ['Ward','Polling Station','Branch','Total Members'],
      ...Object.values(byStation).map(({ info, members: sm }) => [info.ward||'—', info.station||'—', info.branch||'—', sm.length]),
      ['','','GRAND TOTAL', members.length],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(sumRows);
    wsSummary['!cols'] = [{wch:20},{wch:28},{wch:22},{wch:15}];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    XLSX.writeFile(wb, `KNNDCmdb_${new Date().toISOString().slice(0,10)}.xlsx`);
    Toast.show('Export Ready',`${members.length} records exported.`,'success');
    App.logAudit('EXPORT_EXCEL',`Exported ${members.length} records to Excel`, App.currentUser.username);
  },

  exportPDF() {
    window.print();
    App.logAudit('EXPORT_PDF','Printed/exported report as PDF', App.currentUser.username);
  },

  // ══════════════════════════════════════════════════════════
  // ANALYTICS
  // ══════════════════════════════════════════════════════════
  analytics() {
    const members = App.getMembersForUser();
    const stats   = App.getStats();
    document.getElementById('ana-total').textContent    = members.length.toLocaleString();
    document.getElementById('ana-today').textContent    = stats.today;
    document.getElementById('ana-stations').textContent = stats.stations;

    const yest = new Date(); yest.setDate(yest.getDate()-1);
    const yKey = yest.toLocaleDateString('en-GH');
    const yCount = members.filter(m => m.timestamp?.includes(yKey)).length;
    const growth = yCount ? ((stats.today - yCount) / yCount * 100).toFixed(0) : stats.today > 0 ? '∞' : 0;
    document.getElementById('ana-growth').textContent = (typeof growth === 'number' && growth > 0 ? '+' : '') + growth + (growth !== '∞' ? '%' : '');

    setTimeout(() => {
      this._drawBarChart('ana-chart', stats.byDay);
      this._drawPieChart(stats.byStation, members.length);
    }, 100);

    const byOfficer = {};
    members.forEach(m => { byOfficer[m.officer] = (byOfficer[m.officer]||0)+1; });
    const top = Object.entries(byOfficer).sort((a,b)=>b[1]-a[1]).slice(0,8);
    const t = document.getElementById('ana-officers-tbody');
    if (t) t.innerHTML = top.length
      ? top.map(([name,count]) => `<tr>
          <td>${name}</td><td>${count}</td>
          <td><div class="progress-bar"><div class="progress-fill" style="width:${members.length?Math.round(count/members.length*100):0}%"></div></div></td>
          <td>${members.length?Math.round(count/members.length*100):0}%</td>
        </tr>`).join('')
      : '<tr><td colspan="4" style="text-align:center">No data</td></tr>';
  },

  // ══════════════════════════════════════════════════════════
  // AUDIT LOG
  // ══════════════════════════════════════════════════════════
  audit() {
    const q      = document.getElementById('audit-search')?.value?.toLowerCase() || '';
    const filter = document.getElementById('audit-filter')?.value || '';
    // Fresh read from storage
    App.auditLog = JSON.parse(localStorage.getItem('knndc_audit') || '[]');
    let log = App.auditLog;
    if (q)      log = log.filter(e => e.action?.toLowerCase().includes(q) || e.details?.toLowerCase().includes(q) || e.user?.toLowerCase().includes(q));
    if (filter) log = log.filter(e => e.action === filter);

    const container = document.getElementById('audit-entries');
    if (!container) return;
    container.innerHTML = log.slice(0, 300).map(e => {
      const isDanger  = ['DELETE_MEMBER','LOCKOUT','DISABLE_USER'].includes(e.action);
      const isWarning = ['EDIT_MEMBER','EXPORT_EXCEL','EXPORT_PDF','AUTO_LOGOUT'].includes(e.action);
      return `<div class="log-entry ${isDanger?'danger':isWarning?'warning':''}">
        <div class="log-entry-header">
          <span class="badge ${isDanger?'badge-red':isWarning?'badge-amber':'badge-green'}">${e.action}</span>
          <span class="log-user">👤 ${e.user}</span>
          <span class="log-time">${e.timestamp}</span>
        </div>
        <div class="log-details">${e.details||''}</div>
        ${e.reason?`<div class="log-details" style="color:var(--ndc-red);margin-top:4px">⚠️ Reason: ${e.reason}</div>`:''}
      </div>`;
    }).join('') || `<div class="empty-state"><div class="empty-icon">🛡️</div><div class="empty-title">No audit entries found</div></div>`;
  },

  // ══════════════════════════════════════════════════════════
  // USER MANAGEMENT  — with passwords column + assignment modal
  // ══════════════════════════════════════════════════════════
  users() {
    const tbody = document.getElementById('users-tbody');
    if (!tbody) return;
    const roleLabels  = { officer:'Data Entry Officer', ward:'Ward Coordinator', exec:'Constituency Exec', admin:'System Admin' };
    const roleBadges  = { officer:'badge-green', ward:'badge-amber', exec:'badge-blue', admin:'badge-red' };
    // Reload users fresh
    App.users = JSON.parse(localStorage.getItem('knndc_users') || 'null') || DEFAULT_USERS;

    tbody.innerHTML = App.users.map(u => `
      <tr>
        <td>
          <div style="display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:var(--ndc-green);color:white;font-weight:700;font-size:12px;font-family:var(--font-head)">
            ${u.name.split(' ').map(n=>n[0]).slice(0,2).join('')}
          </div>
        </td>
        <td><strong>${u.name}</strong><br><small style="color:var(--gray-400)">${u.username}</small></td>
        <td><span class="badge ${roleBadges[u.role]||'badge-gray'}">${roleLabels[u.role]||u.role}</span></td>
        <td>
          <div style="display:flex;align-items:center;gap:6px">
            <code id="pwd-${u.id}" style="font-family:monospace;font-size:12px;background:var(--gray-100);padding:3px 8px;border-radius:4px;letter-spacing:2px">
              ${'•'.repeat(Math.min(u.password?.length||8, 10))}
            </code>
            <button class="btn btn-sm btn-secondary" style="padding:3px 8px;font-size:11px" onclick="PageRenderers.togglePwd('${u.id}','${u.password?.replace(/'/g,"\\'")}')" title="Show/hide password">👁️</button>
          </div>
        </td>
        <td>${u.ward||'—'}</td>
        <td>
          ${(u.assignedStations||[]).length
            ? (u.assignedStations||[]).map(c=>`<span class="badge badge-blue" style="font-size:10px;margin:1px">${c}</span>`).join('')
            : `<span style="color:var(--gray-400);font-size:12px">None assigned</span>`}
        </td>
        <td><span class="badge ${u.active?'badge-green':'badge-gray'}">${u.active?'Active':'Inactive'}</span></td>
        <td>
          <div class="btn-group" style="flex-wrap:nowrap">
            <button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEditUser('${u.id}')">✏️ Edit</button>
            ${u.role==='officer'?`<button class="btn btn-sm btn-outline" onclick="PageRenderers.openAssignModal('${u.id}')">📍 Assign</button>`:''}
            <button class="btn btn-sm btn-danger"    onclick="PageRenderers.toggleUser('${u.id}')">${u.active?'🚫':'✅'}</button>
          </div>
        </td>
      </tr>`).join('');
  },

  togglePwd(uid, pwd) {
    const el = document.getElementById('pwd-' + uid);
    if (!el) return;
    const isHidden = el.textContent.includes('•');
    el.textContent = isHidden ? pwd : '•'.repeat(Math.min(pwd?.length||8, 10));
    el.style.letterSpacing = isHidden ? 'normal' : '2px';
  },

  openAddUser() {
    document.getElementById('user-modal-title').textContent = 'Add New User';
    document.getElementById('edit-user-id').value = '';
    ['u-name','u-username','u-password','u-ward','u-station','u-branch'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
    document.getElementById('u-role').value = 'officer';
    Modal.open('modal-user');
  },

  openEditUser(id) {
    const u = App.users.find(x => x.id === id);
    if (!u) return;
    document.getElementById('user-modal-title').textContent = 'Edit User';
    document.getElementById('edit-user-id').value = id;
    document.getElementById('u-name').value        = u.name;
    document.getElementById('u-username').value    = u.username;
    document.getElementById('u-password').value    = u.password;
    document.getElementById('u-role').value        = u.role;
    document.getElementById('u-ward').value        = u.ward    || '';
    document.getElementById('u-station').value     = u.station || '';
    document.getElementById('u-branch').value      = u.branch  || '';
    Modal.open('modal-user');
  },

  submitUser() {
    const id = document.getElementById('edit-user-id').value;
    const data = {
      name:     document.getElementById('u-name').value.trim(),
      username: document.getElementById('u-username').value.trim(),
      password: document.getElementById('u-password').value.trim(),
      role:     document.getElementById('u-role').value,
      ward:     document.getElementById('u-ward').value.trim(),
      station:  document.getElementById('u-station').value.trim(),
      branch:   document.getElementById('u-branch').value.trim(),
      active:   true,
    };
    if (!data.name || !data.username || !data.password) { Toast.show('Error','Name, username and password are required.','error'); return; }
    App.users = JSON.parse(localStorage.getItem('knndc_users') || 'null') || DEFAULT_USERS;
    if (id) {
      const idx = App.users.findIndex(u => u.id === id);
      App.users[idx] = { ...App.users[idx], ...data };
      App.logAudit('EDIT_USER', `Edited user: ${data.username} (${data.role})`, App.currentUser.username);
      Toast.show('User Updated','Changes saved.','success');
    } else {
      if (App.users.find(u => u.username === data.username)) { Toast.show('Error','Username already exists.','error'); return; }
      App.users.push({ id:'u'+Date.now(), ...data, assignedStations:[] });
      App.logAudit('ADD_USER', `Created user: ${data.username} (${data.role})`, App.currentUser.username);
      Toast.show('User Created','New user added.','success');
    }
    App.saveUsers();
    Modal.close('modal-user');
    PageRenderers.users();
  },

  toggleUser(id) {
    App.users = JSON.parse(localStorage.getItem('knndc_users') || 'null') || DEFAULT_USERS;
    const u = App.users.find(x => x.id === id);
    if (!u) return;
    u.active = !u.active;
    App.saveUsers();
    App.logAudit(u.active?'ENABLE_USER':'DISABLE_USER', `${u.active?'Enabled':'Disabled'} user: ${u.username}`, App.currentUser.username);
    Toast.show('Status Updated',`${u.name} is now ${u.active?'active':'inactive'}.`, u.active?'success':'warning');
    PageRenderers.users();
  },

  // ── ASSIGNMENT MODAL — checkbox-based station picker ──
  openAssignModal(userId) {
    App.users = JSON.parse(localStorage.getItem('knndc_users') || 'null') || DEFAULT_USERS;
    const u = App.users.find(x => x.id === userId);
    if (!u) return;

    document.getElementById('assign-user-id').value = userId;
    document.getElementById('assign-user-name').textContent  = u.name;
    document.getElementById('assign-user-role').textContent  = 'Data Entry Officer';

    const assigned = u.assignedStations || [];
    const stations = App.pollingStations;

    // Group by ward
    const byWard = {};
    stations.forEach(s => { if (!byWard[s.ward]) byWard[s.ward] = []; byWard[s.ward].push(s); });

    const container = document.getElementById('assign-stations-container');
    container.innerHTML = Object.entries(byWard).map(([ward, sts]) => `
      <div style="margin-bottom:18px">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:700;color:var(--ndc-green-dk);text-transform:none;letter-spacing:0;font-size:13px">
            <input type="checkbox" class="ward-check" data-ward="${ward}"
              ${sts.every(s=>assigned.includes(s.code))?'checked':''}
              onchange="PageRenderers.toggleWardCheck('${ward}',this.checked)"
              style="width:auto;accent-color:var(--ndc-green)">
            🏘️ ${ward}
          </label>
          <span class="badge badge-green" style="font-size:10px">${sts.length} station${sts.length!==1?'s':''}</span>
        </div>
        <div style="padding-left:20px;display:grid;grid-template-columns:1fr 1fr;gap:6px">
          ${sts.map(s => `
            <label style="display:flex;align-items:flex-start;gap:8px;cursor:pointer;padding:8px 10px;border:1px solid var(--gray-200);border-radius:var(--radius);background:white;transition:var(--transition);text-transform:none;letter-spacing:0;font-weight:400;font-size:12px"
              onmouseover="this.style.background='var(--ndc-green-pale)'" onmouseout="this.style.background='white'">
              <input type="checkbox" class="station-check" name="assign_stations" value="${s.code}"
                data-ward="${ward}" ${assigned.includes(s.code)?'checked':''}
                onchange="PageRenderers.updateWardCheckState('${ward}')"
                style="width:auto;accent-color:var(--ndc-green);margin-top:2px">
              <div>
                <div style="font-weight:600;color:var(--gray-800)">${s.name}</div>
                <div style="color:var(--gray-400)">${s.code} · ${s.branchCode}</div>
              </div>
            </label>`).join('')}
        </div>
      </div>`
    ).join('') || '<div class="empty-state"><div class="empty-icon">🏢</div><div class="empty-title">No polling stations configured</div><div class="empty-text">Add stations in Settings first.</div></div>';

    Modal.open('modal-assign');
  },

  toggleWardCheck(ward, checked) {
    document.querySelectorAll(`.station-check[data-ward="${ward}"]`).forEach(cb => cb.checked = checked);
  },

  updateWardCheckState(ward) {
    const all = document.querySelectorAll(`.station-check[data-ward="${ward}"]`);
    const wardCb = document.querySelector(`.ward-check[data-ward="${ward}"]`);
    if (!wardCb) return;
    const checkedCount = [...all].filter(cb => cb.checked).length;
    wardCb.checked       = checkedCount === all.length;
    wardCb.indeterminate = checkedCount > 0 && checkedCount < all.length;
  },

  submitAssignment() {
    const userId = document.getElementById('assign-user-id').value;
    const selected = [...document.querySelectorAll('.station-check:checked')].map(cb => cb.value);

    App.users = JSON.parse(localStorage.getItem('knndc_users') || 'null') || DEFAULT_USERS;
    const u = App.users.find(x => x.id === userId);
    if (!u) return;

    const prev = u.assignedStations || [];
    u.assignedStations = selected;

    // Also update primary station to first assigned (for backward compat)
    if (selected.length > 0) {
      const primary = App.pollingStations.find(s => s.code === selected[0]);
      if (primary) { u.station = primary.code; u.ward = primary.ward; u.branch = primary.branch; }
    }

    App.saveUsers();
    App.logAudit('ASSIGN_STATIONS', `Assigned ${selected.length} station(s) to ${u.username}: [${selected.join(', ')}]. Previous: [${prev.join(', ')}]`, App.currentUser.username);
    Modal.close('modal-assign');
    Toast.show('Assignment Saved', `${u.name} assigned to ${selected.length} station${selected.length!==1?'s':''}.`, 'success');
    PageRenderers.users();
  },

  // ══════════════════════════════════════════════════════════
  // SETTINGS
  // ══════════════════════════════════════════════════════════
  settings() {
    const s = App.settings;
    ['set-app-name','set-constituency','set-sheet-id','set-api-key','set-script-url'].forEach(id => {
      const map = { 'set-app-name':s.appName,'set-constituency':s.constituency,'set-sheet-id':s.sheetId,'set-api-key':s.apiKey,'set-script-url':s.scriptUrl };
      const el = document.getElementById(id); if (el) el.value = map[id] || '';
    });
    PageRenderers._renderStationsList();
  },

  saveGeneralSettings() {
    App.settings.appName      = document.getElementById('set-app-name').value.trim()      || CONFIG.APP_NAME;
    App.settings.constituency = document.getElementById('set-constituency').value.trim();
    App.settings.sheetId      = document.getElementById('set-sheet-id').value.trim();
    App.settings.apiKey       = document.getElementById('set-api-key').value.trim();
    App.settings.scriptUrl    = document.getElementById('set-script-url').value.trim();
    App.saveSettings();
    App.applyAppName();
    App.logAudit('SETTINGS_CHANGE','Updated general settings', App.currentUser.username);
    Toast.show('Settings Saved','Configuration updated.','success');
  },

  _renderStationsList() {
    const c = document.getElementById('stations-list');
    if (!c) return;
    if (!App.pollingStations.length) {
      c.innerHTML = '<div class="empty-state"><div class="empty-icon">🏢</div><div class="empty-title">No stations configured</div><div class="empty-text">Add using the form above.</div></div>';
      return;
    }
    c.innerHTML = `<div class="table-wrap">
      <table>
        <thead><tr><th>#</th><th>Ward Name</th><th>Polling Station Name</th><th>Branch Name</th><th>Station Code</th><th>Branch Code</th><th>Action</th></tr></thead>
        <tbody>
          ${App.pollingStations.map((s,i) => `<tr>
            <td>${i+1}</td>
            <td>${s.ward||'—'}</td>
            <td><strong>${s.name}</strong></td>
            <td>${s.branch}</td>
            <td><span class="badge badge-blue">${s.code}</span></td>
            <td><span class="badge badge-green">${s.branchCode}</span></td>
            <td><button class="btn btn-sm btn-danger" onclick="PageRenderers.removeStation(${i})">🗑️ Remove</button></td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  },

  addStation() {
    const g = id => document.getElementById(id)?.value.trim() || '';
    const ward = g('st-ward'), name = g('st-name'), code = g('st-code'), branch = g('st-branch'), bCode = g('st-bcode');
    if (!ward||!name||!code||!branch||!bCode) { Toast.show('Error','All 5 fields are required.','error'); return; }
    if (App.pollingStations.find(s => s.code === code)) { Toast.show('Duplicate','Station Code already exists.','error'); return; }
    App.pollingStations.push({ ward, name, code, branch, branchCode:bCode });
    App.settings.pollingStations = App.pollingStations;
    App.saveSettings();
    ['st-ward','st-name','st-code','st-branch','st-bcode'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
    PageRenderers._renderStationsList();
    Toast.show('Station Added',`${name} (${code}) saved.`,'success');
    App.logAudit('ADD_STATION',`Added station: ${name} (${code}), Ward: ${ward}`, App.currentUser.username);
  },

  removeStation(i) {
    const s = App.pollingStations[i];
    App.pollingStations.splice(i, 1);
    App.settings.pollingStations = App.pollingStations;
    App.saveSettings();
    PageRenderers._renderStationsList();
    Toast.show('Station Removed',`${s.name} removed.`,'warning');
    App.logAudit('REMOVE_STATION',`Removed: ${s.name} (${s.code})`, App.currentUser.username);
  },
};
