/* ============================================================
   KNNDCmdb – Page Renderers
   ============================================================ */

const PageRenderers = {

  // ─── DASHBOARD ─────────────────────────────────────────────
  dashboard() {
    const s = App.getStats();
    const u = App.currentUser;
    const roleLabel = {officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Executive',admin:'System Administrator'}[u.role]||u.role;

    document.getElementById('dash-welcome').textContent = `Welcome back, ${u.name.split(' ')[0]}`;
    document.getElementById('dash-role').textContent = roleLabel;
    document.getElementById('dash-total').textContent = s.total.toLocaleString();
    document.getElementById('dash-today').textContent = s.today;
    document.getElementById('dash-stations').textContent = s.stations;
    document.getElementById('dash-offline').textContent = App.offlineQueue.length;

    // Recent entries table
    const recent = App.getMembersForUser().slice(0,8);
    const tbody = document.getElementById('dash-recent-tbody');
    if (tbody) {
      tbody.innerHTML = recent.length ? recent.map(m =>
        `<tr>
          <td><strong>${m.lastName}, ${m.firstName}</strong>${m.otherNames?' '+m.otherNames:''}</td>
          <td>${m.partyId||'—'}</td>
          <td>${m.station||'—'}</td>
          <td><span class="badge badge-blue">${m.officer}</span></td>
          <td>${m.timestamp||'—'}</td>
        </tr>`
      ).join('') : `<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">No records yet</div></div></td></tr>`;
    }

    // Chart
    setTimeout(() => this._drawDayChart(s.byDay), 100);
  },

  _drawDayChart(byDay) {
    const canvas = document.getElementById('dash-chart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const labels = Object.keys(byDay);
    const data   = Object.values(byDay);
    const max    = Math.max(...data, 1);
    const W = canvas.width = canvas.offsetWidth;
    const H = canvas.height = 160;
    ctx.clearRect(0,0,W,H);

    const barW  = (W-80) / labels.length;
    const barGap= barW * 0.25;

    // Grid lines
    for (let i=0;i<=4;i++) {
      const y = 20 + (H-50) * i/4;
      ctx.strokeStyle = '#e5e7eb'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(40,y); ctx.lineTo(W-10,y); ctx.stroke();
      ctx.fillStyle='#9ca3af'; ctx.font='10px Inter'; ctx.textAlign='right';
      ctx.fillText(Math.round(max-(max*i/4)), 36, y+4);
    }

    data.forEach((v,i) => {
      const x = 40 + i*barW + barGap/2;
      const bw = barW - barGap;
      const bh = ((v/max) * (H-60)) || 2;
      const y  = H-30-bh;
      // Gradient
      const grd = ctx.createLinearGradient(0,y,0,H-30);
      grd.addColorStop(0,'#1a6b3a');
      grd.addColorStop(1,'rgba(26,107,58,.3)');
      ctx.fillStyle = grd;
      ctx.beginPath();
      ctx.roundRect?.(x,y,bw,bh,3);
      ctx.fill();
      // Label
      ctx.fillStyle='#6b7280'; ctx.font='9px Inter'; ctx.textAlign='center';
      const d = new Date(labels[i].split('/').reverse().join('-'));
      ctx.fillText(isNaN(d)?labels[i]:d.toLocaleDateString('en',{weekday:'short'}), x+bw/2, H-12);
      if (v>0) { ctx.fillStyle='#1a6b3a'; ctx.font='bold 10px Inter'; ctx.fillText(v, x+bw/2, y-4); }
    });
  },

  // ─── DATA ENTRY ────────────────────────────────────────────
  entry() {
    const u = App.currentUser;
    const stationInfo = App.pollingStations.find(s => s.code === u.station) || null;

    // For officers: auto-fill from their assigned station (locked)
    // For admin: branch code search box enabled
    if (stationInfo && u.role === 'officer') {
      this._fillStationFields(stationInfo);
    }
    this._setupBranchCodeSearch(u.role);
  },

  _fillStationFields(s) {
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    set('f-ward-name',    s.ward        || '');
    set('f-station-name', s.name        || '');
    set('f-branch-name',  s.branch      || '');
    set('f-station-code', s.code        || '');
    set('f-branch-code',  s.branchCode  || '');
  },

  _clearStationFields() {
    ['f-ward-name','f-station-name','f-branch-name','f-station-code','f-branch-code']
      .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  },

  _setupBranchCodeSearch(role) {
    const input    = document.getElementById('f-branch-code');
    const dropdown = document.getElementById('branch-code-dropdown');
    if (!input || !dropdown) return;

    // Officers: field is locked (auto-filled). Admins/exec: searchable.
    if (role === 'officer') {
      input.setAttribute('readonly', true);
      input.classList.add('auto-filled');
      return;
    }

    // Make editable and attach search
    input.removeAttribute('readonly');
    input.classList.remove('auto-filled');
    input.placeholder = 'Type branch code to search…';

    const showDropdown = (q) => {
      const matches = App.pollingStations.filter(s =>
        !q ||
        s.branchCode?.toLowerCase().includes(q) ||
        s.branch?.toLowerCase().includes(q) ||
        s.name?.toLowerCase().includes(q) ||
        s.ward?.toLowerCase().includes(q)
      ).slice(0, 20);

      dropdown.innerHTML = matches.length
        ? matches.map(s =>
            `<div class="dropdown-item" data-idx="${App.pollingStations.indexOf(s)}">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
                <div>
                  <strong style="color:var(--ndc-green-dk)">${s.branchCode}</strong>
                  <span style="color:var(--gray-500);font-size:11px;margin-left:6px">${s.name}</span>
                </div>
                <span class="badge badge-green" style="font-size:10px">${s.ward||''}</span>
              </div>
              <div style="font-size:11px;color:var(--gray-400);margin-top:2px">Branch: ${s.branch} &nbsp;·&nbsp; Code: ${s.code}</div>
            </div>`
          ).join('')
        : '<div class="dropdown-item" style="color:var(--gray-400);text-align:center">No matching stations</div>';

      dropdown.classList.toggle('open', true);

      dropdown.querySelectorAll('.dropdown-item[data-idx]').forEach(item => {
        item.addEventListener('click', () => {
          const s = App.pollingStations[parseInt(item.dataset.idx)];
          if (s) {
            PageRenderers._fillStationFields(s);
            Toast.show('Station Selected', `${s.name} · ${s.branchCode}`, 'success');
          }
          dropdown.classList.remove('open');
        });
      });
    };

    // Show all on focus
    input.addEventListener('focus', () => showDropdown(input.value.toLowerCase()));
    input.addEventListener('input', () => showDropdown(input.value.toLowerCase()));
    document.addEventListener('click', e => {
      if (!input.contains(e.target) && !dropdown.contains(e.target)) dropdown.classList.remove('open');
    }, { capture: false });

    // Show all stations initially for admins
    if (role === 'admin' || role === 'exec') {
      setTimeout(() => showDropdown(''), 200);
    }
  },

  submitEntry() {
    const required = ['f-branch-code','f-last-name','f-first-name','f-party-id'];
    let valid = true;
    required.forEach(id => {
      const el = document.getElementById(id);
      if (el && !el.value.trim()) { el.style.borderColor='var(--ndc-red)'; valid=false; }
      else if (el) el.style.borderColor='';
    });
    if (!valid) { Toast.show('Validation Error','Please fill in all required fields.','error'); return; }

    // Also ensure station fields populated
    if (!document.getElementById('f-station-code').value.trim()) {
      Toast.show('Station Required','Please select a polling station via the Branch Code field.','error');
      document.getElementById('f-branch-code').focus();
      return;
    }

    const data = {
      ward:        document.getElementById('f-ward-name').value.trim(),
      station:     document.getElementById('f-station-name').value.trim(),
      stationCode: document.getElementById('f-station-code').value.trim(),
      branch:      document.getElementById('f-branch-name').value.trim(),
      branchCode:  document.getElementById('f-branch-code').value.trim(),
      lastName:    document.getElementById('f-last-name').value.trim(),
      firstName:   document.getElementById('f-first-name').value.trim(),
      otherNames:  document.getElementById('f-other-names').value.trim(),
      partyId:     document.getElementById('f-party-id').value.trim(),
      voterId:     document.getElementById('f-voter-id').value.trim(),
      phone:       document.getElementById('f-phone').value.trim(),
    };

    App.addMember(data);
    Toast.show('Record Saved', `${data.firstName} ${data.lastName} registered successfully.`, 'success');

    // Clear member fields only
    ['f-last-name','f-first-name','f-other-names','f-party-id','f-voter-id','f-phone'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    document.getElementById('f-last-name')?.focus();
  },

  // ─── MY RECORDS ────────────────────────────────────────────
  'my-records'() {
    const members = App.getMembersForUser();
    PageRenderers._renderMembersTable('my-records-tbody', members, false);
    document.getElementById('my-records-count').textContent = `${members.length} record${members.length!==1?'s':''}`;
  },

  // ─── ALL RECORDS ───────────────────────────────────────────
  records() {
    PageRenderers._allRecordsState = PageRenderers._allRecordsState || { q:'', page:1 };
    PageRenderers._renderAllRecords();
  },

  _renderAllRecords() {
    const state = PageRenderers._allRecordsState;
    let members = App.getMembersForUser();
    if (state.q) {
      const q = state.q.toLowerCase();
      members = members.filter(m =>
        m.firstName?.toLowerCase().includes(q) ||
        m.lastName?.toLowerCase().includes(q) ||
        m.partyId?.toLowerCase().includes(q) ||
        m.voterId?.toLowerCase().includes(q) ||
        m.phone?.includes(q) ||
        m.station?.toLowerCase().includes(q)
      );
    }
    const perPage = 20;
    const total   = members.length;
    const pages   = Math.ceil(total/perPage);
    const page    = Math.min(state.page, pages||1);
    const slice   = members.slice((page-1)*perPage, page*perPage);

    document.getElementById('records-count').textContent = `${total.toLocaleString()} record${total!==1?'s':''}`;
    PageRenderers._renderMembersTable('records-tbody', slice, true);

    // Pagination
    const pc = document.getElementById('records-pagination');
    if (pc) {
      pc.innerHTML = '';
      if (pages > 1) {
        const mkBtn = (label,pg,active=false,disabled=false) => {
          const b = document.createElement('button');
          b.className = 'page-btn' + (active?' active':'');
          b.textContent = label;
          b.disabled = disabled;
          b.onclick = () => { PageRenderers._allRecordsState.page=pg; PageRenderers._renderAllRecords(); };
          return b;
        };
        pc.appendChild(mkBtn('‹', page-1, false, page===1));
        for (let i=1;i<=pages;i++) {
          if (i===1||i===pages||Math.abs(i-page)<=2) pc.appendChild(mkBtn(i,i,i===page));
          else if (Math.abs(i-page)===3) { const sp=document.createElement('span'); sp.className='page-btn'; sp.textContent='…'; pc.appendChild(sp); }
        }
        pc.appendChild(mkBtn('›', page+1, false, page===pages));
      }
    }
  },

  _renderMembersTable(tbodyId, members, showActions) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = members.length ? members.map(m =>
      `<tr>
        <td><strong>${m.lastName||''}</strong>, ${m.firstName||''} ${m.otherNames||''}</td>
        <td>${m.partyId||'—'}</td>
        <td>${m.voterId||'—'}</td>
        <td>${m.phone||'—'}</td>
        <td>${m.station||'—'}</td>
        <td>${m.branch||'—'}</td>
        <td>${m.timestamp||'—'}</td>
        ${showActions && ['admin','exec','ward'].includes(App.currentUser?.role)?
        `<td>
          <div class="btn-group">
            <button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEdit('${m.id}')">✏️ Edit</button>
            ${App.currentUser?.role==='admin'?`<button class="btn btn-sm btn-danger" onclick="PageRenderers.confirmDelete('${m.id}')">🗑️</button>`:''}
          </div>
         </td>` : showActions?'<td>—</td>':''}
       </tr>`
    ).join('') :
    `<tr><td colspan="9"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records found</div><div class="empty-text">Try adjusting your search.</div></div></td></tr>`;
  },

  openEdit(id) {
    const m = App.members.find(x=>x.id===id);
    if (!m) return;
    document.getElementById('edit-id').value = id;
    document.getElementById('edit-first').value = m.firstName||'';
    document.getElementById('edit-last').value  = m.lastName||'';
    document.getElementById('edit-other').value = m.otherNames||'';
    document.getElementById('edit-party').value = m.partyId||'';
    document.getElementById('edit-voter').value = m.voterId||'';
    document.getElementById('edit-phone').value = m.phone||'';
    document.getElementById('edit-reason').value= '';
    document.getElementById('edit-current-info').textContent = `Editing: ${m.firstName} ${m.lastName} | Party ID: ${m.partyId} | Added: ${m.timestamp}`;
    Modal.open('modal-edit');
  },

  submitEdit() {
    const id     = document.getElementById('edit-id').value;
    const reason = document.getElementById('edit-reason').value.trim();
    if (!reason) { Toast.show('Reason Required','Please provide a reason for this change.','error'); return; }
    App.updateMember(id, {
      firstName:  document.getElementById('edit-first').value.trim(),
      lastName:   document.getElementById('edit-last').value.trim(),
      otherNames: document.getElementById('edit-other').value.trim(),
      partyId:    document.getElementById('edit-party').value.trim(),
      voterId:    document.getElementById('edit-voter').value.trim(),
      phone:      document.getElementById('edit-phone').value.trim(),
    }, reason);
    Modal.close('modal-edit');
    Toast.show('Record Updated','Changes saved successfully.','success');
    PageRenderers.records();
  },

  confirmDelete(id) {
    const m = App.members.find(x=>x.id===id);
    document.getElementById('del-id').value    = id;
    document.getElementById('del-name').textContent = m ? `${m.firstName} ${m.lastName} (${m.partyId})` : id;
    document.getElementById('del-reason').value = '';
    Modal.open('modal-delete');
  },

  submitDelete() {
    const id     = document.getElementById('del-id').value;
    const reason = document.getElementById('del-reason').value.trim();
    if (!reason) { Toast.show('Reason Required','Please provide a reason for deletion.','error'); return; }
    App.deleteMember(id, reason);
    Modal.close('modal-delete');
    Toast.show('Record Deleted','Member removed from database.','error');
    PageRenderers.records();
  },

  // ─── REPORTS ───────────────────────────────────────────────
  reports() {
    const members = App.getMembersForUser();
    const byStation = {};
    members.forEach(m => {
      if (!byStation[m.station]) byStation[m.station]={'station':m.station,'branch':m.branch||'','count':0};
      byStation[m.station].count++;
    });
    const rows = Object.values(byStation).sort((a,b)=>b.count-a.count);
    const tbody = document.getElementById('report-tbody');
    if (tbody) {
      tbody.innerHTML = rows.length ? rows.map((r,i) =>
        `<tr>
          <td>${i+1}</td>
          <td><strong>${r.station}</strong></td>
          <td>${r.branch}</td>
          <td>${r.count}</td>
          <td>
            <div class="progress-bar" style="width:120px">
              <div class="progress-fill" style="width:${Math.round(r.count/members.length*100)}%"></div>
            </div>
          </td>
          <td>${members.length?Math.round(r.count/members.length*100):0}%</td>
         </tr>`
      ).join('') : `<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--gray-400)">No data</td></tr>`;
    }
    document.getElementById('report-total').textContent = members.length.toLocaleString();

    // Daily report
    const today = new Date().toLocaleDateString('en-GH');
    const todayM = members.filter(m=>m.timestamp?.includes(today));
    document.getElementById('daily-count').textContent = todayM.length;
    document.getElementById('daily-date').textContent  = today;

    setTimeout(() => PageRenderers._drawPieChart(byStation, members.length), 100);
  },

  _drawPieChart(byStation, total) {
    const canvas = document.getElementById('report-chart');
    if (!canvas || !total) return;
    const ctx = canvas.getContext('2d');
    const W = canvas.width = canvas.offsetWidth||300;
    const H = canvas.height = 200;
    ctx.clearRect(0,0,W,H);
    const cx=W/2, cy=H/2, r=Math.min(cx,cy)-20;
    const colors = ['#1a6b3a','#c8102e','#2563eb','#d97706','#7c3aed','#10b981','#f59e0b'];
    let angle = -Math.PI/2;
    Object.values(byStation).forEach((s,i) => {
      const slice = (s.count/total)*Math.PI*2;
      ctx.beginPath(); ctx.moveTo(cx,cy);
      ctx.arc(cx,cy,r,angle,angle+slice);
      ctx.closePath();
      ctx.fillStyle = colors[i%colors.length];
      ctx.fill();
      angle += slice;
    });
    // Donut hole
    ctx.beginPath(); ctx.arc(cx,cy,r*.55,0,Math.PI*2); ctx.fillStyle='white'; ctx.fill();
    ctx.fillStyle='#1f2937'; ctx.font='bold 20px Outfit'; ctx.textAlign='center';
    ctx.fillText(total, cx, cy+4);
    ctx.fillStyle='#6b7280'; ctx.font='10px Inter';
    ctx.fillText('Total', cx, cy+18);
  },

  exportExcel() {
    const members = App.getMembersForUser();
    const constituency = App.settings.constituency || 'Ketu North';

    // Group members by polling station (matching template layout)
    const byStation = {};
    members.forEach(m => {
      const key = m.stationCode || m.station || 'Unknown';
      if (!byStation[key]) byStation[key] = { info: m, members: [] };
      byStation[key].members.push(m);
    });

    // Build workbook using SheetJS (loaded via CDN)
    if (typeof XLSX === 'undefined') {
      // Fallback: load SheetJS then retry
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      script.onload = () => PageRenderers.exportExcel();
      document.head.appendChild(script);
      Toast.show('Loading', 'Preparing export engine…', 'info');
      return;
    }

    const wb = XLSX.utils.book_new();

    // ── Sheet 1: Full Database (all members) ──
    const allData = [
      ['MEMBERSHIP DATABASE', '', '', '', ''],
      [`Polling Station / Branch Name: ALL STATIONS`, '', '', '', `Constituency: ${constituency}`],
      ['', '', '', '', ''],
      ['First Name', 'Surname', 'Party ID Number', 'Voter ID Number', 'Telephone Number'],
      ...members.map(m => [m.firstName, m.lastName, m.partyId, m.voterId, m.phone])
    ];
    const wsAll = XLSX.utils.aoa_to_sheet(allData);
    // Column widths
    wsAll['!cols'] = [{wch:18},{wch:18},{wch:20},{wch:20},{wch:18}];
    // Merge title cell
    wsAll['!merges'] = [
      { s:{r:0,c:0}, e:{r:0,c:3} },
      { s:{r:1,c:0}, e:{r:1,c:3} },
    ];
    XLSX.utils.book_append_sheet(wb, wsAll, 'All Members');

    // ── Sheet 2+: One sheet per polling station ──
    Object.values(byStation).forEach(({ info, members: sMembers }) => {
      const sheetName = (info.station || 'Station').replace(/[\\\/\?\*\[\]:]/g,'').substring(0, 31);
      const stationData = [
        ['MEMBERSHIP DATABASE', '', '', '', ''],
        [`Polling Station / Branch Name: ${info.station || ''} / ${info.branch || ''}`, '', '', '', `Constituency: ${constituency}`],
        ['', '', '', '', ''],
        ['First Name', 'Surname', 'Party ID Number', 'Voter ID Number', 'Telephone Number'],
        ...sMembers.map(m => [m.firstName, m.lastName, m.partyId, m.voterId, m.phone])
      ];
      const ws = XLSX.utils.aoa_to_sheet(stationData);
      ws['!cols'] = [{wch:18},{wch:18},{wch:20},{wch:20},{wch:18}];
      ws['!merges'] = [
        { s:{r:0,c:0}, e:{r:0,c:3} },
        { s:{r:1,c:0}, e:{r:1,c:3} },
      ];
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    // ── Sheet: Summary by Station ──
    const summaryData = [
      ['MEMBERSHIP DATABASE — STATION SUMMARY', '', '', ''],
      [`Constituency: ${constituency}`, '', '', `Export Date: ${new Date().toLocaleDateString('en-GH')}`],
      ['', '', '', ''],
      ['Ward', 'Polling Station', 'Branch', 'Total Members'],
      ...Object.values(byStation).map(({ info, members: sm }) => [
        info.ward || '—', info.station || '—', info.branch || '—', sm.length
      ]),
      ['', '', 'GRAND TOTAL', members.length],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary['!cols'] = [{wch:20},{wch:28},{wch:22},{wch:15}];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    // Write and download
    const filename = `KNNDCmdb_Export_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, filename);
    Toast.show('Export Ready', `${members.length} records exported to Excel.`, 'success');
    App.logAudit('EXPORT_EXCEL', `Exported ${members.length} member records to Excel`, App.currentUser.username);
  },

  exportPDF() {
    window.print();
    App.logAudit('EXPORT_PDF','Exported report to PDF', App.currentUser.username);
  },

  // ─── ANALYTICS ─────────────────────────────────────────────
  analytics() {
    const members = App.getMembersForUser();
    const stats = App.getStats();

    document.getElementById('ana-total').textContent   = members.length.toLocaleString();
    document.getElementById('ana-today').textContent   = stats.today;
    document.getElementById('ana-stations').textContent= stats.stations;

    // Growth rate
    const yesterday = new Date(); yesterday.setDate(yesterday.getDate()-1);
    const yKey = yesterday.toLocaleDateString('en-GH');
    const yCount = members.filter(m=>m.timestamp?.includes(yKey)).length;
    const growth = yCount?((stats.today-yCount)/yCount*100).toFixed(0):0;
    document.getElementById('ana-growth').textContent = (growth>0?'+':'')+growth+'%';

    setTimeout(() => {
      PageRenderers._drawDayChart(stats.byDay);
      PageRenderers._drawPieChart(stats.byStation, members.length);
    }, 100);

    // Top officers
    const byOfficer = {};
    members.forEach(m => { byOfficer[m.officer]=(byOfficer[m.officer]||0)+1; });
    const topOfficers = Object.entries(byOfficer).sort((a,b)=>b[1]-a[1]).slice(0,5);
    const t = document.getElementById('ana-officers-tbody');
    if (t) {
      t.innerHTML = topOfficers.map(([name,count]) =>
        `<tr><td>${name}</td><td>${count}</td>
          <td><div class="progress-bar"><div class="progress-fill" style="width:${Math.round(count/members.length*100)}%"></div></div></td>
          <td>${Math.round(count/members.length*100)}%</td></tr>`
      ).join('') || '<tr><td colspan="4">No data</td></tr>';
    }
  },

  // ─── AUDIT LOG ─────────────────────────────────────────────
  audit() {
    const q      = document.getElementById('audit-search')?.value?.toLowerCase()||'';
    const filter = document.getElementById('audit-filter')?.value||'';
    let log = App.auditLog;
    if (q) log = log.filter(e => e.action?.toLowerCase().includes(q)||e.details?.toLowerCase().includes(q)||e.user?.toLowerCase().includes(q));
    if (filter) log = log.filter(e => e.action===filter);

    const container = document.getElementById('audit-entries');
    if (!container) return;
    container.innerHTML = log.length ? log.slice(0,200).map(e => {
      const isDanger  = ['DELETE_MEMBER','LOGOUT'].includes(e.action);
      const isWarning = ['EDIT_MEMBER','EXPORT','EXPORT_PDF'].includes(e.action);
      return `<div class="log-entry ${isDanger?'danger':isWarning?'warning':''}">
        <div class="log-entry-header">
          <span class="badge ${isDanger?'badge-red':isWarning?'badge-amber':'badge-green'}">${e.action}</span>
          <span class="log-user">👤 ${e.user}</span>
          <span class="log-time">${e.timestamp}</span>
        </div>
        <div class="log-details">${e.details||''}</div>
        ${e.reason?`<div class="log-details" style="color:var(--ndc-red);margin-top:4px">Reason: ${e.reason}</div>`:''}
      </div>`;
    }).join('') :
    `<div class="empty-state"><div class="empty-icon">🛡️</div><div class="empty-title">No audit entries</div></div>`;
  },

  // ─── USERS ─────────────────────────────────────────────────
  users() {
    const tbody = document.getElementById('users-tbody');
    if (!tbody) return;
    const roleLabels = {officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Exec',admin:'System Admin'};
    tbody.innerHTML = App.users.map(u =>
      `<tr>
        <td><div class="user-avatar" style="display:inline-flex;width:30px;height:30px;font-size:11px">${u.name.split(' ').map(n=>n[0]).slice(0,2).join('')}</div></td>
        <td><strong>${u.name}</strong><br><small style="color:var(--gray-400)">${u.username}</small></td>
        <td><span class="badge ${u.role==='admin'?'badge-red':u.role==='exec'?'badge-blue':u.role==='ward'?'badge-amber':'badge-green'}">${roleLabels[u.role]||u.role}</span></td>
        <td>${u.station||'—'}</td>
        <td>${u.branch||'—'}</td>
        <td><span class="badge ${u.active?'badge-green':'badge-gray'}">${u.active?'Active':'Inactive'}</span></td>
        <td>
          <div class="btn-group">
            <button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEditUser('${u.id}')">✏️</button>
            <button class="btn btn-sm btn-danger" onclick="PageRenderers.toggleUser('${u.id}')">${u.active?'🚫 Disable':'✅ Enable'}</button>
          </div>
        </td>
       </tr>`
    ).join('');
  },

  openAddUser() {
    document.getElementById('user-form').reset?.();
    document.getElementById('user-modal-title').textContent = 'Add New User';
    document.getElementById('edit-user-id').value = '';
    Modal.open('modal-user');
  },

  openEditUser(id) {
    const u = App.users.find(x=>x.id===id);
    if (!u) return;
    document.getElementById('user-modal-title').textContent = 'Edit User';
    document.getElementById('edit-user-id').value       = id;
    document.getElementById('u-name').value             = u.name;
    document.getElementById('u-username').value         = u.username;
    document.getElementById('u-password').value         = u.password;
    document.getElementById('u-role').value             = u.role;
    document.getElementById('u-station').value          = u.station||'';
    document.getElementById('u-branch').value           = u.branch||'';
    Modal.open('modal-user');
  },

  submitUser() {
    const id       = document.getElementById('edit-user-id').value;
    const userData = {
      name:     document.getElementById('u-name').value.trim(),
      username: document.getElementById('u-username').value.trim(),
      password: document.getElementById('u-password').value.trim(),
      role:     document.getElementById('u-role').value,
      station:  document.getElementById('u-station').value.trim(),
      branch:   document.getElementById('u-branch').value.trim(),
      active:   true,
    };
    if (!userData.name||!userData.username||!userData.password) { Toast.show('Error','Please fill all required fields.','error'); return; }
    if (id) {
      const idx = App.users.findIndex(u=>u.id===id);
      App.users[idx] = { ...App.users[idx], ...userData };
      App.logAudit('EDIT_USER', `Edited user: ${userData.username}`, App.currentUser.username);
      Toast.show('User Updated','User details saved.','success');
    } else {
      if (App.users.find(u=>u.username===userData.username)) { Toast.show('Error','Username already exists.','error'); return; }
      App.users.push({ id:'u'+Date.now(), ...userData });
      App.logAudit('ADD_USER', `Added user: ${userData.username} (${userData.role})`, App.currentUser.username);
      Toast.show('User Added','New user created.','success');
    }
    App.saveUsers();
    Modal.close('modal-user');
    PageRenderers.users();
  },

  toggleUser(id) {
    const u = App.users.find(x=>x.id===id);
    if (!u) return;
    u.active = !u.active;
    App.saveUsers();
    App.logAudit(u.active?'ENABLE_USER':'DISABLE_USER', `${u.active?'Enabled':'Disabled'} user: ${u.username}`, App.currentUser.username);
    Toast.show('Status Updated', `${u.name} has been ${u.active?'enabled':'disabled'}.`, u.active?'success':'warning');
    PageRenderers.users();
  },

  // ─── SETTINGS ──────────────────────────────────────────────
  settings() {
    // Populate form
    const s = App.settings;
    const fields = {
      'set-app-name':    s.appName,
      'set-constituency':s.constituency,
      'set-sheet-id':    s.sheetId,
      'set-api-key':     s.apiKey,
      'set-script-url':  s.scriptUrl,
    };
    Object.entries(fields).forEach(([id,val]) => { const el=document.getElementById(id); if(el) el.value=val||''; });

    // Polling stations
    PageRenderers._renderStationsList();
  },

  saveGeneralSettings() {
    App.settings.appName     = document.getElementById('set-app-name').value.trim() || CONFIG.APP_NAME;
    App.settings.constituency= document.getElementById('set-constituency').value.trim();
    App.settings.sheetId     = document.getElementById('set-sheet-id').value.trim();
    App.settings.apiKey      = document.getElementById('set-api-key').value.trim();
    App.settings.scriptUrl   = document.getElementById('set-script-url').value.trim();
    App.saveSettings();
    App.applyAppName();
    App.logAudit('SETTINGS_CHANGE','Updated general settings', App.currentUser.username);
    Toast.show('Settings Saved','Configuration updated.','success');
  },

  _renderStationsList() {
    const container = document.getElementById('stations-list');
    if (!container) return;
    if (!App.pollingStations.length) {
      container.innerHTML = '<div class="empty-state"><div class="empty-icon">🏢</div><div class="empty-title">No stations configured</div><div class="empty-text">Add stations using the form above.</div></div>';
      return;
    }
    container.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Ward Name</th>
              <th>Polling Station Name</th>
              <th>Branch Name</th>
              <th>Station Code</th>
              <th>Branch Code</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            ${App.pollingStations.map((s,i) => `
              <tr>
                <td>${i+1}</td>
                <td>${s.ward||'—'}</td>
                <td><strong>${s.name}</strong></td>
                <td>${s.branch}</td>
                <td><span class="badge badge-blue">${s.code}</span></td>
                <td><span class="badge badge-green">${s.branchCode}</span></td>
                <td>
                  <button class="btn btn-sm btn-danger" onclick="PageRenderers.removeStation(${i})">🗑️ Remove</button>
                </td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
  },

  addStation() {
    const ward      = document.getElementById('st-ward').value.trim();
    const name      = document.getElementById('st-name').value.trim();
    const code      = document.getElementById('st-code').value.trim();
    const branch    = document.getElementById('st-branch').value.trim();
    const bCode     = document.getElementById('st-bcode').value.trim();
    if (!ward||!name||!code||!branch||!bCode) { Toast.show('Error','Please fill all station fields (Ward, Station Name, Station Code, Branch Name, Branch Code).','error'); return; }

    // Check for duplicate codes
    if (App.pollingStations.find(s => s.code === code)) { Toast.show('Duplicate','Station Code already exists.','error'); return; }
    if (App.pollingStations.find(s => s.branchCode === bCode && s.name === name)) { Toast.show('Duplicate','This station already exists.','error'); return; }

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
    App.pollingStations.splice(i,1);
    App.settings.pollingStations = App.pollingStations;
    App.saveSettings();
    PageRenderers._renderStationsList();
    Toast.show('Station Removed',`${s.name} removed.`,'warning');
    App.logAudit('REMOVE_STATION',`Removed station: ${s.name}`, App.currentUser.username);
  },
};
