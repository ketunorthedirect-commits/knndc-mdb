/* ============================================================
   KNNDCmdb – Page Renderers  v1.3
   ============================================================ */
'use strict';

const PageRenderers = {

  // ══════════════════════════════════════════════════════════
  // DASHBOARD — with gender distribution
  // ══════════════════════════════════════════════════════════
  dashboard() {
    const s = App.getStats();
    const u = App.currentUser;
    const rl = {officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Executive',admin:'System Administrator'}[u.role]||u.role;
    document.getElementById('dash-welcome').textContent  = `Welcome back, ${u.name.split(' ')[0]}`;
    document.getElementById('dash-role').textContent     = rl;
    document.getElementById('dash-total').textContent    = s.total.toLocaleString();
    document.getElementById('dash-today').textContent    = s.today;
    document.getElementById('dash-stations').textContent = s.stations;
    document.getElementById('dash-offline').textContent  = App.offlineQueue.length;

    // Gender stats
    const male = s.byGender?.Male||0, female = s.byGender?.Female||0, other = s.byGender?.Other||0;
    document.getElementById('dash-male').textContent   = male;
    document.getElementById('dash-female').textContent = female;
    const mPct = s.total ? Math.round(male/s.total*100) : 0;
    const fPct = s.total ? Math.round(female/s.total*100) : 0;
    const mBar = document.getElementById('dash-male-bar');
    const fBar = document.getElementById('dash-female-bar');
    if (mBar) mBar.style.width = mPct + '%';
    if (fBar) fBar.style.width = fPct + '%';
    document.getElementById('dash-male-pct').textContent   = mPct + '%';
    document.getElementById('dash-female-pct').textContent = fPct + '%';

    const recent = App.getMembersForUser().slice(0,8);
    const tbody  = document.getElementById('dash-recent-tbody');
    if (tbody) tbody.innerHTML = recent.length
      ? recent.map(m=>`<tr>
          <td><strong>${m.lastName||''},</strong> ${m.firstName||''} ${m.otherNames||''}</td>
          <td>${m.gender||'—'}</td>
          <td>${m.partyId||'—'}</td>
          <td>${m.station||'—'}</td>
          <td><span class="badge badge-blue">${m.officer||'—'}</span></td>
          <td>${m.timestamp||'—'}</td>
        </tr>`).join('')
      : `<tr><td colspan="6"><div class="empty-state"><div class="empty-icon">📋</div><div class="empty-title">No records yet</div></div></td></tr>`;

    setTimeout(()=>{ this._drawBarChart('dash-chart',s.byDay); this._drawGenderDonut('dash-gender-chart',s.byGender); },100);
  },

  _drawBarChart(canvasId, byDay) {
    const canvas=document.getElementById(canvasId); if(!canvas) return;
    const ctx=canvas.getContext('2d');
    const labels=Object.keys(byDay), data=Object.values(byDay), max=Math.max(...data,1);
    const W=canvas.width=canvas.offsetWidth||400, H=canvas.height=160;
    ctx.clearRect(0,0,W,H);
    const barW=(W-80)/labels.length, gap=barW*0.28;
    for(let i=0;i<=4;i++){
      const y=20+(H-50)*i/4;
      ctx.strokeStyle='#e5e7eb';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(40,y);ctx.lineTo(W-10,y);ctx.stroke();
      ctx.fillStyle='#9ca3af';ctx.font='10px Inter';ctx.textAlign='right';
      ctx.fillText(Math.round(max-(max*i/4)),36,y+4);
    }
    data.forEach((v,i)=>{
      const x=40+i*barW+gap/2,bw=barW-gap,bh=((v/max)*(H-60))||2,y=H-30-bh;
      const g=ctx.createLinearGradient(0,y,0,H-30);
      g.addColorStop(0,'#1a6b3a');g.addColorStop(1,'rgba(26,107,58,.25)');
      ctx.fillStyle=g;
      if(ctx.roundRect){ctx.beginPath();ctx.roundRect(x,y,bw,bh,3);ctx.fill();}else{ctx.fillRect(x,y,bw,bh);}
      ctx.fillStyle='#6b7280';ctx.font='9px Inter';ctx.textAlign='center';
      const d=new Date(labels[i]);
      ctx.fillText(isNaN(d)?labels[i]:d.toLocaleDateString('en',{weekday:'short'}),x+bw/2,H-12);
      if(v>0){ctx.fillStyle='#1a6b3a';ctx.font='bold 10px Inter';ctx.fillText(v,x+bw/2,y-4);}
    });
  },

  _drawGenderDonut(canvasId, byGender) {
    const canvas=document.getElementById(canvasId); if(!canvas) return;
    const ctx=canvas.getContext('2d');
    const W=canvas.width=canvas.offsetWidth||220, H=canvas.height=180;
    ctx.clearRect(0,0,W,H);
    const male=byGender?.Male||0, female=byGender?.Female||0, total=male+female;
    if(!total){ ctx.fillStyle='#e5e7eb';ctx.beginPath();ctx.arc(W/2,H/2,60,0,Math.PI*2);ctx.fill();return; }
    const cx=W/2, cy=H/2, r=Math.min(cx,cy)-15;
    const slices=[
      {val:male,  color:'#1a6b3a',label:'Male'},
      {val:female,color:'#c8102e',label:'Female'},
    ];
    let angle=-Math.PI/2;
    slices.forEach(s=>{
      const slice=(s.val/total)*Math.PI*2;
      ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,r,angle,angle+slice);ctx.closePath();
      ctx.fillStyle=s.color;ctx.fill();
      angle+=slice;
    });
    // Donut hole
    ctx.beginPath();ctx.arc(cx,cy,r*.55,0,Math.PI*2);ctx.fillStyle='white';ctx.fill();
    ctx.fillStyle='#1f2937';ctx.font='bold 16px Outfit';ctx.textAlign='center';
    ctx.fillText(total,cx,cy+4);
    ctx.fillStyle='#6b7280';ctx.font='9px Inter';ctx.fillText('Total',cx,cy+16);
  },

  // ══════════════════════════════════════════════════════════
  // DATA ENTRY — Zone + Gender fields added
  // ══════════════════════════════════════════════════════════
  entry() {
    const u = App.currentUser;
    if (u.role==='officer' && u.assignedStations?.length===1) {
      const s = App.pollingStations.find(ps=>ps.code===u.assignedStations[0]);
      if (s) this._fillStationFields(s);
    }
    // Clear gender selection
    const gEl = document.getElementById('f-gender');
    if (gEl) gEl.value = '';
    this._setupBranchCodeSearch(u.role);
  },

  _fillStationFields(s) {
    const set=(id,val)=>{ const el=document.getElementById(id); if(el) el.value=val||''; };
    set('f-zone',        s.zone);
    set('f-ward-name',   s.ward);
    set('f-station-name',s.name);
    set('f-branch-name', s.branch);
    set('f-station-code',s.code);
    set('f-branch-code', s.branchCode);
  },

  _setupBranchCodeSearch(role) {
    const input    = document.getElementById('f-branch-code');
    const dropdown = document.getElementById('branch-code-dropdown');
    if (!input||!dropdown) return;

    const isLocked = role==='officer' && App.currentUser?.assignedStations?.length===1;
    if (isLocked) { input.setAttribute('readonly',true); input.classList.add('auto-filled'); return; }

    input.removeAttribute('readonly'); input.classList.remove('auto-filled');
    input.placeholder='Type branch code, zone, ward or station name…';

    const u = App.currentUser;
    const getPool = () => {
      if (u.role==='officer' && u.assignedStations?.length)
        return App.pollingStations.filter(s=>u.assignedStations.includes(s.code));
      return App.pollingStations;
    };

    const show = (q) => {
      const hits = getPool().filter(s=>!q||
        s.branchCode?.toLowerCase().includes(q)||s.branch?.toLowerCase().includes(q)||
        s.name?.toLowerCase().includes(q)||s.ward?.toLowerCase().includes(q)||
        s.zone?.toLowerCase().includes(q)||s.code?.toLowerCase().includes(q)
      ).slice(0,30);
      dropdown.innerHTML = hits.length
        ? hits.map(s=>`
            <div class="dropdown-item" data-sidx="${App.pollingStations.indexOf(s)}">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
                <div>
                  <strong style="color:var(--ndc-green-dk)">${s.branchCode}</strong>
                  <span style="color:var(--gray-500);font-size:11px;margin-left:6px">${s.name}</span>
                </div>
                <div style="display:flex;gap:4px">
                  <span class="badge badge-green" style="font-size:10px">${s.zone||''}</span>
                  <span class="badge badge-blue"  style="font-size:10px">${s.ward||''}</span>
                </div>
              </div>
              <div style="font-size:11px;color:var(--gray-400);margin-top:2px">Branch: ${s.branch} · Code: ${s.code}</div>
            </div>`).join('')
        : '<div class="dropdown-item" style="color:var(--gray-400);text-align:center;padding:16px">No matching stations</div>';
      dropdown.classList.add('open');
      dropdown.querySelectorAll('.dropdown-item[data-sidx]').forEach(item=>{
        item.onclick=()=>{
          const s=App.pollingStations[parseInt(item.dataset.sidx)];
          if(s){this._fillStationFields(s);Toast.show('Station Selected',`${s.name} · ${s.branchCode}`,'success');}
          dropdown.classList.remove('open');
        };
      });
    };

    const ni=input.cloneNode(true); input.parentNode.replaceChild(ni,input);
    ni.addEventListener('focus',()=>show(ni.value.toLowerCase().trim()));
    ni.addEventListener('input',()=>show(ni.value.toLowerCase().trim()));
    document.addEventListener('click',e=>{if(!ni.contains(e.target)&&!dropdown.contains(e.target))dropdown.classList.remove('open');});
    if(role!=='officer') setTimeout(()=>show(''),150);
  },

  submitEntry() {
    const required=['f-branch-code','f-station-code','f-last-name','f-first-name','f-party-id','f-gender'];
    let ok=true;
    required.forEach(id=>{
      const el=document.getElementById(id);
      if(el&&!el.value.trim()){el.style.borderColor='var(--ndc-red)';ok=false;}
      else if(el) el.style.borderColor='';
    });
    if(!ok){Toast.show('Validation Error','Please fill all required fields including Gender.','error');return;}
    const g=id=>document.getElementById(id)?.value.trim()||'';
    App.addMember({
      zone:        g('f-zone'),
      ward:        g('f-ward-name'),
      station:     g('f-station-name'),
      stationCode: g('f-station-code'),
      branch:      g('f-branch-name'),
      branchCode:  g('f-branch-code'),
      lastName:    g('f-last-name'),
      firstName:   g('f-first-name'),
      otherNames:  g('f-other-names'),
      gender:      g('f-gender'),
      partyId:     g('f-party-id'),
      voterId:     g('f-voter-id'),
      phone:       g('f-phone'),
    });
    Toast.show('Record Saved',`${g('f-first-name')} ${g('f-last-name')} registered.`,'success');
    ['f-last-name','f-first-name','f-other-names','f-gender','f-party-id','f-voter-id','f-phone']
      .forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
    document.getElementById('f-last-name')?.focus();
  },

  // ══════════════════════════════════════════════════════════
  // MY RECORDS
  // ══════════════════════════════════════════════════════════
  'my-records'() {
    const members = App.getMembersForUser();
    document.getElementById('my-records-count').textContent = `${members.length} record${members.length!==1?'s':''}`;
    this._renderMembersTable('my-records-tbody', members, false);
  },

  // ══════════════════════════════════════════════════════════
  // ALL RECORDS — with ward/station/branch/gender filters
  // ══════════════════════════════════════════════════════════
  records() {
    PageRenderers._allState = PageRenderers._allState || { q:'', page:1, ward:'', station:'', branch:'', gender:'' };
    this._populateRecordFilters();
    this._renderAllRecords();
  },

  _populateRecordFilters() {
    const all = App.getMembersForUser();
    const wards    = [...new Set(all.map(m=>m.ward).filter(Boolean))].sort();
    const stations = [...new Set(all.map(m=>m.station).filter(Boolean))].sort();
    const branches = [...new Set(all.map(m=>m.branch).filter(Boolean))].sort();

    const populate=(id,vals,label)=>{
      const el=document.getElementById(id); if(!el) return;
      const cur=el.value;
      el.innerHTML=`<option value="">All ${label}</option>`+vals.map(v=>`<option value="${v}"${v===cur?' selected':''}>${v}</option>`).join('');
    };
    populate('filter-ward',    wards,    'Wards');
    populate('filter-station', stations, 'Stations');
    populate('filter-branch',  branches, 'Branches');
  },

  _renderAllRecords() {
    const st = PageRenderers._allState;
    let members = App.getMembersForUser();

    if (st.q)       { const q=st.q.toLowerCase(); members=members.filter(m=>m.firstName?.toLowerCase().includes(q)||m.lastName?.toLowerCase().includes(q)||m.partyId?.toLowerCase().includes(q)||m.voterId?.toLowerCase().includes(q)||m.phone?.includes(q)||m.station?.toLowerCase().includes(q)||m.ward?.toLowerCase().includes(q)||m.zone?.toLowerCase().includes(q)); }
    if (st.ward)    members=members.filter(m=>m.ward===st.ward);
    if (st.station) members=members.filter(m=>m.station===st.station);
    if (st.branch)  members=members.filter(m=>m.branch===st.branch);
    if (st.gender)  members=members.filter(m=>m.gender===st.gender);

    const per=20, total=members.length, pages=Math.ceil(total/per)||1;
    const page=Math.min(st.page,pages), slice=members.slice((page-1)*per,page*per);

    document.getElementById('records-count').textContent=`${total.toLocaleString()} record${total!==1?'s':''}`;
    this._renderMembersTable('records-tbody',slice,true);

    const pc=document.getElementById('records-pagination'); if(!pc) return; pc.innerHTML='';
    if(pages>1){
      const mk=(lbl,pg,active,disabled)=>{ const b=document.createElement('button');b.className='page-btn'+(active?' active':'');b.textContent=lbl;b.disabled=disabled;b.onclick=()=>{PageRenderers._allState.page=pg;PageRenderers._renderAllRecords();}; return b; };
      pc.appendChild(mk('‹',page-1,false,page===1));
      for(let i=1;i<=pages;i++){
        if(i===1||i===pages||Math.abs(i-page)<=2) pc.appendChild(mk(i,i,i===page));
        else if(Math.abs(i-page)===3){const s=document.createElement('span');s.className='page-btn';s.textContent='…';pc.appendChild(s);}
      }
      pc.appendChild(mk('›',page+1,false,page===pages));
    }
  },

  _renderMembersTable(tbodyId, members, showActions) {
    const tbody=document.getElementById(tbodyId); if(!tbody) return;
    const canEdit=showActions&&['admin','exec','ward'].includes(App.currentUser?.role);
    const canDel =showActions&&App.currentUser?.role==='admin';
    tbody.innerHTML = members.length
      ? members.map(m=>`<tr>
          <td><strong>${m.lastName||''},</strong> ${m.firstName||''} ${m.otherNames||''}</td>
          <td><span class="badge ${m.gender==='Male'?'badge-blue':'badge-red'}" style="font-size:11px">${m.gender||'—'}</span></td>
          <td>${m.zone||'—'}</td>
          <td>${m.partyId||'—'}</td>
          <td>${m.voterId||'—'}</td>
          <td>${m.phone||'—'}</td>
          <td>${m.ward||'—'}</td>
          <td>${m.station||'—'}</td>
          <td>${m.branch||'—'}</td>
          <td>${m.timestamp||'—'}</td>
          ${showActions?`<td>
            ${canEdit?`<button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEdit('${m.id}')">✏️</button>`:''}
            ${canDel?`<button class="btn btn-sm btn-danger"    onclick="PageRenderers.confirmDelete('${m.id}')">🗑️</button>`:''}
          </td>`:''}
        </tr>`).join('')
      : `<tr><td colspan="11"><div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No records found</div></div></td></tr>`;
  },

  openEdit(id) {
    const m=App.members.find(x=>x.id===id); if(!m) return;
    const map={'edit-id':id,'edit-first':m.firstName,'edit-last':m.lastName,'edit-other':m.otherNames,'edit-gender':m.gender,'edit-party':m.partyId,'edit-voter':m.voterId,'edit-phone':m.phone,'edit-reason':''};
    Object.entries(map).forEach(([fid,val])=>{ const el=document.getElementById(fid); if(el) el.value=val||''; });
    document.getElementById('edit-current-info').textContent=`Editing: ${m.firstName} ${m.lastName} | Party ID: ${m.partyId} | Added: ${m.timestamp}`;
    Modal.open('modal-edit');
  },

  submitEdit() {
    const id=document.getElementById('edit-id').value;
    const reason=document.getElementById('edit-reason').value.trim();
    if(!reason){Toast.show('Reason Required','Please provide a reason for this change.','error');return;}
    App.updateMember(id,{
      firstName: document.getElementById('edit-first').value.trim(),
      lastName:  document.getElementById('edit-last').value.trim(),
      otherNames:document.getElementById('edit-other').value.trim(),
      gender:    document.getElementById('edit-gender').value,
      partyId:   document.getElementById('edit-party').value.trim(),
      voterId:   document.getElementById('edit-voter').value.trim(),
      phone:     document.getElementById('edit-phone').value.trim(),
    },reason);
    Modal.close('modal-edit'); Toast.show('Record Updated','Changes saved.','success');
    PageRenderers.records();
  },

  confirmDelete(id) {
    const m=App.members.find(x=>x.id===id);
    document.getElementById('del-id').value=id;
    document.getElementById('del-name').textContent=m?`${m.firstName} ${m.lastName} (${m.partyId})`:id;
    document.getElementById('del-reason').value='';
    Modal.open('modal-delete');
  },

  submitDelete() {
    const id=document.getElementById('del-id').value;
    const reason=document.getElementById('del-reason').value.trim();
    if(!reason){Toast.show('Reason Required','Please provide a reason.','error');return;}
    App.deleteMember(id,reason); Modal.close('modal-delete');
    Toast.show('Record Deleted','Member removed.','error'); PageRenderers.records();
  },

  // ══════════════════════════════════════════════════════════
  // REPORTS — with filters
  // ══════════════════════════════════════════════════════════
  reports() {
    PageRenderers._repState = PageRenderers._repState || { ward:'', station:'', branch:'', gender:'' };
    this._populateReportFilters();
    this._renderReports();
  },

  _populateReportFilters() {
    const all=App.getMembersForUser();
    const pop=(id,vals,lbl)=>{ const el=document.getElementById(id);if(!el)return;const cur=el.value;el.innerHTML=`<option value="">All ${lbl}</option>`+vals.map(v=>`<option value="${v}"${v===cur?' selected':''}>${v}</option>`).join(''); };
    pop('rep-filter-ward',   [...new Set(all.map(m=>m.ward).filter(Boolean))].sort(),'Wards');
    pop('rep-filter-station',[...new Set(all.map(m=>m.station).filter(Boolean))].sort(),'Stations');
    pop('rep-filter-branch', [...new Set(all.map(m=>m.branch).filter(Boolean))].sort(),'Branches');
  },

  _renderReports() {
    const st=PageRenderers._repState;
    let members=App.getMembersForUser();
    if(st.ward)    members=members.filter(m=>m.ward===st.ward);
    if(st.station) members=members.filter(m=>m.station===st.station);
    if(st.branch)  members=members.filter(m=>m.branch===st.branch);
    if(st.gender)  members=members.filter(m=>m.gender===st.gender);

    const byStation={};
    members.forEach(m=>{
      if(!byStation[m.station]) byStation[m.station]={station:m.station,branch:m.branch||'',ward:m.ward||'',zone:m.zone||'',count:0,male:0,female:0};
      byStation[m.station].count++;
      if(m.gender==='Male') byStation[m.station].male++;
      if(m.gender==='Female') byStation[m.station].female++;
    });
    const rows=Object.values(byStation).sort((a,b)=>b.count-a.count);
    const tbody=document.getElementById('report-tbody');
    if(tbody) tbody.innerHTML=rows.length
      ? rows.map((r,i)=>`<tr>
          <td>${i+1}</td><td>${r.zone||'—'}</td><td>${r.ward}</td><td><strong>${r.station}</strong></td><td>${r.branch}</td>
          <td>${r.count}</td>
          <td><span class="badge badge-blue">${r.male}</span></td>
          <td><span class="badge badge-red">${r.female}</span></td>
          <td><div class="progress-bar" style="width:90px"><div class="progress-fill" style="width:${members.length?Math.round(r.count/members.length*100):0}%"></div></div></td>
          <td>${members.length?Math.round(r.count/members.length*100):0}%</td>
        </tr>`).join('')
      : `<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--gray-400)">No data</td></tr>`;

    document.getElementById('report-total').textContent=members.length.toLocaleString();
    const today=new Date().toLocaleDateString('en-GH');
    document.getElementById('daily-count').textContent=members.filter(m=>m.timestamp?.includes(today)).length;
    document.getElementById('daily-date').textContent=today;

    const male=members.filter(m=>m.gender==='Male').length;
    const female=members.filter(m=>m.gender==='Female').length;
    document.getElementById('rep-male-count').textContent=male;
    document.getElementById('rep-female-count').textContent=female;

    setTimeout(()=>{ this._drawPieChart(byStation,members.length); this._drawGenderDonut('rep-gender-chart',{Male:male,Female:female}); },100);
  },

  _drawPieChart(byStation,total) {
    const canvas=document.getElementById('report-chart'); if(!canvas||!total) return;
    const ctx=canvas.getContext('2d');
    const W=canvas.width=canvas.offsetWidth||280, H=canvas.height=200;
    ctx.clearRect(0,0,W,H);
    const cx=W/2, cy=H/2, r=Math.min(cx,cy)-20;
    const colors=['#1a6b3a','#c8102e','#2563eb','#d97706','#7c3aed','#10b981','#f59e0b','#ef4444','#8b5cf6'];
    let angle=-Math.PI/2;
    Object.values(byStation).forEach((s,i)=>{
      const slice=(s.count/total)*Math.PI*2;
      ctx.beginPath();ctx.moveTo(cx,cy);ctx.arc(cx,cy,r,angle,angle+slice);ctx.closePath();
      ctx.fillStyle=colors[i%colors.length];ctx.fill();angle+=slice;
    });
    ctx.beginPath();ctx.arc(cx,cy,r*.55,0,Math.PI*2);ctx.fillStyle='white';ctx.fill();
    ctx.fillStyle='#1f2937';ctx.font='bold 20px Outfit';ctx.textAlign='center';
    ctx.fillText(total,cx,cy+4);
    ctx.fillStyle='#6b7280';ctx.font='10px Inter';ctx.fillText('Total',cx,cy+18);
  },

  exportExcel() {
    const members=App.getMembersForUser(), constituency=App.settings.constituency||'Ketu North';
    if(typeof XLSX==='undefined'){
      const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      s.onload=()=>PageRenderers.exportExcel();document.head.appendChild(s);Toast.show('Loading','Preparing…','info');return;
    }
    const wb=XLSX.utils.book_new();
    const hdr=['First Name','Surname','Party ID Number','Voter ID Number','Telephone Number'];
    const mkSheet=(rows)=>{ const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[{wch:18},{wch:18},{wch:20},{wch:20},{wch:18}];ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:1,c:0},e:{r:1,c:3}}];return ws; };
    XLSX.utils.book_append_sheet(wb,mkSheet([['MEMBERSHIP DATABASE','','','',''],['Polling Station / Branch Name: ALL STATIONS','','','',`Constituency: ${constituency}`],['','','','',''],hdr,...members.map(m=>[m.firstName,m.lastName,m.partyId,m.voterId,m.phone])]),'All Members');
    const byStation={};
    members.forEach(m=>{ const k=m.stationCode||'UNK';if(!byStation[k])byStation[k]={info:m,members:[]};byStation[k].members.push(m); });
    Object.values(byStation).forEach(({info,members:sm})=>{
      const name=(info.station||'Station').replace(/[\\\/\?\*\[\]:]/g,'').substring(0,31);
      XLSX.utils.book_append_sheet(wb,mkSheet([['MEMBERSHIP DATABASE','','','',''],['Polling Station / Branch Name: '+info.station+' / '+info.branch,'','','',`Constituency: ${constituency}`],['','','','',''],hdr,...sm.map(m=>[m.firstName,m.lastName,m.partyId,m.voterId,m.phone])]),name);
    });
    const wsSummary=XLSX.utils.aoa_to_sheet([['MEMBERSHIP DATABASE — SUMMARY','','',''],['Constituency: '+constituency,'','','Export Date: '+new Date().toLocaleDateString('en-GH')],['','','',''],['Ward','Polling Station','Branch','Total Members'],...Object.values(byStation).map(({info,members:sm})=>[info.ward||'—',info.station||'—',info.branch||'—',sm.length]),['','','GRAND TOTAL',members.length]]);
    wsSummary['!cols']=[{wch:20},{wch:28},{wch:22},{wch:15}];
    XLSX.utils.book_append_sheet(wb,wsSummary,'Summary');
    XLSX.writeFile(wb,`KNNDCmdb_${new Date().toISOString().slice(0,10)}.xlsx`);
    Toast.show('Export Ready',`${members.length} records exported.`,'success');
    App.logAudit('EXPORT_EXCEL',`Exported ${members.length} records`,App.currentUser.username);
  },

  exportPDF() { window.print(); App.logAudit('EXPORT_PDF','Printed/exported as PDF',App.currentUser.username); },

  // ══════════════════════════════════════════════════════════
  // ANALYTICS — with gender distribution and filters
  // ══════════════════════════════════════════════════════════
  analytics() {
    PageRenderers._anaState = PageRenderers._anaState || { ward:'', station:'', branch:'', gender:'' };
    this._populateAnaFilters();
    this._renderAnalytics();
  },

  _populateAnaFilters() {
    const all=App.getMembersForUser();
    const pop=(id,vals,lbl)=>{ const el=document.getElementById(id);if(!el)return;const cur=el.value;el.innerHTML=`<option value="">All ${lbl}</option>`+vals.map(v=>`<option value="${v}"${v===cur?' selected':''}>${v}</option>`).join(''); };
    pop('ana-filter-ward',   [...new Set(all.map(m=>m.ward).filter(Boolean))].sort(),'Wards');
    pop('ana-filter-station',[...new Set(all.map(m=>m.station).filter(Boolean))].sort(),'Stations');
    pop('ana-filter-branch', [...new Set(all.map(m=>m.branch).filter(Boolean))].sort(),'Branches');
  },

  _renderAnalytics() {
    const st=PageRenderers._anaState;
    let members=App.getMembersForUser();
    if(st.ward)    members=members.filter(m=>m.ward===st.ward);
    if(st.station) members=members.filter(m=>m.station===st.station);
    if(st.branch)  members=members.filter(m=>m.branch===st.branch);
    if(st.gender)  members=members.filter(m=>m.gender===st.gender);

    const male=members.filter(m=>m.gender==='Male').length;
    const female=members.filter(m=>m.gender==='Female').length;
    const today=new Date().toLocaleDateString('en-GH');
    const todayCount=members.filter(m=>m.timestamp?.includes(today)).length;
    const byStation={};
    members.forEach(m=>{byStation[m.station]=(byStation[m.station]||0)+1;});

    document.getElementById('ana-total').textContent    = members.length.toLocaleString();
    document.getElementById('ana-today').textContent    = todayCount;
    document.getElementById('ana-stations').textContent = Object.keys(byStation).length;
    document.getElementById('ana-male').textContent     = male;
    document.getElementById('ana-female').textContent   = female;
    document.getElementById('ana-male-pct').textContent = members.length?Math.round(male/members.length*100)+'%':'0%';
    document.getElementById('ana-female-pct').textContent=members.length?Math.round(female/members.length*100)+'%':'0%';

    const yest=new Date();yest.setDate(yest.getDate()-1);
    const yKey=yest.toLocaleDateString('en-GH');
    const yCount=members.filter(m=>m.timestamp?.includes(yKey)).length;
    const growth=yCount?((todayCount-yCount)/yCount*100).toFixed(0):(todayCount>0?'∞':0);
    document.getElementById('ana-growth').textContent=(typeof growth==='number'&&growth>0?'+':'')+growth+(growth!=='∞'?'%':'');

    const byDay={};
    for(let i=6;i>=0;i--){const d=new Date();d.setDate(d.getDate()-i);const k=d.toLocaleDateString('en-GH');byDay[k]=members.filter(m=>m.timestamp?.includes(k)).length;}

    setTimeout(()=>{
      this._drawBarChart('ana-chart',byDay);
      this._drawGenderDonut('ana-gender-chart',{Male:male,Female:female});
      this._drawPieChart(byStation,members.length);
    },100);

    const byOfficer={};
    members.forEach(m=>{byOfficer[m.officer]=(byOfficer[m.officer]||0)+1;});
    const top=Object.entries(byOfficer).sort((a,b)=>b[1]-a[1]).slice(0,8);
    const t=document.getElementById('ana-officers-tbody');
    if(t) t.innerHTML=top.length
      ? top.map(([name,count])=>`<tr><td>${name}</td><td>${count}</td><td><div class="progress-bar"><div class="progress-fill" style="width:${members.length?Math.round(count/members.length*100):0}%"></div></div></td><td>${members.length?Math.round(count/members.length*100):0}%</td></tr>`).join('')
      : '<tr><td colspan="4" style="text-align:center">No data</td></tr>';
  },

  // ══════════════════════════════════════════════════════════
  // AUDIT LOG
  // ══════════════════════════════════════════════════════════
  audit() {
    const q=document.getElementById('audit-search')?.value?.toLowerCase()||'';
    const filter=document.getElementById('audit-filter')?.value||'';
    App.auditLog=JSON.parse(localStorage.getItem('knndc_audit')||'[]');
    let log=App.auditLog;
    if(q)      log=log.filter(e=>e.action?.toLowerCase().includes(q)||e.details?.toLowerCase().includes(q)||e.user?.toLowerCase().includes(q));
    if(filter) log=log.filter(e=>e.action===filter);
    const c=document.getElementById('audit-entries'); if(!c) return;
    c.innerHTML=log.slice(0,300).map(e=>{
      const isDanger=['DELETE_MEMBER','LOCKOUT','DISABLE_USER','CLEAR_DEMO'].includes(e.action);
      const isWarn=['EDIT_MEMBER','EXPORT_EXCEL','EXPORT_PDF','AUTO_LOGOUT','PASSWORD_RESET'].includes(e.action);
      return `<div class="log-entry ${isDanger?'danger':isWarn?'warning':''}">
        <div class="log-entry-header">
          <span class="badge ${isDanger?'badge-red':isWarn?'badge-amber':'badge-green'}">${e.action}</span>
          <span class="log-user">👤 ${e.user}</span>
          <span class="log-time">${e.timestamp}</span>
        </div>
        <div class="log-details">${e.details||''}</div>
        ${e.reason?`<div class="log-details" style="color:var(--ndc-red);margin-top:4px">⚠️ ${e.reason}</div>`:''}
      </div>`;
    }).join('')||`<div class="empty-state"><div class="empty-icon">🛡️</div><div class="empty-title">No entries found</div></div>`;
  },

  // ══════════════════════════════════════════════════════════
  // USER MANAGEMENT — password reset + assign
  // ══════════════════════════════════════════════════════════
  users() {
    const tbody=document.getElementById('users-tbody'); if(!tbody) return;
    const rLabels={officer:'Data Entry Officer',ward:'Ward Coordinator',exec:'Constituency Exec',admin:'System Admin'};
    const rBadge={officer:'badge-green',ward:'badge-amber',exec:'badge-blue',admin:'badge-red'};
    App.users=JSON.parse(localStorage.getItem('knndc_users')||'null')||[];

    tbody.innerHTML=App.users.map(u=>`
      <tr>
        <td><div style="display:inline-flex;align-items:center;justify-content:center;width:34px;height:34px;border-radius:50%;background:var(--ndc-green);color:white;font-weight:700;font-size:12px;font-family:var(--font-head)">${u.name.split(' ').map(n=>n[0]).slice(0,2).join('')}</div></td>
        <td><strong>${u.name}</strong><br><small style="color:var(--gray-400)">${u.username}</small></td>
        <td><span class="badge ${rBadge[u.role]||'badge-gray'}">${rLabels[u.role]||u.role}</span></td>
        <td>
          ${u.mustChangePassword
            ? '<span class="badge badge-amber">⚠️ Must change</span>'
            : '<span class="badge badge-green">✅ Set</span>'}
        </td>
        <td>${u.ward||'—'}</td>
        <td>${(u.assignedStations||[]).length?(u.assignedStations||[]).map(c=>`<span class="badge badge-blue" style="font-size:10px;margin:1px">${c}</span>`).join(''):'<span style="color:var(--gray-400);font-size:12px">None</span>'}</td>
        <td><span class="badge ${u.active?'badge-green':'badge-gray'}">${u.active?'Active':'Inactive'}</span></td>
        <td>
          <div class="btn-group" style="flex-wrap:nowrap;gap:4px">
            <button class="btn btn-sm btn-secondary" onclick="PageRenderers.openEditUser('${u.id}')" title="Edit user">✏️</button>
            <button class="btn btn-sm btn-outline"   onclick="PageRenderers.resetPassword('${u.id}')" title="Reset password to default">🔑</button>
            ${u.role==='officer'?`<button class="btn btn-sm btn-secondary" style="border-color:var(--ndc-green);color:var(--ndc-green)" onclick="PageRenderers.openAssignModal('${u.id}')" title="Assign stations">📍</button>`:''}
            <button class="btn btn-sm btn-danger"    onclick="PageRenderers.toggleUser('${u.id}')" title="${u.active?'Disable':'Enable'}">${u.active?'🚫':'✅'}</button>
          </div>
        </td>
      </tr>`).join('');
  },

  resetPassword(userId) {
    const u=App.users.find(x=>x.id===userId); if(!u) return;
    if(!confirm(`Reset ${u.name}'s password to the system default?\n\nThey will be required to change it on next login.`)) return;
    App.resetPasswordToDefault(userId);
    PageRenderers.users();
  },

  openAddUser() {
    document.getElementById('user-modal-title').textContent='Add New User';
    document.getElementById('edit-user-id').value='';
    ['u-name','u-username','u-password','u-ward','u-station','u-branch'].forEach(id=>{ const el=document.getElementById(id);if(el)el.value=''; });
    document.getElementById('u-role').value='officer';
    Modal.open('modal-user');
  },

  openEditUser(id) {
    const u=App.users.find(x=>x.id===id); if(!u) return;
    document.getElementById('user-modal-title').textContent='Edit User';
    document.getElementById('edit-user-id').value=id;
    document.getElementById('u-name').value    =u.name;
    document.getElementById('u-username').value=u.username;
    document.getElementById('u-password').value=u.password;
    document.getElementById('u-role').value    =u.role;
    document.getElementById('u-ward').value    =u.ward||'';
    document.getElementById('u-station').value =u.station||'';
    document.getElementById('u-branch').value  =u.branch||'';
    Modal.open('modal-user');
  },

  submitUser() {
    const id=document.getElementById('edit-user-id').value;
    const data={name:document.getElementById('u-name').value.trim(),username:document.getElementById('u-username').value.trim(),password:document.getElementById('u-password').value.trim(),role:document.getElementById('u-role').value,ward:document.getElementById('u-ward').value.trim(),station:document.getElementById('u-station').value.trim(),branch:document.getElementById('u-branch').value.trim(),active:true};
    if(!data.name||!data.username||!data.password){Toast.show('Error','Name, username and password are required.','error');return;}
    App.users=JSON.parse(localStorage.getItem('knndc_users')||'null')||[];
    if(id){
      const idx=App.users.findIndex(u=>u.id===id);
      App.users[idx]={...App.users[idx],...data};
      App.logAudit('EDIT_USER',`Edited user: ${data.username} (${data.role})`,App.currentUser.username);
      Toast.show('User Updated','Changes saved.','success');
    } else {
      if(App.users.find(u=>u.username===data.username)){Toast.show('Error','Username already exists.','error');return;}
      App.users.push({id:'u'+Date.now(),...data,assignedStations:[],mustChangePassword:true});
      App.logAudit('ADD_USER',`Created: ${data.username} (${data.role})`,App.currentUser.username);
      Toast.show('User Created','New user added. They must change password on first login.','success');
    }
    App.saveUsers();Modal.close('modal-user');PageRenderers.users();
  },

  toggleUser(id) {
    App.users=JSON.parse(localStorage.getItem('knndc_users')||'null')||[];
    const u=App.users.find(x=>x.id===id); if(!u) return;
    u.active=!u.active;App.saveUsers();
    App.logAudit(u.active?'ENABLE_USER':'DISABLE_USER',`${u.active?'Enabled':'Disabled'}: ${u.username}`,App.currentUser.username);
    Toast.show('Status Updated',`${u.name} is now ${u.active?'active':'inactive'}.`,u.active?'success':'warning');
    PageRenderers.users();
  },

  openAssignModal(userId) {
    App.users=JSON.parse(localStorage.getItem('knndc_users')||'null')||[];
    const u=App.users.find(x=>x.id===userId); if(!u) return;
    document.getElementById('assign-user-id').value=userId;
    document.getElementById('assign-user-name').textContent=u.name;
    document.getElementById('assign-user-role').textContent='Data Entry Officer';
    const assigned=u.assignedStations||[];
    const byWard={};
    App.pollingStations.forEach(s=>{if(!byWard[s.ward])byWard[s.ward]=[];byWard[s.ward].push(s);});
    document.getElementById('assign-stations-container').innerHTML=Object.entries(byWard).map(([ward,sts])=>`
      <div style="margin-bottom:18px">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:700;color:var(--ndc-green-dk);text-transform:none;letter-spacing:0;font-size:13px">
            <input type="checkbox" class="ward-check" data-ward="${ward}" ${sts.every(s=>assigned.includes(s.code))?'checked':''} onchange="PageRenderers.toggleWardCheck('${ward}',this.checked)" style="width:auto;accent-color:var(--ndc-green)">
            🏘️ ${ward}
          </label>
          <span class="badge badge-green" style="font-size:10px">${sts.length} station${sts.length!==1?'s':''}</span>
        </div>
        <div style="padding-left:20px;display:grid;grid-template-columns:1fr 1fr;gap:6px">
          ${sts.map(s=>`<label style="display:flex;align-items:flex-start;gap:8px;cursor:pointer;padding:8px 10px;border:1px solid var(--gray-200);border-radius:var(--radius);background:white;transition:var(--transition);text-transform:none;letter-spacing:0;font-weight:400;font-size:12px">
            <input type="checkbox" class="station-check" name="assign_stations" value="${s.code}" data-ward="${ward}" ${assigned.includes(s.code)?'checked':''} onchange="PageRenderers.updateWardCheckState('${ward}')" style="width:auto;accent-color:var(--ndc-green);margin-top:2px">
            <div><div style="font-weight:600;color:var(--gray-800)">${s.name}</div><div style="color:var(--gray-400)">${s.code} · ${s.branchCode}</div></div>
          </label>`).join('')}
        </div>
      </div>`).join('')||'<div class="empty-state"><div class="empty-icon">🏢</div><div class="empty-title">No stations configured</div></div>';
    Modal.open('modal-assign');
  },

  toggleWardCheck(ward,checked) { document.querySelectorAll(`.station-check[data-ward="${ward}"]`).forEach(cb=>cb.checked=checked); },
  updateWardCheckState(ward) {
    const all=document.querySelectorAll(`.station-check[data-ward="${ward}"]`);
    const wc=document.querySelector(`.ward-check[data-ward="${ward}"]`); if(!wc) return;
    const n=[...all].filter(cb=>cb.checked).length;
    wc.checked=n===all.length; wc.indeterminate=n>0&&n<all.length;
  },

  submitAssignment() {
    const userId=document.getElementById('assign-user-id').value;
    const selected=[...document.querySelectorAll('.station-check:checked')].map(cb=>cb.value);
    App.users=JSON.parse(localStorage.getItem('knndc_users')||'null')||[];
    const u=App.users.find(x=>x.id===userId); if(!u) return;
    const prev=u.assignedStations||[];
    u.assignedStations=selected;
    if(selected.length>0){const p=App.pollingStations.find(s=>s.code===selected[0]);if(p){u.station=p.code;u.ward=p.ward;u.branch=p.branch;}}
    App.saveUsers();
    App.logAudit('ASSIGN_STATIONS',`Assigned ${selected.length} station(s) to ${u.username}: [${selected.join(', ')}]. Prev: [${prev.join(', ')}]`,App.currentUser.username);
    Modal.close('modal-assign');
    Toast.show('Assignment Saved',`${u.name} assigned to ${selected.length} station${selected.length!==1?'s':''}.`,'success');
    PageRenderers.users();
  },

  // ══════════════════════════════════════════════════════════
  // SETTINGS
  // ══════════════════════════════════════════════════════════
  settings() {
    const s=App.settings;
    ['set-app-name','set-constituency','set-sheet-id','set-api-key','set-script-url'].forEach(id=>{
      const map={'set-app-name':s.appName,'set-constituency':s.constituency,'set-sheet-id':s.sheetId,'set-api-key':s.apiKey,'set-script-url':s.scriptUrl};
      const el=document.getElementById(id);if(el)el.value=map[id]||'';
    });
    this._renderStationsList();
  },

  saveGeneralSettings() {
    App.settings.appName     =document.getElementById('set-app-name').value.trim()||CONFIG.APP_NAME;
    App.settings.constituency=document.getElementById('set-constituency').value.trim();
    App.settings.sheetId     =document.getElementById('set-sheet-id').value.trim();
    App.settings.apiKey      =document.getElementById('set-api-key').value.trim();
    App.settings.scriptUrl   =document.getElementById('set-script-url').value.trim();
    App.saveSettings();App.applyAppName();
    App.logAudit('SETTINGS_CHANGE','Updated general settings',App.currentUser.username);
    Toast.show('Settings Saved','Configuration updated.','success');
    // Re-start sync timer if script URL changed
    App._startSyncTimer();
  },

  _renderStationsList() {
    const c=document.getElementById('stations-list'); if(!c) return;
    if(!App.pollingStations.length){c.innerHTML='<div class="empty-state"><div class="empty-icon">🏢</div><div class="empty-title">No stations configured</div></div>';return;}
    c.innerHTML=`<div class="table-wrap"><table>
      <thead><tr><th>#</th><th>Zone</th><th>Ward</th><th>Polling Station</th><th>Branch</th><th>Stn Code</th><th>Branch Code</th><th>Action</th></tr></thead>
      <tbody>${App.pollingStations.map((s,i)=>`<tr>
        <td>${i+1}</td><td>${s.zone||'—'}</td><td>${s.ward||'—'}</td><td><strong>${s.name}</strong></td><td>${s.branch}</td>
        <td><span class="badge badge-blue">${s.code}</span></td>
        <td><span class="badge badge-green">${s.branchCode}</span></td>
        <td><button class="btn btn-sm btn-danger" onclick="PageRenderers.removeStation(${i})">🗑️</button></td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  },

  addStation() {
    const g=id=>document.getElementById(id)?.value.trim()||'';
    const zone=g('st-zone'),ward=g('st-ward'),name=g('st-name'),code=g('st-code'),branch=g('st-branch'),bCode=g('st-bcode');
    if(!zone||!ward||!name||!code||!branch||!bCode){Toast.show('Error','All 6 fields are required.','error');return;}
    if(App.pollingStations.find(s=>s.code===code)){Toast.show('Duplicate','Station Code already exists.','error');return;}
    App.pollingStations.push({zone,ward,name,code,branch,branchCode:bCode});
    App.settings.pollingStations=App.pollingStations;App.saveSettings();
    ['st-zone','st-ward','st-name','st-code','st-branch','st-bcode'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
    this._renderStationsList();
    Toast.show('Station Added',`${name} (${code}) saved.`,'success');
    App.logAudit('ADD_STATION',`Added: ${name} (${code}), Zone: ${zone}, Ward: ${ward}`,App.currentUser.username);
  },

  removeStation(i) {
    const s=App.pollingStations[i];
    App.pollingStations.splice(i,1);App.settings.pollingStations=App.pollingStations;App.saveSettings();
    this._renderStationsList();
    Toast.show('Station Removed',`${s.name} removed.`,'warning');
    App.logAudit('REMOVE_STATION',`Removed: ${s.name} (${s.code})`,App.currentUser.username);
  },
};
