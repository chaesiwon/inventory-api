/* ══════════════════════════════════════
   장기재고 관리 시스템 app.js v5
   ══════════════════════════════════════ */

// ── 단위 변환
const Units = {
  amount:'KRW', weight:'ton',
  load() {
    this.amount = localStorage.getItem('u_amt') || 'KRW';
    this.weight = localStorage.getItem('u_wt')  || 'ton';
    document.getElementById('unit-amount').value = this.amount;
    document.getElementById('unit-weight').value = this.weight;
    this.updateBadge();
  },
  save() {
    this.amount = document.getElementById('unit-amount').value;
    this.weight = document.getElementById('unit-weight').value;
    localStorage.setItem('u_amt', this.amount);
    localStorage.setItem('u_wt',  this.weight);
    this.updateBadge();
    const act = document.querySelector('.sidebar-menu li.active')?.dataset.page;
    if(act==='dashboard')  Dashboard.load();
    if(act==='inventory')  Inventory.load();
    if(act==='plans')      Plans.refreshCurrent();
    if(act==='compare')    Compare.search();
    toast('표기 기준 변경됨','inf');
  },
  updateBadge() {
    document.querySelectorAll('.plan-wt-badge').forEach(el=>el.textContent=this.weight);
    const b=document.getElementById('plan-wt-badge'); if(b) b.textContent=this.weight;
  },
  fmtAmt(v) {
    if(v==null||isNaN(v)) return '-';
    const n=Number(v);
    if(this.amount==='MN') return (n/1e6).toLocaleString('ko-KR',{maximumFractionDigits:2})+' 백만원';
    if(this.amount==='HM') return (n/1e8).toLocaleString('ko-KR',{maximumFractionDigits:2})+' 억원';
    return n.toLocaleString('ko-KR')+'원';
  },
  fmtAmtRaw(v) {
    if(v==null||isNaN(v)) return 0;
    const n=Number(v);
    if(this.amount==='MN') return +(n/1e6).toFixed(2);
    if(this.amount==='HM') return +(n/1e8).toFixed(2);
    return Math.round(n);
  },
  amtLabel() { return this.amount==='MN'?'백만원':this.amount==='HM'?'억원':'원'; },
  fmtWt(v) {
    if(v==null||isNaN(v)) return '-';
    const n=Number(v);
    if(this.weight==='kg') return (n*1000).toLocaleString('ko-KR',{maximumFractionDigits:1})+' kg';
    return n.toFixed(3)+' ton';
  },
  fmtWtRaw(v) {
    if(v==null||isNaN(v)) return 0;
    return this.weight==='kg' ? +(Number(v)*1000).toFixed(1) : +Number(v).toFixed(3);
  },
  wtLabel() { return this.weight; }
};

// ── 공통 유틸
const num  = v => v==null?'-':Number(v).toLocaleString('ko-KR');
const dt_  = v => v?String(v).slice(0,10):'-';
const rd_  = v => {
  if(!v) return '-';
  const s=String(v);
  return s.length===8?`${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`:s;
};
const diffStr = (cur,prv,isAmt=true) => {
  if(cur==null||prv==null) return '';
  const d=Number(cur)-Number(prv);
  const cls = d>0?'diff-up':d<0?'diff-dn':'diff-zero';
  const arrow = d>0?'▲':d<0?'▼':'─';
  const val = isAmt?Units.fmtAmt(Math.abs(d)):`${Math.abs(d).toLocaleString()}`;
  const pct = prv?((d/Math.abs(Number(prv)))*100).toFixed(1)+'%':'';
  return `<span class="${cls}">${arrow} ${val} ${pct}</span>`;
};

async function api(path, opts={}) {
  // localStorage에 저장된 인증 정보를 헤더로 전송 (세션 쿠키 백업)
  const authUser = _getAuthUser();
  const baseHeaders = {'Content-Type':'application/json'};
  if(authUser) {
    baseHeaders['X-User-Id']    = String(authUser.id);
    baseHeaders['X-Auth-Token'] = authUser.token || '';
  }
  const mergedOpts = {...opts, headers:{...baseHeaders, ...(opts.headers||{})}};
  const res = await fetch('/api'+path, mergedOpts);
  if(!res.ok){ const e=await res.json().catch(()=>({detail:res.statusText})); throw new Error(e.detail||'오류'); }
  return res.json();
}

function _getAuthUser() {
  try { return JSON.parse(localStorage.getItem('_auth_user') || 'null'); } catch { return null; }
}
function _setAuthUser(user, token) {
  try { localStorage.setItem('_auth_user', JSON.stringify({...user, token})); } catch {}
}
function _clearAuthUser() {
  try { localStorage.removeItem('_auth_user'); } catch {}
}

function toast(msg, type='suc') {
  const el=document.getElementById('toast');
  el.textContent=msg; el.className='toast '+type;
  el.classList.remove('hidden');
  clearTimeout(el._t); el._t=setTimeout(()=>el.classList.add('hidden'),3000);
}

function paging(wId, total, page, size, fn) {
  const el=document.getElementById(wId);
  const pages=Math.ceil(total/size);
  if(pages<=1){el.innerHTML='';return;}
  const s=Math.max(1,page-3),e=Math.min(pages,page+3);
  let h='';
  if(page>1) h+=`<button onclick="${fn}(${page-1})">‹</button>`;
  for(let i=s;i<=e;i++) h+=`<button class="${i===page?'on':''}" onclick="${fn}(${i})">${i}</button>`;
  if(page<pages) h+=`<button onclick="${fn}(${page+1})">›</button>`;
  el.innerHTML=h;
}

let _refDates=[];
async function loadRefDates() {
  try {
    const d=await api('/inventory/ref-dates');
    _refDates=d.ref_dates||[];
    document.querySelectorAll('#dash-ref,#top20-ref,#inv-ref,#noplan-ref,#plan-ref,#cmp-ref,#bulk-inv-ref').forEach(el=>{
      const cur=el.value;
      el.innerHTML='<option value="">최신</option>';
      _refDates.forEach(r=>el.innerHTML+=`<option ${r===cur?'selected':''} value="${r}">${rd_(r)}</option>`);
    });
  } catch(e){}
}

// ══════════════════════════════════════
// App
// ══════════════════════════════════════
const App = {
  async init() {
    Units.load();
    // 1. 서버 세션 확인
    const d=await api('/auth/me').catch(()=>({logged_in:false}));
    if(d.logged_in) {
      _setAuthUser(d.user, _getAuthUser()?.token || '');
      this.showApp(d.user);
      return;
    }
    // 2. localStorage 복원 시도 (세션 만료/쿠키 미작동 대비)
    const saved = _getAuthUser();
    if(saved && saved.id) {
      try {
        const d2 = await api('/auth/me');
        if(d2.logged_in){ this.showApp(d2.user); return; }
      } catch {}
      // 저장된 정보로 바로 화면 표시 (서버 검증 실패 시에도)
      if(saved.username) { this.showApp(saved); return; }
    }
    document.getElementById('login-screen').classList.remove('hidden');
  },
  async login() {
    document.getElementById('login-err').textContent='';
    const username=document.getElementById('login-user').value.trim();
    const password=document.getElementById('login-pw').value;
    try {
      const d=await api('/auth/login',{method:'POST',body:JSON.stringify({username,password})});
      // 세션 + localStorage 이중 저장 (세션 쿠키 미작동 환경 대비)
      _setAuthUser(d.user, d.token || '');
      this.showApp(d.user);
    } catch(e){ document.getElementById('login-err').textContent=e.message; }
  },
  showApp(user) {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('app').classList.remove('hidden');
    document.getElementById('sidebar-user').textContent='👤 '+(user.display_name||user.username);
    document.querySelectorAll('.sidebar-menu li').forEach(li=>{
      li.onclick=()=>{
        document.querySelectorAll('.sidebar-menu li').forEach(x=>x.classList.remove('active'));
        document.querySelectorAll('.page').forEach(p=>{p.classList.remove('active');p.classList.add('hidden');});
        li.classList.add('active');
        const pg=document.getElementById('page-'+li.dataset.page);
        if(pg){pg.classList.remove('hidden');pg.classList.add('active');}
        if(li.dataset.page==='dashboard') Dashboard.load();
        if(li.dataset.page==='inventory') Inventory.search();
        if(li.dataset.page==='plans')     Plans.init();
        if(li.dataset.page==='compare')   Compare.init();
        if(li.dataset.page==='upload')    Upload.loadHist();
      };
    });
    // admin 전용 메뉴 표시/숨김
    const menuUsers = document.getElementById('menu-users');
    if(menuUsers) menuUsers.style.display = user.role==='admin' ? '' : 'none';

    // 모든 사이드바 메뉴 이벤트 등록 (users 포함)
    document.querySelectorAll('.sidebar-menu li[data-page]').forEach(li=>{
      li.onclick=()=>{
        // admin 아닌데 admin 전용 페이지 클릭 방지
        if(li.dataset.page==='users' && user.role!=='admin') {
          toast('관리자만 접근 가능합니다.','err'); return;
        }
        document.querySelectorAll('.sidebar-menu li').forEach(x=>x.classList.remove('active'));
        document.querySelectorAll('.page').forEach(p=>{p.classList.remove('active');p.classList.add('hidden');});
        li.classList.add('active');
        const pg=document.getElementById('page-'+li.dataset.page);
        if(pg){pg.classList.remove('hidden');pg.classList.add('active');}
        if(li.dataset.page==='dashboard') Dashboard.load();
        if(li.dataset.page==='inventory') Inventory.search();
        if(li.dataset.page==='plans')     Plans.init();
        if(li.dataset.page==='compare')   Compare.init();
        if(li.dataset.page==='upload')    Upload.loadHist();
        if(li.dataset.page==='users')     Users.load();
      };
    });
    loadRefDates().then(()=>Dashboard.load());
  },
  async logout() {
    await api('/auth/logout',{method:'POST'}).catch(()=>{});
    _clearAuthUser();
    location.reload();
  }
};

// ══════════════════════════════════════
// Dashboard
// ══════════════════════════════════════
const Dashboard = {
  charts:{},
  async load() {
    const rd=document.getElementById('dash-ref').value;
    const qs=rd?'?ref_date='+rd:'';
    try {
      const [kpi,trend,planWt,ccSum]=await Promise.all([
        api('/dashboard/kpi'+qs),
        api('/dashboard/monthly-trend'),
        api('/dashboard/plan-weight-trend'),
        api('/dashboard/cost-center-summary'+qs),
      ]);
      this.renderKpi(kpi);
      this.renderTrend(trend.trend||[]);
      this.renderPlanWt(planWt.trend||[]);
      this.renderCcSummary(ccSum.items||[]);
      this.loadTop20();
      this.loadCompare();
    } catch(e){ toast('대시보드 오류: '+e.message,'err'); }
  },
  renderKpi(d) {
    document.getElementById('kpi-grid').innerHTML=`
      <div class="kpi c1">
        <div class="kpi-label">총 장기재고 금액</div>
        <div class="kpi-value">${Units.fmtAmt(d.total_amount)}</div>
        <div class="kpi-sub">${num(d.total_count)}건 | ${Units.fmtWt(d.total_weight_ton)}</div>
      </div>
      <div class="kpi c2">
        <div class="kpi-label">당월 소진 예정</div>
        <div class="kpi-value">${Units.fmtAmt(d.plan_this_month)}</div>
        <div class="kpi-sub">이번달 소진계획금액</div>
      </div>
      <div class="kpi c3">
        <div class="kpi-label">미조치 (계획미등록)</div>
        <div class="kpi-value">${Units.fmtAmt(d.uncompleted_amount)}</div>
        <div class="kpi-sub">계획 미등록 금액</div>
      </div>
      <div class="kpi c4">
        <div class="kpi-label">소진 완료</div>
        <div class="kpi-value">${Units.fmtAmt(d.completed_amount)}</div>
        <div class="kpi-sub">계획+실적 확인됨</div>
      </div>
      <div class="kpi c5">
        <div class="kpi-label">소진금액 (전월대비)</div>
        <div class="kpi-value">${Units.fmtAmt(d.total_consumed_amount)}</div>
        <div class="kpi-sub">전월 대비 감소분</div>
      </div>`;
  },
  _mkChart(id,cfg) {
    if(this.charts[id]) this.charts[id].destroy();
    const el=document.getElementById(id);
    if(el) this.charts[id]=new Chart(el.getContext('2d'),cfg);
  },
  renderTrend(trend) {
    this._mkChart('chart-trend',{type:'line',data:{
      labels:trend.map(r=>rd_(r.ref_date)),
      datasets:[
        {label:`장기재고(${Units.amtLabel()})`,tension:.3,fill:true,
         borderColor:'#1a56db',backgroundColor:'rgba(26,86,219,.08)',
         data:trend.map(r=>Units.fmtAmtRaw(r.total_amount))},
        {label:`소진금액(${Units.amtLabel()})`,tension:.3,type:'bar',
         backgroundColor:'rgba(5,150,105,.5)',
         data:trend.map(r=>Units.fmtAmtRaw(r.total_consumed))},
      ]
    },options:{responsive:true,plugins:{legend:{position:'top'}},scales:{y:{beginAtZero:false}}}});
  },
  renderPlanWt(trend) {
    const b=document.getElementById('plan-wt-badge'); if(b) b.textContent=Units.wtLabel();
    this._mkChart('chart-plan-wt',{type:'bar',data:{
      labels:trend.map(r=>r.plan_month||'-'),
      datasets:[
        {label:`계획중량(${Units.wtLabel()})`,backgroundColor:'rgba(26,86,219,.65)',
         data:trend.map(r=>Units.fmtWtRaw(r.plan_weight_ton))},
        {label:'계획건수',type:'line',yAxisID:'y2',tension:.3,
         borderColor:'#d97706',backgroundColor:'rgba(217,119,6,.2)',
         data:trend.map(r=>r.plan_count)},
      ]
    },options:{responsive:true,scales:{y:{title:{display:true,text:Units.wtLabel()}},
      y2:{position:'right',title:{display:true,text:'건'},grid:{drawOnChartArea:false}}}}});
  },
  renderCcSummary(items) {
    const rows=items.map(r=>`<tr>
      <td>${r.cc_name||'-'}</td>
      <td class="num">${num(r.item_count)}</td>
      <td class="num">${Units.fmtWt(r.total_weight)}</td>
      <td class="num">${Units.fmtAmt(r.total_amount)}</td>
      <td class="num">${Units.fmtAmt(r.consumed_amount)}</td>
      <td class="num">${num(r.plan_count)}</td>
      <td class="num">${num(r.actual_count)}</td>
      <td class="num">${r.item_count?((r.actual_count/r.item_count)*100).toFixed(1)+'%':'0%'}</td>
    </tr>`).join('');
    document.getElementById('cc-body').innerHTML=rows||
      '<tr><td colspan="8" style="text-align:center;padding:20px;color:#9ca3af">데이터 없음</td></tr>';
  },
  async loadTop20() {
    const rd=document.getElementById('top20-ref').value;
    const qs=rd?'?ref_date='+rd:'';
    try {
      const d=await api('/dashboard/top20'+qs);
      const rows=(d.items||[]).map((r,i)=>`<tr>
        <td>${i+1}</td><td>${r.factory}</td><td>${r.item_type}</td>
        <td title="${r.item_name}">${r.item_name.slice(0,18)}</td>
        <td>${r.cc_name||r.cost_center||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td class="num">${Units.fmtAmt(r.amount_consumed)}</td>
        <td>${rd_(r.base_date)}</td>
        <td>${r.plan_type||'-'}</td><td>${r.plan_date||'-'}</td>
        <td><span class="badge ${r.is_completed?'b-ok':'b-warn'}">${r.is_completed?'완료':'미완료'}</span></td>
        <td><button class="btn btn-xs btn-outline"
          onclick="Modal.open('${r.lot_no}','${(r.item_name||'').replace(/'/g,"\\'")}',${r.amount||0},${r.weight_ton||0})">계획입력</button></td>
      </tr>`).join('');
      document.getElementById('top20-wrap').innerHTML=`
        <table class="data-table"><thead><tr>
          <th>순위</th><th>공장</th><th>품목구분</th><th>품명</th><th>원가중심점</th>
          <th>LOT NO</th><th>중량</th><th>금액</th><th>소진금액</th><th>기준일자</th>
          <th>소진계획</th><th>계획기한</th><th>완료여부</th><th>편집</th>
        </tr></thead><tbody>${rows}</tbody></table>`;
    } catch(e){}
  },
  async loadCompare() {
    const rd=document.getElementById('dash-ref').value;
    const mode=document.getElementById('dash-compare-mode').value;
    const qs=new URLSearchParams({ref_date:rd,mode}).toString();
    try {
      const d=await api('/dashboard/period-compare?'+qs);
      const cur=d.current, prv=d.previous, lbl=d.mode_label||'전월';
      if(!cur){document.getElementById('compare-cards').innerHTML='';return;}
      const cards=[
        {label:'총 장기재고 금액',cur:cur.total_amount,prv:prv?.total_amount,isAmt:true},
        {label:'총 중량',cur:cur.total_weight,prv:prv?.total_weight,isAmt:false},
        {label:'소진 확인 금액',cur:cur.action_amount,prv:prv?.action_amount,isAmt:true},
      ];
      document.getElementById('compare-cards').innerHTML=cards.map(c=>`
        <div class="cmp-card">
          <div class="cmp-card-title">${c.label} (${lbl} 대비)</div>
          <div class="cmp-card-value">${c.isAmt?Units.fmtAmt(c.cur):Units.fmtWt(c.cur)}</div>
          <div class="cmp-card-diff">${prv?diffStr(c.cur,c.prv,c.isAmt):'이전 데이터 없음'}</div>
        </div>`).join('');
    } catch(e){}
  }
};

// ══════════════════════════════════════
// Inventory (조회전용)
// ══════════════════════════════════════
const Inventory = {
  pg:1, _t:null,
  get p(){return{
    ref_date:document.getElementById('inv-ref').value,
    factory:document.getElementById('inv-fac').value,
    item_type:document.getElementById('inv-type').value,
    item_code:document.getElementById('inv-code').value.trim(),
    lot_no:document.getElementById('inv-lot').value.trim(),
    item_name:document.getElementById('inv-iname').value.trim(),
    cost_center:document.getElementById('inv-cc').value.trim(),
    dept:document.getElementById('inv-dept').value,
  };},
  search(){this.pg=1;this.load();},
  debounce(){clearTimeout(this._t);this._t=setTimeout(()=>this.search(),400);},
  reset(){
    ['inv-ref','inv-fac','inv-type','inv-dept'].forEach(id=>document.getElementById(id).value='');
    ['inv-code','inv-lot','inv-iname','inv-cc'].forEach(id=>document.getElementById(id).value='');
    this.search();
  },
  async load(){
    const qs=new URLSearchParams({...this.p,page:this.pg,page_size:50}).toString();
    try{
      const d=await api('/inventory?'+qs);
      document.getElementById('inv-info').innerHTML=
        `<b>총 ${num(d.total)}건</b> &nbsp;·&nbsp; 기준일: <b>${rd_(d.ref_date)}</b>`;
      const rows=(d.items||[]).map(r=>`<tr>
        <td>${r.factory}</td><td>${r.item_type}</td>
        <td>${r.item_code}</td><td title="${r.item_name}">${r.item_name.slice(0,22)}</td>
        <td>${r.cc_name||r.cost_center||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td class="num">${Units.fmtAmt(r.amount_consumed)}</td>
        <td>${rd_(r.base_date)}</td><td>${r.months_label||'-'}</td>
        <td><span class="badge ${r.is_new?'b-new':'b-old'}">${r.is_new?'신규':'기존'}</span></td>
        <td>${r.dept||'-'}</td><td>${r.reason||'-'}</td>
        <td>${r.plan_type||'-'}</td><td>${r.plan_date||'-'}</td>
        <td title="${r.detail_plan||''}">${(r.detail_plan||'').slice(0,12)||'-'}</td>
        <td><span class="badge ${r.has_actual?'b-ok':'b-warn'}">${r.has_actual?'있음':'없음'}</span></td>
      </tr>`).join('');
      document.getElementById('inv-tbody').innerHTML=rows||
        '<tr><td colspan="18" style="text-align:center;padding:24px;color:#9ca3af">결과 없음</td></tr>';
      paging('inv-page',d.total,d.page,50,'Inventory.goPage');
    }catch(e){toast('조회 오류: '+e.message,'err');}
  },
  goPage(p){this.pg=p;this.load();},
  exportExcel(){window.location.href='/api/compare/export?'+new URLSearchParams(this.p).toString();},
  exportPpt(){window.location.href='/api/compare/export-ppt?'+new URLSearchParams(this.p).toString();='/api/inventory/export?'+new URLSearchParams(this.p).toString();}
};

// ══════════════════════════════════════
// Plans
// ══════════════════════════════════════
const Plans = {
  currentTab:'noplan', noplanPg:1, regPg:1, bulkInvPg:1,
  bulkSel:new Set(), _nt:null, _rt:null, _bit:null,

  init(){this.switchTab('noplan');},
  refreshCurrent(){
    if(this.currentTab==='noplan')     this.loadNoPlan();
    if(this.currentTab==='registered') this.loadRegistered();
    if(this.currentTab==='bulk')       this.loadBulkInventory();
  },
  switchTab(tab){
    // 탭 전환 시 일괄입력 선택 상태 완전 초기화
    this._resetBulkState();
    this.currentTab=tab;
    document.querySelectorAll('#page-plans .tab-btn').forEach((b,i)=>{
      b.classList.toggle('active',['noplan','registered','bulk'][i]===tab);
    });
    ['noplan','registered','bulk'].forEach((t,i)=>{
      document.getElementById('tab-'+t).classList.toggle('hidden',i!=['noplan','registered','bulk'].indexOf(tab));
    });
    if(tab==='noplan')     this.loadNoPlan();
    if(tab==='registered') this.loadRegistered();
    if(tab==='bulk')       this.loadBulkInventory();
  },

  // ─ 미등록 재고
  noplanDebounce(){clearTimeout(this._nt);this._nt=setTimeout(()=>this.loadNoPlan(),400);},
  async loadNoPlan(){
    const qs=new URLSearchParams({
      ref_date:document.getElementById('noplan-ref').value,
      factory:document.getElementById('noplan-fac').value,
      item_name:document.getElementById('noplan-iname').value.trim(),
      cost_center:document.getElementById('noplan-cc').value.trim(),
      lot_no:document.getElementById('noplan-lot').value.trim(),
      page:this.noplanPg, page_size:50
    }).toString();
    try{
      const d=await api('/plans/no-plan?'+qs);
      document.getElementById('noplan-info').innerHTML=`<b>소진계획 미등록: ${num(d.total)}건</b>`;
      const badge=document.getElementById('noplan-cnt'); if(badge) badge.textContent=d.total||'';
      const rows=(d.items||[]).map(r=>`<tr>
        <td>${r.factory}</td><td>${r.item_type}</td>
        <td>${r.item_code}</td><td title="${r.item_name}">${r.item_name.slice(0,22)}</td>
        <td>${r.cc_name||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td>${rd_(r.base_date)}</td>
        <td><span class="badge ${r.is_new?'b-new':'b-old'}">${r.is_new?'신규':'기존'}</span></td>
        <td><button class="btn btn-xs btn-primary"
          onclick="Modal.open('${r.lot_no}','${(r.item_name||'').replace(/'/g,"\\'")}',${r.amount||0},${r.weight_ton||0})">입력</button></td>
      </tr>`).join('');
      document.getElementById('noplan-tbody').innerHTML=rows||
        '<tr><td colspan="11" style="text-align:center;padding:24px;color:#9ca3af">미등록 재고 없음 ✅</td></tr>';
      paging('noplan-page',d.total,d.page,50,'Plans.noplanGoPage');
    }catch(e){toast('조회 오류: '+e.message,'err');}
  },
  noplanGoPage(p){this.noplanPg=p;this.loadNoPlan();},

  // ─ 등록완료
  regDebounce(){clearTimeout(this._rt);this._rt=setTimeout(()=>this.loadRegistered(),400);},
  async loadRegistered(){
    const qs=new URLSearchParams({
      ref_date:document.getElementById('plan-ref').value,
      factory:document.getElementById('plan-fac').value,
      dept:document.getElementById('plan-dept').value,
      plan_type:document.getElementById('plan-type').value,
      lot_no:document.getElementById('plan-lot').value.trim(),
      cost_center:document.getElementById('plan-cc').value.trim(),
      page:this.regPg, page_size:50
    }).toString();
    try{
      const d=await api('/plans?'+qs);
      document.getElementById('plan-info').innerHTML=`<b>등록 완료: ${num(d.total)}건</b>`;
      const rows=(d.items||[]).map(r=>`<tr>
        <td>${r.factory}</td><td>${r.item_type}</td>
        <td>${r.item_code}</td><td title="${r.item_name}">${r.item_name.slice(0,22)}</td>
        <td>${r.cc_name||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td>${rd_(r.base_date)}</td>
        <td>${r.dept||'-'}</td><td>${r.reason||'-'}</td>
        <td>${r.plan_type||'-'}</td><td>${r.plan_date||'-'}</td>
        <td title="${r.detail_plan||''}">${(r.detail_plan||'').slice(0,12)||'-'}</td>
        <td>${r.created_by_name||'-'}</td>
        <td>${r.plan_created_at?r.plan_created_at.slice(0,16):'-'}</td>
        <td>${r.updated_by_name||'-'}</td>
        <td>${r.plan_updated_at?r.plan_updated_at.slice(0,16):'-'}</td>
        <td style="display:flex;gap:3px">
          <button class="btn btn-xs btn-outline"
            onclick="Modal.open('${r.lot_no}','${(r.item_name||'').replace(/'/g,"\\'")}',${r.amount||0},${r.weight_ton||0})">수정</button>
          <button class="btn btn-xs btn-danger" onclick="Plans.deletePlan('${r.lot_no}')">삭제</button>
        </td>
      </tr>`).join('');
      document.getElementById('plan-tbody').innerHTML=rows||
        '<tr><td colspan="15" style="text-align:center;padding:24px;color:#9ca3af">결과 없음</td></tr>';
      paging('plan-page',d.total,d.page,50,'Plans.regGoPage');
    }catch(e){toast('조회 오류: '+e.message,'err');}
  },
  regGoPage(p){this.regPg=p;this.loadRegistered();},
  async deletePlan(lot_no){
    if(!confirm(`LOT NO: ${lot_no}\n소진계획을 삭제하시겠습니까?`)) return;
    try{
      await api('/plans/'+encodeURIComponent(lot_no),{method:'DELETE'});
      toast('삭제 완료'); this.loadRegistered(); this.loadNoPlan();
    }catch(e){toast('삭제 오류: '+e.message,'err');}
  },

  // ─ 일괄입력용 재고목록
  bulkInvDebounce(){clearTimeout(this._bit);this._bit=setTimeout(()=>this.loadBulkInventory(),400);},
  async loadBulkInventory(){
    // 검색조건 변경 시 기존 선택 상태 초기화 (append 방지)
    this._resetBulkState();
    const qs=new URLSearchParams({
      ref_date:document.getElementById('bulk-inv-ref').value,
      lot_no:document.getElementById('bulk-inv-lot').value.trim(),
      cost_center:document.getElementById('bulk-inv-cc').value.trim(),
      item_name:document.getElementById('bulk-inv-iname').value.trim(),
      page:this.bulkInvPg, page_size:50
    }).toString();
    try{
      const d=await api('/inventory?'+qs);
      const rows=(d.items||[]).map(r=>`<tr>
        <td><input type="checkbox" class="bulk-chk" value="${r.lot_no}"
          ${this.bulkSel.has(r.lot_no)?'checked':''}
          onchange="Plans.toggleBulkItem('${r.lot_no}',this.checked)"></td>
        <td>${r.factory}</td><td>${r.item_code}</td>
        <td title="${r.item_name}">${r.item_name.slice(0,20)}</td>
        <td>${r.cc_name||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td>${rd_(r.base_date)}</td>
      </tr>`).join('');
      document.getElementById('bulk-inv-tbody').innerHTML=rows||
        '<tr><td colspan="9" style="text-align:center;padding:20px;color:#9ca3af">데이터 없음</td></tr>';
      paging('bulk-inv-page',d.total,d.page,50,'Plans.bulkInvGoPage');
    }catch(e){toast('재고 로드 오류: '+e.message,'err');}
  },
  bulkInvGoPage(p){this.bulkInvPg=p;this.loadBulkInventory();},
  toggleBulkItem(lot,checked){ checked?this.bulkSel.add(lot):this.bulkSel.delete(lot); this.syncLots(); },
  syncLots(){
    const ta=document.getElementById('bulk-lots');
    const ex=ta.value.split('\n').map(s=>s.trim()).filter(s=>s);
    const merged=[...new Set([...ex,...this.bulkSel])];
    ta.value=merged.join('\n');
  },
  selectAllBulk(){
    // 현재 페이지에 보이는 체크박스만 선택 (이전 데이터 누적 방지)
    // 먼저 기존 LOT NO 목록 초기화
    this.bulkSel.clear();
    document.getElementById('bulk-lots').value='';
    // 현재 조회된 항목만 선택
    document.querySelectorAll('.bulk-chk').forEach(cb=>{cb.checked=true;this.bulkSel.add(cb.value);});
    this.syncLots();
  },
  clearBulkSelect(){ document.querySelectorAll('.bulk-chk').forEach(cb=>{cb.checked=false;this.bulkSel.delete(cb.value);}); document.getElementById('bulk-lots').value=''; },

  async bulkSave(){
    const dept   =document.getElementById('bulk-dept').value;
    const reason =document.getElementById('bulk-reason').value;
    const ptype  =document.getElementById('bulk-ptype').value;
    const pdate  =document.getElementById('bulk-date').value;
    const detail =document.getElementById('bulk-detail').value;
    const lots   =document.getElementById('bulk-lots').value.split('\n').map(s=>s.trim()).filter(s=>s);
    if(!dept||!reason||!ptype||!pdate){toast('필수 항목을 모두 입력하세요.','err');return;}
    if(!lots.length){toast('LOT NO를 입력하거나 목록에서 선택하세요.','err');return;}
    const res=document.getElementById('bulk-result');
    res.className='inf'; res.textContent=`⏳ ${lots.length}건 저장 중...`; res.classList.remove('hidden');
    let ok=0,fail=0,errors=[];
    for(const lot of lots){
      try{
        await api('/plans/'+encodeURIComponent(lot),{method:'POST',body:JSON.stringify({dept,reason,plan_type:ptype,plan_date:pdate,detail_plan:detail})});
        ok++;
      }catch(e){fail++;errors.push(`${lot}: ${e.message}`);}
    }
    res.className=fail===0?'suc':'err';
    res.innerHTML=`${fail===0?'✅':'⚠️'} 완료 — 성공 <b>${ok}건</b>, 실패 <b>${fail}건</b>`+(errors.length?'<br>'+errors.slice(0,3).join('<br>'):'');
    if(ok>0){toast(`${ok}건 저장 완료`); this.loadNoPlan();}
  },
  _resetBulkState(){
    // 선택 상태 완전 초기화 (조회조건 변경 시 호출)
    this.bulkSel.clear();
    document.querySelectorAll('.bulk-chk').forEach(cb=>cb.checked=false);
    const ta=document.getElementById('bulk-lots');
    if(ta) ta.value='';
  },
  bulkClear(){
    ['bulk-dept','bulk-reason','bulk-ptype'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('bulk-date').value='';
    document.getElementById('bulk-detail').value='';
    this._resetBulkState();
    const res=document.getElementById('bulk-result');
    if(res) res.classList.add('hidden');
  },
  downloadTemplate(){
    const rd=document.getElementById('plan-ref').value;
    window.location.href='/api/plans/export-template'+(rd?'?ref_date='+rd:'');
  },
  async excelUpload(input){
    const f=input.files[0]; if(!f) return;
    const fd=new FormData(); fd.append('file',f);
    try{
      const d=await fetch('/api/plans/bulk-upload',{method:'POST',body:fd}).then(r=>r.json());
      if(d.ok){
        if(d.fail>0) toast(`업로드 완료: 성공 ${d.success}건, 실패 ${d.fail}건`,'inf');
        else toast(`엑셀 업로드 완료 ${d.success}건`);
        if(d.errors&&d.errors.length) console.warn('업로드 오류 상세:', d.errors);
        this.loadRegistered(); this.loadNoPlan();
      } else toast('업로드 실패: '+(d.detail||''),'err');
    }catch(e){toast('업로드 오류: '+e.message,'err');}
    input.value='';
  }
};

// ══════════════════════════════════════
// Compare
// ══════════════════════════════════════
const Compare = {
  currentTab:'summary', pg:1, charts:{},
  init(){this.switchTab('summary');},
  switchTab(tab){
    this.currentTab=tab;
    document.querySelectorAll('#page-compare .tab-btn').forEach((b,i)=>{
      b.classList.toggle('active',['summary','list'][i]===tab);
    });
    document.getElementById('tab-cmp-summary').classList.toggle('hidden',tab!=='summary');
    document.getElementById('tab-cmp-list').classList.toggle('hidden',tab!=='list');
    this.search();
  },
  search(){this.pg=1;this.load();},
  get p(){return{
    ref_date:document.getElementById('cmp-ref').value,
    factory:document.getElementById('cmp-fac').value,
    dept:document.getElementById('cmp-dept').value,
  };},
  async load(){
    try{
      if(this.currentTab==='summary') await this.loadSummary();
      else await this.loadList();
    }catch(e){toast('비교 오류: '+e.message,'err');}
  },
  _mkChart(id,cfg){
    if(this.charts[id]) this.charts[id].destroy();
    const el=document.getElementById(id);
    if(el) this.charts[id]=new Chart(el.getContext('2d'),cfg);
  },
  async loadSummary(){
    const qs=new URLSearchParams(this.p).toString();
    const [summ,planWt]=await Promise.all([
      api('/compare/summary?'+qs),
      api('/dashboard/plan-weight-trend'),
    ]);
    // KPI
    const plan_total=summ.plan_total||1;
    const action_cnt=summ.action_count||0;
    const no_action_cnt=summ.no_action_count||0;
    document.getElementById('cmp-kpi').innerHTML=`
      <div class="kpi c4">
        <div class="kpi-label">달성률 (조치/계획전체)</div>
        <div class="kpi-value">${(action_cnt/plan_total*100).toFixed(1)}%</div>
        <div class="kpi-sub">조치 ${num(action_cnt)} / 전체 ${num(plan_total)}</div>
      </div>
      <div class="kpi c3">
        <div class="kpi-label">미조치 건수</div>
        <div class="kpi-value">${num(no_action_cnt)}건</div>
        <div class="kpi-sub">중량: ${Units.fmtWt(summ.no_action_weight)}</div>
      </div>
      <div class="kpi c3" style="border-top-color:#9333ea">
        <div class="kpi-label">미조치 금액</div>
        <div class="kpi-value">${Units.fmtAmt(summ.no_action_amount)}</div>
        <div class="kpi-sub">계획+실적 미일치</div>
      </div>
      <div class="kpi c1">
        <div class="kpi-label">조치 금액</div>
        <div class="kpi-value">${Units.fmtAmt(summ.action_amount)}</div>
        <div class="kpi-sub">중량: ${Units.fmtWt(summ.action_weight)}</div>
      </div>
      <div class="kpi c5">
        <div class="kpi-label">소진금액 (전월대비)</div>
        <div class="kpi-value">${Units.fmtAmt(summ.consumed_amount)}</div>
        <div class="kpi-sub">LOT 수량 감소분</div>
      </div>`;

    // 검증: 전체=조치+미조치
    const chk=(action_cnt+no_action_cnt===plan_total);
    document.getElementById('cmp-kpi').insertAdjacentHTML('beforeend',`
      <div style="grid-column:1/-1;font-size:12px;color:${chk?'#059669':'#dc2626'};padding:4px 0">
        ✔ 전체(${plan_total}) = 조치(${action_cnt}) + 미조치(${no_action_cnt}) ${chk?'✅':'❌'}
      </div>`);

    // 유형별 건수
    const allTypes=[...new Set([...(summ.plan_by_type||[]).map(r=>r.plan_type||'미등록'),...(summ.actual_by_type||[]).map(r=>r.actual_type||'기타')])];
    const pMap=Object.fromEntries((summ.plan_by_type||[]).map(r=>[r.plan_type||'미등록',r.plan_count]));
    const aMap=Object.fromEntries((summ.actual_by_type||[]).map(r=>[r.actual_type||'기타',r.actual_count]));
    const pwMap=Object.fromEntries((summ.plan_by_type||[]).map(r=>[r.plan_type||'미등록',r.plan_weight]));
    const awMap=Object.fromEntries((summ.actual_by_type||[]).map(r=>[r.actual_type||'기타',r.actual_weight]));
    this._mkChart('chart-cmp-count',{type:'bar',data:{labels:allTypes,datasets:[
      {label:'계획 건수',backgroundColor:'rgba(26,86,219,.7)',data:allTypes.map(t=>pMap[t]||0)},
      {label:'실적 건수',backgroundColor:'rgba(5,150,105,.7)',data:allTypes.map(t=>aMap[t]||0)},
    ]},options:{responsive:true,plugins:{legend:{position:'top'}}}});
    this._mkChart('chart-cmp-wt',{type:'bar',data:{labels:allTypes,datasets:[
      {label:`계획 중량(${Units.wtLabel()})`,backgroundColor:'rgba(26,86,219,.7)',data:allTypes.map(t=>Units.fmtWtRaw(pwMap[t]||0))},
      {label:`실적 중량(${Units.wtLabel()})`,backgroundColor:'rgba(5,150,105,.7)',data:allTypes.map(t=>Units.fmtWtRaw(awMap[t]||0))},
    ]},options:{responsive:true,plugins:{legend:{position:'top'}}}});
    const pw=planWt.trend||[];
    this._mkChart('chart-plan-monthly',{type:'bar',data:{
      labels:pw.map(r=>r.plan_month||'-'),
      datasets:[{label:`계획 중량(${Units.wtLabel()})`,backgroundColor:'rgba(26,86,219,.6)',
        data:pw.map(r=>Units.fmtWtRaw(r.plan_weight_ton))}]
    },options:{responsive:true,scales:{y:{beginAtZero:true}}}});
  },
  async loadList(){
    const qs=new URLSearchParams({...this.p,page:this.pg,page_size:50}).toString();
    const d=await api('/compare?'+qs);
    document.getElementById('cmp-info').innerHTML=
      `<b>총 ${num(d.total)}건</b> (계획 등록 LOT 기준)`;
    const rows=(d.items||[]).map(r=>{
      const match=r.type_match==null?'<span class="badge b-old">-</span>':
        r.type_match?'<span class="badge b-blue">일치</span>':'<span class="badge b-mismatch">불일치</span>';
      const act=r.action_status==='조치'?
        '<span class="badge b-action">조치</span>':'<span class="badge b-noaction">미조치</span>';
      const atype=r.actual_type_manual||r.actual_type||'-';
      return `<tr>
        <td>${r.factory}</td><td>${r.item_type}</td><td>${r.item_code}</td>
        <td title="${r.item_name}">${r.item_name.slice(0,20)}</td>
        <td>${r.cc_name||r.cost_center||'-'}</td><td>${r.lot_no}</td>
        <td class="num">${Units.fmtWt(r.weight_ton)}</td>
        <td class="num">${Units.fmtAmt(r.amount)}</td>
        <td class="num">${Units.fmtAmt(r.amount_consumed)}</td>
        <td>${rd_(r.base_date)}</td><td>${r.dept||'-'}</td>
        <td>${r.plan_type||'-'}</td><td>${r.plan_date||'-'}</td>
        <td>${atype}</td><td>${dt_(r.process_date)}</td>
        <td>${match}</td><td>${act}</td>
        <td>${r.actual_id?`<button class="btn btn-xs btn-outline" onclick="AModal.open(${r.actual_id},'${atype}')">수정</button>`:'-'}</td>
      </tr>`;
    }).join('');
    document.getElementById('cmp-tbody').innerHTML=rows||
      '<tr><td colspan="18" style="text-align:center;padding:24px;color:#9ca3af">결과 없음</td></tr>';
    paging('cmp-page',d.total,d.page,50,'Compare.goPage');
  },
  goPage(p){this.pg=p;this.load();},
  exportExcel(){window.location.href='/api/compare/export?'+new URLSearchParams(this.p).toString();}
};

// ══════════════════════════════════════
// Upload
// ══════════════════════════════════════
const Upload = {
  drop(e){e.preventDefault();const f=e.dataTransfer.files[0];if(f)this._do(f);},
  upload(input){const f=input.files[0];if(f)this._do(f);input.value='';},
  async _do(file){
    const res=document.getElementById('up-result');
    res.className='inf';res.innerHTML='⏳ 업로드 처리 중...';res.classList.remove('hidden');
    const fd=new FormData();fd.append('file',file);
    try{
      const d=await fetch('/api/upload',{method:'POST',body:fd}).then(r=>r.json());
      if(d.ok){
        res.className='suc';
        res.innerHTML=`✅ <b>업로드 완료</b><br>
          기준일: <b>${rd_(d.ref_date)}</b> | 재고 <b>${d.inv_count}건</b> | 재공 <b>${d.wip_count}건</b> | 실적 <b>${d.act_count}건</b><br>
          총 금액: <b>${Units.fmtAmt(d.total_amount)}</b>`;
        toast('업로드 완료!'); loadRefDates(); this.loadHist();
      } else {res.className='err';res.textContent='❌ '+(d.detail||'오류');}
    }catch(e){res.className='err';res.textContent='❌ '+e.message;}
  },
  async deleteOne(uid,fn){
    if(!confirm(`파일: ${fn}\n이 업로드 데이터를 삭제하시겠습니까?`)) return;
    try{await api('/upload/'+uid,{method:'DELETE'});toast('삭제 완료');loadRefDates();this.loadHist();}
    catch(e){toast('삭제 오류: '+e.message,'err');}
  },
  async deleteAll(){
    if(!confirm('⚠️ 전체 업로드 데이터를 삭제합니다.\n소진계획은 유지됩니다.')) return;
    try{await api('/upload/all/data',{method:'DELETE'});toast('전체 삭제 완료');loadRefDates();this.loadHist();}
    catch(e){toast('삭제 오류: '+e.message,'err');}
  },
  async loadHist(){
    try{
      const d=await api('/upload-history');
      const rows=(d.history||[]).map(r=>`<tr>
        <td title="${r.filename}">${(r.filename||'').slice(0,28)}</td>
        <td>${rd_(r.ref_date)}</td>
        <td class="num">${num(r.inv_count)}</td><td class="num">${num(r.wip_count)}</td>
        <td class="num">${num(r.act_count)}</td>
        <td class="num">${Units.fmtAmt(r.total_amount)}</td>
        <td>${r.uploaded_by||'-'}</td><td>${r.created_at||'-'}</td>
        <td><button class="btn btn-xs btn-danger"
          onclick="Upload.deleteOne('${r.upload_id}','${(r.filename||'').replace(/'/g,"\\'")}')">삭제</button></td>
      </tr>`).join('');
      document.getElementById('hist-tbody').innerHTML=rows||
        '<tr><td colspan="9" style="text-align:center;padding:20px;color:#9ca3af">이력 없음</td></tr>';
    }catch(e){}
  }
};

// ══════════════════════════════════════
// Modal / AModal
// ══════════════════════════════════════
const Modal = {
  lot:null,
  async open(lot_no,item_name,amount,wt){
    this.lot=lot_no;
    document.getElementById('modal-info').innerHTML=
      `<b>LOT NO:</b> ${lot_no}<br><b>품명:</b> ${item_name}<br>`+
      `<b>금액:</b> ${Units.fmtAmt(amount)} &nbsp; <b>중량:</b> ${Units.fmtWt(wt)}`;
    try{
      const d=await api('/plans?lot_no_exact='+encodeURIComponent(lot_no));
      const ex=(d.items||[])[0];
      document.getElementById('f-dept').value   =ex?.dept||'';
      document.getElementById('f-reason').value =ex?.reason||'';
      document.getElementById('f-ptype').value  =ex?.plan_type||'';
      document.getElementById('f-date').value   =ex?.plan_date||'';
      document.getElementById('f-detail').value =ex?.detail_plan||'';
    }catch{}
    document.getElementById('modal').classList.remove('hidden');
  },
  close(){document.getElementById('modal').classList.add('hidden');},
  async save(){
    const body={dept:document.getElementById('f-dept').value,reason:document.getElementById('f-reason').value,
      plan_type:document.getElementById('f-ptype').value,plan_date:document.getElementById('f-date').value,
      detail_plan:document.getElementById('f-detail').value};
    if(!body.dept||!body.reason||!body.plan_type||!body.plan_date){toast('필수 항목을 모두 입력하세요.','err');return;}
    try{
      await api('/plans/'+encodeURIComponent(this.lot),{method:'POST',body:JSON.stringify(body)});
      toast('저장 완료!'); this.close();
      const act=document.querySelector('.sidebar-menu li.active')?.dataset.page;
      if(act==='dashboard') Dashboard.load();
      if(act==='plans')     Plans.refreshCurrent();
      if(act==='compare')   Compare.search();
    }catch(e){toast('저장 오류: '+e.message,'err');}
  }
};

const AModal = {
  id:null,
  open(id,cur){this.id=id;document.getElementById('af-type').value=cur||'기타';document.getElementById('amodal').classList.remove('hidden');},
  close(){document.getElementById('amodal').classList.add('hidden');},
  async save(){
    try{
      await api('/actuals/'+this.id+'/type',{method:'PATCH',body:JSON.stringify({actual_type_manual:document.getElementById('af-type').value})});
      toast('수정 완료!');this.close();Compare.search();
    }catch(e){toast('수정 오류: '+e.message,'err');}
  }
};

document.addEventListener('keydown',e=>{
  if(e.key==='Enter'&&!document.getElementById('login-screen').classList.contains('hidden')) App.login();
  if(e.key==='Escape'){Modal.close();AModal.close();}
});

App.init();

// ══════════════════════════════════════
// Users (사용자 관리 - admin only)
// ══════════════════════════════════════
const Users = {
  async load() {
    try {
      const d = await api('/users');
      const roleLabel = {admin:'관리자', user:'일반사용자'};
      const rows = (d.users||[]).map(u=>`<tr>
        <td>${u.id}</td>
        <td><b>${u.username}</b></td>
        <td>${u.display_name||'-'}</td>
        <td><span class="badge ${u.role==='admin'?'b-blue':'b-old'}">${roleLabel[u.role]||u.role}</span></td>
        <td>${u.department||'-'}</td>
        <td><span class="badge ${u.is_active?'b-ok':'b-warn'}">${u.is_active?'활성':'비활성'}</span></td>
        <td>${u.last_login?u.last_login.slice(0,16):'-'}</td>
        <td style="display:flex;gap:3px">
          <button class="btn btn-xs btn-outline" onclick="Users.openEditModal(${JSON.stringify(u).replace(/"/g,'&quot;')})">수정</button>
          <button class="btn btn-xs ${u.is_active?'btn-danger':'btn-outline'}"
            onclick="Users.toggleActive(${u.id},${u.is_active})">${u.is_active?'비활성화':'활성화'}</button>
        </td>
      </tr>`).join('');
      document.getElementById('users-tbody').innerHTML = rows ||
        '<tr><td colspan="8" style="text-align:center;padding:20px;color:#9ca3af">사용자 없음</td></tr>';
    } catch(e) { toast('사용자 목록 조회 실패: '+e.message, 'err'); }
  },
  openCreateModal() {
    document.getElementById('umodal-title').textContent = '사용자 추가';
    document.getElementById('u-id').value = '';
    document.getElementById('u-username').value = '';
    document.getElementById('u-username').disabled = false;
    document.getElementById('u-dname').value = '';
    document.getElementById('u-pw').value = '';
    document.getElementById('u-pw-req').textContent = '*';
    document.getElementById('u-role').value = 'user';
    document.getElementById('u-dept').value = '';
    document.getElementById('u-active').value = '1';
    document.getElementById('umodal').classList.remove('hidden');
  },
  openEditModal(user) {
    document.getElementById('umodal-title').textContent = '사용자 수정';
    document.getElementById('u-id').value = user.id;
    document.getElementById('u-username').value = user.username;
    document.getElementById('u-username').disabled = true;
    document.getElementById('u-dname').value = user.display_name||'';
    document.getElementById('u-pw').value = '';
    document.getElementById('u-pw-req').textContent = '(비워두면 변경 안함)';
    document.getElementById('u-role').value = user.role||'user';
    document.getElementById('u-dept').value = user.department||'';
    document.getElementById('u-active').value = user.is_active??1;
    document.getElementById('umodal').classList.remove('hidden');
  },
  closeModal() { document.getElementById('umodal').classList.add('hidden'); },
  async save() {
    const uid = document.getElementById('u-id').value;
    const isEdit = !!uid;
    const body = {
      display_name: document.getElementById('u-dname').value,
      role:         document.getElementById('u-role').value,
      department:   document.getElementById('u-dept').value,
      is_active:    parseInt(document.getElementById('u-active').value),
    };
    const pw = document.getElementById('u-pw').value;
    if(pw) body.password = pw;
    try {
      if(isEdit) {
        await api('/users/'+uid, {method:'PUT', body:JSON.stringify(body)});
        toast('사용자 정보 수정 완료');
      } else {
        const username = document.getElementById('u-username').value.trim();
        if(!username) { toast('아이디를 입력하세요.','err'); return; }
        if(!pw) { toast('비밀번호를 입력하세요.','err'); return; }
        const dname = document.getElementById('u-dname').value.trim();
        if(!dname) { toast('이름을 입력하세요.','err'); return; }
        await api('/users', {method:'POST', body:JSON.stringify({...body, username, password:pw, display_name:dname})});
        toast('사용자 추가 완료');
      }
      this.closeModal();
      this.load();
    } catch(e) { toast('저장 실패: '+e.message, 'err'); }
  },
  async toggleActive(id, currentActive) {
    const newState = currentActive ? 0 : 1;
    const label = newState ? '활성화' : '비활성화';
    if(!confirm(`사용자를 ${label}하시겠습니까?`)) return;
    try {
      await api('/users/'+id, {method:'PUT', body:JSON.stringify({is_active:newState})});
      toast(`${label} 완료`);
      this.load();
    } catch(e) { toast('변경 실패: '+e.message, 'err'); }
  }
};
