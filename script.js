// Paste the deployed Google Apps Script Web App URL here.
const API_URL = "https://script.google.com/macros/s/AKfycby_A9iKIrfCa22aDZLXlKk7CWz4hsiGOx5ufFREVPRWQ8ds_xfi772D-SDfT6QidYdGvA/exec";
const POLL_MS = 10000;

const FALLBACK_DATA = {
  updatedAt: new Date().toISOString(),
  monthLabel: 'September 2026',
  runners: [
    {name:"Beam",level:"R3",deals:3,targetRevenue:90000,revenue:124500,achievement:138.33,finishPlace:1,finishDays:9},
    {name:"Gorn",level:"R[1]N",deals:3,targetRevenue:37000,revenue:52000,achievement:140.54,finishPlace:2,finishDays:13},
    {name:"Ping",level:"R[1]N",deals:2,targetRevenue:37000,revenue:31000,achievement:83.78},
    {name:"Kwan",level:"R[1]N",deals:1,targetRevenue:37000,revenue:30000,achievement:81.08},
    {name:"Nat",level:"R3",deals:5,targetRevenue:90000,revenue:65750,achievement:73.06},
    {name:"Pik",level:"R5",deals:3,targetRevenue:180000,revenue:129000,achievement:71.67},
    {name:"Baitong",level:"R[1]N",deals:2,targetRevenue:37000,revenue:24667,achievement:66.67},
    {name:"Sarah",level:"R3",deals:2,targetRevenue:90000,revenue:56519,achievement:62.80},
    {name:"Little",level:"R5",deals:2,targetRevenue:180000,revenue:108000,achievement:60.00},
    {name:"Pimdow",level:"R[1]N",deals:1,targetRevenue:37000,revenue:15000,achievement:40.54},
    {name:"Tle",level:"R4",deals:2,targetRevenue:125000,revenue:47000,achievement:37.60},
    {name:"Nick",level:"R4",deals:2,targetRevenue:125000,revenue:38000,achievement:30.40},
    {name:"Imim",level:"R2",deals:1,targetRevenue:65000,revenue:14250,achievement:21.92},
    {name:"Chalis",level:"R[1]N",deals:0,targetRevenue:37000,revenue:0,achievement:0}
  ]
};

const colors = ["#ffd65c","#9be8ff","#7ee29a","#a99cff","#ff8fbe","#ffb66e","#5bd6ce","#ff8181","#a8dd6e","#7cb7ff","#f9db79","#c5a4ff","#8ed7bc","#c7d0de","#ff9d76","#65e4ff","#f4a8ff","#a4f27a"];
const state = { previousAchievements: new Map(), lastGood: null, initialized: false };

function money(v){ return new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(v || 0); }
function pct(v){ return `${Number(v || 0).toFixed(2)}%`; }
function medal(rank){ return rank===1?'🥇':rank===2?'🥈':rank===3?'🥉':String(rank); }
function ordinal(rank){
  const n=Number(rank)||0;
  const mod100=n%100;
  if(mod100>=11 && mod100<=13) return `${n}th`;
  if(n%10===1) return `${n}st`;
  if(n%10===2) return `${n}nd`;
  if(n%10===3) return `${n}rd`;
  return `${n}th`;
}

function installStadiumTweaks(){
  if(document.querySelector('#stadiumTweaks')) return;
  const style=document.createElement('style');
  style.id='stadiumTweaks';
  style.textContent=`
    .stadium-runner{flex-direction:column-reverse !important;align-items:center !important;gap:2px !important;pointer-events:auto !important;cursor:default;}
    .stadium-nameplate{transform:translateY(-2px);transition:transform .16s ease,filter .16s ease,box-shadow .16s ease;}
    .stadium-runner:hover{z-index:999 !important;}
    .stadium-runner:hover .stadium-nameplate{transform:translateY(-2px) scale(1.05);filter:brightness(1.08);box-shadow:0 14px 34px rgba(0,0,0,.55),0 0 0 2px rgba(255,255,255,.45) !important;}
    .finish-line{width:60px !important;}

    .checkpoint{
      font-size:28px !important;
      font-weight:900 !important;
      color:rgba(235,239,245,.72) !important;
      background:rgba(96,103,116,.30) !important;
      border:1px solid rgba(255,255,255,.14) !important;
      padding:8px 13px !important;
      backdrop-filter:blur(3px);
      -webkit-backdrop-filter:blur(3px);
      box-shadow:none !important;
    }
    .checkpoint-25{left:auto !important;right:1% !important;top:50% !important;transform:translate(50%,-50%) !important;}
    .checkpoint-75{right:auto !important;left:1% !important;top:50% !important;transform:translate(-50%,-50%) !important;}
    .checkpoint-100{color:rgba(235,239,245,.72) !important;}

    .mini-race-board{
      max-height:430px;
      overflow-y:auto;
      overflow-x:hidden;
      padding-right:18px !important;
      margin-right:2px;
      scrollbar-gutter:stable;
    }
    .mini-race-board::-webkit-scrollbar{width:8px;}
    .mini-race-board::-webkit-scrollbar-track{background:rgba(255,255,255,.035);border-radius:999px;}
    .mini-race-board::-webkit-scrollbar-thumb{background:rgba(255,255,255,.20);border-radius:999px;}

    .podium-grid{
      display:grid;
      grid-template-columns:repeat(3,minmax(0,1fr));
      gap:14px;
      margin-bottom:16px;
      align-items:stretch;
    }
    .podium-grid .podium-card{margin-bottom:0 !important;min-width:0;}
    .sales-podium-card .podium,
    .deal-podium-card .podium{gap:10px !important;padding-left:8px !important;padding-right:8px !important;}
    .sales-podium-card .podium-slot,
    .deal-podium-card .podium-slot{width:min(170px,31%);}
    .sales-podium-card .podium-medal,
    .deal-podium-card .podium-medal{font-size:54px !important;}
    .sales-podium-card .podium-name,
    .deal-podium-card .podium-name{font-size:30px !important;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
    .sales-podium-card .podium-slot.first .podium-name,
    .deal-podium-card .podium-slot.first .podium-name{font-size:34px !important;}
    .sales-podium-card .podium-days,
    .deal-podium-card .podium-days{font-size:18px !important;white-space:nowrap;}
    .sales-podium-card .podium-slot.winner .podium-person,
    .deal-podium-card .podium-slot.winner .podium-person{animation:none !important;}
    .sales-podium-card .podium-slot.winner::before,
    .sales-podium-card .podium-slot.winner::after,
    .deal-podium-card .podium-slot.winner::before,
    .deal-podium-card .podium-slot.winner::after{display:none !important;content:none !important;}

    @media (max-width:1450px){
      .podium-grid{grid-template-columns:minmax(0,1fr) minmax(0,1fr);}
      .podium-grid .deal-podium-card{grid-column:1 / -1;}
      .sales-podium-card .podium-slot,
      .deal-podium-card .podium-slot{width:min(190px,30%);}
      .sales-podium-card .podium-medal,
      .deal-podium-card .podium-medal{font-size:62px !important;}
      .sales-podium-card .podium-name,
      .deal-podium-card .podium-name{font-size:36px !important;}
      .sales-podium-card .podium-slot.first .podium-name,
      .deal-podium-card .podium-slot.first .podium-name{font-size:40px !important;}
      .sales-podium-card .podium-days,
      .deal-podium-card .podium-days{font-size:19px !important;}
    }

    @media (max-width:1100px){
      .podium-grid{grid-template-columns:1fr;}
      .podium-grid .deal-podium-card{grid-column:auto;}
      .sales-podium-card .podium-slot,
      .deal-podium-card .podium-slot{width:min(210px,27vw);}
      .sales-podium-card .podium-medal,
      .deal-podium-card .podium-medal{font-size:72px !important;}
      .sales-podium-card .podium-name,
      .deal-podium-card .podium-name{font-size:42px !important;}
      .sales-podium-card .podium-slot.first .podium-name,
      .deal-podium-card .podium-slot.first .podium-name{font-size:46px !important;}
      .sales-podium-card .podium-days,
      .deal-podium-card .podium-days{font-size:23px !important;}
    }

    @media (max-width:900px){
      .checkpoint{font-size:18px !important;padding:6px 9px !important;}
      .sales-podium-card .podium-medal,
      .deal-podium-card .podium-medal{font-size:48px !important;}
      .sales-podium-card .podium-name,
      .sales-podium-card .podium-slot.first .podium-name,
      .deal-podium-card .podium-name,
      .deal-podium-card .podium-slot.first .podium-name{font-size:24px !important;}
      .sales-podium-card .podium-days,
      .deal-podium-card .podium-days{font-size:14px !important;}
    }
  `;
  document.head.appendChild(style);
}

function installSalesPodium(){
  if(document.querySelector('#salesPodium')) return;
  const raceCard=document.querySelector('.podium-card');
  if(!raceCard || !raceCard.parentNode) return;

  const grid=document.createElement('section');
  grid.className='podium-grid';
  raceCard.parentNode.insertBefore(grid,raceCard);
  grid.appendChild(raceCard);

  const salesCard=document.createElement('section');
  salesCard.className='podium-card sales-podium-card';
  salesCard.innerHTML=`
    <div class="section-heading">
      <div>
        <div class="section-title">💰 Top Sales Podium</div>
        <div class="section-note">Current top 3 by revenue · changes live</div>
      </div>
    </div>
    <div id="salesPodium" class="podium"></div>`;
  grid.appendChild(salesCard);

  const dealCard=document.createElement('section');
  dealCard.className='podium-card deal-podium-card';
  dealCard.innerHTML=`
    <div class="section-heading">
      <div>
        <div class="section-title">🤝 Top Deals Podium</div>
        <div class="section-note">Current top 3 by deals closed · changes live</div>
      </div>
    </div>
    <div id="dealPodium" class="podium"></div>`;
  grid.appendChild(dealCard);
}

function normalize(payload){
  const runners = (payload.runners || []).map(r => ({
    ...r,
    achievement:Number(r.achievement)||0,
    finishPlace:Number(r.finishPlace)||0,
    finishDays:Number(r.finishDays)||0,
    revenue:Number(r.revenue)||0,
    deals:Number(r.deals)||0,
    targetRevenue:Number(r.targetRevenue)||0
  }));

  runners.sort((a,b)=>{
    if(a.finishPlace && b.finishPlace) return a.finishPlace-b.finishPlace;
    if(a.finishPlace) return -1;
    if(b.finishPlace) return 1;
    return b.achievement-a.achievement || b.revenue-a.revenue;
  });
  runners.forEach((r,i)=>r.rank=i+1);
  return {...payload,runners};
}

function renderKpis(data){
  const rs=data.runners;
  const totalRevenue=rs.reduce((s,r)=>s+r.revenue,0);
  const totalTarget=rs.reduce((s,r)=>s+r.targetRevenue,0);
  const totalDeals=rs.reduce((s,r)=>s+r.deals,0);
  const avg=rs.length?rs.reduce((s,r)=>s+r.achievement,0)/rs.length:0;
  document.querySelector('#kpiGrid').innerHTML = [
    ['Team Revenue',`฿${money(totalRevenue)}`],
    ['Total Target',`฿${money(totalTarget)}`],
    ['Total Deals',money(totalDeals)],
    ['Avg Achievement',pct(avg)]
  ].map(([label,value])=>`<div class="kpi"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div></div>`).join('');
}

function renderPodium(data){
  const podium=document.querySelector('#podium');
  if(!podium) return;
  const finishers=[...data.runners].filter(r=>r.finishPlace>0).sort((a,b)=>a.finishPlace-b.finishPlace);
  const byPlace=new Map(finishers.map(r=>[r.finishPlace,r]));
  const order=[2,1,3];

  podium.innerHTML=order.map(place=>{
    const r=byPlace.get(place);
    const cls=place===1?'first':place===2?'second':'third';
    const name=r?escapeHtml(r.name):'—';
    const days=r&&r.finishDays?`${r.finishDays} day${r.finishDays===1?'':'s'} to finish`:'Waiting for finisher';
    return `<div class="podium-slot ${cls}">
      <div class="podium-person">
        <div class="podium-medal">${medal(place)}</div>
        <div class="podium-name">${name}</div>
        <div class="podium-days">${days}</div>
      </div>
      <div class="podium-block"><span>${place}</span><small>${ordinal(place)} PLACE</small></div>
    </div>`;
  }).join('');
}

function renderSalesPodium(data){
  const podium=document.querySelector('#salesPodium');
  if(!podium) return;
  const ranked=[...data.runners]
    .sort((a,b)=>b.revenue-a.revenue || b.deals-a.deals || a.name.localeCompare(b.name))
    .slice(0,3);
  const byPlace=new Map(ranked.map((r,i)=>[i+1,r]));
  const order=[2,1,3];

  podium.innerHTML=order.map(place=>{
    const r=byPlace.get(place);
    const cls=place===1?'first':place===2?'second':'third';
    const name=r?escapeHtml(r.name):'—';
    const detail=r?`฿${money(r.revenue)}`:'Waiting for sales';
    return `<div class="podium-slot ${cls}">
      <div class="podium-person">
        <div class="podium-medal">${medal(place)}</div>
        <div class="podium-name">${name}</div>
        <div class="podium-days">${detail}</div>
      </div>
      <div class="podium-block"><span>${place}</span><small>${ordinal(place)} PLACE</small></div>
    </div>`;
  }).join('');
}

function renderDealPodium(data){
  const podium=document.querySelector('#dealPodium');
  if(!podium) return;
  const ranked=[...data.runners]
    .sort((a,b)=>b.deals-a.deals || b.revenue-a.revenue || a.name.localeCompare(b.name))
    .slice(0,3);
  const byPlace=new Map(ranked.map((r,i)=>[i+1,r]));
  const order=[2,1,3];

  podium.innerHTML=order.map(place=>{
    const r=byPlace.get(place);
    const cls=place===1?'first':place===2?'second':'third';
    const name=r?escapeHtml(r.name):'—';
    const detail=r?`${money(r.deals)} deal${r.deals===1?'':'s'}`:'Waiting for deals';
    return `<div class="podium-slot ${cls}">
      <div class="podium-person">
        <div class="podium-medal">${medal(place)}</div>
        <div class="podium-name">${name}</div>
        <div class="podium-days">${detail}</div>
      </div>
      <div class="podium-block"><span>${place}</span><small>${ordinal(place)} PLACE</small></div>
    </div>`;
  }).join('');
}

function stadiumPosition(progress){
  const p=Math.max(0,Math.min(100,Number(progress)||0))/100;
  const xLeft=20;
  const xRight=80;
  const yTop=10;
  const yBottom=90;
  const centerY=50;
  const rx=20;
  const ry=40;
  const aspect=1.92;
  const straight=(xRight-xLeft)/100*aspect;
  const radius=ry/100;
  const semicircle=Math.PI*radius;
  const total=2*straight+2*semicircle;
  let distance=p*total;
  const halfStraight=straight/2;

  if(distance<=halfStraight){
    const t=distance/halfStraight;
    return {x:50+(xRight-50)*t,y:yBottom};
  }
  distance-=halfStraight;

  if(distance<=semicircle){
    const t=distance/semicircle;
    const theta=(90-180*t)*Math.PI/180;
    return {x:xRight+rx*Math.cos(theta),y:centerY+ry*Math.sin(theta)};
  }
  distance-=semicircle;

  if(distance<=straight){
    const t=distance/straight;
    return {x:xRight+(xLeft-xRight)*t,y:yTop};
  }
  distance-=straight;

  if(distance<=semicircle){
    const t=distance/semicircle;
    const theta=(270-180*t)*Math.PI/180;
    return {x:xLeft+rx*Math.cos(theta),y:centerY+ry*Math.sin(theta)};
  }
  distance-=semicircle;

  const t=Math.min(1,distance/halfStraight);
  return {x:xLeft+(50-xLeft)*t,y:yBottom};
}

function renderStadium(data){
  const container=document.querySelector('#stadiumRunners');
  if(!container) return;
  const finishers=data.runners.filter(r=>r.finishPlace>0);

  container.innerHTML=data.runners.map((r,i)=>{
    const color=colors[i%colors.length];
    const capped=Math.min(100,r.achievement);
    const pos=stadiumPosition(capped);
    const finished=Boolean(r.finishPlace);
    const finishIndex=finished ? finishers.findIndex(f=>f.name===r.name) : -1;
    const finishOffset=finished ? ((finishIndex%5)-2)*24 : ((i%3)-1)*7;
    const label=finished ? `${ordinal(r.finishPlace)} · ${escapeHtml(r.name)}` : escapeHtml(r.name);
    const status=finished ? `${r.finishDays || ''}${r.finishDays?'d':''}` : `${Math.round(capped)}%`;
    return `<div class="stadium-runner ${finished?'finished':''}" style="left:${pos.x}%;top:${pos.y}%;--runner:${color};--offset:${finishOffset}px" title="${escapeHtml(r.name)} · ${pct(r.achievement)}">
      <div class="stadium-runner-icon">🏃</div>
      <div class="stadium-nameplate"><strong>${label}</strong><span>${status}</span></div>
    </div>`;
  }).join('');
}

function renderMiniRace(data){
  const board=document.querySelector('#miniRaceBoard');
  if(!board) return;
  board.innerHTML=data.runners.map((r,i)=>{
    const color=colors[i%colors.length];
    const capped=Math.max(0,Math.min(100,r.achievement));
    const result=r.finishPlace ? ordinal(r.finishPlace) : `${Math.round(capped)}%`;
    return `<div class="mini-runner" style="--runner:${color}">
      <div class="mini-runner-top"><span class="mini-name">${escapeHtml(r.name)}</span><span class="mini-result">${result}</span></div>
      <div class="mini-track"><div class="mini-progress" style="width:${capped}%"></div></div>
    </div>`;
  }).join('');
}

function renderFinishHistory(data){
  const box=document.querySelector('#finishHistory');
  if(!box) return;
  const finishers=[...data.runners].filter(r=>r.finishPlace>0).sort((a,b)=>a.finishPlace-b.finishPlace);
  if(!finishers.length){
    box.innerHTML='<div class="finish-empty">No finishers yet</div>';
    return;
  }
  box.innerHTML=finishers.map(r=>`<div class="finish-history-row">
    <span class="finish-place">${ordinal(r.finishPlace)}</span>
    <strong>${escapeHtml(r.name)}</strong>
    <span>${r.finishDays?`${r.finishDays}d`:''}</span>
  </div>`).join('');
}

function renderSalesBoard(data){
  const tbody=document.querySelector('#salesBoard');
  if(!tbody) return;
  const ranked=[...data.runners].sort((a,b)=>b.revenue-a.revenue || b.deals-a.deals || a.name.localeCompare(b.name));
  tbody.innerHTML=ranked.map((r,i)=>{
    const rank=i+1;
    const badge=rank<=3?medal(rank):String(rank);
    const topClass=rank<=3?` top-${rank}`:'';
    return `<tr class="sales-row${topClass}">
      <td class="sales-rank">${badge}</td>
      <td><div class="sales-name">${escapeHtml(r.name)}</div></td>
      <td><span class="level-pill">${escapeHtml(r.level || '-')}</span></td>
      <td class="sales-revenue">฿${money(r.revenue)}</td>
      <td class="sales-deals">${money(r.deals)}</td>
      <td><span class="target-pill ${r.achievement>=100?'hit':''}">${pct(r.achievement)}</span></td>
    </tr>`;
  }).join('');
}

function render(data){
  data=normalize(data); state.lastGood=data;
  renderKpis(data);
  renderPodium(data);
  renderSalesPodium(data);
  renderDealPodium(data);
  renderStadium(data);
  renderMiniRace(data);
  renderFinishHistory(data);
  renderSalesBoard(data);

  const d=new Date(data.updatedAt || Date.now());
  document.querySelector('#lastUpdated').textContent=`Last updated: ${d.toLocaleTimeString('en-GB',{hour12:false})}`;
  const monthLabel = data.monthLabel || data.sourceTab || 'Current Month';
  const eyebrow = document.querySelector('#monthEyebrow');
  if (eyebrow) eyebrow.textContent = `${String(monthLabel).toUpperCase()} · FIRST TO 100% WINS`;

  data.runners.forEach(r=>{
    const prev=state.previousAchievements.get(r.name);
    if(state.initialized && prev != null && prev < 100 && r.achievement >= 100) launchConfetti();
    state.previousAchievements.set(r.name,r.achievement);
  });
  state.initialized=true;
}

async function fetchData(){
  const warning=document.querySelector('#warning');
  if(!API_URL){ render(FALLBACK_DATA); warning.classList.remove('hidden'); warning.textContent='Demo mode · add Apps Script URL for live data'; return; }
  try{
    const url = `${API_URL}${API_URL.includes('?')?'&':'?'}t=${Date.now()}`;
    const res=await fetch(url,{cache:'no-store'}); if(!res.ok) throw new Error(`HTTP ${res.status}`);
    const data=await res.json(); render(data); warning.classList.add('hidden');
  }catch(err){ console.error(err); warning.classList.remove('hidden'); warning.textContent='Live data temporarily unavailable'; if(!state.lastGood) render(FALLBACK_DATA); }
}

function escapeHtml(s){ return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c])); }
function launchConfetti(){ const layer=document.querySelector('#confettiLayer'); const palette=['#ffd65c','#62e89a','#60cfff','#ff78b8','#b89bff']; for(let i=0;i<70;i++){ const e=document.createElement('i'); e.className='confetti'; e.style.left=`${Math.random()*100}%`; e.style.background=palette[i%palette.length]; e.style.setProperty('--drift',`${(Math.random()-.5)*260}px`); e.style.animationDelay=`${Math.random()*.45}s`; layer.appendChild(e); setTimeout(()=>e.remove(),2600); } }

document.querySelector('#fullscreenBtn').addEventListener('click',()=>{ if(!document.fullscreenElement) document.documentElement.requestFullscreen?.(); else document.exitFullscreen?.(); });
installStadiumTweaks();
installSalesPodium();
fetchData(); setInterval(fetchData,POLL_MS);
