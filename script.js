// Paste the deployed Google Apps Script Web App URL here.
const API_URL = "https://script.google.com/macros/s/AKfycby_A9iKIrfCa22aDZLXlKk7CWz4hsiGOx5ufFREVPRWQ8ds_xfi772D-SDfT6QidYdGvA/exec";
const POLL_MS = 10000;

const FALLBACK_DATA = {
  updatedAt: new Date().toISOString(),
  monthLabel: 'September 2026',
  runners: [
    {name:"Gorn",level:"R[1]N",deals:3,targetRevenue:37000,revenue:52000,achievement:140.54},
    {name:"Beam",level:"R3",deals:3,targetRevenue:90000,revenue:124500,achievement:138.33},
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

const colors = ["#ffd65c","#9be8ff","#7ee29a","#a99cff","#ff8fbe","#ffb66e","#5bd6ce","#ff8181","#a8dd6e","#7cb7ff","#f9db79","#c5a4ff","#8ed7bc","#c7d0de"];
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

function finishBadge(r){
  if(!r.finishPlace) return '';
  const place = r.finishPlace<=3 ? medal(r.finishPlace) : '✅';
  const days = r.finishDays ? ` · ${r.finishDays} day${r.finishDays===1?'':'s'}` : '';
  return `<span class="finish-badge">${place} FINISHED${days}</span>`;
}

function renderRace(data){
  const board=document.querySelector('#raceBoard');
  const oldRects=new Map([...board.children].map(el=>[el.dataset.name,el.getBoundingClientRect()]));

  board.innerHTML=data.runners.map((r,i)=>{
    const color=colors[i%colors.length];
    const progress=Math.max(0,Math.min(100,r.achievement));
    const finished=Boolean(r.finishPlace);
    const rankDisplay = finished ? medal(r.finishPlace) : String(r.rank);
    const meta=`${r.level} · ฿${money(r.revenue)} / ฿${money(r.targetRevenue)}`;
    const marker = finished
      ? `<div class="place-banner">${ordinal(r.finishPlace)} PLACE</div>`
      : `<div class="runner-marker" style="left:${progress}%"><span class="runner-icon">🏃</span><span class="nameplate">${escapeHtml(r.name)}</span></div>`;
    const racePct = Math.min(100, r.achievement);
    return `<div class="runner-row" data-name="${escapeHtml(r.name)}" style="--runner:${color}">
      <div class="rank ${finished && r.finishPlace<=3?'medal':''}">${rankDisplay}</div>
      <div class="identity"><div class="name">${escapeHtml(r.name)}</div><div class="meta">${escapeHtml(meta)}</div>${finishBadge(r)}</div>
      <div class="track ${finished?'finished-track':''}"><div class="progress" style="width:${progress}%"></div>${marker}</div>
      <div class="achievement">${pct(racePct)}</div>
    </div>`;
  }).join('');

  [...board.children].forEach(el=>{
    const old=oldRects.get(el.dataset.name); if(!old) return;
    const now=el.getBoundingClientRect(); const dy=old.top-now.top;
    if(Math.abs(dy)>1){ el.style.transition='none'; el.style.transform=`translateY(${dy}px)`; requestAnimationFrame(()=>{requestAnimationFrame(()=>{el.style.transition='transform .65s cubic-bezier(.2,.8,.2,1), opacity .3s'; el.style.transform='';});}); }
  });

  data.runners.forEach(r=>{
    const prev=state.previousAchievements.get(r.name);
    if(state.initialized && prev != null && prev < 100 && r.achievement >= 100) launchConfetti();
    state.previousAchievements.set(r.name,r.achievement);
  });
  state.initialized=true;
}

function renderSalesBoard(data){
  const tbody=document.querySelector('#salesBoard');
  if(!tbody) return;

  const ranked=[...data.runners]
    .sort((a,b)=>b.revenue-a.revenue || b.deals-a.deals || a.name.localeCompare(b.name));

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
  renderKpis(data); renderRace(data); renderSalesBoard(data);
  const d=new Date(data.updatedAt || Date.now());
  document.querySelector('#lastUpdated').textContent=`Last updated: ${d.toLocaleTimeString('en-GB',{hour12:false})}`;
  const monthLabel = data.monthLabel || data.sourceTab || 'Current Month';
  const eyebrow = document.querySelector('#monthEyebrow');
  if (eyebrow) eyebrow.textContent = `${String(monthLabel).toUpperCase()} · FIRST TO 100% WINS`;
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

function escapeHtml(s){ return String(s??'').replace(/[&<>'\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#039;','\"':'&quot;'}[c])); }
function launchConfetti(){ const layer=document.querySelector('#confettiLayer'); const palette=['#ffd65c','#62e89a','#60cfff','#ff78b8','#b89bff']; for(let i=0;i<70;i++){ const e=document.createElement('i'); e.className='confetti'; e.style.left=`${Math.random()*100}%`; e.style.background=palette[i%palette.length]; e.style.setProperty('--drift',`${(Math.random()-.5)*260}px`); e.style.animationDelay=`${Math.random()*.45}s`; layer.appendChild(e); setTimeout(()=>e.remove(),2600); } }

document.querySelector('#fullscreenBtn').addEventListener('click',()=>{ if(!document.fullscreenElement) document.documentElement.requestFullscreen?.(); else document.exitFullscreen?.(); });
fetchData(); setInterval(fetchData,POLL_MS);
