const API_URL="https://script.google.com/macros/s/AKfycby_A9iKIrfCa22aDZLXlKk7CWz4hsiGOx5ufFREVPRWQ8ds_xfi772D-SDfT6QidYdGvA/exec";
const $=s=>document.querySelector(s);
const money=v=>`฿${new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(Number(v||0))}`;
const pct=v=>v===null||v===undefined||v===''?'—':`${Number(v||0).toFixed(2)}%`;
const tokenKey='pcDashboardToken';
const cacheKey='pcDashboardCache';

async function post(params){
  const res=await fetch(API_URL,{method:'POST',body:new URLSearchParams(params)});
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const data=await res.json();
  if(!data.ok) throw new Error(data.error||'Request failed');
  return data;
}

function getToken(){
  let token=localStorage.getItem(tokenKey);
  if(!token){
    token=sessionStorage.getItem(tokenKey);
    if(token){localStorage.setItem(tokenKey,token);sessionStorage.removeItem(tokenKey);}
  }
  return token;
}
function setToken(token){localStorage.setItem(tokenKey,token);sessionStorage.removeItem(tokenKey);}
function clearToken(){localStorage.removeItem(tokenKey);sessionStorage.removeItem(tokenKey);localStorage.removeItem(cacheKey);}
function saveCache(data){try{localStorage.setItem(cacheKey,JSON.stringify(data));}catch(e){}}
function getCache(){try{return JSON.parse(localStorage.getItem(cacheKey)||'null');}catch(e){return null;}}
function isSessionError(err){return /session expired|invalid session|unauthori[sz]ed/i.test(String(err&&err.message||err||''));}

function setLoginMode(loggedIn){
  $('#loginCard').classList.toggle('hidden',loggedIn);
  $('#dashboard').classList.toggle('hidden',!loggedIn);
  $('#logoutBtn').classList.toggle('hidden',!loggedIn);
}

function statusClass(code){
  if(['ELIGIBLE','SECURE','FAST_TRACK'].includes(code)) return 'good';
  if(['RISK','BELOW_RETENTION'].includes(code)) return 'risk';
  return 'warn';
}

function completedHistory(data){
  const currentLabel=String(data.currentMonthLabel||'').trim().toLowerCase();
  return (data.history||[])
    .filter(r=>String(r.monthLabel||'').trim().toLowerCase()!==currentLabel)
    .slice(0,3);
}

function avgCompletedKpi(rows){
  const vals=rows
    .filter(r=>r.conversionAvailable && r.kpi!==null && r.kpi!==undefined && r.kpi!=='')
    .map(r=>Number(r.kpi))
    .filter(Number.isFinite);
  return vals.length===3 ? vals.reduce((a,b)=>a+b,0)/3 : null;
}

function levelRank(level){
  const s=String(level||'').trim().toUpperCase().replace(/\s+/g,'');
  if(s==='RN'||s==='R1'||s==='R[1]N') return 1;
  return ({R2:2,R3:3,R4:4,R5:5})[s]||0;
}

function monthlyLevelResult(row,index,history,current){
  const next=index===0?current:history[index-1];
  if(!next) return {label:'—',cls:''};
  const from=levelRank(row.level), to=levelRank(next.level);
  if(!from||!to) return {label:'—',cls:''};
  if(to>from) return {label:'↑ Level Up',cls:'up'};
  if(to<from) return {label:'↓ Level Down',cls:'down'};
  return {label:'Stay',cls:'stay'};
}

function render(data){
  const c=data.current;
  const history=completedHistory(data);
  const completedAvg=avgCompletedKpi(history);

  $('#welcomeName').textContent=`Welcome, ${data.name}`;
  $('#monthLabel').textContent=data.currentMonthLabel||'CURRENT MONTH';
  $('#currentLevel').textContent=c.level||'—';
  $('#revenue').textContent=money(c.revenue);
  $('#revenueTarget').textContent=`Target ${money(c.targetRevenue)}`;
  $('#deals').textContent=new Intl.NumberFormat('en-US',{maximumFractionDigits:2}).format(c.deals||0);
  $('#dealsTarget').textContent=`Target ${new Intl.NumberFormat('en-US',{maximumFractionDigits:2}).format(c.targetDeals||0)}`;
  $('#conversion').textContent=c.conversionAvailable?pct(c.conversion):'Pending';
  $('#kpi').textContent=c.conversionAvailable && c.kpiAvailable!==false?pct(c.kpi):'Pending';

  $('#historyBody').innerHTML=history.map((r,i)=>{
    const result=monthlyLevelResult(r,i,history,c);
    return `<tr>
      <td>${r.monthLabel}</td><td>${r.level||'—'}</td><td><span class="history-result ${result.cls}">${result.label}</span></td><td>${money(r.revenue)}</td><td>${Number(r.deals||0).toFixed(0)}</td>
      <td>${r.conversionAvailable?pct(r.conversion):'Pending'}</td><td>${r.conversionAvailable && r.kpiAvailable!==false?pct(r.kpi):'Pending'}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="7">No completed history yet.</td></tr>';
  $('#avgKpi').textContent=`3M Avg ${completedAvg==null?'—':pct(completedAvg)}`;

  const s=data.levelStatus||{};
  $('#statusBadge').textContent=s.badge||'ESTIMATE';
  $('#statusBadge').className=`status-badge ${statusClass(s.code)}`;
  $('#levelHeadline').textContent=s.headline||'Month in progress';
  $('#levelMessage').textContent=s.message||'Final status will be confirmed after month-end conversion is entered.';

  let facts=(s.facts||[]).filter(f=>!/^2M Avg KPI$/i.test(f.label||'') && !/^3M Avg KPI$/i.test(f.label||''));
  facts=facts.filter(f=>!/^Minimum rev before level down$/i.test(f.label||''));

  let retentionValue='Waiting for 2 completed months';
  const retain=Number(s.retentionMinimum);
  const prev=history[0], prev2=history[1];
  if(Number.isFinite(retain) && prev && prev2 && prev.conversionAvailable && prev2.conversionAvailable){
    const k1=Number(prev.kpi), k2=Number(prev2.kpi);
    if(Number.isFinite(k1) && Number.isFinite(k2)){
      const need=Math.max(0,retain*3-k1-k2);
      const rev=Number(c.targetRevenue||0)*need/100;
      retentionValue=rev<=0?'Already covered':`≈ ${money(rev)} this month`;
    }
  } else if(s.retentionRevenueNeeded!==undefined && s.retentionRevenueNeeded!==null){
    retentionValue=Number(s.retentionRevenueNeeded)<=0?'Already covered':`≈ ${money(s.retentionRevenueNeeded)} this month`;
  }
  facts.push({label:'Minimum rev before level down',value:retentionValue});
  $('#levelFacts').innerHTML=facts.map(f=>`<div class="fact"><span>${f.label}</span><strong>${f.value}</strong></div>`).join('');

  if(s.nextLevel){
    $('#nextLevelTitle').textContent=`Estimated target to reach ${s.nextLevel}`;
    $('#estimatedRevenueTarget').textContent=s.estimatedRevenueTarget==null?'—':money(s.estimatedRevenueTarget);
    if(s.revenueGap==null){
      $('#revenueGap').textContent='Waiting for enough completed history';
      $('#revenueGap').className='gap-line';
    } else if(s.revenueGap<=0){
      $('#revenueGap').textContent='Revenue estimate already reached';
      $('#revenueGap').className='gap-line positive';
    } else {
      $('#revenueGap').textContent=`${money(s.revenueGap)} more needed`;
      $('#revenueGap').className='gap-line negative';
    }
    $('#estimateNote').textContent='Revenue equivalent only. Final level eligibility is confirmed after month-end conversion is entered.';
  } else {
    $('#nextLevelTitle').textContent='Current level';
    $('#estimatedRevenueTarget').textContent='Top level';
    $('#revenueGap').textContent='No higher level configured';
    $('#revenueGap').className='gap-line positive';
    $('#estimateNote').textContent='Retention is based on the last 3 completed months.';
  }

  setLoginMode(true);
}

function showRefreshProblem(){
  setLoginMode(true);
  if($('#statusBadge')){
    $('#statusBadge').textContent='RETRYING';
    $('#statusBadge').className='status-badge warn';
  }
}

async function loadDashboard({silent=false}={}){
  const token=getToken();
  if(!token){setLoginMode(false);return;}

  const cached=getCache();
  if(cached && !silent){
    render(cached);
  } else if(!silent){
    setLoginMode(true);
    $('#statusBadge').textContent='LOADING';
    $('#statusBadge').className='status-badge warn';
  }

  try{
    const data=await post({action:'dashboard',token});
    saveCache(data);
    render(data);
  }catch(e){
    if(isSessionError(e)){
      clearToken();
      setLoginMode(false);
      $('#loginError').textContent='Session expired. Please log in again.';
      $('#loginError').classList.remove('hidden');
      return;
    }
    console.warn('Dashboard refresh failed:',e);
    if(cached) render(cached);
    showRefreshProblem();
    setTimeout(()=>loadDashboard({silent:true}),5000);
  }
}

$('#loginForm').addEventListener('submit',async e=>{
  e.preventDefault();
  $('#loginError').classList.add('hidden');
  const btn=e.currentTarget.querySelector('button');
  btn.disabled=true;btn.textContent='Checking…';
  try{
    const data=await post({action:'login',username:$('#username').value.trim(),pin:$('#pin').value.trim()});
    setToken(data.token);
    setLoginMode(true);
    $('#statusBadge').textContent='LOADING';
    $('#statusBadge').className='status-badge warn';
    btn.disabled=false;btn.textContent='Open dashboard';
    loadDashboard({silent:true});
    return;
  }catch(err){
    $('#loginError').textContent=err.message==='Invalid login'?'Username or PIN is incorrect.':err.message;
    $('#loginError').classList.remove('hidden');
  }finally{
    btn.disabled=false;btn.textContent='Open dashboard';
  }
});

$('#logoutBtn').addEventListener('click',async()=>{
  const token=getToken();
  clearToken();
  if(token){try{await post({action:'logout',token});}catch(e){}}
  setLoginMode(false);
});

loadDashboard();
