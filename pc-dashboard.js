const API_URL="https://script.google.com/macros/s/AKfycby_A9iKIrfCa22aDZLXlKk7CWz4hsiGOx5ufFREVPRWQ8ds_xfi772D-SDfT6QidYdGvA/exec";
const $=s=>document.querySelector(s);
const money=v=>`฿${new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(Number(v||0))}`;
const pct=v=>v===null||v===undefined||v===''?'—':`${Number(v||0).toFixed(2)}%`;
const tokenKey='pcDashboardToken';

async function post(params){
  const res=await fetch(API_URL,{method:'POST',body:new URLSearchParams(params)});
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const data=await res.json();
  if(!data.ok) throw new Error(data.error||'Request failed');
  return data;
}

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

function render(data){
  const c=data.current;
  $('#welcomeName').textContent=`Welcome, ${data.name}`;
  $('#monthLabel').textContent=data.currentMonthLabel||'CURRENT MONTH';
  $('#currentLevel').textContent=c.level||'—';
  $('#revenue').textContent=money(c.revenue);
  $('#revenueTarget').textContent=`Target ${money(c.targetRevenue)}`;
  $('#deals').textContent=new Intl.NumberFormat('en-US',{maximumFractionDigits:2}).format(c.deals||0);
  $('#dealsTarget').textContent=`Target ${new Intl.NumberFormat('en-US',{maximumFractionDigits:2}).format(c.targetDeals||0)}`;
  $('#conversion').textContent=c.conversionAvailable?pct(c.conversion):'Pending';
  $('#kpi').textContent=c.kpiAvailable?pct(c.kpi):'Pending';

  $('#historyBody').innerHTML=(data.history||[]).map(r=>`<tr>
    <td>${r.monthLabel}</td><td>${r.level||'—'}</td><td>${money(r.revenue)}</td><td>${Number(r.deals||0).toFixed(0)}</td>
    <td>${r.conversionAvailable?pct(r.conversion):'Pending'}</td><td>${r.kpiAvailable?pct(r.kpi):'Pending'}</td>
  </tr>`).join('')||'<tr><td colspan="6">No completed history yet.</td></tr>';
  $('#avgKpi').textContent=`3M Avg ${data.threeMonthAvgKpi==null?'—':pct(data.threeMonthAvgKpi)}`;

  const s=data.levelStatus;
  $('#statusBadge').textContent=s.badge;
  $('#statusBadge').className=`status-badge ${statusClass(s.code)}`;
  $('#levelHeadline').textContent=s.headline;
  $('#levelMessage').textContent=s.message;
  $('#levelFacts').innerHTML=(s.facts||[]).map(f=>`<div class="fact"><span>${f.label}</span><strong>${f.value}</strong></div>`).join('');

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

async function loadDashboard(){
  const token=sessionStorage.getItem(tokenKey);
  if(!token){setLoginMode(false);return;}
  try{
    const data=await post({action:'dashboard',token});
    render(data);
  }catch(e){
    sessionStorage.removeItem(tokenKey);
    setLoginMode(false);
    $('#loginError').textContent='Session expired. Please log in again.';
    $('#loginError').classList.remove('hidden');
  }
}

$('#loginForm').addEventListener('submit',async e=>{
  e.preventDefault();
  $('#loginError').classList.add('hidden');
  const btn=e.currentTarget.querySelector('button');
  btn.disabled=true;btn.textContent='Checking…';
  try{
    const data=await post({action:'login',username:$('#username').value.trim(),pin:$('#pin').value.trim()});
    sessionStorage.setItem(tokenKey,data.token);
    await loadDashboard();
  }catch(err){
    $('#loginError').textContent=err.message==='Invalid login'?'Username or PIN is incorrect.':err.message;
    $('#loginError').classList.remove('hidden');
  }finally{btn.disabled=false;btn.textContent='Open dashboard';}
});

$('#logoutBtn').addEventListener('click',async()=>{
  const token=sessionStorage.getItem(tokenKey);
  sessionStorage.removeItem(tokenKey);
  if(token){try{await post({action:'logout',token});}catch(e){}}
  setLoginMode(false);
});

loadDashboard();
