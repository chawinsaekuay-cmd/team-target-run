/**
 * Add this file to the SAME Apps Script project as google-apps-script.gs,
 * then deploy a new Web App version. It adds secure nickname + PIN login
 * and a private rolling performance endpoint without changing the public leaderboard API.
 */

const PC_LOGIN_SHEET_ = 'PC_Login';
const PC_SESSION_SECONDS_ = 21600; // 6 hours
const PC_CONVERSION_TARGET_ = 12;

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || '').trim().toLowerCase();
    if (action === 'login') return pcJson_(pcLogin_(p.username, p.pin));
    if (action === 'dashboard') return pcJson_(pcDashboard_(p.token));
    if (action === 'logout') return pcJson_(pcLogout_(p.token));
    return pcJson_({ ok:false, error:'Unknown action' });
  } catch (err) {
    return pcJson_({ ok:false, error:String(err && err.message || err) });
  }
}

function pcLogin_(username, pin) {
  const user = String(username || '').trim();
  const pass = String(pin || '').trim();
  if (!user || !pass) return { ok:false, error:'Invalid login' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PC_LOGIN_SHEET_);
  if (!sheet || sheet.getLastRow() < 2) return { ok:false, error:'Login is not configured' };

  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
  const match = rows.find(r => String(r[0] || '').trim().toLowerCase() === user.toLowerCase());
  if (!match || match[4] === false || String(match[4]).toLowerCase() === 'false') return { ok:false, error:'Invalid login' };

  const storedHash = String(match[3] || '').trim().toLowerCase();
  const candidateHash = pcHashPin_(user, pass, String(match[2] || ''));
  if (!storedHash || candidateHash !== storedHash) return { ok:false, error:'Invalid login' };

  const staffCode = String(match[1] || '').trim();
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put('pcsession:' + token, JSON.stringify({
    username:user,
    staffCode:staffCode,
    issuedAt:new Date().toISOString()
  }), PC_SESSION_SECONDS_);
  return { ok:true, token:token, username:user };
}

function pcLogout_(token) {
  const t = String(token || '').trim();
  if (t) CacheService.getScriptCache().remove('pcsession:' + t);
  return { ok:true };
}

function pcSession_(token) {
  const t = String(token || '').trim();
  if (!t) throw new Error('Session expired');
  const raw = CacheService.getScriptCache().get('pcsession:' + t);
  if (!raw) throw new Error('Session expired');
  CacheService.getScriptCache().put('pcsession:' + t, raw, PC_SESSION_SECONDS_);
  return JSON.parse(raw);
}

function pcHashPin_(username, pin, salt) {
  const text = String(username || '').trim().toLowerCase() + ':' + String(pin || '') + ':' + String(salt || '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function pcDashboard_(token) {
  const session = pcSession_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthly = pcMonthlySheets_(ss);
  if (!monthly.length) throw new Error('No monthly TP tabs found');

  const all = monthly
    .map(item => {
      const row = pcReadMonthlyRow_(item.sheet, session.staffCode);
      return row ? Object.assign(row, {
        tab:item.sheet.getName(),
        monthLabel:item.parsed.monthName + ' ' + item.parsed.year,
        sortKey:item.sortKey
      }) : null;
    })
    .filter(Boolean);

  if (!all.length) throw new Error('No performance data found for this PC');

  const current = all[0];
  const history = all.slice(0,3);
  const threeMonthAvgKpi = pcAverage_(history.map(r => r.kpi));
  const levelStatus = pcLevelStatus_(current, all, threeMonthAvgKpi);

  return {
    ok:true,
    name:current.name || session.username,
    staffCode:session.staffCode,
    currentMonthLabel:current.monthLabel,
    current:current,
    history:history,
    threeMonthAvgKpi:threeMonthAvgKpi,
    levelStatus:levelStatus,
    breakdown:pcBreakdown_(current)
  };
}

function pcMonthlySheets_(ss) {
  return ss.getSheets().map(sheet => {
    const parsed = pcParseMonth_(sheet.getName());
    if (!parsed) return null;
    return { sheet:sheet, parsed:parsed, sortKey:parsed.year * 12 + parsed.monthIndex };
  }).filter(Boolean).sort((a,b) => b.sortKey - a.sortKey);
}

function pcParseMonth_(name) {
  const aliases = {
    jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,
    jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,september:8,
    oct:9,october:9,nov:10,november:10,dec:11,december:11
  };
  const names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const m = String(name || '').trim().match(/^TP([A-Za-z]+)(\d{4})$/);
  if (!m) return null;
  const key = m[1].toLowerCase();
  if (!Object.prototype.hasOwnProperty.call(aliases,key)) return null;
  const monthIndex = aliases[key];
  return { monthIndex:monthIndex, year:Number(m[2]), monthName:names[monthIndex] };
}

function pcReadMonthlyRow_(sheet, staffCode) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.min(Math.max(sheet.getLastColumn(),1),40);
  if (lastRow < 2) return null;
  const probeRows = Math.min(lastRow,12);
  const probe = sheet.getRange(1,1,probeRows,lastCol).getDisplayValues();
  let headerIndex = -1;
  for (let r=0;r<probe.length;r++) {
    const norm = probe[r].map(pcNormHeader_);
    if (norm.indexOf('staff code') >= 0 && norm.indexOf('pc') >= 0 && norm.indexOf('level') >= 0) { headerIndex = r; break; }
  }
  if (headerIndex < 0) return null;

  const headers = probe[headerIndex].map(pcNormHeader_);
  const rowStart = headerIndex + 2;
  if (rowStart > lastRow) return null;
  const rows = sheet.getRange(rowStart,1,lastRow-rowStart+1,lastCol).getDisplayValues();
  const staffCol = pcFindCol_(headers,['staff code','staffcode']);
  if (staffCol < 0) return null;
  const row = rows.find(r => String(r[staffCol] || '').trim() === String(staffCode || '').trim());
  if (!row) return null;

  const get = names => {
    const i = pcFindCol_(headers,names);
    return i >= 0 ? row[i] : '';
  };
  const conversionRaw = get(['conversion (l60%)','conversion','conversion %']);
  const kpiRaw = get(['kpi']);
  return {
    name:String(get(['pc']) || '').trim(),
    level:pcNormalizeLevel_(get(['level'])),
    deals:pcNum_(get(['deals'])),
    targetDeals:pcNum_(get(['target d','target deals'])),
    targetRevenue:pcNum_(get(['target r','target revenue'])),
    revenue:pcNum_(get(['revenue'])),
    revenuePct:pcPct_(get(['%','revenue kpi'])),
    conversion:pcNum_(conversionRaw),
    conversionAvailable:String(conversionRaw || '').trim() !== '',
    kpi:pcPct_(kpiRaw),
    kpiAvailable:String(kpiRaw || '').trim() !== ''
  };
}

function pcNormHeader_(v) {
  return String(v || '').trim().toLowerCase().replace(/\s+/g,' ');
}
function pcFindCol_(headers,names) {
  for (let n=0;n<names.length;n++) {
    const wanted = pcNormHeader_(names[n]);
    const i = headers.indexOf(wanted);
    if (i >= 0) return i;
  }
  return -1;
}
function pcNum_(v) {
  if (typeof v === 'number') return v;
  return Number(String(v || '').replace(/[฿,%\s,]/g,'')) || 0;
}
function pcPct_(v) { return pcNum_(v); }
function pcAverage_(arr) {
  const nums = (arr || []).filter(v => Number.isFinite(Number(v))).map(Number);
  return nums.length ? nums.reduce((a,b) => a+b,0) / nums.length : null;
}
function pcNormalizeLevel_(v) {
  const s = String(v || '').trim().toUpperCase().replace(/\s+/g,'');
  if (s === 'RN' || s === 'R1' || s === 'R[1]N') return 'R[1]N';
  return s;
}
function pcFmtPct_(v) { return v == null ? '—' : Number(v).toFixed(2) + '%'; }
function pcFmtMoney_(v) { return '฿' + Math.round(Number(v || 0)).toLocaleString('en-US'); }

function pcLevelStatus_(current, all, threeMonthAvg) {
  const level = pcNormalizeLevel_(current.level);
  const nextByLevel = {'R[1]N':'R2','R2':'R3','R3':'R4','R4':'R5'};
  const upThreshold = {'R[1]N':100,'R2':115,'R3':125,'R4':135};
  const downThreshold = {'R[1]N':0,'R2':90,'R3':90,'R4':85,'R5':80};
  const nextLevel = nextByLevel[level] || null;
  const previous = all.length > 1 ? all[1] : null;
  const avg2 = previous ? pcAverage_([current.kpi, previous.kpi]) : null;
  const retain = downThreshold[level];

  let consecutive = 0;
  for (let i=0;i<all.length;i++) {
    if (pcNormalizeLevel_(all[i].level) === level) consecutive++;
    else break;
  }
  const tenureRequired = (level === 'R3' || level === 'R4') ? 3 : 0;
  const fastTrack = avg2 != null && avg2 >= 200;
  const tenureMet = tenureRequired === 0 || consecutive >= tenureRequired || fastTrack;
  const finalMonth = !!current.conversionAvailable;
  const facts = [];

  if (nextLevel) facts.push({label:'Level-up rule',value:'2M Avg ≥ ' + upThreshold[level] + '%'});
  facts.push({label:'2M Avg KPI',value:pcFmtPct_(avg2)});
  facts.push({label:'3M Avg KPI',value:pcFmtPct_(threeMonthAvg)});
  facts.push({label:'Retention minimum',value:retain == null ? '—' : retain + '%'});
  if (tenureRequired) facts.push({label:'Time at ' + level,value:consecutive + ' / ' + tenureRequired + ' months'});

  let code='SECURE', badge='LEVEL SECURE', headline='Current level secure', message='Your rolling performance is above the retention requirement.';
  if (!finalMonth) {
    code='IN_PROGRESS'; badge='ESTIMATE'; headline=nextLevel ? 'Working toward ' + nextLevel : 'Month in progress';
    message='Final status will be confirmed after month-end conversion is entered.';
  } else if (nextLevel && avg2 != null && avg2 >= upThreshold[level] && tenureMet) {
    code=fastTrack && tenureRequired ? 'FAST_TRACK' : 'ELIGIBLE';
    badge=fastTrack && tenureRequired ? 'FAST-TRACK' : 'ELIGIBLE';
    headline='Eligible for ' + nextLevel;
    message=fastTrack && tenureRequired ? 'The accelerated level-up condition is met.' : 'All configured level-up requirements are met.';
  } else if (nextLevel && avg2 != null && avg2 >= upThreshold[level] && !tenureMet) {
    code='TENURE'; badge='WAITING'; headline='Performance target met';
    message=(tenureRequired-consecutive) + ' more month' + ((tenureRequired-consecutive)===1?'':'s') + ' at ' + level + ' required before moving to ' + nextLevel + '.';
  } else if (threeMonthAvg != null && retain != null && threeMonthAvg < retain) {
    code='RISK'; badge='AT RISK'; headline='At risk of level down';
    message='Your 3-month KPI average is below the ' + retain + '% retention requirement for ' + level + '.';
  } else if (nextLevel) {
    code='ALMOST'; badge='IN PROGRESS'; headline='Working toward ' + nextLevel;
    message='Keep building this month toward the 2-month average requirement.';
  }

  let requiredCurrentKpi = null, estimatedRevenueTarget = null, revenueGap = null;
  if (nextLevel && previous) {
    requiredCurrentKpi = Math.max(0, upThreshold[level] * 2 - Number(previous.kpi || 0));
    estimatedRevenueTarget = Number(current.targetRevenue || 0) * requiredCurrentKpi / 100;
    revenueGap = Math.max(0, estimatedRevenueTarget - Number(current.revenue || 0));
    facts.push({label:'This month KPI needed',value:'~' + requiredCurrentKpi.toFixed(2) + '%'});
  }

  return {
    code:code,badge:badge,headline:headline,message:message,nextLevel:nextLevel,
    twoMonthAvgKpi:avg2,threeMonthAvgKpi:threeMonthAvg,retentionMinimum:retain,
    tenureMonths:consecutive,tenureRequired:tenureRequired,fastTrack:fastTrack,
    requiredCurrentKpi:requiredCurrentKpi,
    estimatedRevenueTarget:estimatedRevenueTarget,
    revenueGap:revenueGap,
    facts:facts
  };
}

function pcBreakdown_(current) {
  const revenuePct = Number(current.revenuePct || (current.targetRevenue ? current.revenue/current.targetRevenue*100 : 0));
  const dealsPct = current.targetDeals ? Number(current.deals || 0) / Number(current.targetDeals) * 100 : 0;
  const conversionAvailable = !!current.conversionAvailable;
  return [
    {label:'Revenue',current:revenuePct.toFixed(2)+'%',target:'100%',status:revenuePct>=100?'✓':'Below',statusClass:revenuePct>=100?'ok':'bad'},
    {label:'Deals',current:dealsPct.toFixed(2)+'%',target:'100%',status:dealsPct>=100?'✓':'Below',statusClass:dealsPct>=100?'ok':'bad'},
    {label:'Conversion',current:conversionAvailable?Number(current.conversion||0).toFixed(2)+'%':'Pending',target:PC_CONVERSION_TARGET_.toFixed(0)+'%',status:conversionAvailable?(Number(current.conversion)>=PC_CONVERSION_TARGET_?'✓':'Below'):'Pending',statusClass:conversionAvailable?(Number(current.conversion)>=PC_CONVERSION_TARGET_?'ok':'bad'):'pending'}
  ];
}

function pcJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
