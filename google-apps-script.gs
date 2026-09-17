/**
 * Deploy as: Web app
 * Execute as: Me
 * Who has access: Anyone
 *
 * Monthly tab naming convention:
 *   TPJan2026, TPFeb2026, TPMar2026 ... TPSep2026, TPOct2026, etc.
 *
 * Historical browsing starts from September 2026 onward.
 * The API defaults to the latest month, but also accepts ?month=TPSep2026.
 *
 * Race logic:
 * - First time someone reaches 100%, that original finish is permanently recorded.
 * - They only keep an active podium position while their current achievement stays >=100%.
 * - If a broken deal drops them below 100%, they lose the active podium position and the
 *   remaining qualified finishers move up.
 * - If they later reach 100% again, they re-qualify at that later time and go behind people
 *   who remained qualified. Original first-finish history is never deleted.
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthlySheets = getMonthlySheets_(ss);
  if (!monthlySheets.length) throw new Error('No monthly TP tab found from September 2026 onward.');

  const latest = monthlySheets[0];
  const requestedMonth = String(e && e.parameter && e.parameter.month || '').trim();
  const selected = monthlySheets.find(item => item.sheet.getName() === requestedMonth) || latest;
  const sheet = selected.sheet;
  const isCurrentMonth = sheet.getName() === latest.sheet.getName();

  const values = sheet.getRange(7, 1, Math.max(sheet.getLastRow() - 6, 1), 17).getDisplayValues();

  const runners = values
    .filter(row => {
      const team = String(row[0] || '').trim();
      return team.startsWith('Chawin') || team.startsWith('Junior');
    })
    .filter(row => String(row[2] || '').trim())
    .map(row => ({
      team: row[0],
      staffCode: row[1],
      name: row[2],
      level: row[3],
      deals: toNumber(row[4]),
      targetDeals: toNumber(row[5]),
      targetRevenue: toNumber(row[6]),
      revenue: toNumber(row[7]),
      achievement: toPercent(row[9]),
      kpi: row[12] || '',
      workMode: row[13] || ''
    }));

  const suspension = getSuspensionStatus_(ss);
  const leadQuota = getLeadQuotaStatus_(ss);
  runners.forEach(r => {
    const staffCode = String(r.staffCode || '').trim();
    const s = suspension.byStaffCode[staffCode];
    const lq = leadQuota.byStaffCode[staffCode];
    r.suspended = s ? !!s.suspended : false;
    r.suspensionCheckedAt = s ? s.checkedAt : suspension.checkedAt;
    r.leadQuota = lq ? lq.leadQuota : null;
    r.leadsReceived = lq ? lq.leadsReceived : null;
    r.leadQuotaCheckedAt = lq ? lq.checkedAt : leadQuota.checkedAt;
  });

  const parsed = selected.parsed;
  const finishMap = isCurrentMonth
    ? syncFinishLog_(ss, sheet.getName(), parsed, runners)
    : getFinishLog_(ss, sheet.getName());

  runners.forEach(r => {
    const finish = finishMap[raceKey_(r)];
    if (finish) {
      r.finishPlace = finish.place;
      r.finishAt = finish.finishedAt;
      r.finishDate = finish.finishDate;
      r.finishDays = finish.daysToFinish;
      r.originalFinishPlace = finish.originalPlace || finish.place;
    }
  });

  runners.sort((a,b) => {
    const ap = Number(a.finishPlace) || 0;
    const bp = Number(b.finishPlace) || 0;
    if (ap && bp) return ap - bp;
    if (ap) return -1;
    if (bp) return 1;
    return b.achievement - a.achievement || b.revenue - a.revenue;
  });
  runners.forEach((r,i) => r.rank = i + 1);

  const availableMonths = monthlySheets.map(item => ({
    tab: item.sheet.getName(),
    label: `${item.parsed.monthName} ${item.parsed.year}`,
    isCurrent: item.sheet.getName() === latest.sheet.getName()
  }));

  return ContentService
    .createTextOutput(JSON.stringify({
      updatedAt: new Date().toISOString(),
      suspensionCheckedAt: suspension.checkedAt,
      leadQuotaCheckedAt: leadQuota.checkedAt,
      sourceTab: sheet.getName(),
      currentTab: latest.sheet.getName(),
      isCurrentMonth,
      monthLabel: `${parsed.monthName} ${parsed.year}`,
      availableMonths,
      runners
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSuspensionStatus_(ss) {
  const sheet = ss.getSheetByName('Suspended Status');
  const result = { checkedAt:'', byStaffCode:{} };
  if (!sheet || sheet.getLastRow() < 2) return result;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 5)).getDisplayValues();
  rows.forEach(row => {
    const checkedAt = String(row[0] || '').trim();
    const staffCode = String(row[1] || '').trim();
    if (!staffCode) return;
    const raw = String(row[4] || '').trim().toLowerCase();
    const suspended = raw === 'true' || raw === 'yes' || raw === '1' || raw === 'suspended';
    result.byStaffCode[staffCode] = { suspended, checkedAt };
    if (checkedAt && (!result.checkedAt || checkedAt > result.checkedAt)) result.checkedAt = checkedAt;
  });
  return result;
}

function getLeadQuotaStatus_(ss) {
  const sheet = ss.getSheetByName('Lead Quota Status');
  const result = { checkedAt:'', byStaffCode:{} };
  if (!sheet || sheet.getLastRow() < 2) return result;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 6)).getDisplayValues();
  rows.forEach(row => {
    const checkedAt = String(row[0] || '').trim();
    const staffCode = String(row[1] || '').trim();
    if (!staffCode) return;
    result.byStaffCode[staffCode] = {
      leadQuota: toNumber(row[4]),
      leadsReceived: toNumber(row[5]),
      checkedAt
    };
    if (checkedAt && (!result.checkedAt || checkedAt > result.checkedAt)) result.checkedAt = checkedAt;
  });
  return result;
}

function ensureFinishLog_(ss) {
  let log = ss.getSheetByName('RaceFinishLog');
  if (!log) {
    log = ss.insertSheet('RaceFinishLog');
    log.getRange(1, 1, 1, 10).setValues([[
      'Month Tab','Original Place','Name','Staff Code','First Finished At','First Finish Date','First Days To Finish','Achievement When First Logged','Active Qualified At','Qualified Now'
    ]]);
    log.setFrozenRows(1);
    log.hideSheet();
    return log;
  }

  // Upgrade older 8-column logs in place without losing history.
  if (log.getMaxColumns() < 10) log.insertColumnsAfter(log.getMaxColumns(), 10 - log.getMaxColumns());
  const headers = log.getRange(1, 1, 1, 10).getValues()[0];
  const wanted = [
    'Month Tab','Original Place','Name','Staff Code','First Finished At','First Finish Date','First Days To Finish','Achievement When First Logged','Active Qualified At','Qualified Now'
  ];
  let changed = false;
  wanted.forEach((v,i) => {
    if (headers[i] !== v) { headers[i] = v; changed = true; }
  });
  if (changed) log.getRange(1,1,1,10).setValues([headers]);
  return log;
}

function getFinishLog_(ss, sourceTab) {
  const log = ss.getSheetByName('RaceFinishLog');
  if (!log || log.getLastRow() <= 1) return {};

  const width = Math.min(Math.max(log.getLastColumn(), 8), 10);
  const rows = log.getRange(2, 1, log.getLastRow() - 1, width).getValues();
  const active = [];

  rows.forEach(row => {
    if (String(row[0]) !== sourceTab) return;
    const key = `${String(row[3] || '').trim()}|${String(row[2] || '').trim().toLowerCase()}`;
    const originalPlace = Number(row[1]) || 0;
    const firstAt = row[4] instanceof Date ? row[4] : (row[4] ? new Date(row[4]) : null);
    const activeAtRaw = width >= 9 ? row[8] : '';
    const qualifiedRaw = width >= 10 ? row[9] : '';
    const qualified = qualifiedRaw === '' ? true : String(qualifiedRaw).toLowerCase() !== 'false';
    if (!qualified) return;

    const activeAt = activeAtRaw instanceof Date ? activeAtRaw : (activeAtRaw ? new Date(activeAtRaw) : firstAt);
    active.push({
      key,
      originalPlace,
      activeAt,
      finishedAt: firstAt ? firstAt.toISOString() : '',
      finishDate: String(row[5] || ''),
      daysToFinish: Number(row[6]) || 0
    });
  });

  active.sort((a,b) => {
    const at = a.activeAt ? a.activeAt.getTime() : 0;
    const bt = b.activeAt ? b.activeAt.getTime() : 0;
    return at - bt || a.originalPlace - b.originalPlace;
  });

  const result = {};
  active.forEach((item,index) => {
    result[item.key] = {
      place:index+1,
      originalPlace:item.originalPlace,
      finishedAt:item.finishedAt,
      finishDate:item.finishDate,
      daysToFinish:item.daysToFinish
    };
  });
  return result;
}

function syncFinishLog_(ss, sourceTab, parsed, runners) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const log = ensureFinishLog_(ss);
    const lastRow = log.getLastRow();
    const rows = lastRow > 1 ? log.getRange(2, 1, lastRow - 1, 10).getValues() : [];
    const current = {};
    const rowIndexByKey = {};
    let maxOriginalPlace = 0;

    rows.forEach((row, index) => {
      if (String(row[0]) !== sourceTab) return;
      const key = `${String(row[3] || '').trim()}|${String(row[2] || '').trim().toLowerCase()}`;
      const originalPlace = Number(row[1]) || 0;
      const firstAt = row[4] instanceof Date ? row[4] : (row[4] ? new Date(row[4]) : null);
      const activeAt = row[8] instanceof Date ? row[8] : (row[8] ? new Date(row[8]) : firstAt);
      const qualified = row[9] === '' ? true : String(row[9]).toLowerCase() !== 'false';
      current[key] = {
        originalPlace,
        firstAt,
        finishDate:String(row[5] || ''),
        daysToFinish:Number(row[6]) || 0,
        activeAt,
        qualified
      };
      rowIndexByKey[key] = index + 2;
      maxOriginalPlace = Math.max(maxOriginalPlace, originalPlace);
    });

    const manualFinish = sourceTab === 'TPSep2026' ? {
      beam: { place: 1, finishDate: '9 Sep 2026', daysToFinish: 9, iso: '2026-09-09T12:00:00+07:00' },
      gorn: { place: 2, finishDate: '13 Sep 2026', daysToFinish: 13, iso: '2026-09-13T12:00:00+07:00' }
    } : {};

    // Preserve user-confirmed first-finish history for September 2026.
    runners.forEach(r => {
      const override = manualFinish[String(r.name || '').trim().toLowerCase()];
      if (!override) return;
      const key = raceKey_(r);
      const rowNumber = rowIndexByKey[key];
      if (!current[key]) return;

      current[key].originalPlace = override.place;
      current[key].firstAt = new Date(override.iso);
      current[key].finishDate = override.finishDate;
      current[key].daysToFinish = override.daysToFinish;
      if (!current[key].activeAt) current[key].activeAt = new Date(override.iso);

      if (rowNumber) {
        log.getRange(rowNumber, 2, 1, 6).setValues([[
          override.place,
          r.name,
          r.staffCode,
          new Date(override.iso),
          override.finishDate,
          override.daysToFinish
        ]]);
      }
      maxOriginalPlace = Math.max(maxOriginalPlace, override.place);
    });

    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Bangkok';
    const daysToFinish = parsed ? getRaceDay_(now, parsed.year, parsed.monthIndex, tz) : '';
    const newRows = [];

    // First-time finishers: permanently record their original finish event.
    runners
      .filter(r => r.achievement >= 100 && !current[raceKey_(r)])
      .sort((a,b) => {
        const ao = (manualFinish[String(a.name || '').trim().toLowerCase()] || {}).place || 9999;
        const bo = (manualFinish[String(b.name || '').trim().toLowerCase()] || {}).place || 9999;
        if (ao !== bo) return ao - bo;
        return b.achievement - a.achievement || b.revenue - a.revenue;
      })
      .forEach(r => {
        const override = manualFinish[String(r.name || '').trim().toLowerCase()];
        const originalPlace = override ? override.place : (maxOriginalPlace + 1);
        maxOriginalPlace = Math.max(maxOriginalPlace, originalPlace);
        const firstAt = override ? new Date(override.iso) : now;
        const finishDate = override ? override.finishDate : Utilities.formatDate(now, tz, 'd MMM yyyy');
        const finishDays = override ? override.daysToFinish : daysToFinish;
        const key = raceKey_(r);

        current[key] = {
          originalPlace,
          firstAt,
          finishDate,
          daysToFinish:Number(finishDays) || 0,
          activeAt:firstAt,
          qualified:true
        };
        newRows.push([
          sourceTab,
          originalPlace,
          r.name,
          r.staffCode,
          firstAt,
          finishDate,
          finishDays,
          r.achievement,
          firstAt,
          true
        ]);
      });

    if (newRows.length) log.getRange(log.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);

    // Broken-deal rule: qualification is live. Dropping below 100 removes the active podium spot.
    // Crossing back to 100 later gets a fresh active qualification time, so they re-enter behind
    // everyone who stayed qualified.
    runners.forEach(r => {
      const key = raceKey_(r);
      const record = current[key];
      if (!record) return;
      const isQualifiedNow = r.achievement >= 100;
      const rowNumber = rowIndexByKey[key] || findLogRow_(log, sourceTab, r);

      if (!isQualifiedNow && record.qualified) {
        record.qualified = false;
        if (rowNumber) log.getRange(rowNumber, 10).setValue(false);
      } else if (isQualifiedNow && !record.qualified) {
        record.qualified = true;
        record.activeAt = now;
        if (rowNumber) log.getRange(rowNumber, 9, 1, 2).setValues([[now, true]]);
      } else if (isQualifiedNow && record.qualified && !record.activeAt) {
        record.activeAt = record.firstAt || now;
        if (rowNumber) log.getRange(rowNumber, 9, 1, 2).setValues([[record.activeAt, true]]);
      }
    });

    const active = Object.keys(current)
      .map(key => ({ key, ...current[key] }))
      .filter(item => item.qualified)
      .sort((a,b) => {
        const at = a.activeAt ? a.activeAt.getTime() : 0;
        const bt = b.activeAt ? b.activeAt.getTime() : 0;
        return at - bt || a.originalPlace - b.originalPlace;
      });

    const result = {};
    active.forEach((item,index) => {
      result[item.key] = {
        place:index+1,
        originalPlace:item.originalPlace,
        finishedAt:item.firstAt ? item.firstAt.toISOString() : '',
        finishDate:item.finishDate,
        daysToFinish:item.daysToFinish
      };
    });
    return result;
  } finally {
    lock.releaseLock();
  }
}

function findLogRow_(log, sourceTab, runner) {
  if (log.getLastRow() <= 1) return 0;
  const rows = log.getRange(2,1,log.getLastRow()-1,4).getValues();
  const key = raceKey_(runner);
  for (let i=0;i<rows.length;i++) {
    const rowKey = `${String(rows[i][3] || '').trim()}|${String(rows[i][2] || '').trim().toLowerCase()}`;
    if (String(rows[i][0]) === sourceTab && rowKey === key) return i+2;
  }
  return 0;
}

function raceKey_(runner) {
  return `${String(runner.staffCode || '').trim()}|${String(runner.name || '').trim().toLowerCase()}`;
}

function getRaceDay_(date, year, monthIndex, timezone) {
  const localDay = Number(Utilities.formatDate(date, timezone, 'd'));
  const localMonth = Number(Utilities.formatDate(date, timezone, 'M')) - 1;
  const localYear = Number(Utilities.formatDate(date, timezone, 'yyyy'));
  if (localYear === year && localMonth === monthIndex) return localDay;

  const start = new Date(year, monthIndex, 1);
  return Math.max(1, Math.ceil((date.getTime() - start.getTime()) / 86400000));
}

function getMonthlySheets_(ss) {
  const historyStartKey = 2026 * 12 + 8; // September 2026
  return ss.getSheets()
    .map(sheet => {
      const parsed = parseMonthlyTab_(sheet.getName());
      if (!parsed) return null;
      const sortKey = parsed.year * 12 + parsed.monthIndex;
      return sortKey >= historyStartKey ? { sheet, parsed, sortKey } : null;
    })
    .filter(Boolean)
    .sort((a,b) => b.sortKey - a.sortKey);
}

function getLatestMonthlySheet_(ss) {
  const candidates = getMonthlySheets_(ss);
  return candidates.length ? candidates[0].sheet : null;
}

function parseMonthlyTab_(name) {
  const months = {
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
    jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
  };
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const match = String(name || '').trim().match(/^TP(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})$/i);
  if (!match) return null;

  const monthIndex = months[match[1].toLowerCase()];
  const year = Number(match[2]);
  return { monthIndex, year, monthName: monthNames[monthIndex] };
}

function toNumber(value) {
  if (typeof value === 'number') return value;
  return Number(String(value || '').replace(/[฿,%\s,]/g, '')) || 0;
}

function toPercent(value) {
  return toNumber(value);
}