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
 * Finish order is permanently logged in a hidden RaceFinishLog sheet and is
 * only updated while viewing the current month.
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
      sourceTab: sheet.getName(),
      currentTab: latest.sheet.getName(),
      isCurrentMonth,
      monthLabel: `${parsed.monthName} ${parsed.year}`,
      availableMonths,
      runners
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getFinishLog_(ss, sourceTab) {
  const log = ss.getSheetByName('RaceFinishLog');
  if (!log || log.getLastRow() <= 1) return {};

  const rows = log.getRange(2, 1, log.getLastRow() - 1, 8).getValues();
  const current = {};
  rows.forEach(row => {
    if (String(row[0]) !== sourceTab) return;
    const key = `${String(row[3] || '').trim()}|${String(row[2] || '').trim().toLowerCase()}`;
    current[key] = {
      place: Number(row[1]) || 0,
      finishedAt: row[4] instanceof Date ? row[4].toISOString() : String(row[4] || ''),
      finishDate: String(row[5] || ''),
      daysToFinish: Number(row[6]) || 0
    };
  });
  return current;
}

function syncFinishLog_(ss, sourceTab, parsed, runners) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let log = ss.getSheetByName('RaceFinishLog');
    if (!log) {
      log = ss.insertSheet('RaceFinishLog');
      log.getRange(1, 1, 1, 8).setValues([[
        'Month Tab','Place','Name','Staff Code','Finished At','Finish Date','Days To Finish','Achievement When Logged'
      ]]);
      log.setFrozenRows(1);
      log.hideSheet();
    }

    const lastRow = log.getLastRow();
    const rows = lastRow > 1 ? log.getRange(2, 1, lastRow - 1, 8).getValues() : [];
    const current = {};
    const rowIndexByKey = {};
    let maxPlace = 0;

    rows.forEach((row, index) => {
      if (String(row[0]) !== sourceTab) return;
      const key = `${String(row[3] || '').trim()}|${String(row[2] || '').trim().toLowerCase()}`;
      const place = Number(row[1]) || 0;
      current[key] = {
        place,
        finishedAt: row[4] instanceof Date ? row[4].toISOString() : String(row[4] || ''),
        finishDate: String(row[5] || ''),
        daysToFinish: Number(row[6]) || 0
      };
      rowIndexByKey[key] = index + 2;
      maxPlace = Math.max(maxPlace, place);
    });

    const manualFinish = sourceTab === 'TPSep2026' ? {
      beam: { place: 1, finishDate: '9 Sep 2026', daysToFinish: 9, iso: '2026-09-09T12:00:00+07:00' },
      gorn: { place: 2, finishDate: '13 Sep 2026', daysToFinish: 13, iso: '2026-09-13T12:00:00+07:00' }
    } : {};

    runners.forEach(r => {
      const override = manualFinish[String(r.name || '').trim().toLowerCase()];
      if (!override) return;
      const key = raceKey_(r);
      const rowNumber = rowIndexByKey[key];

      if (current[key]) {
        current[key].place = override.place;
        current[key].finishedAt = override.iso;
        current[key].finishDate = override.finishDate;
        current[key].daysToFinish = override.daysToFinish;
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
      }
      maxPlace = Math.max(maxPlace, override.place);
    });

    const newFinishers = runners
      .filter(r => r.achievement >= 100 && !current[raceKey_(r)])
      .sort((a,b) => {
        const ao = (manualFinish[String(a.name || '').trim().toLowerCase()] || {}).place || 9999;
        const bo = (manualFinish[String(b.name || '').trim().toLowerCase()] || {}).place || 9999;
        if (ao !== bo) return ao - bo;
        return b.achievement - a.achievement || b.revenue - a.revenue;
      });

    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Bangkok';
    const daysToFinish = parsed ? getRaceDay_(now, parsed.year, parsed.monthIndex, tz) : '';
    const newRows = [];

    newFinishers.forEach(r => {
      const override = manualFinish[String(r.name || '').trim().toLowerCase()];
      const place = override ? override.place : (maxPlace + 1);
      maxPlace = Math.max(maxPlace, place);
      const finishedAt = override ? new Date(override.iso) : now;
      const finishDate = override ? override.finishDate : Utilities.formatDate(now, tz, 'd MMM yyyy');
      const finishDays = override ? override.daysToFinish : daysToFinish;
      const key = raceKey_(r);
      current[key] = {
        place,
        finishedAt: finishedAt.toISOString(),
        finishDate,
        daysToFinish: Number(finishDays) || 0
      };
      newRows.push([
        sourceTab,
        place,
        r.name,
        r.staffCode,
        finishedAt,
        finishDate,
        finishDays,
        r.achievement
      ]);
    });

    if (newRows.length) {
      log.getRange(log.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
    }

    return current;
  } finally {
    lock.releaseLock();
  }
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
