/**
 * Deploy as: Web app
 * Execute as: Me
 * Who has access: Anyone
 *
 * Monthly tab naming convention:
 *   TPJan2026, TPFeb2026, TPMar2026 ... TPSep2026, TPOct2026, etc.
 *
 * The script automatically uses the latest month/year tab it can find.
 * Finish order is permanently logged in a hidden RaceFinishLog sheet.
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getLatestMonthlySheet_(ss);
  if (!sheet) throw new Error('No monthly TP tab found. Expected names like TPSep2026 or TPOct2026.');

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

  const parsed = parseMonthlyTab_(sheet.getName());
  const finishMap = syncFinishLog_(ss, sheet.getName(), parsed, runners);

  runners.forEach(r => {
    const finish = finishMap[raceKey_(r)];
    if (finish) {
      r.finishPlace = finish.place;
      r.finishAt = finish.finishedAt;
      r.finishDate = finish.finishDate;
      r.finishDays = finish.daysToFinish;
    }
  });

  // Once someone finishes, their finishing place is fixed forever for that month.
  // Finishers stay above active racers in finish order; active racers remain sorted by achievement.
  runners.sort((a,b) => {
    const ap = Number(a.finishPlace) || 0;
    const bp = Number(b.finishPlace) || 0;
    if (ap && bp) return ap - bp;
    if (ap) return -1;
    if (bp) return 1;
    return b.achievement - a.achievement || b.revenue - a.revenue;
  });
  runners.forEach((r,i) => r.rank = i + 1);

  return ContentService
    .createTextOutput(JSON.stringify({
      updatedAt: new Date().toISOString(),
      sourceTab: sheet.getName(),
      monthLabel: parsed ? `${parsed.monthName} ${parsed.year}` : sheet.getName(),
      runners
    }))
    .setMimeType(ContentService.MimeType.JSON);
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

    // Historical correction for September 2026: Beam crossed 100% first, then Gorn.
    // This only fixes their finishing places; finish date/day remains whatever is already logged.
    const manualOrder = sourceTab === 'TPSep2026' ? { beam: 1, gorn: 2 } : {};
    runners.forEach(r => {
      const wantedPlace = manualOrder[String(r.name || '').trim().toLowerCase()];
      if (!wantedPlace) return;
      const key = raceKey_(r);
      if (current[key] && current[key].place !== wantedPlace) {
        current[key].place = wantedPlace;
        const rowNumber = rowIndexByKey[key];
        if (rowNumber) log.getRange(rowNumber, 2).setValue(wantedPlace);
      }
      maxPlace = Math.max(maxPlace, wantedPlace);
    });

    // If several people are already over 100% the first time this version runs,
    // historical manual order is respected first; otherwise current leaderboard order is used.
    const newFinishers = runners
      .filter(r => r.achievement >= 100 && !current[raceKey_(r)])
      .sort((a,b) => {
        const ao = manualOrder[String(a.name || '').trim().toLowerCase()] || 9999;
        const bo = manualOrder[String(b.name || '').trim().toLowerCase()] || 9999;
        if (ao !== bo) return ao - bo;
        return b.achievement - a.achievement || b.revenue - a.revenue;
      });

    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Bangkok';
    const daysToFinish = parsed ? getRaceDay_(now, parsed.year, parsed.monthIndex, tz) : '';
    const newRows = [];

    newFinishers.forEach(r => {
      const manualPlace = manualOrder[String(r.name || '').trim().toLowerCase()] || 0;
      const place = manualPlace || (maxPlace + 1);
      maxPlace = Math.max(maxPlace, place);
      const finishDate = Utilities.formatDate(now, tz, 'd MMM yyyy');
      const key = raceKey_(r);
      current[key] = {
        place,
        finishedAt: now.toISOString(),
        finishDate,
        daysToFinish: Number(daysToFinish) || 0
      };
      newRows.push([
        sourceTab,
        place,
        r.name,
        r.staffCode,
        now,
        finishDate,
        daysToFinish,
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

function getLatestMonthlySheet_(ss) {
  const candidates = ss.getSheets()
    .map(sheet => {
      const parsed = parseMonthlyTab_(sheet.getName());
      return parsed ? { sheet, sortKey: parsed.year * 12 + parsed.monthIndex } : null;
    })
    .filter(Boolean)
    .sort((a,b) => b.sortKey - a.sortKey);

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
