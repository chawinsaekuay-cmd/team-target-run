/**
 * Deploy as: Web app
 * Execute as: Me
 * Who has access: Anyone
 *
 * Monthly tab naming convention:
 *   TPJan2026, TPF​​eb2026, TPMar2026 ... TPSep2026, TPOct2026, etc.
 *
 * The script automatically uses the latest month/year tab it can find.
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

  runners.sort((a,b) => b.achievement - a.achievement || b.revenue - a.revenue);
  runners.forEach((r,i) => r.rank = i + 1);

  const parsed = parseMonthlyTab_(sheet.getName());

  return ContentService
    .createTextOutput(JSON.stringify({
      updatedAt: new Date().toISOString(),
      sourceTab: sheet.getName(),
      monthLabel: parsed ? `${parsed.monthName} ${parsed.year}` : sheet.getName(),
      runners
    }))
    .setMimeType(ContentService.MimeType.JSON);
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
