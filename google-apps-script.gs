/**
 * Deploy as: Web app
 * Execute as: Me
 * Who has access: Anyone
 * Then paste the /exec URL into API_URL in script.js.
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TPSep2026');
  if (!sheet) throw new Error('TPSep2026 not found');

  const values = sheet.getRange(7, 1, Math.max(sheet.getLastRow() - 6, 1), 17).getDisplayValues();

  const runners = values
    .filter(row => String(row[0] || '').trim().startsWith('Chawin'))
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

  return ContentService
    .createTextOutput(JSON.stringify({
      updatedAt: new Date().toISOString(),
      sourceTab: 'TPSep2026',
      runners
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNumber(value) {
  if (typeof value === 'number') return value;
  return Number(String(value || '').replace(/[฿,%\s,]/g, '')) || 0;
}

function toPercent(value) {
  const n = toNumber(value);
  return n;
}
