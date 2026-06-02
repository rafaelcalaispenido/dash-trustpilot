function doGet(e) {
  const SHEET_ID = '14Yr7ND4E1XP-7EWNiuhzy2qpz-Gf1A_9ZiMbrgIN9Xc';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Trustpilot');
  const raw = sh.getDataRange().getValues();
  const headers = raw[0].map(h => h.toString().trim());
  const rows = raw.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd');
      } else if (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}/)) {
        v = v.substring(0, 10);
      } else {
        v = (v === null || v === undefined) ? '' : v.toString().trim();
      }
      obj[h] = v;
    });
    return obj;
  });
  const cb = (e.parameter && e.parameter.callback) ? e.parameter.callback : 'callback';
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify({ trustpilot: rows }) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
