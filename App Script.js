function doGet(e) {
  const output = handleRequest();
  return output;
}

function doPost(e) {
  const output = handleRequest();
  return output;
}

function handleRequest() {
  const SHEET_ID = '14Yr7ND4E1XP-7EWNiuhzy2qpz-Gf1A_9ZiMbrgIN9Xc';
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Aba Trustpilot ──────────────────────────────────
  const shTP = ss.getSheetByName('Trustpilot');
  const rawTP = shTP.getDataRange().getValues();
  const headersTP = rawTP[0].map(h => h.toString().trim());
  const rowsTP = rawTP.slice(1).map(row => {
    const obj = {};
    headersTP.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd');
