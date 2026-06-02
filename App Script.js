function doGet(e) {
  const SHEET_ID = '14Yr7ND4E1XP-7EWNiuhzy2qpz-Gf1A_9ZiMbrgIN9Xc';
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Aba Trustpilot ──────────────────────────────────────
  const shTP = ss.getSheetByName('Trustpilot');
  const rawTP = shTP.getDataRange().getValues();
  const headersTP = rawTP[0].map(h => h.toString().trim());
  const rowsTP = rawTP.slice(1).map(row => {
    const obj = {};
    headersTP.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd');
      obj[h] = v === null || v === undefined ? '' : v.toString().trim();
    });
    return obj;
  });

  // ── Aba Visitor Insights (cria se não existir) ──────────
  let visitorData = [];
  let shVI = ss.getSheetByName('Visitor Insights');
  if (!shVI) {
    shVI = ss.insertSheet('Visitor Insights');
    shVI.getRange(1,1,1,6).setValues([['Mês','Visitas','Principal País','Tempo Médio (seg)','Reviews Visualizadas (méd)','Observações']]);
    shVI.setFrozenRows(1);
  }
  const rawVI = shVI.getDataRange().getValues();
  if (rawVI.length > 1) {
    const headVI = rawVI[0].map(h => h.toString().trim());
    visitorData = rawVI.slice(1).map(row => {
      const obj = {};
      headVI.forEach((h, i) => { obj[h] = row[i] ? row[i].toString().trim() : ''; });
      return obj;
    });
  }

  // ── Aba Invitation (cria se não existir) ───────────────
  let inviteData = [];
  let shINV = ss.getSheetByName('Invitation');
  if (!shINV) {
    shINV = ss.insertSheet('Invitation');
    shINV.getRange(1,1,1,6).setValues([['Mês','Invites Enviados','Taxa de Abertura (%)','Cliques','Conversão (%)','Observações']]);
    shINV.setFrozenRows(1);
  }
  const rawINV = shINV.getDataRange().getValues();
  if (rawINV.length > 1) {
    const headINV = rawINV[0].map(h => h.toString().trim());
    inviteData = rawINV.slice(1).map(row => {
      const obj = {};
      headINV.forEach((h, i) => { obj[h] = row[i] ? row[i].toString().trim() : ''; });
      return obj;
    });
  }

  // ── Reply Insights (cria se não existir) ───────────────
  let replyData = [];
  let shRP = ss.getSheetByName('Reply Insights');
  if (!shRP) {
    shRP = ss.insertSheet('Reply Insights');
    shRP.getRange(1,1,1,5).setValues([['Mês','Volume de Replies','Taxa de Reply (%)','Tempo Médio Resposta','Observações']]);
    shRP.setFrozenRows(1);
  }
  const rawRP = shRP.getDataRange().getValues();
  if (rawRP.length > 1) {
    const headRP = rawRP[0].map(h => h.toString().trim());
    replyData = rawRP.slice(1).map(row => {
      const obj = {};
      headRP.forEach((h, i) => { obj[h] = row[i] ? row[i].toString().trim() : ''; });
      return obj;
    });
  }

  const output = ContentService
    .createTextOutput(JSON.stringify({
      trustpilot: rowsTP,
      visitor: visitorData,
      invitation: inviteData,
      reply: replyData
    }))
    .setMimeType(ContentService.MimeType.JSON);

  return output;
}