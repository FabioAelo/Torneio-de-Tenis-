const SPREADSHEET_ID = '11j3X1nisX87z7_0Cxqi_DV5EO3Y7Z3yRG_ex5SLtkmM';
const MATCH_SHEET = 'Lançamento Único';
const LIVE_SHEET = 'Placar Ao Vivo';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Placar ao vivo — Torneio Vog Ville')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupPlacar() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(LIVE_SHEET);
  if (!sh) sh = ss.insertSheet(LIVE_SHEET);
  sh.getRange(1, 1, 1, 4).setValues([['Linha da partida', 'Estado JSON', 'Atualizado em', 'Status']]);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 4);
  return 'Aba Placar Ao Vivo preparada.';
}

function getMatches(pin) {
  requirePin_(pin);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(MATCH_SHEET);
  if (!sh) throw new Error('Não encontrei a aba ' + MATCH_SHEET + '.');
  const values = sh.getRange(5, 1, Math.max(sh.getLastRow() - 4, 1), 13).getDisplayValues();
  const matches = [];
  values.forEach((r, i) => {
    const rowNum = i + 5;
    const modality = (r[3] || '').trim();
    const a = (r[5] || '').trim();
    const b = (r[7] || '').trim();
    if (!a || !b || !modality) return;
    matches.push({
      row: rowNum,
      id: r[0] || String(rowNum - 4),
      date: r[1] || '', time: r[2] || '', modality,
      category: r[4] || '',
      playerA: modality.toLowerCase().includes('dupla') ? [a, r[6]].filter(Boolean).join(' / ') : a,
      playerB: modality.toLowerCase().includes('dupla') ? [b, r[8]].filter(Boolean).join(' / ') : b,
      finished: r[9] !== '' && r[10] !== '', scoreA: r[9], scoreB: r[10]
    });
  });
  return matches;
}

function getLiveState(pin, row) {
  requirePin_(pin);
  const sh = getLiveSheet_();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(row) && data[i][1]) {
      try { return JSON.parse(data[i][1]); } catch (e) {}
    }
  }
  return null;
}

function saveLiveState(pin, row, state) {
  requirePin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getLiveSheet_();
    const data = sh.getDataRange().getValues();
    let target = 0;
    for (let i = 1; i < data.length; i++) if (Number(data[i][0]) === Number(row)) { target = i + 1; break; }
    if (!target) target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, 4).setValues([[Number(row), JSON.stringify(state), new Date(), state.status || 'Em andamento']]);
    return { ok: true, updatedAt: new Date().toISOString() };
  } finally { lock.releaseLock(); }
}

function finishMatch(pin, row, state) {
  requirePin_(pin);
  if (!state || state.status !== 'Finalizado' || !state.winner) throw new Error('A partida ainda não está finalizada.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(MATCH_SHEET);
    const r = Number(row);
    const current = sh.getRange(r, 10, 1, 2).getValues()[0];
    if (current[0] !== '' || current[1] !== '') throw new Error('Esta partida já tem resultado na planilha. Para evitar sobrescrever dados, confira a aba Lançamento Único antes de continuar.');
    const totalA = Number(state.totalGamesA || 0);
    const totalB = Number(state.totalGamesB || 0);
    // J:K são os campos usados pelas fórmulas atuais de classificação.
    sh.getRange(r, 10, 1, 2).setValues([[totalA, totalB]]);
    saveLiveStateInternal_(getLiveSheet_(), r, state);
    return { ok: true, totalGamesA: totalA, totalGamesB: totalB };
  } finally { lock.releaseLock(); }
}

function getLiveSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(LIVE_SHEET);
  if (!sh) { setupPlacar(); sh = ss.getSheetByName(LIVE_SHEET); }
  return sh;
}
function saveLiveStateInternal_(sh, row, state) {
  const data = sh.getDataRange().getValues();
  let target = 0;
  for (let i = 1; i < data.length; i++) if (Number(data[i][0]) === Number(row)) { target = i + 1; break; }
  if (!target) target = sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, 4).setValues([[Number(row), JSON.stringify(state), new Date(), state.status || 'Finalizado']]);
}
function requirePin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('PLACAR_PIN');
  if (!expected) throw new Error('O administrador precisa definir PLACAR_PIN nas Propriedades do script antes de usar o placar.');
  if (String(pin || '') !== expected) throw new Error('PIN incorreto.');
}
