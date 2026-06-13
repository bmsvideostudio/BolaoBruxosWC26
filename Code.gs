// ============================================================
//  BOLÃO DOS BRUXOS — World Cup 2026 | Google Apps Script
// ============================================================

const SHEET_APOSTAS     = "Apostas";
const SHEET_RESULTADOS  = "Resultados";
const SHEET_PLAYERS     = "Players";
const SHEET_COMENTARIOS = "Comentarios";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === "salvarAposta")     return salvarAposta(data);
    if (action === "salvarResultado")  return salvarResultado(data);
    if (action === "salvarComentario") return salvarComentario(data);
    if (action === "getRanking")       return getRanking();
    if (action === "getApostas")       return getApostas(data.player);
    if (action === "getResultados")    return getResultados();
    if (action === "getComentarios")   return getComentarios(data.player);
    if (action === "getPlayers")       return getPlayers();
    return resp({ ok: false, msg: "Ação desconhecida" });
  } catch (err) {
    return resp({ ok: false, msg: err.toString() });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "getRanking")      return getRanking();
  if (action === "getResultados")   return getResultados();
  if (action === "getPlayers")      return getPlayers();
  if (action === "getApostas")      return getApostas(e.parameter.player);
  if (action === "getComentarios")  return getComentarios(e.parameter.player);
  return resp({ ok: false, msg: "Ação desconhecida" });
}

// ---------- APOSTAS ----------

function salvarAposta(data) {
  const sheet = getOrCreateSheet(SHEET_APOSTAS,
    ["timestamp","player","jogo_id","time1","time2","gols1","gols2"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.player && rows[i][2] === data.jogo_id) {
      sheet.getRange(i+1,1).setValue(new Date().toISOString());
      sheet.getRange(i+1,6).setValue(data.gols1);
      sheet.getRange(i+1,7).setValue(data.gols2);
      return resp({ ok: true, msg: "Aposta atualizada!" });
    }
  }
  sheet.appendRow([new Date().toISOString(), data.player, data.jogo_id,
    data.time1, data.time2, data.gols1, data.gols2]);
  registrarPlayer(data.player);
  return resp({ ok: true, msg: "Aposta salva!" });
}

function getApostas(player) {
  const sheet = getOrCreateSheet(SHEET_APOSTAS,
    ["timestamp","player","jogo_id","time1","time2","gols1","gols2"]);
  const rows = sheet.getDataRange().getValues();
  const apostas = {};
  rows.slice(1).forEach((r) => {
    if (!player || r[1] === player)
      apostas[r[2]] = { gols1: r[5], gols2: r[6] };
  });
  return resp({ ok: true, apostas });
}

// ---------- COMENTÁRIOS ----------

function salvarComentario(data) {
  const sheet = getOrCreateSheet(SHEET_COMENTARIOS,
    ["timestamp","player","jogo_id","comentario"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.player && rows[i][2] === data.jogo_id) {
      sheet.getRange(i+1,1).setValue(new Date().toISOString());
      sheet.getRange(i+1,4).setValue(data.comentario);
      return resp({ ok: true, msg: "Comentário atualizado!" });
    }
  }
  sheet.appendRow([new Date().toISOString(), data.player, data.jogo_id, data.comentario]);
  return resp({ ok: true, msg: "Comentário salvo!" });
}

function getComentarios(player) {
  const sheet = getOrCreateSheet(SHEET_COMENTARIOS,
    ["timestamp","player","jogo_id","comentario"]);
  const rows = sheet.getDataRange().getValues();
  const comentarios = {};
  rows.slice(1).forEach((r) => {
    if (!player || r[1] === player)
      comentarios[r[2]] = r[3];
  });
  return resp({ ok: true, comentarios });
}

// ---------- RESULTADOS ----------

function salvarResultado(data) {
  const sheet = getOrCreateSheet(SHEET_RESULTADOS,
    ["jogo_id","gols1","gols2","atualizado"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.jogo_id) {
      sheet.getRange(i+1,2).setValue(data.gols1);
      sheet.getRange(i+1,3).setValue(data.gols2);
      sheet.getRange(i+1,4).setValue(new Date().toISOString());
      return resp({ ok: true, msg: "Resultado atualizado!" });
    }
  }
  sheet.appendRow([data.jogo_id, data.gols1, data.gols2, new Date().toISOString()]);
  return resp({ ok: true, msg: "Resultado salvo!" });
}

function getResultados() {
  const sheet = getOrCreateSheet(SHEET_RESULTADOS,
    ["jogo_id","gols1","gols2","atualizado"]);
  const rows = sheet.getDataRange().getValues();
  const resultados = {};
  rows.slice(1).forEach((r) => {
    resultados[r[0]] = { gols1: r[1], gols2: r[2] };
  });
  return resp({ ok: true, resultados });
}

// ---------- RANKING ----------

function getRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName(SHEET_APOSTAS);
  const sheetR = ss.getSheetByName(SHEET_RESULTADOS);
  const sheetP = ss.getSheetByName(SHEET_PLAYERS);
  if (!sheetA || !sheetR || !sheetP) return resp({ ok: true, ranking: [] });

  const apostasRows    = sheetA.getDataRange().getValues().slice(1);
  const resultadosRows = sheetR.getDataRange().getValues().slice(1);
  const playersRows    = sheetP.getDataRange().getValues().slice(1);

  const resultados = {};
  resultadosRows.forEach((r) => {
    resultados[r[0]] = { gols1: Number(r[1]), gols2: Number(r[2]) };
  });

  const pontos = {};
  playersRows.forEach((r) => {
    if (r[0]) pontos[r[0]] = { exato: 0, resultado: 0, gols: 0, total: 0 };
  });

  apostasRows.forEach((r) => {
    const player = r[1];
    const jogoId = r[2];
    const ag1 = Number(r[5]);
    const ag2 = Number(r[6]);
    if (!pontos[player]) pontos[player] = { exato: 0, resultado: 0, gols: 0, total: 0 };
    const res = resultados[jogoId];
    if (res === undefined) return;
    const rg1 = res.gols1;
    const rg2 = res.gols2;
    let pts = 0;
    if (ag1 === rg1 && ag2 === rg2) {
      pts += 3; pontos[player].exato++;
    } else {
      const vA = ag1 > ag2 ? 1 : ag1 < ag2 ? 2 : 0;
      const vR = rg1 > rg2 ? 1 : rg1 < rg2 ? 2 : 0;
      if (vA === vR) { pts += 1; pontos[player].resultado++; }
    }
    if (ag1 === rg1 || ag2 === rg2) { pts += 1; pontos[player].gols++; }
    pontos[player].total += pts;
  });

  const ranking = Object.entries(pontos)
    .map(([nome, p]) => ({ nome, ...p }))
    .sort((a, b) => b.total - a.total || b.exato - a.exato);

  return resp({ ok: true, ranking });
}

// ---------- PLAYERS ----------

function registrarPlayer(nome) {
  const sheet = getOrCreateSheet(SHEET_PLAYERS, ["nome","desde"]);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === nome) return;
  }
  sheet.appendRow([nome, new Date().toISOString()]);
}

function getPlayers() {
  const sheet = getOrCreateSheet(SHEET_PLAYERS, ["nome","desde"]);
  const rows = sheet.getDataRange().getValues();
  const players = rows.slice(1).map((r) => r[0]).filter(Boolean);
  return resp({ ok: true, players });
}

// ---------- HELPERS ----------

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
