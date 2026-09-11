// ============================================================================
// SERIE A ANALYTICS ENGINE — porting fedele della logica del workbook Excel
// (PLAYERS, TEAM_STATS, ROLE_CONCESSIONS, BASELINE). Nessuna formula: solo
// funzioni pure su array di oggetti, pensate per girare in un browser.
// ============================================================================

function buildIndexes(matchLog) {
  // Ordina per data e assegna Player_Seq / PlayerVenue_Seq (mirror delle
  // colonne calcolate di MATCH_LOG nel file originale).
  const rows = [...matchLog].sort((a, b) => a.Data.localeCompare(b.Data));
  const seqAll = {};
  const seqVenue = {};
  for (const r of rows) {
    seqAll[r.Player_ID] = (seqAll[r.Player_ID] || 0) + 1;
    r.Player_Seq = seqAll[r.Player_ID];
    const key = r.Player_ID + '|' + r.Venue;
    seqVenue[key] = (seqVenue[key] || 0) + 1;
    r.PlayerVenue_Seq = seqVenue[key];
  }
  return rows;
}

function filterRows(matchLog, pred) {
  return matchLog.filter(pred);
}

function sum(rows, col) {
  return rows.reduce((a, r) => a + (Number(r[col]) || 0), 0);
}

function countGE(rows, col, n) {
  return rows.filter((r) => Number(r[col]) >= n).length;
}

// ---------------------------------------------------------------------------
// PLAYERS — statistiche per singolo giocatore, segmento TOTALE/HOME/AWAY
// ---------------------------------------------------------------------------
function computePlayerStats(matchLog, playerId, segmento) {
  let rows = filterRows(matchLog, (r) => r.Player_ID === playerId);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'H' : 'A';
    rows = rows.filter((r) => r.Venue === v);
  }
  const partite = rows.length;
  const minuti = sum(rows, 'Minuti');
  const per90 = (col) => (minuti ? (sum(rows, col) / minuti) * 90 : null);
  const hit = (col, n) => (partite ? countGE(rows, col, n) / partite : null);

  const out = {
    Partite: partite,
    Minuti: minuti,
    Minuti_Medi: partite ? minuti / partite : null,
    Tiri: sum(rows, 'Tiri'),
    Tiri_90: per90('Tiri'),
    SOT: sum(rows, 'Tiri_in_porta'),
    SOT_90: per90('Tiri_in_porta'),
    FC: sum(rows, 'Falli_commessi'),
    FC_90: per90('Falli_commessi'),
    FD: sum(rows, 'Falli_subiti'),
    FD_90: per90('Falli_subiti'),
    Gialli: sum(rows, 'Gialli'),
    Rossi: sum(rows, 'Rossi'),
  };
  for (let n = 1; n <= 4; n++) out['Hit_Tiri_' + n] = hit('Tiri', n);
  for (let n = 1; n <= 3; n++) out['Hit_SOT_' + n] = hit('Tiri_in_porta', n);
  for (let n = 1; n <= 3; n++) out['Hit_FC_' + n] = hit('Falli_commessi', n);
  for (let n = 1; n <= 3; n++) out['Hit_FD_' + n] = hit('Falli_subiti', n);
  out.Cartellini = out.Gialli + out.Rossi;
  out.Cartellini_90 = minuti ? (out.Cartellini / minuti) * 90 : null;
  out.Hit_Cart_1 = partite
    ? (countGE(rows, 'Gialli', 1) +
        rows.filter((r) => Number(r.Gialli) === 0 && Number(r.Rossi) >= 1).length) /
      partite
    : null;

  for (const [tag, N] of [['L5', 5], ['L10', 10]]) {
    const seqCol = segmento === 'TOTALE' ? 'Player_Seq' : 'PlayerVenue_Seq';
    const sub = rows.filter((r) => r[seqCol] > partite - N);
    const subMin = sum(sub, 'Minuti');
    const subN = Math.min(N, partite);
    out[tag + '_N'] = subN;
    out[tag + '_Minuti'] = subMin;
    const p90 = (col) => (subMin ? (sum(sub, col) / subMin) * 90 : null);
    out[tag + '_Tiri_90'] = p90('Tiri');
    out[tag + '_SOT_90'] = p90('Tiri_in_porta');
    out[tag + '_FC_90'] = p90('Falli_commessi');
    out[tag + '_FD_90'] = p90('Falli_subiti');
    out[tag + '_Cart_90'] = subMin
      ? ((sum(sub, 'Gialli') + sum(sub, 'Rossi')) / subMin) * 90
      : null;
  }
  return out;
}

// ---------------------------------------------------------------------------
// TEAM_STATS — squadra x segmento. "Fatti" = diretto sulla propria venue;
// "Subiti" = sulle righe dell'avversario, con venue INVERTITA (verificato
// contro il calendario reale: per le partite in CASA di una squadra, la riga
// dell'avversario nel log riporta Venue="A", perché lì Team_ID è l'avversario
// che gioca in trasferta).
// ---------------------------------------------------------------------------
function teamRowsFatti(matchLog, team, segmento) {
  let rows = filterRows(matchLog, (r) => r.Team_ID === team);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'H' : 'A';
    rows = rows.filter((r) => r.Venue === v);
  }
  return rows;
}
function teamRowsSubiti(matchLog, team, segmento) {
  let rows = filterRows(matchLog, (r) => r.Opp_ID === team);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'A' : 'H'; // inversione verificata
    rows = rows.filter((r) => r.Venue === v);
  }
  return rows;
}

function computeTeamStats(matchLog, team, segmento) {
  const fatti = teamRowsFatti(matchLog, team, segmento);
  const subiti = teamRowsSubiti(matchLog, team, segmento);
  const partite = new Set(fatti.map((r) => r.Match_ID)).size;
  const per90 = (rows, col) => (partite ? sum(rows, col) / partite : null);

  const possessoRows = [];
  const seen = new Set();
  for (const r of fatti) {
    const k = r.Match_ID + '|' + r.Team_ID;
    if (r.Possesso_Team != null && r.Possesso_Team !== '' && !seen.has(k)) {
      seen.add(k);
      possessoRows.push(Number(r.Possesso_Team));
    }
  }

  const out = { Team_ID: team, Segmento: segmento, Partite: partite };
  for (const [label, col] of [
    ['Tiri', 'Tiri'],
    ['SOT', 'Tiri_in_porta'],
    ['Falli', 'Falli_commessi'], // "Falli_Fatti" nel workbook
  ]) {
    out[label + '_Fatti'] = sum(fatti, col);
    out[label + '_Fatti_90'] = per90(fatti, col);
  }
  // Falli_Subiti nel workbook usa la colonna Falli_subiti del log (falli subiti
  // dalla squadra che attacca) sommata sulle proprie righe (Fatti-side)
  out['FalliSubiti_Fatti'] = sum(fatti, 'Falli_subiti');
  out['FalliSubiti_Fatti_90'] = per90(fatti, 'Falli_subiti');

  for (const [label, col] of [
    ['Tiri', 'Tiri'],
    ['SOT', 'Tiri_in_porta'],
  ]) {
    out[label + '_Subiti'] = sum(subiti, col);
    out[label + '_Subiti_90'] = per90(subiti, col);
  }
  // Falli commessi dagli avversari contro questa squadra (=Falli_commessi sulle
  // righe Opp_ID=team, venue invertita) — mappa a "Falli_Subiti" del workbook
  out['Falli_Subiti'] = sum(subiti, 'Falli_commessi');
  out['Falli_Subiti_90'] = per90(subiti, 'Falli_commessi');

  out.Gialli = sum(fatti, 'Gialli');
  out.Rossi = sum(fatti, 'Rossi');
  out.Cartellini = out.Gialli + out.Rossi;
  out.Cartellini_90 = per90(fatti, null) && null; // non usato a valle
  out.Possesso_Medio = possessoRows.length
    ? possessoRows.reduce((a, b) => a + b, 0) / partite
    : null;
  return out;
}

// Rank discendente su un elenco di {Team_ID, value, partite}, valido solo se
// tutte le 20 squadre sono eleggibili (campione minimo) — altrimenti "N/D".
function rankDesc(list, minSample) {
  const eligible = list.filter((x) => x.partite >= minSample && x.value != null);
  const ranked = {};
  if (eligible.length !== 20) {
    for (const x of list) ranked[x.key] = 'N/D';
    return ranked;
  }
  for (const x of list) {
    if (x.partite < minSample || x.value == null) {
      ranked[x.key] = 'N/D';
      continue;
    }
    const higher = eligible.filter((y) => y.value > x.value).length;
    ranked[x.key] = 1 + higher;
  }
  return ranked;
}

// Rank PROVVISORIO: come rankDesc, ma calcolato solo tra le squadre che hanno
// già raggiunto il campione minimo, qualunque sia il loro numero (non blocca
// tutto a "N/D" se non sono ancora 20/20). Va sempre mostrato con l'etichetta
// "X eleggibili su 20", perché prima di fine stagione può cambiare parecchio
// partita dopo partita — non è il rank ufficiale, è una stima di comodo.
function rankProvisional(list, minSample) {
  const eligible = list.filter((x) => x.partite >= minSample && x.value != null);
  const out = {};
  for (const x of list) {
    if (x.partite < minSample || x.value == null) {
      out[x.key] = { rank: 'N/D', eligible: eligible.length };
      continue;
    }
    const higher = eligible.filter((y) => y.value > x.value).length;
    out[x.key] = { rank: 1 + higher, eligible: eligible.length };
  }
  return out;
}

// ---------------------------------------------------------------------------
// ROLE_CONCESSIONS — squadra x segmento x ruolo: cosa concede quella squadra
// ai giocatori che occupano un certo ruolo tattico (venue invertita, come le
// colonne "Subiti" di TEAM_STATS).
// ---------------------------------------------------------------------------
function computeRoleConcessions(matchLog, team, segmento, ruolo, teamPartite) {
  let rows = filterRows(matchLog, (r) => r.Opp_ID === team && r.Posizione_Analitica === ruolo);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'A' : 'H';
    rows = rows.filter((r) => r.Venue === v);
  }
  const nPrestazioni = rows.length;
  const per90 = (col) => (teamPartite ? sum(rows, col) / teamPartite : null);
  return {
    Team_ID: team,
    Segmento: segmento,
    Posizione_Analitica: ruolo,
    N_Prestazioni_Ruolo: nPrestazioni,
    Tiri_Concessi_90: per90('Tiri'),
    SOT_Concessi_90: per90('Tiri_in_porta'),
    Falli_Contro_Ruolo_90: per90('Falli_commessi'),
  };
}

// ---------------------------------------------------------------------------
// BASELINE — media pesata Stagione 40% / Venue 30% / L10 20% / L5 10%,
// normalizzata sui soli pesi delle componenti disponibili.
// ---------------------------------------------------------------------------
const BASELINE_WEIGHTS = { stag: 40, venue: 30, l10: 20, l5: 10 };
const BASELINE_THRESHOLDS = { minMinutiStagione: 90, minNVenue: 2, minL5: 2, minL10: 3 };

function weightedBaseline(stag90, venue90, l1090, l590, weights) {
  let num = 0;
  let den = 0;
  if (stag90 != null) { num += stag90 * weights.stag; den += weights.stag; }
  if (venue90 != null) { num += venue90 * weights.venue; den += weights.venue; }
  if (l1090 != null) { num += l1090 * weights.l10; den += weights.l10; }
  if (l590 != null) { num += l590 * weights.l5; den += weights.l5; }
  return den ? num / den : null;
}

function computeBaseline(totale, home, away) {
  // totale/home/away sono gli oggetti restituiti da computePlayerStats
  const statoHome = (() => {
    if (totale.Partite === 0) return 'NO DATI STAGIONE';
    if (home.Partite === 0) return 'NO HOME';
    if (
      totale.Minuti < BASELINE_THRESHOLDS.minMinutiStagione ||
      home.Partite < BASELINE_THRESHOLDS.minNVenue ||
      totale.L5_N < BASELINE_THRESHOLDS.minL5 ||
      totale.L10_N < BASELINE_THRESHOLDS.minL10
    )
      return 'CAMPIONE RIDOTTO';
    return 'OK';
  })();
  const statoAway = (() => {
    if (totale.Partite === 0) return 'NO DATI STAGIONE';
    if (away.Partite === 0) return 'NO AWAY';
    if (
      totale.Minuti < BASELINE_THRESHOLDS.minMinutiStagione ||
      away.Partite < BASELINE_THRESHOLDS.minNVenue ||
      totale.L5_N < BASELINE_THRESHOLDS.minL5 ||
      totale.L10_N < BASELINE_THRESHOLDS.minL10
    )
      return 'CAMPIONE RIDOTTO';
    return 'OK';
  })();

  const metrics = ['Tiri', 'SOT', 'FC', 'FD', 'Cart'];
  const out = { Stato_HOME: statoHome, Stato_AWAY: statoAway };
  for (const m of metrics) {
    const stag90 = totale[m + '_90'];
    const home90 = home[m + '_90'];
    const away90 = away[m + '_90'];
    const l1090 = totale['L10_' + m + '_90'];
    const l590 = totale['L5_' + m + '_90'];
    out[m + '_BASE_HOME'] =
      totale.Partite && home.Partite ? weightedBaseline(stag90, home90, l1090, l590, BASELINE_WEIGHTS) : null;
    out[m + '_BASE_AWAY'] =
      totale.Partite && away.Partite ? weightedBaseline(stag90, away90, l1090, l590, BASELINE_WEIGHTS) : null;
  }
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = {
    buildIndexes,
    computePlayerStats,
    computeTeamStats,
    rankDesc,
    computeRoleConcessions,
    computeBaseline,
    rankProvisional,
  };
}
