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
// PLAYERS — statistiche per singolo giocatore, segmento TOTALE/HOME/AWAY.
// titolareFiltro: null/undefined = tutte le presenze (default, comportamento
// invariato rispetto a prima); 'TITOLARE' = solo presenze da titolare;
// 'SUBENTRATO' = solo presenze da subentrato. Filtro applicato PRIMA del
// segmento Venue, quindi combinabile liberamente con TOTALE/HOME/AWAY.
// ---------------------------------------------------------------------------
function computePlayerStats(matchLog, playerId, segmento, titolareFiltro) {
  let rows = filterRows(matchLog, (r) => r.Player_ID === playerId);
  if (titolareFiltro === 'TITOLARE') rows = rows.filter((r) => r.Titolare === 'Sì');
  if (titolareFiltro === 'SUBENTRATO') rows = rows.filter((r) => r.Titolare === 'No');
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

  // Consistenza (coefficiente di variazione): quanto oscilla la prestazione
  // partita per partita, a parità di media. CV = deviazione standard / media,
  // calcolato sul conteggio per partita (non /90) — richiede almeno 2 partite,
  // altrimenti non è misurabile e resta null.
  function cvOf(extract) {
    if (partite < 2) return { mean: partite ? extract(rows[0]) : null, sd: null, cv: null };
    const vals = rows.map(extract);
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / (vals.length - 1);
    const sd = Math.sqrt(variance);
    return { mean, sd, cv: mean ? sd / mean : null };
  }
  const cvTiri = cvOf((r) => Number(r.Tiri) || 0);
  const cvSOT = cvOf((r) => Number(r.Tiri_in_porta) || 0);
  const cvFC = cvOf((r) => Number(r.Falli_commessi) || 0);
  const cvFD = cvOf((r) => Number(r.Falli_subiti) || 0);
  const cvCart = cvOf((r) => (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0));
  out.Tiri_CV = cvTiri.cv; out.Tiri_SD = cvTiri.sd;
  out.SOT_CV = cvSOT.cv; out.SOT_SD = cvSOT.sd;
  out.FC_CV = cvFC.cv; out.FC_SD = cvFC.sd;
  out.FD_CV = cvFD.cv; out.FD_SD = cvFD.sd;
  out.Cart_CV = cvCart.cv; out.Cart_SD = cvCart.sd;

  for (const [tag, N] of [['L3', 3], ['L5', 5], ['L10', 10]]) {
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
const SHRINK_K = 3; // a n=3 il peso della componente è dimezzato; a n=1 è un quarto; a n=10 è ~77%.

// Shrinkage bayesiano: il peso "di diritto" di ogni componente (venue, L10, L5)
// viene ridotto in proporzione a quanti dati ha davvero dietro (n/(n+k)).
// Il peso "perso" da una componente con campione piccolo non sparisce: viene
// spostato sulla Stagione, che è la componente col campione più ampio e serve
// da riferimento più stabile quando le altre sono ancora deboli.
function weightedBaselineShrunk(stag90, venue90, l1090, l590, nVenue, nL10, nL5, weights) {
  const shrink = (w, n) => w * (n / (n + SHRINK_K));
  const wVenueEff = venue90 != null ? shrink(weights.venue, nVenue) : 0;
  const wL10Eff = l1090 != null ? shrink(weights.l10, nL10) : 0;
  const wL5Eff = l590 != null ? shrink(weights.l5, nL5) : 0;
  const lost =
    (venue90 != null ? weights.venue - wVenueEff : 0) +
    (l1090 != null ? weights.l10 - wL10Eff : 0) +
    (l590 != null ? weights.l5 - wL5Eff : 0);
  const wStagEff = stag90 != null ? weights.stag + lost : 0;

  let num = 0;
  let den = 0;
  if (stag90 != null) { num += stag90 * wStagEff; den += wStagEff; }
  if (venue90 != null) { num += venue90 * wVenueEff; den += wVenueEff; }
  if (l1090 != null) { num += l1090 * wL10Eff; den += wL10Eff; }
  if (l590 != null) { num += l590 * wL5Eff; den += wL5Eff; }
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
      totale.Partite && home.Partite
        ? weightedBaselineShrunk(stag90, home90, l1090, l590, home.Partite, totale.L10_N, totale.L5_N, BASELINE_WEIGHTS)
        : null;
    out[m + '_BASE_AWAY'] =
      totale.Partite && away.Partite
        ? weightedBaselineShrunk(stag90, away90, l1090, l590, away.Partite, totale.L10_N, totale.L5_N, BASELINE_WEIGHTS)
        : null;
  }
  return out;
}

// Intervallo di confidenza di Wilson (95%) per un hit rate su campione n.
// Più corretto del normale su campioni piccoli (mai limiti assurdi tipo
// negativi o sopra 100%). Ritorna [lower, upper] in frazione (0-1).
function wilsonInterval(phat, n) {
  if (!n || phat == null) return [null, null];
  const z = 1.96;
  const denom = 1 + (z * z) / n;
  const center = phat + (z * z) / (2 * n);
  const adj = z * Math.sqrt((phat * (1 - phat)) / n + (z * z) / (4 * n * n));
  return [(center - adj) / denom, (center + adj) / denom];
}

// ---------------------------------------------------------------------------
// ARBITRI — statistiche aggregate per arbitro: falli fischiati e cartellini
// mostrati per partita diretta, con split "squadra di casa" vs "squadra in
// trasferta" (bias arbitrale casa/trasferta), calcolati sui dati grezzi di
// MATCH_LOG. NESSUNA soglia minima blocca il calcolo — a differenza dei rank
// squadra/ruolo — ma il numero di partite dirette (N) va SEMPRE mostrato
// accanto al valore: con 1-2 partite dirette il numero non è una media
// affidabile, è il dato grezzo di poche gare travestito da statistica.
// ---------------------------------------------------------------------------
function computeArbitroStats(matchLog, arbitri, arbitro) {
  const matchIds = arbitri.filter((a) => a.Arbitro === arbitro).map((a) => a.Match_ID);
  const nPartite = matchIds.length;
  const rowsCasa = filterRows(matchLog, (r) => matchIds.includes(r.Match_ID) && r.Venue === 'H');
  const rowsTrasferta = filterRows(matchLog, (r) => matchIds.includes(r.Match_ID) && r.Venue === 'A');

  const avg = (rows, col, n) => (n ? sum(rows, col) / n : null);
  const cartellini = (rows) => sum(rows, 'Gialli') + sum(rows, 'Rossi');

  return {
    Arbitro: arbitro,
    N_Partite: nPartite,
    Falli_Casa_Medi: avg(rowsCasa, 'Falli_commessi', nPartite),
    Falli_Trasferta_Medi: avg(rowsTrasferta, 'Falli_commessi', nPartite),
    Falli_Totali_Medi: nPartite
      ? (sum(rowsCasa, 'Falli_commessi') + sum(rowsTrasferta, 'Falli_commessi')) / nPartite
      : null,
    Cartellini_Casa_Medi: nPartite ? cartellini(rowsCasa) / nPartite : null,
    Cartellini_Trasferta_Medi: nPartite ? cartellini(rowsTrasferta) / nPartite : null,
    Cartellini_Totali_Medi: nPartite ? (cartellini(rowsCasa) + cartellini(rowsTrasferta)) / nPartite : null,
  };
}

// Statistiche per tutti gli arbitri presenti nel foglio ARBITRI, una riga per
// arbitro. Utile per popolare una tabella "Statistiche arbitri" nello Scanner.
function computeAllArbitroStats(matchLog, arbitri) {
  const nomi = [...new Set(arbitri.map((a) => a.Arbitro))];
  return nomi.map((nome) => computeArbitroStats(matchLog, arbitri, nome));
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
    wilsonInterval,
    computeArbitroStats,
    computeAllArbitroStats,
  };
}
