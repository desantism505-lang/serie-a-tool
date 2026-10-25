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
// titolareFiltro: null/undefined = tutte le presenze; 'TITOLARE'/'SUBENTRATO'
// come sopra. teamFiltro: se passato (un Team_ID), limita le righe a quelle
// giocate CON QUELLA squadra — serve per i giocatori trasferiti a stagione in
// corso: senza questo filtro, le partite con il club precedente si
// mescolerebbero silenziosamente in "Tiri stagione" del club attuale. Se
// null/undefined, comportamento invariato (tutte le presenze, qualunque
// squadra) — utile per uno storico complessivo esplicito.
function computePlayerStats(matchLog, playerId, segmento, titolareFiltro, teamFiltro) {
  let rows = filterRows(matchLog, (r) => r.Player_ID === playerId);
  if (teamFiltro) rows = rows.filter((r) => r.Team_ID === teamFiltro);
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
  out.Cartellini_90 = partite ? out.Cartellini / partite : null;
  out.Possesso_Medio = possessoRows.length
    ? possessoRows.reduce((a, b) => a + b, 0) / possessoRows.length // divide per le sole partite col dato
    : null;
  out.Possesso_N = possessoRows.length;
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
// colonne "Subiti" di TEAM_STATS). Le righe filtrate sono quelle DEGLI
// AVVERSARI di quel ruolo, quindi:
// - Falli_Contro_Ruolo_90 = falli che il ruolo COMMETTE contro questa squadra
//   (colonna Falli_commessi sulle loro righe — aggressività del ruolo)
// - FalliSubiti_Ruolo_90 = falli che il ruolo SUBISCE da questa squadra
//   (colonna Falli_subiti sulle loro righe — quanto la squadra fallisce su
//   quel ruolo, il concetto opposto al precedente)
// - Cartellini_Ruolo_90 = cartellini (gialli+rossi) ricevuti dal ruolo
//   quando affronta questa squadra
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
    FalliSubiti_Ruolo_90: per90('Falli_subiti'),
    Cartellini_Ruolo_90: teamPartite ? (sum(rows, 'Gialli') + sum(rows, 'Rossi')) / teamPartite : null,
  };
}

// ---------------------------------------------------------------------------
// ROLE_PRODUCTION — speculare di ROLE_CONCESSIONS: cosa produce la squadra
// coi PROPRI giocatori di un certo ruolo tattico (venue DIRETTA, non
// invertita — qui Team_ID è la squadra stessa, non l'avversario come in
// ROLE_CONCESSIONS). Serve per l'incrocio "chi produce / chi concede" tra le
// due squadre di una partita (vedi computeRoleCrossExtremes più sotto):
// prima si guardava solo la concessione (un lato), non se l'avversario di
// oggi produce davvero in quel ruolo rispetto al resto della lega.
// ---------------------------------------------------------------------------
function computeRoleProduction(matchLog, team, segmento, ruolo, teamPartite) {
  let rows = filterRows(matchLog, (r) => r.Team_ID === team && r.Posizione_Analitica === ruolo);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'H' : 'A';
    rows = rows.filter((r) => r.Venue === v);
  }
  const nPrestazioni = rows.length;
  const per90 = (col) => (teamPartite ? sum(rows, col) / teamPartite : null);
  return {
    Team_ID: team,
    Segmento: segmento,
    Posizione_Analitica: ruolo,
    N_Prestazioni_Ruolo: nPrestazioni,
    Tiri_Prodotti_90: per90('Tiri'),
    SOT_Prodotti_90: per90('Tiri_in_porta'),
    Falli_Dal_Ruolo_90: per90('Falli_commessi'),
  };
}

// ---------------------------------------------------------------------------
// SCOMPOSIZIONE PER LATO (Lato_Campo) — richiesta Marco, 25/10. NON è un
// secondo rank: il campione per-lato (dimezzato/triplicato) non reggerebbe
// un confronto a 20 squadre, farebbe lo stesso errore di pareggio di massa
// già corretto altrove. È solo la scomposizione in numeri grezzi del totale
// che già sostiene un fatto "Matchup ruolo" esistente (computeRoleConcessions
// / fattiRuoloVsAvversario) — un arricchimento del fatto, non un fatto nuovo.
// Applicabile ai ruoli strutturalmente larghi (sempre) e a Centrocampista
// centrale (solo nelle partite-squadra con centrocampo a 3 titolari, vedi
// sotto) — per tutti gli altri ruoli centrali Lato_Campo resta un artefatto
// di dove si è trovato il giocatore in quel momento, non va mai mostrato.
// Ritorna null (il chiamante allora non mostra nulla, mai un N/D forzato) se:
// ruolo non idoneo, totale sotto soglia minima, nessun lato chiaramente
// dominante, o il dataset non ha affatto Lato_Campo valorizzato (es. UCL
// finché non avrà quella colonna — il motore resta identico tra le repo, la
// funzione si attiva da sola quando i dati arrivano).
// Ruoli strutturalmente larghi (sx/dx sempre presente, qualunque modulo):
// qui il lato è sempre potenzialmente informativo, nessuna condizione extra.
const WIDE_ROLES_LATO = ['Esterno offensivo', 'Terzino', 'Esterno/Quinto'];
// Ruoli "centrali condizionati": aggiunto 25/10/2026 su richiesta Marco. Qui
// il lato NON è strutturale come sopra — è reale solo nelle partite in cui
// quella specifica squadra ha schierato 3 titolari in quel ruolo (modulo a
// 3): centrocampo a 3 per Centrocampista centrale (verifica sui dati: Inter
// 7/7 partite con 3 CC titolari, Lato_Campo valorizzato in modo reale),
// difesa a 3 per Centrale difensivo (verifica: 69 partite-squadra su 15
// squadre con 3 CB titolari, Lato_Campo pieno su quelle righe — nessuna riga
// "altro"). Con un numero diverso di titolari in quel ruolo (centrocampo a
// 4, difesa a 4) Lato_Campo torna ad essere un artefatto del momento come
// per ogni altro ruolo centrale — quelle righe vanno escluse dal conteggio,
// non trattate come se il lato fosse stabile.
const CONDITIONAL_TRIPLE_ROLES = ['Centrocampista centrale', 'Centrale difensivo'];
const LATO_MIN_SAMPLE = 5; // sotto questa soglia, scomporre il totale è rumore, non informazione
// Soglia di dominanza — aggiunta 25/10/2026 dopo segnalazione di Marco: un
// fatto reale (Frosinone esterni offensivi, 6 tiri in porta: 3 sinistra/3
// destra/0 centro) mostrava una scomposizione perfettamente in pareggio,
// zero segnale su quale lato attaccare — lo stesso errore di principio
// della regola anti-pareggio sui Fatti salienti (MAX_TIE_FOR_FACT), qui non
// coperto perché quella regola agisce sul confronto fra squadre, non sulla
// scomposizione interna di un singolo totale. Senza un lato chiaramente
// dominante, la scomposizione non è un'informazione azionabile: va omessa,
// non mostrata con un pareggio che il lettore deve scoprire da solo.
const LATO_SKEW_MIN = 0.6; // il lato dominante deve coprire almeno il 60% del totale
// Cache (per riferimento dell'array matchLog, invariato per tutta la durata
// di una sessione nel browser): per ciascun ruolo condizionato, set di
// chiavi Team_ID|Match_ID in cui quella squadra ha schierato esattamente 3
// titolari in quel ruolo. Costruita una volta sola per ruolo, non ad ogni
// chiamata — richiamata potenzialmente decine di volte per scansione
// (Scanner su una giornata intera).
const _tripleRoleCache = new WeakMap(); // matchLog -> Map(ruolo -> Set)
function getTripleRoleSet(matchLog, ruolo) {
  let byRuolo = _tripleRoleCache.get(matchLog);
  if (!byRuolo) { byRuolo = new Map(); _tripleRoleCache.set(matchLog, byRuolo); }
  let set = byRuolo.get(ruolo);
  if (set) return set;
  const counts = {};
  for (const r of matchLog) {
    if (r.Posizione_Analitica !== ruolo || r.Titolare !== 'Sì') continue;
    const key = r.Team_ID + '|' + r.Match_ID;
    counts[key] = (counts[key] || 0) + 1;
  }
  set = new Set();
  for (const key in counts) {
    if (counts[key] === 3) set.add(key);
  }
  byRuolo.set(ruolo, set);
  return set;
}
function roleConcessionsSideBreakdown(matchLog, team, segmento, ruolo, rawCol) {
  if (!WIDE_ROLES_LATO.includes(ruolo) && !CONDITIONAL_TRIPLE_ROLES.includes(ruolo)) return null;
  let rows = filterRows(matchLog, (r) => r.Opp_ID === team && r.Posizione_Analitica === ruolo);
  if (segmento !== 'TOTALE') {
    const v = segmento === 'HOME' ? 'A' : 'H';
    rows = rows.filter((r) => r.Venue === v);
  }
  if (CONDITIONAL_TRIPLE_ROLES.includes(ruolo)) {
    const tripleSet = getTripleRoleSet(matchLog, ruolo);
    rows = rows.filter((r) => tripleSet.has(r.Team_ID + '|' + r.Match_ID));
  }
  const hasLato = rows.some((r) => r.Lato_Campo === 'Sinistra' || r.Lato_Campo === 'Destra' || r.Lato_Campo === 'Centro');
  if (!hasLato) return null;
  const by = { Sinistra: 0, Destra: 0, Centro: 0 };
  let totale = 0;
  for (const r of rows) {
    const v = Number(r[rawCol]) || 0;
    totale += v;
    if (by[r.Lato_Campo] != null) by[r.Lato_Campo] += v;
  }
  if (totale < LATO_MIN_SAMPLE) return null;
  const dominante = Math.max(by.Sinistra, by.Destra, by.Centro);
  if (dominante / totale < LATO_SKEW_MIN) return null; // nessun lato chiaramente prevalente: niente da segnalare
  return { totale, sinistra: by.Sinistra, destra: by.Destra, centro: by.Centro };
}

// ---------------------------------------------------------------------------
// LATO PREVALENTE DEL GIOCATORE — richiesta Marco, 25/10. Diverso da
// roleConcessionsSideBreakdown sopra: qui non scomponiamo un totale di
// squadra, cerchiamo UN lato stabile per UN giocatore, da mostrare come
// etichetta (es. nello Scanner, sotto il nome). Conta le PRESENZE per lato
// (non eventi/tiri come sopra), perché la domanda è "dove gioca di solito",
// non "quanto produce".
// Soglie: minimo 3 presenze valorizzate, lato dominante almeno 70% — sotto
// soglia il giocatore non ha un lato stabile (ruota, o campione insufficiente)
// e la funzione ritorna null: nessuna etichetta forzata, mai un "misto"
// spacciato per dato.
const PLAYER_SIDE_MIN_SAMPLE = 3;
const PLAYER_SIDE_MIN_SHARE = 0.7;
function playerPrevalentSide(matchLog, playerId, ruolo, teamId) {
  if (!WIDE_ROLES_LATO.includes(ruolo) && !CONDITIONAL_TRIPLE_ROLES.includes(ruolo)) return null;
  let rows = filterRows(matchLog, (r) => r.Player_ID === playerId && r.Posizione_Analitica === ruolo);
  if (teamId) rows = rows.filter((r) => r.Team_ID === teamId);
  if (CONDITIONAL_TRIPLE_ROLES.includes(ruolo)) {
    const tripleSet = getTripleRoleSet(matchLog, ruolo);
    rows = rows.filter((r) => tripleSet.has(r.Team_ID + '|' + r.Match_ID));
  }
  rows = rows.filter((r) => r.Lato_Campo === 'Sinistra' || r.Lato_Campo === 'Destra' || r.Lato_Campo === 'Centro');
  if (rows.length < PLAYER_SIDE_MIN_SAMPLE) return null;
  const counts = { Sinistra: 0, Destra: 0, Centro: 0 };
  for (const r of rows) counts[r.Lato_Campo]++;
  const totale = rows.length;
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const [ladoTop, nTop] = entries[0];
  if (nTop / totale < PLAYER_SIDE_MIN_SHARE) return null;
  return ladoTop; // 'Sinistra' | 'Destra' | 'Centro'
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
// STATISTICHE SQUADRA (pannello "Statistiche squadra" di index.html): totali
// per singola partita (non per90) di Tiri/SOT/Falli/Cartellini/Corner, usati
// per hit-rate su linea libera, confronto diretto (duel) e proiezione
// moltiplicativa. Nessuna di queste richiede il risultato della partita.
//
// NOTA su Corner: a differenza di Tiri/SOT/Falli/Cartellini (che sono per
// giocatore, in MATCH_LOG, e vanno sommati riga per riga), Corner è già un
// totale-partita in CALENDAR (Corner_Home/Corner_Away) — non esiste a
// livello di singolo giocatore. Per questo le funzioni sotto accettano un
// parametro `calendar` opzionale: se non passato, si comportano come prima
// (solo Tiri/SOT/Falli/Cart); se passato, aggiungono anche Corner.
// ---------------------------------------------------------------------------

// Corner "fatti" e "concessi" di una squadra, partita per partita — solo
// sulle CONCLUSA con Corner_Home/Corner_Away compilati (nessun dato inventato
// sulle partite dove manca).
function teamCornerMatches(calendar, teamId, segmento) {
  let rows = calendar.filter(
    (c) =>
      (c.Home_ID === teamId || c.Away_ID === teamId) &&
      c.Stato === 'CONCLUSA' &&
      c.Corner_Home != null &&
      c.Corner_Away != null
  );
  if (segmento !== 'TOTALE') {
    rows = rows.filter((c) => (segmento === 'HOME' ? c.Home_ID === teamId : c.Away_ID === teamId));
  }
  return rows.map((c) => {
    const isHome = c.Home_ID === teamId;
    return {
      Match_ID: c.Match_ID,
      Corner: isHome ? c.Corner_Home : c.Corner_Away,
      CornerConcessi: isHome ? c.Corner_Away : c.Corner_Home,
    };
  });
}

// Totali per-partita di UNA squadra nel segmento richiesto: un oggetto per
// Match_ID con la somma dei valori di tutti i giocatori di quella squadra in
// quella partita, più Corner se viene passato `calendar`.
function computeTeamMatchTotals(matchLog, teamId, segmento, calendar) {
  const rows = teamRowsFatti(matchLog, teamId, segmento);
  const byMatch = {};
  for (const r of rows) {
    const m = byMatch[r.Match_ID] || (byMatch[r.Match_ID] = { Match_ID: r.Match_ID, Tiri: 0, SOT: 0, Falli: 0, Cart: 0 });
    m.Tiri += Number(r.Tiri) || 0;
    m.SOT += Number(r.Tiri_in_porta) || 0;
    m.Falli += Number(r.Falli_commessi) || 0;
    m.Cart += (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0);
  }
  const result = Object.values(byMatch);
  if (calendar) {
    const cornerByMatch = {};
    for (const c of teamCornerMatches(calendar, teamId, segmento)) cornerByMatch[c.Match_ID] = c.Corner;
    for (const m of result) m.Corner = cornerByMatch[m.Match_ID] != null ? cornerByMatch[m.Match_ID] : null;
  }
  return result;
}

// Hit rate: quante partite (su quelle con un valore valido per la metrica) la
// squadra ha superato (Over) o non raggiunto (Under) una linea libera. Salta
// le partite dove la metrica è null (es. Corner mancante) invece di contarle
// come "non superata". mode: 'over' (default) o 'under'.
function teamHitRate(totals, key, line, mode) {
  if (line == null || !Number.isFinite(line)) return { n: 0, hits: 0, rate: null };
  const valid = totals.filter((t) => t[key] != null);
  const n = valid.length;
  const hits = mode === 'under'
    ? valid.filter((t) => t[key] < line).length
    : valid.filter((t) => t[key] > line).length;
  // pushes = partite esattamente sulla linea (possibili solo con linee intere):
  // Over + Under + Push = n. Con linee .5 pushes e' sempre 0.
  const pushes = valid.filter((t) => t[key] === line).length;
  return { n, hits, pushes, rate: n ? hits / n : null };
}

// Media di lega di una metrica, per-partita, sul segmento richiesto — media
// dei totali-partita di TUTTE le squadre (ogni partita conta una volta per
// ciascuna prospettiva attacco, coerente con una "media attacco" di lega).
function leagueAverageMetric(matchLog, teamsAll, segmento, key, calendar) {
  let s = 0, n = 0;
  for (const t of teamsAll) {
    for (const m of computeTeamMatchTotals(matchLog, t, segmento, calendar)) {
      if (m[key] != null) { s += m[key]; n++; }
    }
  }
  return n ? s / n : null;
}

// Duel totals: per ogni partita della squadra nel segmento, il proprio totale
// e quello dell'avversario nella STESSA partita (per il "confronto diretto").
// Per Corner, il "concesso" dall'avversario è già disponibile in CALENDAR
// (CornerConcessi) — non richiede somma su MATCH_LOG.
function computeTeamDuelTotals(matchLog, teamId, segmento, calendar) {
  const ownTotals = computeTeamMatchTotals(matchLog, teamId, segmento, calendar);
  const oppRows = teamRowsSubiti(matchLog, teamId, segmento); // righe avversario, venue già invertita
  const oppByMatch = {};
  for (const r of oppRows) {
    const m = oppByMatch[r.Match_ID] || (oppByMatch[r.Match_ID] = { Tiri: 0, SOT: 0, Falli: 0, Cart: 0 });
    m.Tiri += Number(r.Tiri) || 0;
    m.SOT += Number(r.Tiri_in_porta) || 0;
    m.Falli += Number(r.Falli_commessi) || 0;
    m.Cart += (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0);
  }
  if (calendar) {
    for (const c of teamCornerMatches(calendar, teamId, segmento)) {
      const m = oppByMatch[c.Match_ID] || (oppByMatch[c.Match_ID] = { Tiri: 0, SOT: 0, Falli: 0, Cart: 0 });
      m.Corner = c.CornerConcessi;
    }
  }
  return ownTotals
    .filter((own) => oppByMatch[own.Match_ID])
    .map((own) => ({ Match_ID: own.Match_ID, own, opp: oppByMatch[own.Match_ID] }));
}

// Record duello su una metrica: quante volte il proprio totale ha superato
// quello dell'avversario nella stessa partita (salta le partite senza un
// valore valido per entrambi i lati).
function teamDuelRecord(duels, key) {
  const valid = duels.filter((d) => d.own[key] != null && d.opp[key] != null);
  const n = valid.length;
  const wins = valid.filter((d) => d.own[key] > d.opp[key]).length;
  return { n, wins, rate: n ? wins / n : null };
}

// Proiezione moltiplicativa: media di lega × forza attacco (proprio totale
// medio / media lega, con shrinkage) × debolezza difesa avversario (concesso
// medio dall'avversario / media lega, con shrinkage). Shrinkage verso 1.0 sui
// campioni piccoli, stesso k della Baseline (SHRINK_K=3) per coerenza.
function multiplicativeProjection(matchLog, teamId, venue, oppTeamId, oppVenue, key, leagueAvg, calendar) {
  const ownTotals = computeTeamMatchTotals(matchLog, teamId, venue, calendar)
    .map((m) => m[key])
    .filter((v) => v != null);
  const nA = ownTotals.length;
  const ownAvgA = nA ? ownTotals.reduce((a, b) => a + b, 0) / nA : null;

  let concedeVals;
  if (key === 'Corner' && calendar) {
    concedeVals = teamCornerMatches(calendar, oppTeamId, oppVenue)
      .map((c) => c.CornerConcessi)
      .filter((v) => v != null);
  } else {
    const concedeRows = teamRowsSubiti(matchLog, oppTeamId, oppVenue);
    const byMatch = {};
    for (const r of concedeRows) {
      const m = byMatch[r.Match_ID] || (byMatch[r.Match_ID] = { Tiri: 0, SOT: 0, Falli: 0, Cart: 0 });
      m.Tiri += Number(r.Tiri) || 0;
      m.SOT += Number(r.Tiri_in_porta) || 0;
      m.Falli += Number(r.Falli_commessi) || 0;
      m.Cart += (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0);
    }
    concedeVals = Object.values(byMatch).map((m) => m[key]);
  }
  const nB = concedeVals.length;
  const concedeAvgB = nB ? concedeVals.reduce((a, b) => a + b, 0) / nB : null;

  if (leagueAvg == null) return { ownAvgA, nA, concedeAvgB, nB, projection: null };
  const shrink = (avg, n) => (avg == null || !leagueAvg ? 1 : 1 + (avg / leagueAvg - 1) * (n / (n + SHRINK_K)));
  const projection = leagueAvg * shrink(ownAvgA, nA) * shrink(concedeAvgB, nB);
  return { ownAvgA, nA, concedeAvgB, nB, projection };
}

// Giorni di riposo prima della partita indicata: differenza in giorni tra la
// Data della partita e la Data della partita precedente della stessa
// squadra (qualunque venue) — null se non c'è una partita precedente.
function restDaysBefore(calendar, teamId, dataPartita) {
  const precedenti = calendar.filter(
    (r) => (r.Home_ID === teamId || r.Away_ID === teamId) && r.Data < dataPartita
  );
  if (!precedenti.length) return null;
  precedenti.sort((a, b) => b.Data.localeCompare(a.Data));
  const d1 = new Date(precedenti[0].Data);
  const d2 = new Date(dataPartita);
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

// STREAK basate sul RISULTATO (segna/subisce/vince/perde/pareggia): calcolate
// sulle sole partite CONCLUSA con Home_Score/Away_Score compilati, ordinate
// dalla più recente, contando quante di fila (a partire dall'ultima) rispettano
// la condizione — si ferma alla prima partita che la rompe.
function computeTeamStreak(calendar, teamId, venue, key, n) {
  let rows = calendar.filter(
    (r) =>
      (r.Home_ID === teamId || r.Away_ID === teamId) &&
      r.Stato === 'CONCLUSA' &&
      r.Home_Score != null &&
      r.Away_Score != null
  );
  if (venue !== 'TOTALE') {
    rows = rows.filter((r) => (venue === 'HOME' ? r.Home_ID === teamId : r.Away_ID === teamId));
  }
  rows.sort((a, b) => b.Data.localeCompare(a.Data)); // più recente prima
  const partite = rows.length;

  let streak = 0;
  for (const r of rows) {
    const isHome = r.Home_ID === teamId;
    const golFatti = isHome ? r.Home_Score : r.Away_Score;
    const golSubiti = isHome ? r.Away_Score : r.Home_Score;
    let cond;
    switch (key) {
      case 'segna': cond = golFatti > 0; break;
      case 'nonSegna': cond = golFatti === 0; break;
      case 'subisce': cond = golSubiti > 0; break;
      case 'nonSubisce': cond = golSubiti === 0; break;
      case 'nonVince': cond = golFatti <= golSubiti; break;
      case 'nonPerde': cond = golFatti >= golSubiti; break;
      case 'nonPareggia': cond = golFatti !== golSubiti; break;
      case 'noNGolFatti': cond = golFatti !== n; break; // "non segna ESATTAMENTE n gol da..."
      case 'noNGolSubiti': cond = golSubiti !== n; break;
      case 'noOver25': cond = golFatti + golSubiti <= 2; break; // partita NON sopra i 2.5 gol totali
      case 'noUnder25': cond = golFatti + golSubiti >= 3; break; // partita NON sotto i 2.5 gol totali
      default: cond = false;
    }
    if (cond) streak++;
    else break;
  }
  return { streak, partite };
}

// ---------------------------------------------------------------------------
// LINEA RICORRENTE SUI GOL — richiesto da Marco, 24/10: stesso principio di
// "linea ricorrente" già usato per Tiri/SOT/Falli/Cartellini/Corner in
// lineFacts() di index.html (sempre sopra o sempre sotto una soglia, in
// quel venue, con campione minimo), ma per gol fatti/subiti/totale-partita,
// letti da CALENDAR (Home_Score/Away_Score) invece che da MATCH_LOG.
// Differenza deliberata rispetto a lineFacts(): QUI non c'è il filtro "linea
// vicina alla media di lega" — con i gol le linee di mercato reali (0.5,
// 1.5, 2.5, 3.5) sono già standard e significative di per sé, quel filtro
// serviva solo a scartare soglie banali su metriche senza linee standard.
// minSample di default 3 = stessa soglia minima decisa da Marco per gli
// arbitri, applicata qui per coerenza.
// ---------------------------------------------------------------------------
const GOAL_LINE_MIN_SAMPLE = 3;
function teamGoalMatches(calendar, teamId, venue) {
  let rows = calendar.filter(
    (r) => (r.Home_ID === teamId || r.Away_ID === teamId) && r.Stato === 'CONCLUSA' &&
      r.Home_Score !== '' && r.Home_Score != null && r.Away_Score !== '' && r.Away_Score != null
  );
  if (venue !== 'TOTALE') rows = rows.filter((r) => (venue === 'HOME' ? r.Home_ID === teamId : r.Away_ID === teamId));
  return rows.map((r) => {
    const isHome = r.Home_ID === teamId;
    const fatti = Number(isHome ? r.Home_Score : r.Away_Score);
    const subiti = Number(isHome ? r.Away_Score : r.Home_Score);
    return { fatti, subiti, totale: fatti + subiti };
  });
}
// Testo diretto, comparativo e immediatamente comprensibile — richiesto da
// Marco, 24/10: via le formulazioni burocratiche/ambigue ("ha sempre
// segnato più di 0,5 gol fatti... minimo 1"), dentro linguaggio naturale di
// frequenza ("in tutte le X partite" / "in nessuna delle X partite") con
// l'etichetta Over/Under mantenuta in testa, perché è quella che rende la
// riga immediatamente riconoscibile come mercato di scommessa. Il campione
// (n/n partite) resta sempre visibile: un 3/3 è un dato storico, non una
// garanzia per la prossima partita.
function fmtLinea(x) { return String(x).replace('.', ','); }
function goalLineOverText(col, label, n, venueLabel, lowL, mn) {
  const tag = `Over ${fmtLinea(lowL)} ${label}`;
  if (lowL === 0.5) {
    if (col === 'fatti') return `${tag}: non è mai rimasta a secco ${venueLabel}, in tutte le ${n} partite (min. registrato: ${mn})`;
    if (col === 'subiti') return `${tag}: ha sempre subito almeno un gol ${venueLabel}, in tutte le ${n} partite (min. registrato: ${mn})`;
    return `${tag}: tutte le ${n} partite ${venueLabel} hanno prodotto almeno un gol complessivo (min. registrato: ${mn})`;
  }
  if (col === 'fatti') return `${tag}: ha sempre segnato più di ${fmtLinea(lowL)} gol in tutte le ${n} partite ${venueLabel} (min. registrato: ${mn})`;
  if (col === 'subiti') return `${tag}: ha sempre subito più di ${fmtLinea(lowL)} gol in tutte le ${n} partite ${venueLabel} (min. registrato: ${mn})`;
  return `${tag}: tutte le ${n} partite ${venueLabel} hanno sempre superato ${fmtLinea(lowL)} gol complessivi (min. registrato: ${mn})`;
}
function goalLineUnderText(col, label, n, venueLabel, highL, mx) {
  const tag = `Under ${fmtLinea(highL)} ${label}`;
  const soglia = Math.floor(highL);
  if (col === 'fatti') return `${tag}: non ha mai segnato più di ${soglia} ${soglia === 1 ? 'gol' : 'gol'} in nessuna delle ${n} partite ${venueLabel} (max. registrato: ${mx})`;
  if (col === 'subiti') return `${tag}: non ha mai subito più di ${soglia} gol in nessuna delle ${n} partite ${venueLabel} (max. registrato: ${mx})`;
  return `${tag}: nessuna delle ${n} partite ${venueLabel} ha superato i ${soglia} gol complessivi (max. registrato: ${mx})`;
}
function computeGoalLineFacts(calendar, teamId, venue, venueLabel, minSample) {
  minSample = minSample == null ? GOAL_LINE_MIN_SAMPLE : minSample;
  const matches = teamGoalMatches(calendar, teamId, venue);
  if (matches.length < minSample) return [];
  const facts = [];
  const metrics = [
    ['fatti', 'gol fatti'],
    ['subiti', 'gol subiti'],
    ['totale', 'gol totali'],
  ];
  for (const [col, label] of metrics) {
    const vs = matches.map((m) => m[col]);
    const mn = Math.min(...vs), mx = Math.max(...vs);
    const n = vs.length;
    const lowL = Math.ceil(mn) - 0.5, highL = Math.floor(mx) + 0.5;
    if (mn >= 1) {
      facts.push({ type: 'goalLine', col, over: true, line: lowL, n, min: mn,
        txt: goalLineOverText(col, label, n, venueLabel, lowL, mn) });
    }
    if (mx <= 4) { // sopra 4 la soglia "sempre sotto" diventa poco interessante come linea di mercato
      facts.push({ type: 'goalLine', col, over: false, line: highL, n, max: mx,
        txt: goalLineUnderText(col, label, n, venueLabel, highL, mx) });
    }
  }
  return facts;
}

// Normalizzazione nome arbitro: l'inserimento è manuale, quindi lo stesso
// arbitro può comparire con maiuscole/minuscole diverse ("ROSSI", "Rossi",
// "rossi") — vanno raggruppati come UNA persona, non contati come arbitri
// diversi. normalizeArbitroKey fa il confronto (case/spazi-insensitive),
// titleCaseArbitro sceglie una forma di visualizzazione unica e leggibile.
function normalizeArbitroKey(nome) {
  return String(nome || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/\s*\.\s*/g, '.'); // uniforma "J.L." vs "J. L." vs "J . L ."
}
function titleCaseArbitro(nome) {
  return String(nome || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .split(' ')
    .map((w) => w.replace(/(^|\.)([a-zà-ÿ])/g, (m, sep, ch) => sep + ch.toUpperCase()))
    .join(' ');
}

// Un nome arbitro inserito a mano può contenere errori di battitura/copia —
// in particolare un orario copiato per sbaglio nella colonna sbagliata
// (es. "h. 20.45"). Non è un nome valido: niente sequenza di 2+ lettere, o
// è letteralmente un orario. Le righe non valide vengono escluse dal conteggio
// invece di comparire come un "arbitro" fantasma.
function isArbitroNomeValido(nome) {
  const s = String(nome || '').trim();
  if (!s) return false;
  if (/^h\.?\s*\d{1,2}[.:]\d{2}$/i.test(s)) return false;
  if (/^\d{1,2}[.:]\d{2}$/.test(s)) return false;
  return /[a-zà-ÿ]{2,}/i.test(s);
}

// ---------------------------------------------------------------------------
// CARTELLINO IN RITARDO (pannello "Serie & Ritardi"): un giocatore che fa
// molti falli ma non prende un giallo da un po' — "in ritardo" per
// un'ammonizione. Finestra di ritardo = dall'ultimo cartellino (giallo o
// rosso) in poi, su tutta la stagione (non spezzata per casa/trasferta,
// perché è una striscia continua nel tempo).
// ---------------------------------------------------------------------------
// teamFiltro: se passato, limita la finestra di ritardo alle sole presenze
// con quella squadra (stesso principio già applicato a computePlayerStats —
// un giocatore appena trasferito non deve trascinarsi minuti/falli fatti con
// la maglia precedente dentro la valutazione per la squadra attuale).
function computeCardDrought(matchLog, playerId, teamFiltro) {
  let rows = matchLog.filter((r) => r.Player_ID === playerId);
  if (teamFiltro) rows = rows.filter((r) => r.Team_ID === teamFiltro);
  rows = rows.sort((a, b) => a.Player_Seq - b.Player_Seq);
  let lastCardSeq = 0; // 0 = mai ammonito/espulso in stagione -> finestra = tutte le presenze
  for (const r of rows) {
    if ((Number(r.Gialli) || 0) >= 1 || (Number(r.Rossi) || 0) >= 1) lastCardSeq = r.Player_Seq;
  }
  const finestra = rows.filter((r) => r.Player_Seq > lastCardSeq);
  const minuti = finestra.reduce((a, r) => a + (Number(r.Minuti) || 0), 0);
  const falli = finestra.reduce((a, r) => a + (Number(r.Falli_commessi) || 0), 0);
  const partite = finestra.length;
  const indiceFalli = minuti ? (falli / minuti) * 90 : null;
  return { minuti, falli, partite, indiceFalli };
}

// Relazione reale (non stimata) tra falli commessi IN UNA PARTITA e
// probabilità di essere ammonito/espulso IN QUELLA STESSA PARTITA, calcolata
// su tutte le presenze di MATCH_LOG (tutta la lega). Bucket per numero intero
// di falli, "5+" accorpa i casi rari (poche righe sopra i 5 falli a partita).
function computeFoulCardProbabilityTable(matchLog) {
  const buckets = {};
  for (const r of matchLog) {
    const f = Math.round(Number(r.Falli_commessi) || 0);
    const key = f >= 5 ? '5+' : String(f);
    if (!buckets[key]) buckets[key] = { n: 0, carded: 0 };
    buckets[key].n++;
    if ((Number(r.Gialli) || 0) >= 1 || (Number(r.Rossi) || 0) >= 1) buckets[key].carded++;
  }
  const table = {};
  for (const key in buckets) table[key] = buckets[key].n ? buckets[key].carded / buckets[key].n : null;
  return table;
}

// Stima la probabilità di ammonizione per un giocatore dato il suo indice
// falli/90 nella finestra di ritardo, guardando la riga del bucket più vicino
// nella tabella reale (arrotonda all'intero, "5+" per 5 o più).
function estimateCardProbability(avgFalli, probTable) {
  if (avgFalli == null || !Number.isFinite(avgFalli)) return null;
  const f = Math.max(0, Math.round(avgFalli));
  const key = f >= 5 ? '5+' : String(f);
  if (probTable[key] != null) return probTable[key];
  const disponibili = Object.keys(probTable).filter((k) => probTable[k] != null);
  if (!disponibili.length) return null;
  const numOf = (k) => (k === '5+' ? 5 : Number(k));
  disponibili.sort((a, b) => Math.abs(numOf(a) - f) - Math.abs(numOf(b) - f));
  return probTable[disponibili[0]];
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
function computeArbitroStats(matchLog, arbitri, arbitro, calendar) {
  const targetKey = normalizeArbitroKey(arbitro);
  const matchIds = arbitri.filter((a) => normalizeArbitroKey(a.Arbitro) === targetKey).map((a) => a.Match_ID);
  const nPartite = matchIds.length;
  const rowsCasa = filterRows(matchLog, (r) => matchIds.includes(r.Match_ID) && r.Venue === 'H');
  const rowsTrasferta = filterRows(matchLog, (r) => matchIds.includes(r.Match_ID) && r.Venue === 'A');

  const avg = (rows, col, n) => (n ? sum(rows, col) / n : null);
  const cartellini = (rows) => sum(rows, 'Gialli') + sum(rows, 'Rossi');

  // Gol medi nelle partite dirette: letti dal CALENDAR (Home_Score/Away_Score),
  // non da MATCH_LOG — solo le partite CONCLUSA hanno il punteggio compilato.
  let golMedi = null;
  if (calendar) {
    const cal = calendar.filter((c) => matchIds.includes(c.Match_ID) && c.Stato === 'CONCLUSA' && c.Home_Score !== '' && c.Home_Score != null);
    if (cal.length) golMedi = cal.reduce((a, c) => a + (Number(c.Home_Score) || 0) + (Number(c.Away_Score) || 0), 0) / cal.length;
  }

  return {
    Arbitro: titleCaseArbitro(arbitro),
    N_Partite: nPartite,
    Falli_Casa_Medi: avg(rowsCasa, 'Falli_commessi', nPartite),
    Falli_Trasferta_Medi: avg(rowsTrasferta, 'Falli_commessi', nPartite),
    Falli_Totali_Medi: nPartite
      ? (sum(rowsCasa, 'Falli_commessi') + sum(rowsTrasferta, 'Falli_commessi')) / nPartite
      : null,
    Cartellini_Casa_Medi: nPartite ? cartellini(rowsCasa) / nPartite : null,
    Cartellini_Trasferta_Medi: nPartite ? cartellini(rowsTrasferta) / nPartite : null,
    Cartellini_Totali_Medi: nPartite ? (cartellini(rowsCasa) + cartellini(rowsTrasferta)) / nPartite : null,
    Gol_Medi: golMedi,
  };
}

// Statistiche per tutti gli arbitri presenti nel foglio ARBITRI, una riga per
// arbitro — deduplicati per nome normalizzato (non per stringa esatta), così
// varianti di maiuscole/minuscole dello stesso arbitro non vengono contate
// come persone diverse. Il nome mostrato è la prima occorrenza, in forma
// Title Case.
function computeAllArbitroStats(matchLog, arbitri, calendar) {
  const canonici = new Map(); // chiave normalizzata -> nome da mostrare
  for (const a of arbitri) {
    if (!isArbitroNomeValido(a.Arbitro)) continue;
    const key = normalizeArbitroKey(a.Arbitro);
    if (key && !canonici.has(key)) canonici.set(key, titleCaseArbitro(a.Arbitro));
  }
  return [...canonici.values()].map((nome) => computeArbitroStats(matchLog, arbitri, nome, calendar));
}

// ---------------------------------------------------------------------------
// RANK ARBITRI — richiesto da Marco, 24/10: segnala in Fatti salienti quando
// l'arbitro assegnato a QUESTA partita è tra i primi 3 o gli ultimi 3 arbitri
// della stagione su una metrica, con soglia minima di partite dirette (come
// gli altri rank del tool, mai un rank spacciato per definitivo su 1-2 gare).
// Casa/Trasferta sono per costruzione specifiche a un lato della partita;
// Totali e Gol sono "comuni" (riguardano l'arbitro, non una delle due
// squadre) — il chiamante decide se mostrarle su un lato o su entrambi.
// ---------------------------------------------------------------------------
const ARBITRO_RANK_MIN_PARTITE = 3;
const ARBITRO_RANK_METRICS = [
  ['Falli_Totali_Medi', 'falli fischiati a partita', 'falli', 'comune'],
  ['Falli_Casa_Medi', 'falli fischiati contro la squadra di casa', 'falli', 'casa'],
  ['Falli_Trasferta_Medi', 'falli fischiati contro la squadra in trasferta', 'falli', 'trasferta'],
  ['Cartellini_Totali_Medi', 'cartellini a partita', 'cartellini', 'comune'],
  ['Gol_Medi', 'gol complessivi a partita', 'gol', 'comune'],
];
function rankArbitriFacts(matchLog, arbitri, calendar, minPartite) {
  minPartite = minPartite == null ? ARBITRO_RANK_MIN_PARTITE : minPartite;
  const all = computeAllArbitroStats(matchLog, arbitri, calendar);
  const out = {}; // Arbitro (titleCase) -> [{ col, label, kind, rank, eligible, value, top }]
  for (const [col, label, metricKind, lato] of ARBITRO_RANK_METRICS) {
    const eligible = all.filter((a) => a.N_Partite >= minPartite && a[col] != null);
    if (eligible.length < 6) continue; // stesso minimo "eleggibili" usato altrove in Fatti salienti
    // Media di lega sullo stesso pool di eleggibili — richiesta da Marco,
    // 24/10, per mostrare "media X vs media di lega Y" nel testo del fatto.
    const leagueAvg = eligible.reduce((a, x) => a + x[col], 0) / eligible.length;
    const sorted = [...eligible].sort((a, b) => b[col] - a[col]);
    sorted.forEach((a, idx) => {
      const rank = idx + 1;
      const top = rank <= 3, bottom = rank > sorted.length - 3;
      if (!top && !bottom) return;
      (out[a.Arbitro] || (out[a.Arbitro] = [])).push({
        col, label, metricKind, lato, rank, eligible: sorted.length, value: a[col], top, n: a.N_Partite, leagueAvg,
      });
    });
  }
  return out;
}

// ============================================================================
// MODELLO v2 — proiezione giocatore (parametri scelti con backtest rolling-origin
// su data.json, giornate 2-5 previste con le precedenti; vedi BACKTEST.md).
//
// Idea: stima per minuto "tirata" verso la media di lega del RUOLO, con peso
// proporzionale ai minuti giocati dal giocatore; fattore avversario per ruolo
// molto smorzato; minuti attesi separati dal tasso. Un giocatore NON viene mai
// escluso: se non ha dati vale la media del ruolo (e l'affidabilita' e' "bassa").
// ============================================================================
const MODEL_PARAMS = {
  K_MIN: 600,   // pseudo-minuti del prior di ruolo (~6,7 partite intere). Ottimo ~450-720 nel backtest
  // Fattore avversario a due livelli (entrambi i pesi crescono da soli col campione, nessuna taratura manuale):
  //  - SQUADRA: concessione totale dell'avversario (tutti i ruoli) vs media di lega; peso = n/(n+K_OPP_TEAM), n = partite dell'avversario
  //  - RUOLO: scostamento ulteriore dello stesso ruolo; peso = minuti/(minuti+K_OPP), molto smorzato
  K_OPP_TEAM: 6,   // partite-equivalenti (backtest: ottimo 3-6, ~45% a 5 partite)
  K_OPP: 6000,     // pseudo-minuti del livello ruolo (segnale debole, cresce con la stagione)
  OPP_TEAM_METRICS: ['Tiri', 'SOT', 'FC', 'FD'],
  // sovradispersione extra-Poisson stimata sulle presenze (var = mu + alpha*mu^2)
  ALPHA: { Tiri: 0.18, SOT: 0.16, FC: 0.06, FD: 0.02, Cart: 0 },
  MIN_AFFIDABILITA_ALTA: 450,   // minuti del giocatore (5 partite intere)
  MIN_AFFIDABILITA_MEDIA: 180,  // 2 partite intere
  MIN_HITRATE: 60,              // l'hit rate usa solo presenze con almeno 60'
  // sovradispersione dei TOTALI di squadra (100 squadra-partita; include anche differenze tra squadre, quindi e' un limite alto)
  TEAM_ALPHA: { Tiri: 0.135, SOT: 0.037, Falli: 0.039, Cart: 0.01, Corner: 0.224 },
  K_TEAM: 4,                    // partite-equivalenti per il fattore forza-squadra sul prior (backtest: ottimo 2-8)
  TEAM_FACTOR_METRICS: ['Tiri', 'SOT', 'FC'], // dove migliora la previsione; su FD/Cart no
};
const MODEL_METRICS = ['Tiri', 'SOT', 'FC', 'FD', 'Cart'];
const METRIC_COL = { Tiri: 'Tiri', SOT: 'Tiri_in_porta', FC: 'Falli_commessi', FD: 'Falli_subiti' };
function metricValue(r, m) {
  if (m === 'Cart') return (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0);
  return Number(r[METRIC_COL[m]]) || 0;
}

// Aggregati di lega: per ruolo, per (avversario, ruolo), minuti medi titolari/
// subentrati e rapporto totale-squadra / solo-titolari (per stimare la panchina).
function buildLeagueModel(matchLog) {
  const role = {}, opp = {}, teamTot = {}, teamMatches = {}, oppT = {}, oppMatches = {}, lgAll = { min: 0 };
  const roleV = {}, oppRV = {}, oppTV = {}, oppMatchesV = {}, lgAllV = {}; // stessi aggregati divisi per venue del GIOCATORE (H casa, A trasferta)
  let stMin = 0, stN = 0, sbMin = 0, sbN = 0;
  const shareAll = {}, shareSt = {};
  for (const m of MODEL_METRICS) { shareAll[m] = 0; shareSt[m] = 0; }
  const add = (obj, key, r) => {
    const o = obj[key] || (obj[key] = { min: 0 });
    o.min += Number(r.Minuti) || 0;
    for (const m of MODEL_METRICS) o[m] = (o[m] || 0) + metricValue(r, m);
  };
  for (const r of matchLog) {
    add(role, r.Posizione_Analitica, r);
    add(opp, r.Opp_ID + '|' + r.Posizione_Analitica, r);
    add(oppT, r.Opp_ID, r);
    add(roleV, r.Posizione_Analitica + '|' + r.Venue, r);
    add(oppRV, r.Opp_ID + '|' + r.Posizione_Analitica + '|' + r.Venue, r);
    add(oppTV, r.Opp_ID + '|' + r.Venue, r);
    add(lgAllV, r.Venue, r);
    (oppMatchesV[r.Opp_ID + '|' + r.Venue] || (oppMatchesV[r.Opp_ID + '|' + r.Venue] = new Set())).add(r.Match_ID);
    add({ L: lgAll }, 'L', r);
    (oppMatches[r.Opp_ID] || (oppMatches[r.Opp_ID] = new Set())).add(r.Match_ID);
    const tt = teamTot[r.Team_ID] || (teamTot[r.Team_ID] = {});
    for (const m of MODEL_METRICS) tt[m] = (tt[m] || 0) + metricValue(r, m);
    (teamMatches[r.Team_ID] || (teamMatches[r.Team_ID] = new Set())).add(r.Match_ID);
    if (r.Titolare === 'Sì') { stMin += Number(r.Minuti) || 0; stN++; } else { sbMin += Number(r.Minuti) || 0; sbN++; }
    for (const m of MODEL_METRICS) { const v = metricValue(r, m); shareAll[m] += v; if (r.Titolare === 'Sì') shareSt[m] += v; }
  }
  const benchFactor = {};
  for (const m of MODEL_METRICS) benchFactor[m] = shareSt[m] ? shareAll[m] / shareSt[m] : 1;
  // media di lega per squadra-partita (base del fattore forza-squadra)
  const nTM = Object.values(teamMatches).reduce((a, s2) => a + s2.size, 0);
  const lgPerTeamMatch = {};
  for (const m of MODEL_METRICS) {
    let t = 0; for (const k in teamTot) t += teamTot[k][m] || 0;
    lgPerTeamMatch[m] = nTM ? t / nTM : null;
  }
  return {
    roleV, oppRV, oppTV, lgAllV, oppNV: Object.fromEntries(Object.entries(oppMatchesV).map(([k, v]) => [k, v.size])),
    role, opp, oppT, lgAll, oppN: Object.fromEntries(Object.entries(oppMatches).map(([k, v]) => [k, v.size])),
    teamTot, teamN: Object.fromEntries(Object.entries(teamMatches).map(([k, v]) => [k, v.size])), lgPerTeamMatch,
    startMinMean: stN ? stMin / stN : 88,
    subMinMean: sbN ? sbMin / sbN : 25,
    benchFactor, // moltiplicatore "totale squadra / solo titolari" osservato in lega
  };
}

function rate90Of(agg, m) { return agg && agg.min ? (agg[m] / agg.min) * 90 : null; }

// Righe di un giocatore con la squadra attuale (niente mescolanza di maglie).
function playerRows(matchLog, playerId, teamId) {
  return matchLog.filter((r) => r.Player_ID === playerId && (!teamId || r.Team_ID === teamId));
}

// Tasso/90 stimato: (X + prior*K/90) / ((minuti + K)/90). Con 0 minuti = prior.
function teamStrengthFactor(model, teamId, m) {
  if (!MODEL_PARAMS.TEAM_FACTOR_METRICS.includes(m)) return 1;
  const n = model.teamN[teamId], lg = model.lgPerTeamMatch[m];
  if (!n || !lg) return 1;
  const ratio = model.teamTot[teamId][m] / n / lg;
  return 1 + (ratio - 1) * (n / (n + MODEL_PARAMS.K_TEAM));
}

function shrunkRate90(model, rows, ruolo, m, teamId) {
  let prior = rate90Of(model.role[ruolo], m);
  if (prior != null && teamId) prior *= teamStrengthFactor(model, teamId, m);
  let X = 0, mn = 0;
  for (const r of rows) { X += metricValue(r, m); mn += Number(r.Minuti) || 0; }
  if (prior == null) return { rate90: mn ? (X / mn) * 90 : null, prior90: null, minutes: mn, wPlayer: 1 };
  const K = MODEL_PARAMS.K_MIN;
  return { rate90: (X + (prior * K) / 90) / ((mn + K) / 90), prior90: prior, minutes: mn, wPlayer: mn / (mn + K) };
}

// Metrica "speculare" di un avversario specifico: quanti falli SUBISCE un
// giocatore (FD) e' cio' che fa capire quanti falli COMMETTE chi gli gioca
// contro (FC), e viceversa. Usata per il segnale sui titolari dell'avversario
// (vedi oppLineupSignal piu' sotto). Nessun mirror per Tiri/SOT/Cart: non c'e'
// una statistica individuale equivalente di "quanto concedo" per quelle.
const MIRROR_METRIC = { FC: 'FD', FD: 'FC' };

// Segnale personalizzato sui titolari REALI dell'avversario (richiesto da
// Marco, 24/10): invece di guardare solo quanto l'avversario concede in
// media storica per ruolo, guarda quanto i suoi titolari di OGGI, uno per
// uno, tendono personalmente a subire/commettere falli (stessa pipeline di
// shrinkage del giocatore che si sta proiettando: shrunkRate90 verso il
// proprio ruolo, pesata sui propri minuti). Poi aggrega sul minutaggio atteso
// di ciascuno. Non sostituisce il segnale storico di squadra/ruolo, si somma:
// combinati piu' sotto, pesati sul proprio campione (minuti personali vs
// minuti del campione storico), cosi' un titolare con poco storico personale
// (nuovo, appena rientrato) non azzera il segnale di squadra gia' noto.
function oppLineupSignal(model, matchLog, oppLineup, m) {
  const mm = MIRROR_METRIC[m];
  if (!mm || !oppLineup || !oppLineup.length) return null;
  let sumMu = 0, sumMin = 0;
  for (const op of oppLineup) {
    const rows = playerRows(matchLog, op.id, op.team);
    const base = shrunkRate90(model, rows, op.ruolo, mm, op.team);
    if (base.rate90 == null) continue;
    const expMin = expectedMinutes(model, rows, true);
    sumMu += base.rate90 * expMin;
    sumMin += expMin;
  }
  if (!sumMin) return null;
  return { rate90: (sumMu / sumMin) * (90 / 90), minutes: sumMin }; // rate90 = media/90 della formazione avversaria su quella metrica speculare
}

// Fattore avversario a due livelli. Squadra: quanto l'avversario concede in totale
// (tutti i ruoli) rispetto alla lega. Ruolo: scostamento ulteriore dello stesso ruolo,
// molto piu' smorzato, ora raffinato con il segnale sui titolari reali dell'avversario
// se disponibile (oppLineup/matchLog passati) — storico di ruolo e segnale titolari si
// combinano in un'unica media pesata sui rispettivi minuti, poi tirata verso la lega con
// lo stesso K_OPP di prima (nessun nuovo parametro da calibrare).
function opponentFactor(model, oppId, ruolo, m, pv, matchLog, oppLineup) {
  // pv ('H' | 'A'): se presente usa solo le partite in cui i giocatori affrontavano l'avversario nello stesso venue
  // di questa partita, confrontate con la lega nello stesso venue. Il campione e' piu' piccolo e i pesi
  // (n/(n+K)) si abbassano da soli.
  let fTeam = 1, rTeam = null, wTeam = 0;
  const n = pv ? model.oppNV[oppId + '|' + pv] : model.oppN[oppId], at = pv ? model.oppTV[oppId + '|' + pv] : model.oppT[oppId], lg = rate90Of(pv ? model.lgAllV[pv] : model.lgAll, m);
  if (MODEL_PARAMS.OPP_TEAM_METRICS.includes(m) && n && at && at.min && lg) {
    rTeam = rate90Of(at, m) / lg;
    wTeam = n / (n + MODEL_PARAMS.K_OPP_TEAM);
    fTeam = 1 + (rTeam - 1) * wTeam;
  }
  let fRole = 1, rRole = null, wRole = 0, minutes = 0, rLineup = null, lineupMinutes = 0;
  const a = pv ? model.oppRV[oppId + '|' + ruolo + '|' + pv] : model.opp[oppId + '|' + ruolo], prior = rate90Of(pv ? model.roleV[ruolo + '|' + pv] : model.role[ruolo], m);
  if (a && a.min && prior) {
    rRole = rate90Of(a, m) / prior;
    wRole = a.min / (a.min + MODEL_PARAMS.K_OPP);
    minutes = a.min;
  }
  if (matchLog && oppLineup && prior) {
    const sig = oppLineupSignal(model, matchLog, oppLineup, m);
    if (sig) { rLineup = sig.rate90 / prior; lineupMinutes = sig.minutes; }
  }
  // Combina lo storico di ruolo (rRole, peso su a.min) e il segnale sui titolari
  // reali (rLineup, peso su lineupMinutes) in un unico rapporto, media pesata sui
  // rispettivi campioni — nessuno dei due azzera l'altro, pesano in base a quanti
  // minuti di dati ciascuno porta.
  let rCombined = rRole, wBasis = minutes;
  if (rLineup != null) {
    wBasis = minutes + lineupMinutes;
    rCombined = rRole != null ? (rRole * minutes + rLineup * lineupMinutes) / wBasis : rLineup;
  }
  if (rCombined != null) {
    const wCombined = wBasis / (wBasis + MODEL_PARAMS.K_OPP);
    fRole = 1 + (rCombined - 1) * wCombined;
    wRole = wCombined;
  }
  return { factor: fTeam * fRole, factorTeam: fTeam, factorRole: fRole, ratio: rCombined, ratioTeam: rTeam,
           ratioRoleStorico: rRole, ratioLineup: rLineup, lineupMinutes,
           minutes: wBasis, weight: wRole, weightTeam: wTeam, matches: n || 0 };
}

// Minuti attesi: media delle proprie presenze nello stesso ruolo di impiego
// (titolare/subentrato) se ce n'e' almeno una, altrimenti media di lega.
// Per i subentrati nel backtest la media di lega e' piu' stabile di quella personale.
function expectedMinutes(model, rows, titolare) {
  if (titolare) {
    const st = rows.filter((r) => r.Titolare === 'Sì');
    if (st.length) return st.reduce((a, r) => a + (Number(r.Minuti) || 0), 0) / st.length;
    return model.startMinMean;
  }
  return model.subMinMean;
}

function reliabilityOf(minutes) {
  if (minutes >= MODEL_PARAMS.MIN_AFFIDABILITA_ALTA) return 'alta';
  if (minutes >= MODEL_PARAMS.MIN_AFFIDABILITA_MEDIA) return 'media';
  return 'bassa';
}

// Proiezione completa di un giocatore su una metrica, PER QUESTA PARTITA.
// player = { id, team, ruolo }; opp = Team_ID avversario; titolare = bool.
function projectPlayer(model, matchLog, player, oppId, titolare, m, pv, oppLineup) {
  const rows = playerRows(matchLog, player.id, player.team);
  const base = shrunkRate90(model, rows, player.ruolo, m, player.team);
  const of = opponentFactor(model, oppId, player.ruolo, m, pv, matchLog, oppLineup);
  const expMin = expectedMinutes(model, rows, titolare);
  const rate90 = base.rate90 == null ? null : base.rate90 * of.factor;
  return {
    metric: m,
    presenze: rows.length,
    minutiStagione: base.minutes,
    baseline90: base.rate90,       // tasso/90 tirato verso il ruolo (senza avversario)
    prior90: base.prior90,         // media di lega del ruolo
    pesoGiocatore: base.wPlayer,   // 0 = tutto ruolo, 1 = tutto giocatore
    fattoreAvversario: of.factor,  // squadra x ruolo
    fattoreSquadra: of.factorTeam, fattoreRuolo: of.factorRole,
    pesoSquadra: of.weightTeam, pesoRuolo: of.weight, rapportoSquadra: of.ratioTeam, partiteAvversario: of.matches,
    rapportoGrezzoAvversario: of.ratio,
    minutiAvversarioRuolo: of.minutes,
    rate90,                        // baseline90 * fattore
    minutiAttesi: expMin,
    mu: rate90 == null ? null : (rate90 * expMin) / 90, // valore atteso in questa partita
    affidabilita: reliabilityOf(base.minutes),
  };
}

// --- Distribuzioni: Poisson (alpha=0) e Binomiale Negativa (var = mu + alpha*mu^2)
function pmfNB(k, mu, alpha) {
  if (mu <= 0) return k === 0 ? 1 : 0;
  if (!alpha) { // Poisson
    let p = Math.exp(-mu);
    for (let i = 1; i <= k; i++) p *= mu / i;
    return p;
  }
  const r = 1 / alpha, p0 = r / (r + mu); // P(0) = p0^r
  let pk = Math.pow(p0, r);
  for (let i = 0; i < k; i++) pk *= ((r + i) / (i + 1)) * (1 - p0);
  return pk;
}
// P(X >= k): k = soglia intera. "Over 2,5" = P(X>=3); "Under 2,5" = 1 - P(X>=3).
function probAtLeast(mu, k, metric) {
  return probAtLeastAlpha(mu, k, MODEL_PARAMS.ALPHA[metric] || 0);
}
// Variante con alpha esplicito (usata per i totali di squadra: TEAM_ALPHA).
function probAtLeastAlpha(mu, k, a) {
  if (mu == null || !Number.isFinite(mu)) return null;
  if (k <= 0) return 1;
  let cdf = 0;
  for (let i = 0; i < k; i++) cdf += pmfNB(i, mu, a);
  return Math.max(0, Math.min(1, 1 - cdf));
}
function minOddsFromProb(p) { return p && p > 0 ? 1 / p : null; }

// Hit rate di un giocatore su soglia intera k (k o piu'), solo presenze con
// almeno MIN_HITRATE minuti, con intervallo di Wilson. Mostrare SEMPRE n.
function playerHitRateK(rows, metric, k) {
  const valid = rows.filter((r) => (Number(r.Minuti) || 0) >= MODEL_PARAMS.MIN_HITRATE);
  const n = valid.length;
  const hits = valid.filter((r) => metricValue(r, metric) >= k).length;
  const rate = n ? hits / n : null;
  const [lo, hi] = wilsonInterval(rate, n);
  return { n, hits, rate, lo, hi };
}

// Cartellini attesi in una partita dai falli attesi: falli * (cartellini per fallo
// di lega). Poi Poisson: P(almeno 1) = 1 - exp(-lambda).
function cardsPerFoulLeague(matchLog) {
  let c = 0, f = 0;
  for (const r of matchLog) { c += metricValue(r, 'Cart'); f += metricValue(r, 'FC'); }
  return f ? c / f : null;
}
function probCardFromFouls(expFouls, cardsPerFoul) {
  if (expFouls == null || cardsPerFoul == null) return null;
  return 1 - Math.exp(-expFouls * cardsPerFoul);
}

// Totale squadra da una formazione. players: [{id, team, ruolo}]. Restituisce
// somma sui titolari, stima con panchina e il CONTEGGIO dei giocatori contati,
// cosi' un totale incompleto non passa mai inosservato.
function projectLineupTotal(model, matchLog, players, oppId, m, extraFactor, pv, oppLineup) {
  const ef = extraFactor == null ? 1 : extraFactor; // es. fattore arbitro per Falli/Cartellini
  let sumStarters = 0, counted = 0, lowRel = 0;
  const rows = [];
  for (const p of players) {
    const pr = projectPlayer(model, matchLog, p, oppId, true, m, pv, oppLineup);
    rows.push({ player: p, proj: pr });
    if (pr.mu != null) { sumStarters += pr.mu * ef; counted++; if (pr.affidabilita === 'bassa') lowRel++; }
  }
  return {
    metric: m,
    startersTotal: sumStarters,
    withBench: sumStarters * (model.benchFactor[m] || 1),
    counted, expected: players.length, lowReliability: lowRel,
    complete: counted === players.length && players.length === 11,
    rows,
  };
}

// =============================================================================
// GENERATORE NARRATIVO — pannello "Analisi partita" nel Matchup Engine.
// Usa solo funzioni già testate sopra (computeRoleConcessions, projectPlayer,
// computeTeamStreak, ecc.); qui si limita a orchestrare e a scegliere COSA
// mostrare (ruolo più bucato, nomi con matchup migliore/peggiore, sequenze
// per venue). Non introduce nuova logica di calcolo statistico oltre al
// modello Poisson con shrinkage per i gol (projectGoals), che NON è lo
// stesso motore testato dei tiri/SOT/falli/cartellini — va trattato con una
// cautela in più, come dichiarato nel testo generato.
// =============================================================================
const NARRATIVE_ROLES = ['Punta centrale', 'Trequartista', 'Esterno/Quinto', 'Centrocampista centrale', 'Terzino', 'Centrale difensivo'];

// Articolo corretto per ogni squadra di Serie A — in italiano calcistico ogni nome di
// squadra richiede l'articolo ("il Bologna", "l'Inter", "la Juventus"), mai il nome nudo
// come soggetto di frase né con "di" semplice ("di Bologna" è sbagliato, "del Bologna" è giusto).
const NARR_TEAM_ART = {
  Atalanta: 'f_v', Bologna: 'm_c', Cagliari: 'm_c', Como: 'm_c', Fiorentina: 'f_c', Frosinone: 'm_c',
  Genoa: 'm_c', Inter: 'f_v', Juventus: 'f_c', Lazio: 'f_c', Lecce: 'm_c', Milan: 'm_c', Monza: 'm_c',
  Napoli: 'm_c', Parma: 'm_c', Roma: 'f_c', Sassuolo: 'm_c', Torino: 'm_c', Udinese: 'f_v', Venezia: 'm_c',
};
function narrArtKind(teamName) { return NARR_TEAM_ART[teamName] || 'm_c'; }
// Articolo determinativo isolato ("il"/"la"/"l'"), con maiuscola se richiesta (inizio frase).
function narrArt(teamName, cap) {
  const k = narrArtKind(teamName);
  const a = k === 'f_v' ? "l'" : k === 'f_c' ? 'la ' : 'il ';
  return cap ? a.charAt(0).toUpperCase() + a.slice(1) : a;
}
// Squadra come soggetto/oggetto con articolo: "il Bologna", "l'Inter", "la Juventus".
function narrTeam(teamName, cap) { return narrArt(teamName, cap) + teamName; }
// Forma possessiva "del/dell'/della X", per sostituire ogni "di X" errato.
function narrTeamOf(teamName) {
  const k = narrArtKind(teamName);
  const p = k === 'f_v' ? "dell'" : k === 'f_c' ? 'della ' : 'del ';
  return p + teamName;
}
// Forma "dal/dall'/dalla X" (complemento d'agente), per sostituire ogni "da X" errato.
function narrTeamDa(teamName) {
  const k = narrArtKind(teamName);
  const p = k === 'f_v' ? "dall'" : k === 'f_c' ? 'dalla ' : 'dal ';
  return p + teamName;
}

// Media di lega di quanto un RUOLO specifico concede quando una squadra gioca nel venue V —
// stesso identico calcolo di narrWeakestRoleExtra ma mediato su tutte le squadre col campione
// minimo, per poter distinguere una VERA debolezza (sopra media) da un dato solo "il meno buono
// tra i propri ruoli" ma comunque nella norma.
function narrLeagueRoleAvg(matchLog, ruolo, venueLetter, metricKey) {
  const key = metricKey || 'Tiri_Concessi_90';
  const teamIds = [...new Set(matchLog.map((r) => r.Team_ID))];
  const segmento = venueLetter === 'H' ? 'HOME' : 'AWAY';
  let s = 0, n = 0;
  for (const t of teamIds) {
    const partite = narrPartiteInVenue(matchLog, t, venueLetter);
    if (!partite) continue;
    const row = computeRoleConcessions(matchLog, t, segmento, ruolo, partite);
    if (row.N_Prestazioni_Ruolo >= 3 && row[key] != null) { s += row[key]; n++; }
  }
  return n ? s / n : null;
}

function narrStandingsTable(calendar) {
  const played = calendar.filter((c) => c.Stato === 'CONCLUSA' && c.Home_Score !== '' && c.Home_Score != null);
  const t = {};
  const ensure = (id, name) => { if (!t[id]) t[id] = { id, name, pts: 0, gf: 0, ga: 0, w: 0, d: 0, l: 0, pg: 0 }; };
  for (const m of played) {
    ensure(m.Home_ID, m.Home); ensure(m.Away_ID, m.Away);
    const hs = Number(m.Home_Score), as = Number(m.Away_Score);
    t[m.Home_ID].gf += hs; t[m.Home_ID].ga += as; t[m.Home_ID].pg++;
    t[m.Away_ID].gf += as; t[m.Away_ID].ga += hs; t[m.Away_ID].pg++;
    if (hs > as) { t[m.Home_ID].pts += 3; t[m.Home_ID].w++; t[m.Away_ID].l++; }
    else if (hs < as) { t[m.Away_ID].pts += 3; t[m.Away_ID].w++; t[m.Home_ID].l++; }
    else { t[m.Home_ID].pts += 1; t[m.Away_ID].pts += 1; t[m.Home_ID].d++; t[m.Away_ID].d++; }
  }
  const arr = Object.values(t).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
  arr.forEach((x, i) => (x.pos = i + 1));
  return arr;
}

function narrVenueRecord(calendar, teamId, venue) {
  const rows = calendar.filter((c) => c.Stato === 'CONCLUSA' && c.Home_Score !== '' && c.Home_Score != null &&
    (venue === 'H' ? c.Home_ID === teamId : c.Away_ID === teamId));
  let w = 0, d = 0, l = 0, gf = 0, ga = 0;
  for (const r of rows) {
    const isHome = r.Home_ID === teamId;
    const f = Number(isHome ? r.Home_Score : r.Away_Score), a = Number(isHome ? r.Away_Score : r.Home_Score);
    gf += f; ga += a;
    if (f > a) w++; else if (f < a) l++; else d++;
  }
  return { n: rows.length, w, d, l, gf, ga };
}

// Ruolo più "bucato" di una squadra nel proprio venue di oggi. segmento passato a
// computeRoleConcessions è l'OPPOSTO del venue reale (stessa inversione di
// ARCHITECTURE.md: per "cosa concede X mentre gioca in V" serve segmento = not(V)).
function narrWeakestRole(matchLog, teamId, venue, teamPartiteInVenue) {
  const r = narrWeakestRoleMulti(matchLog, teamId, venue, teamPartiteInVenue).primary;
  return r ? r.row : null;
}

// Versione multi-metrica: la "vera falla" di una squadra in un venue non è sempre sui
// tiri concessi — può emergere altrettanto (o più) nettamente sui falli SUBITI da un
// ruolo avversario (FalliSubiti_Ruolo_90, la vera debolezza disciplinare della squadra).
// Si valutano entrambe le metriche su tutti i ruoli e si tiene quella con lo scarto
// relativo (ratio sulla media di lega stesso ruolo/venue) più marcato: così la Chiave
// di lettura non è forzata sui tiri in ogni partita. candidate = {row, metric, value, avg, ratio}.
function narrWeakestRoleMulti(matchLog, teamId, venue, teamPartiteInVenue) {
  const segmento = venue === 'H' ? 'HOME' : 'AWAY';
  const rows = NARRATIVE_ROLES.map((r) => computeRoleConcessions(matchLog, teamId, segmento, r, teamPartiteInVenue))
    .filter((x) => x.N_Prestazioni_Ruolo >= 3);
  if (!rows.length) return { primary: null, secondary: null, sotWeak: null };
  const cands = [];
  for (const row of rows) {
    if (row.Tiri_Concessi_90 != null) {
      const avg = narrLeagueRoleAvg(matchLog, row.Posizione_Analitica, venue, 'Tiri_Concessi_90');
      cands.push({ row, metric: 'Tiri', value: row.Tiri_Concessi_90, avg, ratio: avg ? row.Tiri_Concessi_90 / avg : null });
    }
    if (row.FalliSubiti_Ruolo_90 != null) {
      const avg = narrLeagueRoleAvg(matchLog, row.Posizione_Analitica, venue, 'FalliSubiti_Ruolo_90');
      cands.push({ row, metric: 'Falli', value: row.FalliSubiti_Ruolo_90, avg, ratio: avg ? row.FalliSubiti_Ruolo_90 / avg : null });
    }
  }
  // I falli subiti per ruolo hanno varianza di lega più bassa dei tiri: anche uno scarto modesto
  // produce un ratio enorme, che altrimenti farebbe vincere i falli quasi sempre invece di "variare"
  // davvero la chiave come richiesto. Si passa ai falli solo se il loro scarto è CHIARAMENTE più
  // marcato (almeno il 20% in più) di quello sui tiri dello stesso venue — altrimenti i tiri restano
  // il riferimento di default, e i falli emergono solo quando sono davvero il segnale più forte.
  const bestTiri = cands.filter((c) => c.metric === 'Tiri' && c.ratio != null).sort((a, b) => b.ratio - a.ratio)[0] || null;
  const bestFalli = cands.filter((c) => c.metric === 'Falli' && c.ratio != null).sort((a, b) => b.ratio - a.ratio)[0] || null;
  let primary;
  if (bestFalli && (!bestTiri || bestFalli.ratio >= bestTiri.ratio * 1.5)) {
    primary = bestFalli;
  } else if (bestTiri) {
    primary = bestTiri;
  } else {
    const byTiri = rows.filter((r) => r.Tiri_Concessi_90 != null).sort((a, b) => b.Tiri_Concessi_90 - a.Tiri_Concessi_90);
    primary = byTiri[0] ? { row: byTiri[0], metric: 'Tiri', value: byTiri[0].Tiri_Concessi_90, avg: null, ratio: null } : null;
  }
  if (!primary) return { primary: null, secondary: null, sotWeak: null };
  // seconda debolezza: altro ruolo, STESSA metrica del primario, valore vicino (>=75%) — non marginale.
  const sameMetric = cands.filter((c) => c.metric === primary.metric && c.row.Posizione_Analitica !== primary.row.Posizione_Analitica);
  const secondCand = sameMetric.sort((a, b) => b.value - a.value)[0];
  const secondary = secondCand && secondCand.value >= primary.value * 0.75 ? secondCand : null;
  // la lettura specifica sui tiri IN PORTA ha senso solo quando la falla principale è sui tiri.
  let sotWeak = null;
  if (primary.metric === 'Tiri') {
    const bySot = rows.filter((x) => x.SOT_Concessi_90 != null).sort((a, b) => b.SOT_Concessi_90 - a.SOT_Concessi_90);
    sotWeak = bySot[0] && bySot[0].Posizione_Analitica !== primary.row.Posizione_Analitica ? bySot[0] : null;
  }
  return { primary, secondary, sotWeak };
}

// Media di una metrica (Corner/Cart/...) sui totali-partita di una squadra in un venue.
function narrAvgOf(totals, key) {
  const vs = totals.map((t) => t[key]).filter((v) => v != null);
  return vs.length ? { avg: vs.reduce((a, b) => a + b, 0) / vs.length, n: vs.length } : null;
}

function narrPartiteInVenue(matchLog, teamId, venueLetter) {
  return new Set(matchLog.filter((r) => r.Team_ID === teamId && r.Venue === venueLetter).map((r) => r.Match_ID)).size;
}

function narrPlayersInRole(rosters, matchLog, teamId, ruolo, minMin) {
  const roster = rosters.filter((r) => r.Team_ID_Rif === teamId && r.Attivo_SerieA === 'SI' && r.Posizione_Analitica === ruolo);
  return roster.map((r) => {
    const min = matchLog.filter((x) => x.Player_ID === r.Player_ID).reduce((a, x) => a + (Number(x.Minuti) || 0), 0);
    return { id: r.Player_ID, nome: r.Nome_completo, ruolo: r.Posizione_Analitica, min };
  }).filter((p) => p.min >= minMin).sort((a, b) => b.min - a.min);
}

// Le 5 statistiche disponibili per giocatore/ruolo, già quelle usate nei filtri dello
// Scanner: chiave campo in match_log, etichetta, e campo corrispondente prodotto da
// computeRoleConcessions (per la classifica di Punto 2).
const NARR_ROLE_METRICS = [
  { key: 'Tiri', label: 'tiri', concessField: 'Tiri_Concessi_90' },
  { key: 'Tiri_in_porta', label: 'tiri in porta', concessField: 'SOT_Concessi_90' },
  { key: 'Falli_commessi', label: 'falli commessi', concessField: 'Falli_Contro_Ruolo_90' },
  { key: 'Falli_subiti', label: 'falli subiti', concessField: 'FalliSubiti_Ruolo_90' },
  { key: 'Cartellini', label: 'cartellini', concessField: 'Cartellini_Ruolo_90' },
];

// Punto 1 — cosa fa SQUADRA col RUOLO contro QUESTO avversario specifico (solo nel
// venue di questa giornata), rispetto a cosa fa (stesso ruolo, stesso venue) contro
// tutte le altre squadre. Fatto storico diretto (non media di lega): con 1 solo
// precedente il campione resta minimo ma il dato è comunque un fatto vero, va solo
// etichettato come tale (n=1). metricKey una delle chiavi di NARR_ROLE_METRICS.
function narrRoleVsOpponent(matchLog, teamId, ruolo, venueLetter, opponentId, metricKey) {
  const base = matchLog.filter((r) => r.Team_ID === teamId && r.Posizione_Analitica === ruolo && r.Venue === venueLetter);
  const vsOpp = base.filter((r) => r.Opp_ID === opponentId);
  const vsAltri = base.filter((r) => r.Opp_ID !== opponentId);
  if (!vsOpp.length || !vsAltri.length) return null;
  const val = (r) => metricKey === 'Cartellini' ? (Number(r.Gialli) || 0) + (Number(r.Rossi) || 0) : (Number(r[metricKey]) || 0);
  const per90 = (rows) => {
    const n = new Set(rows.map((r) => r.Match_ID)).size;
    return { n, v90: rows.reduce((s, r) => s + val(r), 0) / n };
  };
  const a = per90(vsOpp), b = per90(vsAltri);
  if (!b.v90) return null;
  return { ruolo, metricKey, n: a.n, v90VsOpp: a.v90, v90Altri: b.v90, ratio: a.v90 / b.v90 };
}

// Punto 2 — posizione in classifica (1°-N) di TEAM tra tutte le squadre della lega per
// quanto concede al RUOLO nel VENUE indicato, per una delle 5 statistiche (stessa
// metrica di "Chiave di lettura", qui espressa come rank assoluto invece che come
// ratio-alla-media). Usa tutta la stagione: nessun problema di campione N=1.
function narrRoleConcessionRank(matchLog, teams, teamId, ruolo, venueLetter, metricKey) {
  const segmento = venueLetter === 'H' ? 'HOME' : 'AWAY';
  const field = (NARR_ROLE_METRICS.find((m) => m.key === metricKey) || {}).concessField;
  if (!field) return null;
  const rows = [];
  for (const t of teams) {
    const partite = narrPartiteInVenue(matchLog, t, venueLetter);
    if (!partite) continue;
    const c = computeRoleConcessions(matchLog, t, segmento, ruolo, partite);
    if (!c || !c.N_Prestazioni_Ruolo) continue;
    rows.push({ team: t, value: c[field], partite });
  }
  if (!rows.length) return null;
  rows.sort((a, b) => b.value - a.value);
  const idx = rows.findIndex((r) => r.team === teamId);
  if (idx < 0) return null;
  const mediaLega = rows.reduce((s, r) => s + r.value, 0) / rows.length;
  return { rank: idx + 1, totale: rows.length, value: rows[idx].value, mediaLega, partite: rows[idx].partite };
}

// ---------------------------------------------------------------------------
// Fatti salienti "ruolo vs avversario" per lo Scanner — STESSI criteri di
// computePlayerVsOpponentFacts() in index.html (Fatti salienti):
//  - ruoli = quelli della rosa attiva della squadra in esame, Portiere escluso
//  - metriche: tiri, tiri in porta, falli (le stesse 3 di index.html)
//  - confronto con tutte le squadre della lega nello stesso campo
//  - rankProvisional con campione minimo, almeno 6 squadre idonee,
//    scartato se piu' di FATTI_MAX_TIE squadre hanno lo stesso valore
//  - SOLO top3 (rank <= 3): niente "ultimi 3"/"meno" in questa categoria —
//    qui i fatti alimentano direttamente lo Scanner (filtra tabella → giocata
//    Over), e un fatto "concede meno" suggerirebbe Under, non utile qui
//    (22/10: prima non c'era soglia, un 7°/20 passava; poi provata soglia
//    simmetrica top3/ultimi3 ma scartata su richiesta esplicita — solo top3)
// NB: questa regola (solo top3, mai "meno") vale SOLO per questa categoria
// dentro lo Scanner. In index.html, computePlayerVsOpponentFacts() (Fatti
// salienti, tag "Matchup ruolo") resta allineata: stessa soglia rank<=3.
// Se si cambia un criterio in index.html va cambiato anche qui.
// FATTI_RUOLO_MIN_SAMPLE_VENUE: Serie A 2 partite per campo (come minSample venue di Fatti salienti in index.html); UCL 1.
// ---------------------------------------------------------------------------
const FATTI_MAX_TIE = 2; // = MAX_TIE_FOR_FACT in index.html
const FATTI_RUOLO_MIN_SAMPLE_VENUE = 2;
const FATTI_RUOLO_METRICS = [
  { key: 'Tiri', label: 'tiri', col: 'Tiri_Concessi_90' },
  { key: 'Tiri_in_porta', label: 'tiri in porta', col: 'SOT_Concessi_90' },
  { key: 'Falli_commessi', label: 'falli', col: 'Falli_Contro_Ruolo_90' },
];
// oppVenueLetter: campo in cui gioca l'avversario ('H' casa, 'A' trasferta).
// minSample opzionale: serve solo ai test di allineamento con index.html.
function fattiRuoloVsAvversario(matchLog, rosters, teamId, oppId, oppVenueLetter, minSample = FATTI_RUOLO_MIN_SAMPLE_VENUE) {
  const oppVenue = oppVenueLetter === 'H' ? 'HOME' : 'AWAY';
  const teamsAll = [...new Set(rosters.map((r) => r.Team_ID_Rif))];
  const roles = [...new Set(rosters.filter((p) => p.Attivo_SerieA === 'SI' && p.Team_ID_Rif === teamId).map((p) => p.Posizione_Analitica))]
    .filter((r) => r !== 'Portiere');
  const out = [];
  for (const ruolo of roles) {
    for (const m of FATTI_RUOLO_METRICS) {
      const list = teamsAll.map((t) => {
        const ts = computeTeamStats(matchLog, t, oppVenue);
        const rc = computeRoleConcessions(matchLog, t, oppVenue, ruolo, ts.Partite);
        return { key: t, value: rc[m.col], partite: ts.Partite };
      });
      const { rank, eligible } = rankProvisional(list, minSample)[oppId];
      const me = list.find((x) => x.key === oppId);
      if (rank === 'N/D' || eligible < 6) continue;
      if (rank > 3) continue; // solo top3 (concede di più): mai "meno" in questa categoria, suggerirebbe Under
      const pool = list.filter((x) => x.partite >= minSample && x.value != null);
      if (pool.filter((x) => x.value === me.value).length > FATTI_MAX_TIE) continue;
      const top = rank <= (eligible - rank + 1); // = isTopEnd() di index.html
      out.push({
        ruolo, metricKey: m.key, metricLabel: m.label, rank, totale: eligible, top, value: me.value,
        mediaLega: pool.reduce((a, x) => a + x.value, 0) / pool.length,
        sotto: pool.filter((x) => x.value < me.value).length, // squadre con valore strettamente minore
        sopra: rank - 1,                                      // squadre con valore strettamente maggiore
        lato: roleConcessionsSideBreakdown(matchLog, oppId, oppVenue, ruolo, m.key), // null se ruolo non largo / sotto soglia / dataset senza Lato_Campo
      });
    }
  }

  out.sort((a, b) => Math.min(a.rank, a.totale + 1 - a.rank) - Math.min(b.rank, b.totale + 1 - b.rank));
  return out;
}

// ---------------------------------------------------------------------------
// INCROCIO A DOPPIO ESTREMO (richiesta Marco, 25/10) — fino a qui il "Matchup
// ruolo" guardava UN solo lato: quanto la squadra che difende concede a un
// ruolo (computeRoleConcessions), confrontato con la media di lega. Non
// diceva se la squadra che attacca in quel ruolo è REALMENTE forte lì
// (rank di lega sulla PRODUZIONE, computeRoleProduction) o è nella media e
// il mismatch è solo apparente.
//
// computeRoleCrossExtremes guarda ENTRAMBI i lati per lo stesso ruolo/metrica/
// venue reale della partita e tiene solo le combinazioni dove TUTTE E DUE le
// squadre sono agli estremi della classifica di lega (rank <= EXTREME_RANK
// su entrambi i lati, stesso criterio "solo top3" già in uso per i fatti di
// ruolo in index.html/fattiRuoloVsAvversario — non reinventato). Non
// un rank nuovo: ESATTAMENTE gli stessi due rank già mostrati separatamente
// nei pannelli Role matchup (concessione) — qui aggiunta solo la produzione
// e il confronto incrociato tra i due.
// ---------------------------------------------------------------------------
const CROSS_EXTREME_RANK = 3; // = "top3" di fattiRuoloVsAvversario, stessa soglia
const CROSS_EXTREME_MIN_ELIGIBLE = 6; // = eligible>=6 già richiesto altrove per i fatti di ruolo
const CROSS_EXTREME_METRICS = [
  { key: 'Tiri', label: 'tiri', prodCol: 'Tiri_Prodotti_90', concCol: 'Tiri_Concessi_90', modelMetric: 'Tiri' },
  { key: 'Tiri_in_porta', label: 'tiri in porta', prodCol: 'SOT_Prodotti_90', concCol: 'SOT_Concessi_90', modelMetric: 'SOT' },
  { key: 'Falli_commessi', label: 'falli commessi', prodCol: 'Falli_Dal_Ruolo_90', concCol: 'Falli_Contro_Ruolo_90', modelMetric: 'FC' },
];
const CROSS_INTERPRETER_MIN_MIN = 300; // = soglia minuti già in uso per gli "exploiters" in generateMatchNarrative (withMuTiri/narrPlayersInRole), non nuova

// Possibile interprete (richiesta Marco 25/10): tra i giocatori della squadra
// che ATTACCA, nel ruolo dell'incrocio, quello con mu più alto dal MODELLO già
// testato altrove (projectPlayer — stesso usato in "Nomi che contano"), non una
// media grezza nuova. Nessun numero inventato: se il modello non produce un mu
// per nessun giocatore di quel ruolo con minuti sufficienti, ritorna null e la
// card resta senza interprete (meglio nessun nome che un nome debole).

// Rank di una squadra in classifica di lega per una metrica di ruolo (produzione
// O concessione, stesso identico schema di rankProvisional già in uso) — ritorna
// anche l'intero elenco ordinato, serve a verificare gli "eligible" per il
// controllo pareggio (FATTI_MAX_TIE, la stessa regola anti-pareggio-di-massa
// usata ovunque in questo motore, non una soglia nuova).
function roleMetricLeagueRank(matchLog, teams, ruolo, venueLetter, metricCol, side) {
  const segmento = venueLetter === 'H' ? 'HOME' : 'AWAY';
  const computeFn = side === 'produzione' ? computeRoleProduction : computeRoleConcessions;
  const minSample = 2; // = minSample HOME/AWAY già in uso ovunque nel motore (vedi rankProvisional/minSample), non una soglia nuova
  const rows = [];
  for (const t of teams) {
    const partite = narrPartiteInVenue(matchLog, t, venueLetter);
    if (partite < minSample) continue;
    const c = computeFn(matchLog, t, segmento, ruolo, partite);
    if (!c || c[metricCol] == null) continue;
    rows.push({ team: t, value: c[metricCol], partite, nPrestazioni: c.N_Prestazioni_Ruolo });
  }
  if (rows.length < CROSS_EXTREME_MIN_ELIGIBLE) return { rank: null, eligible: rows.length, rows };
  rows.sort((a, b) => b.value - a.value);
  return { rank: rows, eligible: rows.length, rows };
}

// Per UNA direzione (squadra che attacca in un ruolo, venue V_att; squadra che
// difende quello stesso ruolo, venue reale opposto V_dif di questa gara) — tutte
// le combinazioni ruolo+metrica dove ATTACCANTE è top3 in PRODUZIONE e
// DIFENDENTE è top3 in CONCESSIONE, contemporaneamente.
function crossInterpreter(model, matchLog, rosters, attId, ruolo, difId, attVenueLetter, modelMetric) {
  const roster = rosters.filter((r) => r.Team_ID_Rif === attId && r.Attivo_SerieA === 'SI' && r.Posizione_Analitica === ruolo);
  let best = null;
  for (const r of roster) {
    const rows = matchLog.filter((x) => x.Player_ID === r.Player_ID);
    const min = rows.reduce((a, x) => a + (Number(x.Minuti) || 0), 0);
    if (min < CROSS_INTERPRETER_MIN_MIN) continue;
    const player = { id: r.Player_ID, team: attId, ruolo };
    const pr = projectPlayer(model, matchLog, player, difId, true, modelMetric, attVenueLetter);
    if (pr.mu == null) continue;
    if (!best || pr.mu > best.mu) best = { id: r.Player_ID, nome: r.Nome_completo, mu: pr.mu, min, affidabilita: pr.affidabilita };
  }
  return best;
}

function crossExtremesOneSide(matchLog, rosters, teamsAll, attId, attVenueLetter, difId, difVenueLetter, model) {
  const roles = [...new Set(rosters.filter((p) => p.Attivo_SerieA === 'SI' && p.Team_ID_Rif === attId).map((p) => p.Posizione_Analitica))]
    .filter((r) => r !== 'Portiere');
  const out = [];
  for (const ruolo of roles) {
    for (const m of CROSS_EXTREME_METRICS) {
      const prod = roleMetricLeagueRank(matchLog, teamsAll, ruolo, attVenueLetter, m.prodCol, 'produzione');
      const conc = roleMetricLeagueRank(matchLog, teamsAll, ruolo, difVenueLetter, m.concCol, 'concessione');
      if (prod.eligible < CROSS_EXTREME_MIN_ELIGIBLE || conc.eligible < CROSS_EXTREME_MIN_ELIGIBLE) continue;
      const prodRows = prod.rank, concRows = conc.rank; // array ordinato desc
      const prodIdx = prodRows.findIndex((x) => x.team === attId);
      const concIdx = concRows.findIndex((x) => x.team === difId);
      if (prodIdx < 0 || concIdx < 0) continue;
      const prodRank = prodIdx + 1, concRank = concIdx + 1;
      if (prodRank > CROSS_EXTREME_RANK || concRank > CROSS_EXTREME_RANK) continue;
      // regola anti-pareggio-di-massa: se troppe squadre condividono lo stesso valore estremo, scarta
      const prodTied = prodRows.filter((x) => x.value === prodRows[prodIdx].value).length;
      const concTied = concRows.filter((x) => x.value === concRows[concIdx].value).length;
      if (prodTied > FATTI_MAX_TIE || concTied > FATTI_MAX_TIE) continue;
      const difSegmento = difVenueLetter === 'H' ? 'HOME' : 'AWAY';
      // Lato campo (richiesta Marco 25/10, rafforzata il 25/10: "deve esserci sempre quando
      // disponibile") da cui chi difende concede quella metrica a quel ruolo. Sulla colonna
      // specifica della metrica (es. Tiri_in_porta) il campione è spesso troppo piccolo per
      // una scomposizione per lato: se così, ripieghiamo sul dato grezzo 'Tiri' (stesso ruolo/
      // venue/avversario), che è il segnale posizionale più ampio disponibile — non applicato
      // ai falli, dove il lato campo non è un dato pertinente (vedi mockup approvato).
      let lato = roleConcessionsSideBreakdown(matchLog, difId, difSegmento, ruolo, m.key);
      if (!lato && m.key !== 'Falli_commessi') lato = roleConcessionsSideBreakdown(matchLog, difId, difSegmento, ruolo, 'Tiri');
      const interprete = crossInterpreter(model, matchLog, rosters, attId, ruolo, difId, attVenueLetter, m.modelMetric);
      // Corrispondenza interprete <-> dato di contesto (richiesta Marco 25/10: "controlli
      // corrispondenza tra possibile interprete e dati di contesto"): il lato è una tendenza
      // di SQUADRA (da dove chi difende concede di più, storicamente, contro tutti gli
      // avversari) — non necessariamente il lato in cui gioca l'interprete nominato. Se sappiamo
      // con campione sufficiente che l'interprete gioca prevalentemente dal lato OPPOSTO, il dato
      // di contesto lo contraddice: meglio non mostrarlo che affiancarlo a un nome con cui non
      // corrisponde davvero.
      if (lato && interprete) {
        const domSide = lato.sinistra >= lato.destra && lato.sinistra >= lato.centro ? 'Sinistra' : (lato.destra >= lato.centro ? 'Destra' : 'Centro');
        const playerSide = playerPrevalentSide(matchLog, interprete.id, ruolo, attId);
        if (playerSide && playerSide !== domSide) lato = null;
      }
      out.push({
        attId, difId, ruolo, metricKey: m.key, metricLabel: m.label,
        prodRank, prodEligible: prod.eligible, prodValue: prodRows[prodIdx].value,
        concRank, concEligible: conc.eligible, concValue: concRows[concIdx].value,
        combinedScore: (prod.eligible + 1 - prodRank) + (conc.eligible + 1 - concRank), // più alto = più estremo su entrambi i lati
        interprete,
        lato,
      });
    }
  }
  return out;
}

// Entry point per una partita: incroci a doppio estremo in ENTRAMBE le direzioni
// (casa-attacca/trasferta-difende e trasferta-attacca/casa-difende), ordinati
// dal più estremo. Venue sempre quella reale della gara (mai TOTALE): un
// incrocio vale solo nel venue in cui si gioca davvero quella partita.
function computeRoleCrossExtremes(matchLog, rosters, homeId, awayId) {
  const teamsAll = [...new Set(rosters.map((r) => r.Team_ID_Rif))];
  const model = buildLeagueModel(matchLog); // stesso modello di "Nomi che contano", costruito una volta
  const homeAttacks = crossExtremesOneSide(matchLog, rosters, teamsAll, homeId, 'H', awayId, 'A', model);
  const awayAttacks = crossExtremesOneSide(matchLog, rosters, teamsAll, awayId, 'A', homeId, 'H', model);
  const all = [...homeAttacks, ...awayAttacks];
  all.sort((a, b) => b.combinedScore - a.combinedScore);
  return all;
}

// ---------------------------------------------------------------------------
// INCROCIO A DOPPIO ESTREMO SULLE LINEE (richiesta Marco 25/10: "magari anche
// su linee se una tira tanto e l'altra subisce tanto") — stessa identica logica
// di computeRoleCrossExtremes, ma a livello di SQUADRA invece che di ruolo:
// una squadra che PRODUCE molto (rank<=3 di lega) e l'altra che CONCEDE molto
// (rank<=3) sulla stessa metrica, nel venue reale di questa gara. Usa
// computeTeamStats già testato altrove, nessun numero nuovo.
// ---------------------------------------------------------------------------
const LINE_EXTREME_METRICS = [
  { key: 'Tiri', label: 'tiri', prodCol: 'Tiri_Fatti_90', concCol: 'Tiri_Subiti_90' },
  { key: 'SOT', label: 'tiri in porta', prodCol: 'SOT_Fatti_90', concCol: 'SOT_Subiti_90' },
  { key: 'Falli', label: 'falli', prodCol: 'Falli_Fatti_90', concCol: 'Falli_Subiti_90' },
];
function lineMetricLeagueRank(matchLog, teams, venueSegmento, col) {
  const minSample = 2;
  const list = teams.map((t) => {
    const ts = computeTeamStats(matchLog, t, venueSegmento);
    return { key: t, value: ts[col], partite: ts.Partite };
  });
  return { ranks: rankProvisional(list, minSample), list };
}
function lineExtremesOneSide(matchLog, teamsAll, attId, attSeg, difId, difSeg) {
  const out = [];
  for (const m of LINE_EXTREME_METRICS) {
    const prod = lineMetricLeagueRank(matchLog, teamsAll, attSeg, m.prodCol);
    const conc = lineMetricLeagueRank(matchLog, teamsAll, difSeg, m.concCol);
    const pr = prod.ranks[attId], cr = conc.ranks[difId];
    if (!pr || !cr || pr.rank === 'N/D' || cr.rank === 'N/D') continue;
    if (pr.eligible < CROSS_EXTREME_MIN_ELIGIBLE || cr.eligible < CROSS_EXTREME_MIN_ELIGIBLE) continue;
    if (pr.rank > CROSS_EXTREME_RANK || cr.rank > CROSS_EXTREME_RANK) continue;
    const prodVal = prod.list.find((x) => x.key === attId).value;
    const concVal = conc.list.find((x) => x.key === difId).value;
    const prodTied = prod.list.filter((x) => x.partite >= 2 && x.value === prodVal).length;
    const concTied = conc.list.filter((x) => x.partite >= 2 && x.value === concVal).length;
    if (prodTied > FATTI_MAX_TIE || concTied > FATTI_MAX_TIE) continue;
    out.push({
      attId, difId, metricKey: m.key, metricLabel: m.label,
      prodRank: pr.rank, prodEligible: pr.eligible, prodValue: prodVal,
      concRank: cr.rank, concEligible: cr.eligible, concValue: concVal,
      combinedScore: (pr.eligible + 1 - pr.rank) + (cr.eligible + 1 - cr.rank),
    });
  }
  return out;
}
function computeTeamLineCrossExtremes(matchLog, teamsAll, homeId, awayId) {
  const out = [
    ...lineExtremesOneSide(matchLog, teamsAll, homeId, 'HOME', awayId, 'AWAY'),
    ...lineExtremesOneSide(matchLog, teamsAll, awayId, 'AWAY', homeId, 'HOME'),
  ];
  out.sort((a, b) => b.combinedScore - a.combinedScore);
  return out;
}

function narrShootingEfficiency(matchLog, calendar, teamId, venue) {
  const totals = computeTeamMatchTotals(matchLog, teamId, venue === 'H' ? 'HOME' : 'AWAY', calendar).filter((m) => m.Tiri > 0);
  if (!totals.length) return null;
  const clinical = totals.filter((m) => m.SOT / m.Tiri >= 0.4).length;
  return { n: totals.length, clinical, totals };
}

function narrShrink(avg, n, lg) { return avg == null || !lg ? 1 : 1 + (avg / lg - 1) * (n / (n + SHRINK_K)); }

// Proiezione gol con shrinkage — stessa logica di multiplicativeProjection ma
// su Home_Score/Away_Score di CALENDAR invece che su MATCH_LOG; NON è parte
// del motore tiri/SOT/falli/cartellini testato, va dichiarato nel testo.
function narrProjectGoals(calendar, homeId, awayId) {
  const played = calendar.filter((c) => c.Stato === 'CONCLUSA' && c.Home_Score !== '' && c.Home_Score != null);
  if (!played.length) return null;
  const lgHome = played.reduce((a, c) => a + Number(c.Home_Score), 0) / played.length;
  const lgAway = played.reduce((a, c) => a + Number(c.Away_Score), 0) / played.length;
  const homeAtt = narrVenueRecord(calendar, homeId, 'H');
  const awayDef = narrVenueRecord(calendar, awayId, 'A');
  const awayAtt = narrVenueRecord(calendar, awayId, 'A');
  const homeDef = narrVenueRecord(calendar, homeId, 'H');
  const homeAttAvg = homeAtt.n ? homeAtt.gf / homeAtt.n : null;
  const awayDefAvg = awayDef.n ? awayDef.ga / awayDef.n : null;
  const awayAttAvg = awayAtt.n ? awayAtt.gf / awayAtt.n : null;
  const homeDefAvg = homeDef.n ? homeDef.ga / homeDef.n : null;
  const muHome = lgHome * narrShrink(homeAttAvg, homeAtt.n, lgHome) * narrShrink(awayDefAvg, awayDef.n, lgAway);
  const muAway = lgAway * narrShrink(awayAttAvg, awayAtt.n, lgAway) * narrShrink(homeDefAvg, homeDef.n, lgHome);
  return { muHome, muAway, homeAtt, awayAtt, homeDef, awayDef, bandaPiuPlausibile: narrGoalBandaPiuPlausibile(muHome, muAway) };
}

// Fasce "multigol" (banda di gol più probabile, in %) derivate dalla stessa mu Poisson
// di narrProjectGoals — NON un nuovo modello, solo pmfNB (alpha=0, coerente col commento
// sopra: "modello Poisson con shrinkage per i gol", non il motore NB di tiri/SOT/falli.
// Tre candidati: gol totali di partita (muHome+muAway) e gol della singola squadra (muHome,
// muAway) — la banda restituita è quella con probabilità più alta tra i tre, quindi può
// riferirsi anche a UNA SOLA delle due squadre quando il suo segnale è più netto di quello
// del totale partita (es. una squadra che segna con grande regolarità, l'altra imprevedibile).
const NARR_BANDE_TOTALE = [{ min: 0, max: 1, label: '0-1' }, { min: 2, max: 3, label: '2-3' }, { min: 4, max: 99, label: '4+' }];
const NARR_BANDE_SQUADRA = [{ min: 0, max: 0, label: '0' }, { min: 1, max: 1, label: '1' }, { min: 2, max: 99, label: '2+' }];
function narrBandaProb(mu, banda) {
  if (mu == null || !Number.isFinite(mu)) return null;
  let p = 0;
  for (let k = banda.min; k <= Math.min(banda.max, 30); k++) p += pmfNB(k, mu, 0);
  return p;
}
function narrMiglioreBanda(mu, bande) {
  if (mu == null || !Number.isFinite(mu)) return null;
  let best = null;
  for (const b of bande) {
    const p = narrBandaProb(mu, b);
    if (p != null && (!best || p > best.prob)) best = { label: b.label, prob: p };
  }
  return best;
}
function narrGoalBandaPiuPlausibile(muHome, muAway) {
  const candidati = [];
  if (muHome != null && muAway != null && Number.isFinite(muHome) && Number.isFinite(muAway)) {
    const tot = narrMiglioreBanda(muHome + muAway, NARR_BANDE_TOTALE);
    if (tot) candidati.push({ ambito: 'totale', ...tot });
  }
  const bHome = narrMiglioreBanda(muHome, NARR_BANDE_SQUADRA);
  if (bHome) candidati.push({ ambito: 'home', ...bHome });
  const bAway = narrMiglioreBanda(muAway, NARR_BANDE_SQUADRA);
  if (bAway) candidati.push({ ambito: 'away', ...bAway });
  if (!candidati.length) return null;
  candidati.sort((a, b) => b.prob - a.prob);
  return candidati[0]; // { ambito: 'totale'|'home'|'away', label, prob }
}

// venueLetter (richiesta Marco 25/10: "tutte le info nel pannello devono essere riportate
// esclusivamente al venue") — prima questa funzione guardava TUTTA la stagione del
// giocatore, mescolando casa e trasferta; ora solo le presenze nel venue reale di questa gara.
function narrFoulCardWatch(matchLog, rosters, teamId, venueLetter) {
  const roster = rosters.filter((r) => r.Team_ID_Rif === teamId && r.Attivo_SerieA === 'SI');
  const agg = roster.map((r) => {
    const rows = matchLog.filter((x) => x.Player_ID === r.Player_ID && x.Venue === venueLetter);
    const min = rows.reduce((a, x) => a + (Number(x.Minuti) || 0), 0);
    const falli = rows.reduce((a, x) => a + (Number(x.Falli_commessi) || 0), 0);
    const gialli = rows.reduce((a, x) => a + (Number(x.Gialli) || 0), 0);
    const rossi = rows.reduce((a, x) => a + (Number(x.Rossi) || 0), 0);
    const partite = new Set(rows.map((x) => x.Match_ID)).size;
    return { nome: r.Nome_completo, min, falli, gialli, rossi, partite, falli90: min ? (falli / min) * 90 : null };
  });
  const hotCard = agg.filter((p) => p.min >= 200 && p.gialli >= 2).sort((a, b) => (b.gialli / b.partite) - (a.gialli / a.partite))[0] || null;
  const dueDiligence = agg.filter((p) => p.min >= 150 && p.gialli === 0 && p.rossi === 0 && p.falli90 != null).sort((a, b) => b.falli90 - a.falli90)[0] || null;
  return { hotCard, dueDiligence };
}

function narrStreaksFor(calendar, teamId, venue) {
  const keys = ['segna', 'subisce', 'nonSubisce', 'nonPerde', 'nonVince', 'nonPareggia'];
  return keys.map((k) => ({ key: k, ...computeTeamStreak(calendar, teamId, venue, k) })).filter((s) => s.partite >= 2);
}

// Ruolo difensivo/di centrocampo che tipicamente incontra ciascun ruolo offensivo (approssimazione
// posizionale, non un dato di marcatura reale — usata solo per proporre incroci individuali
// plausibili). Comprende anche le coppie speculari (il terzino che si propone trova l'esterno
// avversario, i due centrocampisti centrali si incontrano direttamente) per avere più candidati.
const NARR_ROLE_MARCATURA = {
  'Esterno/Quinto': 'Terzino',
  'Punta centrale': 'Centrale difensivo',
  'Trequartista': 'Centrocampista centrale',
  'Terzino': 'Esterno/Quinto',
  'Centrocampista centrale': 'Centrocampista centrale',
};

// Incrocio individuale sui falli: cerca, tra i ruoli offensivi con una controparte difensiva
// definita, i giocatori che SUBISCONO più falli (FD, atteso in questa partita) il cui marcatore
// naturale è chi NE COMMETTE più (FC) — tipo "Zaccagni subisce 2,5 falli a partita e gioca vs
// Celik che ne commette 1". Richiede entrambe le proiezioni sopra una soglia minima per non
// forzare l'incrocio quando i numeri non dicono nulla. Restituisce FINO A 2 coppie (ruoli
// diversi, nessun giocatore ripetuto), non solo la migliore, per dare più carne individuale.
function narrFoulMatchup(model, matchLog, rosters, attackTeamId, attackPv, defendTeamId, defendPv) {
  const cands = [];
  for (const role in NARR_ROLE_MARCATURA) {
    const defRole = NARR_ROLE_MARCATURA[role];
    const attackers = narrPlayersInRole(rosters, matchLog, attackTeamId, role, 300).slice(0, 3);
    const defenders = narrPlayersInRole(rosters, matchLog, defendTeamId, defRole, 300).slice(0, 2);
    if (!attackers.length || !defenders.length) continue;
    for (const a of attackers) {
      const fd = projectPlayer(model, matchLog, { id: a.id, team: attackTeamId, ruolo: role }, defendTeamId, true, 'FD', attackPv);
      if (fd.mu == null || fd.mu < 1.5) continue;
      for (const def of defenders) {
        const fc = projectPlayer(model, matchLog, { id: def.id, team: defendTeamId, ruolo: defRole }, attackTeamId, true, 'FC', defendPv);
        if (fc.mu == null || fc.mu < 1.0) continue;
        cands.push({ attacker: a, defender: def, role, defRole, fdMu: fd.mu, fcMu: fc.mu, score: fd.mu + fc.mu });
      }
    }
  }
  cands.sort((a, b) => b.score - a.score);
  const picked = [];
  const usedRoles = new Set();
  for (const c of cands) {
    if (usedRoles.has(c.role)) continue; // niente due coppie sullo stesso ruolo offensivo
    picked.push(c);
    usedRoles.add(c.role);
    if (picked.length >= 2) break;
  }
  return picked;
}

// Incrocio individuale sui cartellini: un attaccante che subisce molti falli (FD) — quindi
// spesso al centro di contatti/proteste — contro un difensore già nervoso sui cartellini
// (ammonito con frequenza, da narrFoulCardWatch.hotCard): due profili che insieme aumentano
// il rischio di un'ammonizione, non solo una somma statistica scollegata.
function narrCardMatchup(model, matchLog, rosters, attackTeamId, attackPv, defendTeamId, defendPv, defendFoulWatch) {
  const hotCard = defendFoulWatch && defendFoulWatch.hotCard;
  if (!hotCard || hotCard.partite < 2) return null;
  let best = null;
  for (const role in NARR_ROLE_MARCATURA) {
    const defRole = NARR_ROLE_MARCATURA[role];
    const defenders = narrPlayersInRole(rosters, matchLog, defendTeamId, defRole, 300);
    if (!defenders.some((p) => p.nome === hotCard.nome)) continue; // il giocatore a rischio deve essere proprio in quel ruolo di marcatura
    const attackers = narrPlayersInRole(rosters, matchLog, attackTeamId, role, 300).slice(0, 3);
    for (const a of attackers) {
      const fd = projectPlayer(model, matchLog, { id: a.id, team: attackTeamId, ruolo: role }, defendTeamId, true, 'FD', attackPv);
      if (fd.mu == null || fd.mu < 1.5) continue;
      if (!best || fd.mu > best.fdMu) best = { attacker: a, defender: { nome: hotCard.nome }, role, defRole, fdMu: fd.mu, cardRate: hotCard.gialli / hotCard.partite };
    }
  }
  return best;
}

// Forma recente (ultime N gare, qualunque venue) — punti a partita e record, per intercettare
// una squadra in slancio o in calo che il solo dato sul venue di oggi non racconta.
function narrRecentForm(calendar, teamId, n) {
  const rows = calendar.filter((r) => (r.Home_ID === teamId || r.Away_ID === teamId) &&
    r.Stato === 'CONCLUSA' && r.Home_Score !== '' && r.Home_Score != null);
  rows.sort((a, b) => b.Data.localeCompare(a.Data));
  const last = rows.slice(0, n);
  let w = 0, d = 0, l = 0, pts = 0, gf = 0, ga = 0;
  for (const r of last) {
    const isHome = r.Home_ID === teamId;
    const f = Number(isHome ? r.Home_Score : r.Away_Score), a = Number(isHome ? r.Away_Score : r.Home_Score);
    gf += f; ga += a;
    if (f > a) { w++; pts += 3; } else if (f < a) { l++; } else { d++; pts += 1; }
  }
  return { n: last.length, w, d, l, pts, gf, ga, ppg: last.length ? pts / last.length : null };
}

function generateMatchNarrative(matchLog, calendar, rosters, match) {
  const homeId = match.Home_ID, awayId = match.Away_ID;
  const homeName = match.Home, awayName = match.Away;

  // --- Preview: classifica generale vs venue reale ---
  const standings = narrStandingsTable(calendar);
  const sHome = standings.find((s) => s.id === homeId);
  const sAway = standings.find((s) => s.id === awayId);
  const homeAtHome = narrVenueRecord(calendar, homeId, 'H');
  const awayAtAway = narrVenueRecord(calendar, awayId, 'A');
  const homeAtAway = narrVenueRecord(calendar, homeId, 'A');
  const awayAtHome = narrVenueRecord(calendar, awayId, 'H');

  // Corner e cartellini medi per venue (stesso schema dei gol, nessun nuovo calcolo:
  // riusa computeTeamMatchTotals già testato per il Matchup Engine).
  const homeTotalsHome = computeTeamMatchTotals(matchLog, homeId, 'HOME', calendar);
  const awayTotalsAway = computeTeamMatchTotals(matchLog, awayId, 'AWAY', calendar);
  const homeCornerCart = { corner: narrAvgOf(homeTotalsHome, 'Corner'), cart: narrAvgOf(homeTotalsHome, 'Cart') };
  const awayCornerCart = { corner: narrAvgOf(awayTotalsAway, 'Corner'), cart: narrAvgOf(awayTotalsAway, 'Cart') };

  // "Dominanza fragile": il record di risultato nel venue (V/N/P) confrontato col record del
  // duello sui tiri partita-su-partita (chi dei due ha tirato di più, stessa gara) — una squadra
  // può vincere le partite e perdere comunque il duello sui tiri più spesso di quanto vinca.
  const homeDuelsHome = computeTeamDuelTotals(matchLog, homeId, 'HOME', calendar);
  const awayDuelsAway = computeTeamDuelTotals(matchLog, awayId, 'AWAY', calendar);
  const homeShotDuel = teamDuelRecord(homeDuelsHome, 'Tiri');
  const awayShotDuel = teamDuelRecord(awayDuelsAway, 'Tiri');

  // Mismatch volume/qualità dei tiri: volume = tiri a partita, qualità = quota che arriva in
  // porta (SOT/Tiri) — confrontato direttamente tra i due venue reali di oggi.
  const sumVals = (totals, key) => { const vs = totals.map((t) => t[key]).filter((v) => v != null); return vs.length ? vs.reduce((a, b) => a + b, 0) / vs.length : null; };
  const homeShotProfile = { volume: sumVals(homeTotalsHome, 'Tiri'), sot: sumVals(homeTotalsHome, 'SOT'), n: homeTotalsHome.length };
  const awayShotProfile = { volume: sumVals(awayTotalsAway, 'Tiri'), sot: sumVals(awayTotalsAway, 'SOT'), n: awayTotalsAway.length };
  homeShotProfile.qualita = homeShotProfile.volume ? homeShotProfile.sot / homeShotProfile.volume : null;
  awayShotProfile.qualita = awayShotProfile.volume ? awayShotProfile.sot / awayShotProfile.volume : null;

  // Trend recente: ultime 5 gare giocate, qualunque venue — la forma del momento, non solo il venue di oggi.
  const homeRecentForm = narrRecentForm(calendar, homeId, 5);
  const awayRecentForm = narrRecentForm(calendar, awayId, 5);

  // --- Chiave di lettura: ruolo più bucato di ciascuna squadra nel proprio venue (su Tiri
  // O su Falli subiti dal ruolo, qualunque dei due scarti rispetto alla media di lega sia più
  // marcato), più un'eventuale seconda debolezza vicina e una debolezza specifica sui tiri in porta ---
  const homePartiteHome = narrPartiteInVenue(matchLog, homeId, 'H');
  const awayPartiteAway = narrPartiteInVenue(matchLog, awayId, 'A');
  const homeWeakX = homePartiteHome ? narrWeakestRoleMulti(matchLog, homeId, 'H', homePartiteHome) : { primary: null, secondary: null, sotWeak: null };
  const awayWeakX = awayPartiteAway ? narrWeakestRoleMulti(matchLog, awayId, 'A', awayPartiteAway) : { primary: null, secondary: null, sotWeak: null };
  const homeWeak = homeWeakX.primary ? homeWeakX.primary.row : null;
  const awayWeak = awayWeakX.primary ? awayWeakX.primary.row : null;
  const homeWeakMetric = homeWeakX.primary ? homeWeakX.primary.metric : null;
  const awayWeakMetric = awayWeakX.primary ? awayWeakX.primary.metric : null;
  // Quanto quel ruolo è DAVVERO sopra la media di lega per lo stesso ruolo/metrica/venue —
  // distingue una vera falla (ratio alto) dal "meno buono tra i propri ruoli" ma comunque nella norma.
  const homeWeakRatio = homeWeakX.primary ? homeWeakX.primary.ratio : null;
  const awayWeakRatio = awayWeakX.primary ? awayWeakX.primary.ratio : null;

  // interprete della squadra ospite che sfrutta il ruolo debole di casa, e viceversa — con numeri
  // individuali (mu Tiri attesi in questa partita, metrica coerente con la falla sui tiri) invece
  // del solo nome, così l'incrocio dice qualcosa in più di "gioca in quel ruolo".
  const modelEarly = buildLeagueModel(matchLog);
  function withMuTiri(list, teamId, oppId, pv) {
    return list.map((p) => {
      const pr = projectPlayer(modelEarly, matchLog, { id: p.id, team: teamId, ruolo: p.ruolo }, oppId, true, 'Tiri', pv);
      const venueRows = matchLog.filter(x => x.Player_ID === p.id && x.Team_ID === teamId && x.Venue === pv);
      const venueMinutes = venueRows.reduce((total,x)=>total+(Number(x.Minuti)||0),0);
      const venueShots = venueRows.reduce((total,x)=>total+(Number(x.Tiri)||0),0);
      return { ...p, mu: pr.mu, venueMinutes, venueShots,
        venueGames: venueRows.length, venueTiri90: venueMinutes >= 90 ? 90*venueShots/venueMinutes : null };
    }).sort((a, b) => (b.mu || 0) - (a.mu || 0));
  }
  const awayExploiters = homeWeak ? withMuTiri(narrPlayersInRole(rosters, matchLog, awayId, homeWeak.Posizione_Analitica, 300), awayId, homeId, 'A') : [];
  const homeExploiters = awayWeak ? withMuTiri(narrPlayersInRole(rosters, matchLog, homeId, awayWeak.Posizione_Analitica, 300), homeId, awayId, 'H') : [];
  // Il ruolo "debole" di una squadra conta solo se L'AVVERSARIO DI QUESTA PARTITA schiera
  // davvero un giocatore in quel ruolo — altrimenti è un dato vero ma non sfruttabile oggi
  // (bug corretto: prima si segnalava la falla anche contro avversari senza quel ruolo).
  const homeWeakOppHasRole = homeWeak ? narrPlayersInRole(rosters, matchLog, awayId, homeWeak.Posizione_Analitica, 0).length > 0 : null;
  const awayWeakOppHasRole = awayWeak ? narrPlayersInRole(rosters, matchLog, homeId, awayWeak.Posizione_Analitica, 0).length > 0 : null;

  const shootEff = narrShootingEfficiency(matchLog, calendar, awayId, 'A');

  // --- Nomi che contano: per ogni giocatore si valutano Tiri/SOT/Falli commessi/Falli subiti e si
  // tiene la metrica dove il giocatore è davvero sopra il proprio ruolo (baseline90/prior90 più alto),
  // non sempre i tiri. La soglia mostrata è proporzionale al valore atteso (mu): soglia intera =
  // floor(mu) [minimo 1], riga comunicata = soglia - 0,5 (es. mu=4,14 -> "almeno 3,5"; mu=3,19 -> "almeno 2,5"). ---
  const model = modelEarly; // stesso modello già costruito sopra per gli exploiters, non ricalcolato
  function topNames(teamId, oppId, pv, excludeGK) {
    const roster = rosters.filter((r) => r.Team_ID_Rif === teamId && r.Attivo_SerieA === 'SI' && (!excludeGK || r.Posizione_Analitica !== 'Portiere'));
    return roster.map((r) => {
      const rows = matchLog.filter((x) => x.Player_ID === r.Player_ID);
      const min = rows.reduce((a, x) => a + (Number(x.Minuti) || 0), 0);
      if (min < 150) return null;
      const player = { id: r.Player_ID, team: teamId, ruolo: r.Posizione_Analitica };
      const metricsAll = ['Tiri', 'SOT', 'FC', 'FD'].map((m) => {
        const p = projectPlayer(model, matchLog, player, oppId, true, m, pv);
        const ratio = p.baseline90 != null && p.prior90 ? p.baseline90 / p.prior90 : null;
        return { metric: m, p, ratio };
      }).filter((x) => x.p.mu != null && x.p.mu > 0);
      if (!metricsAll.length) return null;
      const withRatio = metricsAll.filter((x) => x.ratio != null);
      const standout = withRatio.length
        ? withRatio.sort((a, b) => b.ratio - a.ratio)[0]
        : metricsAll.sort((a, b) => b.p.mu - a.p.mu)[0];
      const tiriEntry = metricsAll.find((x) => x.metric === 'Tiri');
      const muTiri = tiriEntry ? tiriEntry.p.mu : null;
      // Arrotondamento alla soglia PIÙ VICINA (non sempre verso il basso): mu=1,99 deve dare
      // soglia 2 (riga 1,5), non soglia 1 (riga 0,5) come farebbe un floor puro.
      const soglia = Math.max(1, Math.round(standout.p.mu));
      return {
        nome: r.Nome_completo, ruolo: r.Posizione_Analitica, min,
        mu: muTiri, // resta sui Tiri: usato per ordinamento e per l'incrocio col ruolo debole (volume di tiro)
        affidabilita: (tiriEntry || standout).p.affidabilita,
        metric: standout.metric, metricMu: standout.p.mu, metricRatio: standout.ratio,
        soglia, sogliaLine: soglia - 0.5,
        pSoglia: probAtLeastAlpha(standout.p.mu, soglia, MODEL_PARAMS.ALPHA[standout.metric] || 0),
      };
    }).filter(Boolean).sort((a, b) => (b.mu || 0) - (a.mu || 0));
  }
  const homeNames = topNames(homeId, awayId, 'H', true);
  const awayNames = topNames(awayId, homeId, 'A', true);

  // --- Radar disciplinare (anticipato: serve anche ai matchup individuali qui sotto) ---
  const foulHome = narrFoulCardWatch(matchLog, rosters, homeId, 'H');
  const foulAway = narrFoulCardWatch(matchLog, rosters, awayId, 'A');

  // --- Matchup individuali: più stat, più giocatori, non solo il migliore in assoluto ---
  // 1) Falli: un attaccante che SUBISCE molti falli (FD) contro il difensore del ruolo che
  //    tipicamente lo marca, se quel difensore NE COMMETTE molti (FC) — es. Zaccagni/Celik.
  //    Fino a 2 coppie per lato (ruoli diversi). Soglie: FD atteso >= 1,5, FC atteso >= 1,0.
  const foulMatchupHome = narrFoulMatchup(model, matchLog, rosters, homeId, 'H', awayId, 'A');
  const foulMatchupAway = narrFoulMatchup(model, matchLog, rosters, awayId, 'A', homeId, 'H');
  // 2) Cartellini: lo stesso tipo di attaccante (FD alto) contro un difensore già nervoso sui
  //    cartellini (ammonito spesso, narrFoulCardWatch.hotCard) nel ruolo che lo marcherebbe.
  const cardMatchupHome = narrCardMatchup(model, matchLog, rosters, homeId, 'H', awayId, 'A', foulAway);
  const cardMatchupAway = narrCardMatchup(model, matchLog, rosters, awayId, 'A', homeId, 'H', foulHome);

  // nome principale per squadra: il migliore per mu
  const homeLead = homeNames[0] || null;
  const awayLead = awayNames[0] || null;
  // nome "da cautela": miglior mu la cui concessione di ruolo (lato avversario, venue corretto) ha 0 precedenti
  function zeroPrecedentPick(names, oppWeakRows) {
    return names.find((n) => {
      const row = oppWeakRows.find((r) => r.Posizione_Analitica === n.ruolo);
      return row && row.N_Prestazioni_Ruolo === 0;
    }) || null;
  }
  // per questo serve concessione per-ruolo completa (non solo la peggiore) di chi difende
  function allRoleConcessions(teamId, venue, teamPartite) {
    const segmento = venue === 'H' ? 'HOME' : 'AWAY';
    return NARRATIVE_ROLES.map((r) => computeRoleConcessions(matchLog, teamId, segmento, r, teamPartite));
  }
  const homeDefRows = homePartiteHome ? allRoleConcessions(homeId, 'H', homePartiteHome) : [];
  const awayDefRows = awayPartiteAway ? allRoleConcessions(awayId, 'A', awayPartiteAway) : [];
  const cautelaAway = zeroPrecedentPick(awayNames, homeDefRows); // giocatore ospite in ruolo mai visto dalla difesa di casa
  const cautelaHome = zeroPrecedentPick(homeNames, awayDefRows);

  // nome "matchup sfavorevole": alta affidabilità ma il suo ruolo è il MENO concesso dall'avversario
  function sfavoritoPick(names, defRows) {
    const validDef = defRows.filter((r) => r.N_Prestazioni_Ruolo >= 3 && r.Tiri_Concessi_90 != null);
    if (!validDef.length) return null;
    const minRow = validDef.reduce((a, b) => (b.Tiri_Concessi_90 < a.Tiri_Concessi_90 ? b : a));
    return names.find((n) => n.ruolo === minRow.Posizione_Analitica && n.affidabilita === 'alta') || null;
  }
  const sfavoritoHome = sfavoritoPick(homeNames, awayDefRows);
  const sfavoritoAway = sfavoritoPick(awayNames, homeDefRows);

  // --- Multigol ---
  const goals = narrProjectGoals(calendar, homeId, awayId);

  // --- Sequenze per venue ---
  const seqHome = narrStreaksFor(calendar, homeId, 'HOME');
  const seqAway = narrStreaksFor(calendar, awayId, 'AWAY');

  const result = {
    homeId, awayId, homeName, awayName,
    standings: { home: sHome, away: sAway },
    venue: {
      homeAtHome, awayAtAway, homeAtAway, awayAtHome, homeCornerCart, awayCornerCart,
      homeShotDuel, awayShotDuel, homeShotProfile, awayShotProfile, homeRecentForm, awayRecentForm,
    },
    chiaveLettura: {
      homeWeak, awayWeak, homeWeakMetric, awayWeakMetric, homeWeakRatio, awayWeakRatio,
      homeWeakOppHasRole, awayWeakOppHasRole,
      homeWeak2: homeWeakX.secondary ? homeWeakX.secondary.row : null,
      awayWeak2: awayWeakX.secondary ? awayWeakX.secondary.row : null,
      homeWeakSOT: homeWeakX.sotWeak, awayWeakSOT: awayWeakX.sotWeak,
      awayExploiters, homeExploiters, shootEff,
      // Lato campo (richiesta Marco 25/10: evidenziare se terzino/esterno destro o sinistro, non
      // solo il ruolo generico) — solo per la metrica Tiri, unica con una colonna raw affidabile
      // (stesso raw col 'Tiri' già usato da computeRoleCrossExtremes); null altrove, nessun dato inventato.
      homeWeakLato: (homeWeak && homeWeakMetric === 'Tiri') ? roleConcessionsSideBreakdown(matchLog, homeId, 'HOME', homeWeak.Posizione_Analitica, 'Tiri') : null,
      awayWeakLato: (awayWeak && awayWeakMetric === 'Tiri') ? roleConcessionsSideBreakdown(matchLog, awayId, 'AWAY', awayWeak.Posizione_Analitica, 'Tiri') : null,
    },
    nomi: {
      homeLead, awayLead, cautelaHome, cautelaAway, sfavoritoHome, sfavoritoAway, homeNames, awayNames,
      foulMatchupHome, foulMatchupAway, cardMatchupHome, cardMatchupAway,
    },
    multigol: goals,
    sequenze: { home: seqHome, away: seqAway },
    disciplina: { home: foulHome, away: foulAway },
  };
  result.temi = narrDiscoverThemes(result);
  result.preview = narrDynamicPreview(result, result.temi);
  result.chiaveHtml = narrChiaveHtml(result);
  return result;
}

const NARR_ROLE_PLURAL = {
  'Punta centrale': 'punte centrali', 'Trequartista': 'trequartisti', 'Esterno/Quinto': 'esterni',
  'Centrocampista centrale': 'centrocampisti centrali', 'Terzino': 'terzini', 'Centrale difensivo': 'centrali difensivi',
};
// Forme con articolo (anche contratto con preposizione) per ogni ruolo — "punta centrale" è
// femminile ("le punte centrali"), gli altri sono maschili; servono le tre forme perché "a"/"su"
// si contraggono con l'articolo ("ai trequartisti", non "a i trequartisti").
const NARR_ROLE_BARE_ART = {
  'Punta centrale': 'le punte centrali', 'Trequartista': 'i trequartisti', 'Esterno/Quinto': 'gli Esterni/Quinti',
  'Centrocampista centrale': 'i centrocampisti centrali', 'Terzino': 'i terzini', 'Centrale difensivo': 'i centrali difensivi',
};
const NARR_ROLE_A = {
  'Punta centrale': 'alle punte centrali', 'Trequartista': 'ai trequartisti', 'Esterno/Quinto': 'agli Esterni/Quinti',
  'Centrocampista centrale': 'ai centrocampisti centrali', 'Terzino': 'ai terzini', 'Centrale difensivo': 'ai centrali difensivi',
};
const NARR_ROLE_SU = {
  'Punta centrale': 'sulle punte centrali', 'Trequartista': 'sui trequartisti', 'Esterno/Quinto': 'sugli Esterni/Quinti',
  'Centrocampista centrale': 'sui centrocampisti centrali', 'Terzino': 'sui terzini', 'Centrale difensivo': 'sui centrali difensivi',
};

// Confronta, per UNA squadra, il rendimento nel venue di oggi con quello nell'ALTRO suo venue
// (non con l'avversario) — è questo lo scarto che racconta se la squadra sta sopra o sotto la
// sua stessa media abituale, il cuore della "vera analisi": non solo i numeri, ma cosa dicono
// confrontati con l'altra faccia della stessa squadra.
function narrVenueSplitClause(teamName, ownRec, otherRec, ownWord, otherWord) {
  const ppg = (r) => (r && r.n ? (r.w * 3 + r.d) / r.n : null);
  if (!ownRec || !otherRec || ownRec.n < 2 || otherRec.n < 2) return null;
  const ownPPG = ppg(ownRec), otherPPG = ppg(otherRec);
  if (ownPPG == null || otherPPG == null) return null;
  const gap = ownPPG - otherPPG;
  if (Math.abs(gap) < 0.6) return null;
  const recTxt = (r) => `${r.w}V ${r.d}N ${r.l}P in ${r.n} gare (${r.gf}-${r.ga})`;
  if (gap > 0) {
    const totLosses = ownRec.l + otherRec.l;
    const lossNote = (totLosses >= 3 && otherRec.l / totLosses >= 0.6)
      ? ` ${otherRec.l} delle ${totLosses} sconfitte stagionali pesano proprio ${otherWord}, non ${ownWord}.`
      : '';
    return { dir: 'up', gap, text: `<b>${narrTeam(teamName, true)}</b> ${ownWord} ${recTxt(ownRec)} (${(ownPPG).toFixed(2).replace('.', ',')} punti a partita) — meglio di quanto renda persino ${otherWord} (${recTxt(otherRec)}).${lossNote}` };
  }
  return { dir: 'down', gap, text: `<b>${narrTeam(teamName, true)}</b> ${ownWord} non è la squadra che vediamo ${otherWord}: lì fa ${recTxt(otherRec)}, ma ${ownWord} scende a ${recTxt(ownRec)} (${(ownPPG).toFixed(2).replace('.', ',')} punti a partita) — un calo netto, non un dettaglio.` };
}

// "Vera analisi": quando la classifica generale racconta una storia (es. 17ª contro 4ª) ma lo
// spacchettamento per venue ne racconta un'altra — ogni squadra confrontata con SE STESSA
// nell'altro venue, non solo le due squadre tra loro — si scrive il paragrafo lungo in stile
// analista, non la semplice frase di verdetto. Richiede un vero scarto di classifica (>=6
// posizioni) E almeno una delle due squadre con un salto di rendimento reale tra i suoi venue:
// altrimenti resta la versione sobria più sotto, per non forzare questo arco narrativo ovunque.
function narrAnalisiVera(d) {
  const posH = d.standings.home && d.standings.home.pos, posA = d.standings.away && d.standings.away.pos;
  if (posH == null || posA == null || Math.abs(posH - posA) < 6) return null;
  const homeClause = narrVenueSplitClause(d.homeName, d.venue.homeAtHome, d.venue.homeAtAway, 'in casa', 'in trasferta');
  const awayClause = narrVenueSplitClause(d.awayName, d.venue.awayAtAway, d.venue.awayAtHome, 'in trasferta', 'in casa');
  if (!homeClause && !awayClause) return null;
  const hi = Math.max(posH, posA), lo = Math.min(posH, posA);
  // Lessico variegato: lo stesso identico incipit su 190 partite su 380 sarebbe più stucchevole
  // del problema che risolve. La scelta è deterministica (stessa gara → stesso testo se richiamata
  // due volte) ma varia da partita a partita, pescando dal contenuto reale (posizioni, direzione
  // dei due scarti), non da un indice casuale scollegato dal match.
  const dir = (homeClause ? (homeClause.dir === 'up' ? 1 : -1) : 0) + (awayClause ? (awayClause.dir === 'up' ? 1 : -1) : 0);
  const leads = [
    `Il dato che ribalta la lettura prima ancora di entrare nei dettagli: presa per intero, è una sfida tra ${hi}ª e ${lo}ª della classe. Scomposta per venue, è tutt'altra storia.`,
    `Sulla carta è ${hi}ª contro ${lo}ª, un verdetto che parrebbe già scritto. Scomposta per venue, la gara cambia faccia.`,
    `Tra le due c'è lo scarto netto di una ${hi}ª contro una ${lo}ª in classifica generale — ma è un quadro che il dettaglio dei venue smentisce quasi subito.`,
    `${hi}ª contro ${lo}ª: la classifica generale indicherebbe una gara senza storia. I numeri scomposti per venue raccontano una partita diversa.`,
    `Il piazzamento (${hi}ª contro ${lo}ª) suggerisce un copione già scritto. Il rendimento reale nei rispettivi venue lo riscrive.`,
  ];
  const idx = Math.abs(hi * 7 + lo * 3 + dir * 11) % leads.length;
  const parts = [leads[idx]];
  if (awayClause) parts.push(awayClause.text);
  if (homeClause) parts.push(homeClause.text);
  return parts.join(' ');
}

// --- PREVIEW: verdetto sul reale favorito stasera (confronto diretto dei venue, non ciascuno
// contro la propria media stagionale) + eventuale sorpresa classifica/rendimento. Niente debolezze
// di ruolo qui — quelle sono materia della Chiave di lettura, per non ripetersi tra le due sezioni.

/* Preview 2.0: scoperta temi, ordinamento per rilievo e campione.
 * I punteggi sono priorità editoriali, NON probabilità di una scommessa. */
function narrDiscoverThemes(d) {
  const v=d.venue, c=d.chiaveLettura, found=[];
  const f=x=>Number(x).toFixed(1).replace('.', ',');
  const ppg=r=>r&&r.n?(r.w*3+r.d)/r.n:null;
  const add=(id,domain,title,evidence,limit,signal,n)=>{
    if(n<2 || !Number.isFinite(signal)) return;
    const score=Math.round(55*Math.min(1,signal)+20+15*Math.min(1,n/6)+10);
    found.push({id,domain,title,evidence,limit,n,score});
  };
  const h=v.homeShotProfile, a=v.awayShotProfile;
  if(h&&a&&h.n>=3&&a.n>=3&&h.volume!=null&&a.volume!=null){
    const dv=h.volume-a.volume;
    if(Math.abs(dv)>=2.5){
      const team=dv>0?d.homeName:d.awayName;
      add('volume','tiri','Tiri: '+(dv>0?d.homeName:d.awayName)+' nettamente sopra',
        narrTeam(team,true)+' produce '+f(Math.max(h.volume,a.volume))+' tiri a gara nel venue della sfida contro '+f(Math.min(h.volume,a.volume))+' dell’altra squadra: '+f(Math.abs(dv))+' conclusioni di scarto.',
        'Si confrontano i volumi prodotti, non le concessioni difensive: non è una previsione dei tiri di questa gara.',
        Math.abs(dv)/6,Math.min(h.n,a.n));
    }
    if(h.qualita!=null&&a.qualita!=null&&Math.abs(h.qualita-a.qualita)>=.07){
      const team=h.qualita>a.qualita?d.homeName:d.awayName;
      add('precisione','precisione','Tiri in porta: '+(h.qualita>a.qualita?d.homeName:d.awayName)+' più precisa',
        narrTeam(team,true)+' manda in porta il '+Math.round(Math.max(h.qualita,a.qualita)*100)+'% dei propri tiri, contro il '+Math.round(Math.min(h.qualita,a.qualita)*100)+'% dell’avversaria nei rispettivi venue.',
        'La percentuale di tiri nello specchio non è la qualità delle occasioni né la probabilità di segnare.',
        Math.abs(h.qualita-a.qualita)/.18,Math.min(h.n,a.n));
    }
  }
  const ph=ppg(v.homeAtHome),pa=ppg(v.awayAtAway);
  if(ph!=null&&pa!=null&&v.homeAtHome.n>=2&&v.awayAtAway.n>=2&&Math.abs(ph-pa)>=.65){
    const team=ph>pa?d.homeName:d.awayName;
    add('venue','rendimento','Rendimento nel venue: '+team+' avanti',
      narrTeam(team,true)+' ha ottenuto '+f(Math.max(ph,pa))+' punti per gara '+(ph>pa?'in casa':'in trasferta')+', contro '+f(Math.min(ph,pa))+' dell’altra squadra nel suo venue.',
      'Il risultato storico è influenzato anche dal calendario degli avversari: non dimostra superiorità nel gioco.',
      Math.abs(ph-pa)/1.5,Math.min(v.homeAtHome.n,v.awayAtAway.n));
  }
  const roles=[
    {team:d.homeName,opp:d.awayName,weak:c.homeWeak,ratio:c.homeWeakRatio,metric:c.homeWeakMetric,ok:c.homeWeakOppHasRole,n:v.homeAtHome?v.homeAtHome.n:0,players:c.awayExploiters,venueWord:'in casa',lato:c.homeWeakLato},
    {team:d.awayName,opp:d.homeName,weak:c.awayWeak,ratio:c.awayWeakRatio,metric:c.awayWeakMetric,ok:c.awayWeakOppHasRole,n:v.awayAtAway?v.awayAtAway.n:0,players:c.homeExploiters,venueWord:'in trasferta',lato:c.awayWeakLato}
  ];
  for(const z of roles){
    if(!z.weak||z.ok!==true||z.ratio==null||z.ratio<1.15) continue;
    const role=NARR_ROLE_A[z.weak.Posizione_Analitica]||('al ruolo '+z.weak.Posizione_Analitica);
    const field=z.metric==='Falli'?'FalliSubiti_Ruolo_90':'Tiri_Concessi_90';
    const value=z.weak[field];
    const player=z.metric==='Tiri'&&z.players&&z.players.length
      ? z.players.filter(p=>p.venueMinutes>=120&&p.venueShots>=1)
          .sort((a,b)=>(b.venueTiri90||0)-(a.venueTiri90||0))[0]||null : null;
    const playerNote=player
      ? ' Nel venue di questa gara, '+player.nome+' ('+z.opp+') ha prodotto '+player.venueShots+' tiri in '+player.venueGames+' partite ('+player.venueMinutes+' minuti, '+f(player.venueTiri90)+' tiri/90).'
      : '';
    // Lato campo (richiesta Marco 25/10: dire se destro/sinistro, non solo "terzino") — solo
    // quando il ruolo è largo e il dato supera soglia/dominanza (roleConcessionsSideBreakdown).
    const latoNote=z.lato
      ? ' Lato: su '+z.lato.totale+' episodi, '+Math.max(z.lato.sinistra,z.lato.destra,z.lato.centro)+' da '+(z.lato.sinistra>=z.lato.destra&&z.lato.sinistra>=z.lato.centro?'sinistra':z.lato.destra>=z.lato.centro?'destra':'centro')+'.'
      : '';
    const title=z.metric==='Falli'
      ? z.opp+': '+(z.weak.Posizione_Analitica||'ruolo')+' e falli contro '+z.team
      : z.opp+': '+(z.weak.Posizione_Analitica||'ruolo')+' contro la difesa '+z.team;
    // Lessico semplificato (richiesta Marco 25/10): una sola frase che nomina chi concede/dove e
    // chi beneficia, invece delle due frasi ridondanti "Squadra che concede / Squadra che può
    // beneficiare" che dicevano la stessa cosa in due modi diversi.
    add('role'+z.team,'matchup',title,
      narrTeam(z.team,true)+' concede '+f(value)+' '+(z.metric==='Falli'?'falli commessi contro ':'tiri concessi ')+role+' a partita '+z.venueWord+' (+'+Math.round((z.ratio-1)*100)+'% sulla media di lega per lo stesso ruolo/venue): un varco di cui '+narrTeam(z.opp)+' può provare ad approfittare.'+latoNote+playerNote,
      player ? 'L’incrocio merita verifica sui tiri di '+player.nome+' solo se confermato in quel ruolo e con minuti sufficienti; nessuna indicazione di valore senza quota.'
        : 'Il segnale di reparto non è sostenuto da un interprete con volume di tiri sufficiente nel venue: non promuoverlo a opportunità individuale.',
      Math.min(1,(z.ratio-1)/.65)*(player?Math.min(1.25,Math.max(.25,player.venueTiri90/1.4)):0.25),z.n);
  }
  const hc=v.homeCornerCart,ac=v.awayCornerCart;
  if(hc&&ac&&hc.corner!=null&&ac.corner!=null&&v.homeAtHome&&v.awayAtAway&&v.homeAtHome.n>=2&&v.awayAtAway.n>=2){
    const diff=hc.corner-ac.corner;
    if(Math.abs(diff)>=2){
      const team=diff>0?d.homeName:d.awayName;
      add('corner','corner','Corner prodotti: '+team+' avanti',
        narrTeam(team,true)+' conquista '+f(Math.max(hc.corner,ac.corner))+' corner per partita, contro '+f(Math.min(hc.corner,ac.corner))+' dell’avversaria nei rispettivi venue.',
        'Sono corner prodotti, non concessi: il confronto da solo non stima i corner della gara.',
        Math.abs(diff)/4,Math.min(v.homeAtHome.n,v.awayAtAway.n));
    }
  }
  found.sort((x,y)=>y.score-x.score||x.id.localeCompare(y.id));
  const selected=[],domains=new Set();
  for(const t of found){
    if(domains.has(t.domain))continue;
    selected.push(t);domains.add(t.domain);
    if(selected.length===3)break;
  }
  return selected;
}
function narrDynamicPreview(d,temi){
  const home=narrTeam(d.homeName),away=narrTeam(d.awayName);
  if(!temi.length)return [
    'Il confronto tra '+home+' in casa e '+away+' in trasferta non restituisce un vantaggio statistico dominante con i dati attualmente disponibili. Non emerge un tema abbastanza solido per costruire una lettura univoca della gara.',
    'Per il momento la valutazione resta aperta: saranno decisive le formazioni e le concessioni effettive per ruolo. Non utilizziamo dati delle partite nel venue opposto per riempire la Preview.'
  ];
  const first=temi[0],second=temi.find(x=>x.domain!==first.domain);
  const paragraphs=[];
  if(first.domain==='matchup'){
    paragraphs.push('La chiave statistica di '+d.homeName+'–'+d.awayName+' è un incrocio tra ruolo offensivo e difesa avversaria. '+first.evidence+' Se chi occupa quel ruolo viene impiegato con continuità, quel reparto potrebbe trovare più occasioni per la propria statistica specifica, pur senza rendere automatico il rendimento dei singoli.');
  } else if(first.domain==='tiri'){
    paragraphs.push('Il primo squilibrio della sfida riguarda la frequenza delle conclusioni. '+first.evidence+' Questo suggerisce una diversa propensione a cercare il tiro, non necessariamente un maggiore numero di occasioni pulite: sarà la risposta difensiva dell’avversaria a determinare se il divario si ripeterà.');
  } else if(first.domain==='precisione'){
    paragraphs.push('La lettura della gara parte da come le due squadre trasformano i tentativi in tiri nello specchio. '+first.evidence+' La differenza è interessante perché il semplice totale dei tiri potrebbe raccontare una partita diversa rispetto alle conclusioni che impegnano il portiere.');
  } else if(first.domain==='rendimento'){
    paragraphs.push('Il confronto tra il rendimento casalingo di '+home+' e quello esterno di '+away+' propone una prima gerarchia. '+first.evidence+' I risultati, però, vanno letti insieme alla produzione offensiva e a ciò che ciascuna squadra concede: da soli non descrivono il controllo della gara.');
  } else {
    paragraphs.push('Un dato distingue questa partita dalle altre: '+first.evidence+' '+first.limit);
  }
  if(second){
    const conjunction=second.domain==='matchup'
      ? 'Sul piano dei singoli ruoli emerge anche una seconda possibile direzione della gara. '
      : second.domain==='rendimento' ? 'Il quadro dei risultati nel venue aggiunge un elemento diverso. '
      : second.domain==='tiri' ? 'Guardando invece al volume offensivo, emerge un secondo contrasto. '
      : second.domain==='precisione' ? 'C’è però una distinzione importante tra volume di gioco e precisione delle conclusioni. '
      : 'Un secondo segnale merita attenzione. ';
    paragraphs.push(conjunction+second.evidence);
  }
  const n=Math.min(...temi.slice(0,second?2:1).map(t=>t.n));
  let counter='Il punto che potrebbe smentire questa lettura è la scarsa profondità del campione ('+n+' gare per almeno uno degli indicatori).';
  if(first.domain==='matchup') counter+=' Inoltre, senza la conferma del ruolo e dei minuti del giocatore segnalato, il mismatch resta soltanto potenziale.';
  else if(first.domain==='tiri'||first.domain==='precisione') counter+=' Il dato sui tiri prodotti non incorpora automaticamente quelli concessi dall’avversaria.';
  else counter+=' La forza degli avversari già incontrati può aver influenzato il rendimento.';
  paragraphs.push(counter);
  return paragraphs;
}

function narrPreview(d) {
  const out = [];
  const ppgGen = (s) => (s && s.pg ? s.pts / s.pg : null);
  const ppgVenue = (rec) => (rec && rec.n ? (rec.w * 3 + rec.d) / rec.n : null);
  const hOwn = d.venue.homeAtHome, aOwn = d.venue.awayAtAway;
  const ppgH = ppgVenue(hOwn), ppgA = ppgVenue(aOwn);

  const analisiVera = narrAnalisiVera(d);
  if (analisiVera) {
    out.push(analisiVera);
    if (ppgH != null && ppgA != null) {
      const diff = ppgH - ppgA;
      out.push(Math.abs(diff) <= 0.3
        ? `Tradotto in termini da campo: più equilibrio di quanto suggerisca la sola classifica, senza un favorito netto imposto dal rendimento reale nei rispettivi venue.`
        : `Tradotto in termini da campo: oggi non si affrontano una corazzata e una squadra allo sbando, ma il rendimento reale nei rispettivi venue — il favore del pronostico resta dalla parte ${narrTeamOf(diff > 0 ? d.homeName : d.awayName)} — su una cornice tattica più equilibrata di quanto il piazzamento in classifica suggerisca.`);
    }
  } else if (hOwn && aOwn && hOwn.n >= 2 && aOwn.n >= 2 && ppgH != null && ppgA != null) {
    const diff = ppgH - ppgA;
    const recTxt = (name, rec) => `${rec.w}V ${rec.d}N ${rec.l}P in ${rec.n} gare (${rec.gf}-${rec.ga})`;
    if (Math.abs(diff) <= 0.3) {
      out.push(`Nei rispettivi venue il rendimento reale è quasi identico — <b>${narrTeam(d.homeName)}</b> in casa ${recTxt(d.homeName, hOwn)}, <b>${narrTeam(d.awayName)}</b> in trasferta ${recTxt(d.awayName, aOwn)} — più equilibrio di quanto suggerisca il solo piazzamento in classifica.`);
    } else if (diff >= 0.8) {
      out.push(`Chi arriva meglio stasera non è deciso dalla classifica ma dal venue: <b>${narrTeam(d.homeName)}</b> vale <b>${ppgH.toFixed(2)}</b> punti a partita in casa (${recTxt(d.homeName, hOwn)}) contro i <b>${ppgA.toFixed(2)}</b> <b>${narrTeamOf(d.awayName)}</b> in trasferta (${recTxt(d.awayName, aOwn)}): sulla forma reale è ${narrTeam(d.homeName)} la favorita.`);
    } else if (diff <= -0.8) {
      out.push(`Chi arriva meglio stasera non è deciso dalla classifica ma dal venue: <b>${narrTeam(d.awayName)}</b> vale <b>${ppgA.toFixed(2)}</b> punti a partita in trasferta (${recTxt(d.awayName, aOwn)}) contro i <b>${ppgH.toFixed(2)}</b> <b>${narrTeamOf(d.homeName)}</b> in casa (${recTxt(d.homeName, hOwn)}): sulla forma reale è ${narrTeam(d.awayName)} la favorita, nonostante giochi fuori casa.`);
    } else {
      out.push(`Nei rispettivi venue <b>${narrTeam(d.homeName)}</b> fa <b>${ppgH.toFixed(2)}</b> punti a partita in casa (${recTxt(d.homeName, hOwn)}) e <b>${narrTeam(d.awayName)}</b> <b>${ppgA.toFixed(2)}</b> in trasferta (${recTxt(d.awayName, aOwn)}): un margine leggero, non una gara già scritta.`);
    }
  } else {
    out.push(`Campione ancora minimo nel venue di oggi: servono più partite prima che il rendimento reale casa/trasferta diventi un indizio affidabile.`);
  }

  // Sorpresa classifica vs rendimento venue — solo se il gap è marcato, con frase completa (niente
  // ellissi tipo "X vale N": si spiega SEMPRE punti-a-partita generale vs punti-a-partita nel venue.
  // Saltata se è già scattata la "vera analisi" sopra: direbbe la stessa cosa due volte.
  const cands = analisiVera ? [] : [
    { team: d.homeName, gen: ppgGen(d.standings.home), ven: ppgH, side: 'in casa', pos: d.standings.home && d.standings.home.pos },
    { team: d.awayName, gen: ppgGen(d.standings.away), ven: ppgA, side: 'in trasferta', pos: d.standings.away && d.standings.away.pos },
  ].filter((x) => x.gen != null && x.ven != null).map((x) => ({ ...x, gap: x.ven - x.gen }));
  if (cands.length) {
    cands.sort((a, b) => Math.abs(b.gap) - Math.abs(a.gap));
    const g = cands[0];
    if (Math.abs(g.gap) >= 0.6) {
      out.push(g.gap > 0
        ? `Da segnalare: <b>${narrTeam(g.team)}</b> è ${g.pos}ª in classifica generale con una media di <b>${g.gen.toFixed(2)}</b> punti a partita, ma ${g.side} rende molto di più — <b>${g.ven.toFixed(2)}</b> punti a partita — un dato migliore di quanto il piazzamento lasci intuire.`
        : `Da segnalare: <b>${narrTeam(g.team)}</b> è ${g.pos}ª in classifica generale con una media di <b>${g.gen.toFixed(2)}</b> punti a partita, ma ${g.side} scende a <b>${g.ven.toFixed(2)}</b> — un rendimento inferiore a quanto il piazzamento lasci intuire.`);
    }
  }

  // --- Dominanza fragile: vince le partite ma perde il duello sui tiri più spesso di quanto
  // vinca — un vantaggio che sulla carta dei tiri non si vede, e che può non ripetersi stasera.
  function dominanzaFragile(teamName, venRec, duel, side) {
    if (!venRec || venRec.n < 3 || !duel || duel.n < 3 || duel.rate == null) return null;
    const winRate = venRec.n ? venRec.w / venRec.n : null;
    if (winRate == null) return null;
    const gap = winRate - duel.rate;
    if (gap < 0.3 || winRate < 0.45) return null; // serve un vero scarto, e che la squadra vinca davvero spesso
    return `<b>Dominanza fragile</b>: ${narrTeam(teamName)} ${side} ha vinto <b>${Math.round(winRate * 100)}%</b> delle gare (${venRec.w}/${venRec.n}), ma ha tirato di più dell'avversario solo nel <b>${Math.round(duel.rate * 100)}%</b> dei casi (${duel.wins}/${duel.n}) — un rendimento che sul piano dei tiri è meno solido di quanto dica il risultato, e che un episodio può ribaltare.`;
  }
  const domHome = dominanzaFragile(d.homeName, hOwn, d.venue.homeShotDuel, 'in casa');
  const domAway = dominanzaFragile(d.awayName, aOwn, d.venue.awayShotDuel, 'in trasferta');
  if (domHome) out.push(domHome);
  if (domAway) out.push(domAway);

  // --- Mismatch volume/qualità dei tiri: chi tira più spesso contro chi è più preciso —
  // due profili diversi (molto volume, poca precisione / poco volume, molto cinismo) dicono
  // più cose di un confronto puro sulla media tiri.
  const hp = d.venue.homeShotProfile, ap = d.venue.awayShotProfile;
  if (hp && ap && hp.n >= 3 && ap.n >= 3 && hp.volume != null && ap.volume != null && hp.qualita != null && ap.qualita != null) {
    const volDiff = hp.volume - ap.volume, qualDiff = hp.qualita - ap.qualita;
    if (Math.abs(volDiff) >= 2.5 && (volDiff > 0) !== (qualDiff > 0) && Math.abs(qualDiff) >= 0.05) {
      const volTeam = volDiff > 0 ? d.homeName : d.awayName, volSide = volDiff > 0 ? 'in casa' : 'in trasferta';
      const qualTeam = qualDiff > 0 ? d.homeName : d.awayName, qualSide = qualDiff > 0 ? 'in casa' : 'in trasferta';
      const volVal = volDiff > 0 ? hp.volume : ap.volume, qualVal = qualDiff > 0 ? hp.qualita : ap.qualita;
      out.push(`<b>Mismatch volume/qualità dei tiri</b>: ${narrTeam(volTeam)} ${volSide} è la squadra che tira più spesso (<b>${volVal.toFixed(1).replace('.', ',')}</b> tiri a partita), ma è ${narrTeam(qualTeam)} ${qualSide} quella più precisa, con il <b>${Math.round(qualVal * 100)}%</b> dei tiri che arriva in porta — due modi opposti di proporsi in avanti, non lo stesso tipo di minaccia.`);
    }
  }

  // --- Trend recente: ultime 5 gare, qualunque venue — la forma del momento può contraddire
  // il quadro del solo venue di oggi (una squadra in crisi o in striscia positiva a prescindere da dove gioca).
  function trendText(teamName, form) {
    if (!form || form.n < 4) return null;
    return `${narrTeam(teamName)} nelle ultime <b>${form.n}</b> gare (qualunque venue) fa <b>${form.w}V ${form.d}N ${form.l}P</b> (${form.gf}-${form.ga}, <b>${form.ppg.toFixed(2).replace('.', ',')}</b> punti a partita)`;
  }
  const trH = trendText(d.homeName, d.venue.homeRecentForm), trA = trendText(d.awayName, d.venue.awayRecentForm);
  if (trH || trA) {
    const parts = [trH, trA].filter(Boolean);
    out.push(`<b>Trend recente</b>: ${parts.join('; ')}.`);
  }

  return out;
}

// --- CHIAVE DI LETTURA: non sempre "debolezza" — solo se il ruolo individuato concede/subisce
// davvero più della media di lega (ratio >= 1.15); altrimenti si dice esplicitamente che non c'è
// una vera falla. La metrica può essere Tiri concessi O Falli subiti dal ruolo avversario (quella
// con lo scarto più marcato), non è più forzata sempre sui tiri.
const NARR_WEAK_FIELD = { Tiri: 'Tiri_Concessi_90', Falli: 'FalliSubiti_Ruolo_90' };
function narrWeakSentence(teamName, venueWord, weak, ratio, metric, oppHasRole) {
  if (!weak) return `<b>${narrTeam(teamName, true)}</b> ${venueWord} non ha un campione sufficiente, in nessun ruolo, per isolare un punto debole affidabile.`;
  const field = NARR_WEAK_FIELD[metric] || 'Tiri_Concessi_90';
  const val = weak[field];
  const unit = metric === 'Falli' ? 'falli a partita' : 'tiri a partita';
  const roleA = NARR_ROLE_A[weak.Posizione_Analitica];
  const roleSu = NARR_ROLE_SU[weak.Posizione_Analitica];
  const roleBare = NARR_ROLE_BARE_ART[weak.Posizione_Analitica];
  // La falla è vera (ratio alto) ma inutile IN QUESTA PARTITA se l'avversario di oggi non
  // schiera nessun giocatore in quel ruolo — va detto chiaro, non lasciato implicito.
  const nonSfruttabile = oppHasRole === false
    ? ' Resta però un dato teorico in questa partita: l\'avversario di oggi non schiera un giocatore in quel ruolo.'
    : '';
  if (ratio == null) {
    const verbo = metric === 'Falli' ? 'commette più falli contro' : 'concede più tiri';
    return `<b>${narrTeam(teamName, true)}</b> ${venueWord} ${verbo} a partita ${roleA} (<span class="narr-hl">${val.toFixed(1).replace('.', ',')}</span>) che a chiunque altro — non basta ancora il campione di lega per dire se è una vera falla o solo il meno solido dei suoi reparti.`;
  }
  if (ratio >= 1.3) {
    const extreme = ratio >= 2.0 ? ' — scarto così ampio da trattare con cautela: a questo punto della stagione la media di lega su quel ruolo/venue si basa ancora su poche squadre con campione sufficiente.' : '';
    const desc = metric === 'Falli'
      ? `ha una falla vera sulla disciplina ${roleSu}: commette <span class="narr-hl-bad">${val.toFixed(1).replace('.', ',')} falli a partita</span> su quel ruolo avversario`
      : `ha una falla vera ${roleSu}: concede <span class="narr-hl-bad">${val.toFixed(1).replace('.', ',')} tiri a partita</span>`;
    return `<b>${narrTeam(teamName, true)}</b> ${venueWord} ${desc}, il <b>${Math.round((ratio - 1) * 100)}%</b> in più della media di lega per quel ruolo in questo venue.${extreme}${nonSfruttabile}`;
  }
  if (ratio >= 1.15) {
    const desc = metric === 'Falli'
      ? `commette sopra media falli ${roleA} (<span class="narr-hl">${val.toFixed(1).replace('.', ',')}</span> falli a partita, circa il <b>${Math.round((ratio - 1) * 100)}%</b> sopra la media di lega)`
      : `concede sopra media ${roleA} (<span class="narr-hl">${val.toFixed(1).replace('.', ',')}</span> tiri a partita, circa il <b>${Math.round((ratio - 1) * 100)}%</b> sopra la media di lega)`;
    return `<b>${narrTeam(teamName, true)}</b> ${venueWord} ${desc} — un punto da sfruttare, non ancora una voragine.${nonSfruttabile}`;
  }
  return `<b>${narrTeam(teamName, true)}</b> ${venueWord} non ha in realtà una vera debolezza di reparto: anche il suo dato più esposto (${roleBare}, <span class="narr-hl">${val.toFixed(1).replace('.', ',')}</span> ${unit}) resta in linea con la media di lega per quel ruolo in questo venue.`;
}
function narrChiaveHtml(d) {
  const c = d.chiaveLettura;
  const hw = c.homeWeak, aw = c.awayWeak;
  if (!hw && !aw) return '<p>Campione ancora troppo piccolo per isolare una debolezza di ruolo affidabile su una delle due difese in questo venue.</p>';
  const sentences = [
    narrWeakSentence(d.homeName, 'in casa', hw, c.homeWeakRatio, c.homeWeakMetric, c.homeWeakOppHasRole),
    narrWeakSentence(d.awayName, 'in trasferta', aw, c.awayWeakRatio, c.awayWeakMetric, c.awayWeakOppHasRole),
  ];
  const homeGenuine = hw && c.homeWeakRatio != null && c.homeWeakRatio >= 1.15;
  const awayGenuine = aw && c.awayWeakRatio != null && c.awayWeakRatio >= 1.15;
  if (homeGenuine && c.homeWeak2) {
    const field = NARR_WEAK_FIELD[c.homeWeakMetric] || 'Tiri_Concessi_90';
    const unit = c.homeWeakMetric === 'Falli' ? 'falli a partita' : 'tiri a partita';
    sentences.push(`Per ${narrTeam(d.homeName)} non è un buco isolato: il dato resta alto (<span class="narr-hl">${c.homeWeak2[field].toFixed(1).replace('.', ',')}</span> ${unit}) anche ${NARR_ROLE_A[c.homeWeak2.Posizione_Analitica]}.`);
  }
  if (awayGenuine && c.awayWeak2) {
    const field = NARR_WEAK_FIELD[c.awayWeakMetric] || 'Tiri_Concessi_90';
    const unit = c.awayWeakMetric === 'Falli' ? 'falli a partita' : 'tiri a partita';
    sentences.push(`Anche per ${narrTeam(d.awayName)} non è un buco isolato: il dato resta alto (<span class="narr-hl">${c.awayWeak2[field].toFixed(1).replace('.', ',')}</span> ${unit}) anche ${NARR_ROLE_A[c.awayWeak2.Posizione_Analitica]}.`);
  }
  if (homeGenuine && c.homeWeakSOT) sentences.push(`Sui tiri <b>in porta</b> concessi la falla ${narrTeamOf(d.homeName)} in casa è diversa: pesa soprattutto ${NARR_ROLE_SU[c.homeWeakSOT.Posizione_Analitica]} (<span class="narr-hl">${c.homeWeakSOT.SOT_Concessi_90.toFixed(1).replace('.', ',')} SOT a partita</span>).`);
  if (awayGenuine && c.awayWeakSOT) sentences.push(`Sui tiri <b>in porta</b> concessi ${narrTeam(d.awayName)} in trasferta sconta soprattutto ${NARR_ROLE_BARE_ART[c.awayWeakSOT.Posizione_Analitica]} (<span class="narr-hl">${c.awayWeakSOT.SOT_Concessi_90.toFixed(1).replace('.', ',')} SOT a partita</span>).`);

  // Exploiter con numero individuale (mu Tiri attesi in questa partita), non solo il nome: dice
  // quanto quel giocatore specifico è messo per sfruttare la falla, non solo che gioca lì.
  function exploiterText(teamName, list) {
    const withMu = list.filter((p) => p.mu != null).slice(0, 2);
    if (!withMu.length) return `${narrTeam(teamName, true)} ha in rosa ${list.slice(0, 2).map((p) => p.nome).join(' e ')}, proprio nel ruolo che sfrutta quel buco.`;
    const names = withMu.map((p) => `<b>${p.nome}</b> (${fmtNum(p.mu)} tiri attesi)`).join(' e ');
    return `${narrTeam(teamName, true)} risponde con ${names}, proprio nel ruolo che sfrutta quel buco.`;
  }
  function fmtNum(n) { return n.toFixed(1).replace('.', ','); }
  const exploiters = [];
  if (homeGenuine && c.awayExploiters.length) exploiters.push(exploiterText(d.awayName, c.awayExploiters));
  if (awayGenuine && c.homeExploiters.length) exploiters.push(exploiterText(d.homeName, c.homeExploiters));

  // Incrocio col nome più atteso: si dice qualcosa SOLO quando coincide davvero col ruolo debole
  // genuino — citare un nome solo per dire che non è lui a sfruttare il buco non aggiunge nulla.
  const crosses = [];
  if (awayGenuine && d.nomi.homeLead && d.nomi.homeLead.ruolo === aw.Posizione_Analitica) {
    crosses.push(`Non è un caso che il nome più atteso ${narrTeamOf(d.homeName)} sia <b>${d.nomi.homeLead.nome}</b>: gioca esattamente nel ruolo dove ${narrTeam(d.awayName)} concede più spazio in trasferta.`);
  }
  if (homeGenuine && d.nomi.awayLead && d.nomi.awayLead.ruolo === hw.Posizione_Analitica) {
    crosses.push(`Stesso discorso per ${narrTeam(d.awayName)}: <b>${d.nomi.awayLead.nome}</b> gioca esattamente nel ruolo dove ${narrTeam(d.homeName)} concede più spazio in casa.`);
  }

  // Conferma indipendente: campione minimo 3 gare, e il verbo segue il dato reale invece di
  // affermare sempre la stessa cosa a prescindere dal risultato (bug corretto).
  let confText = '';
  if (c.shootEff && c.shootEff.n >= 3) {
    const rate = c.shootEff.clinical / c.shootEff.n;
    confText = rate >= 0.5
      ? `Conferma indipendente: ${narrTeam(d.awayName)} in trasferta ha messo in porta almeno metà dei tiri tentati in ${c.shootEff.clinical} gare su ${c.shootEff.n} — non spreca quando trova lo spazio.`
      : `Da tenere d'occhio: ${narrTeam(d.awayName)} in trasferta ha messo in porta almeno metà dei tiri tentati solo in ${c.shootEff.clinical} gare su ${c.shootEff.n} — anche trovando spazio, fatica a essere clinico.`;
  }

  const lead = homeGenuine || awayGenuine
    ? (homeGenuine && awayGenuine ? 'Due falle di reparto vere, non solo la peggiore tra le proprie.' : 'Una sola vera falla di reparto tra le due squadre — l\'altra difesa tiene meglio di quanto sembri.')
    : 'Nessuna vera falla di reparto da nessuna delle due parti in questo venue: differenze minime, nella norma di lega.';

  // Matchup individuali sui falli e sui cartellini (calcolati a parte, in d.nomi): qui diventano
  // frasi con nomi e numeri, non solo l'elenco di ruoli deboli.
  const matchupSentences = [];
  for (const fm of [...(d.nomi.foulMatchupHome || []), ...(d.nomi.foulMatchupAway || [])]) {
    matchupSentences.push(`<b>Match-up individuale sui falli</b>: <b>${fm.attacker.nome}</b> (${fm.role.toLowerCase()}) subisce in media <span class="narr-hl">${fm.fdMu.toFixed(1).replace('.', ',')} falli</span> a partita, e trova di fronte <b>${fm.defender.nome}</b> (${fm.defRole.toLowerCase()}), che ne commette <span class="narr-hl">${fm.fcMu.toFixed(1).replace('.', ',')}</span>.`);
  }
  for (const cm of [d.nomi.cardMatchupHome, d.nomi.cardMatchupAway]) {
    if (!cm) continue;
    matchupSentences.push(`<b>Match-up sui cartellini</b>: <b>${cm.attacker.nome}</b> (${cm.role.toLowerCase()}) subisce in media <span class="narr-hl">${cm.fdMu.toFixed(1).replace('.', ',')} falli</span> a partita e incontra <b>${cm.defender.nome}</b>, già ammonito in <span class="narr-hl-warn">${Math.round(cm.cardRate * 100)}%</span> delle sue gare — un contatto che può pesare su un cartellino già in bilico.`);
  }

  return `<p class="narr-lead">${lead}</p><p>${sentences.join(' ')}</p>${exploiters.length ? `<p>${exploiters.join(' ')}</p>` : ''}${crosses.length ? `<p>${crosses.join(' ')}</p>` : ''}${confText ? `<p>${confText}</p>` : ''}${matchupSentences.length ? `<p>${matchupSentences.join(' ')}</p>` : ''}`;
}

if (typeof module !== 'undefined') {
  module.exports = {
    buildIndexes,
    computePlayerStats,
    computeTeamStats,
    rankDesc,
    computeRoleConcessions,
    computeRoleProduction,
    computeRoleCrossExtremes,
    roleConcessionsSideBreakdown,
    playerPrevalentSide,
    computeBaseline,
    rankProvisional,
    wilsonInterval,
    computeArbitroStats,
    computeAllArbitroStats,
    rankArbitriFacts,
    computeGoalLineFacts,
    teamGoalMatches,
    computeTeamMatchTotals,
    teamHitRate,
    leagueAverageMetric,
    narrPartiteInVenue,
    narrPlayersInRole,
    narrRoleVsOpponent,
    narrRoleConcessionRank,
    fattiRuoloVsAvversario,
    FATTI_RUOLO_MIN_SAMPLE_VENUE,
    NARRATIVE_ROLES,
    NARR_ROLE_METRICS,
    computeTeamDuelTotals,
    teamDuelRecord,
    multiplicativeProjection,
    restDaysBefore,
    computeTeamStreak,
    normalizeArbitroKey,
    titleCaseArbitro,
    isArbitroNomeValido,
    teamCornerMatches,
    computeCardDrought,
    computeFoulCardProbabilityTable,
    estimateCardProbability,
    BASELINE_THRESHOLDS,
    SHRINK_K,
    // --- v2
    MODEL_PARAMS, MODEL_METRICS, metricValue, buildLeagueModel, playerRows,
    shrunkRate90, teamStrengthFactor, opponentFactor, expectedMinutes, reliabilityOf, projectPlayer,
    pmfNB, probAtLeast, probAtLeastAlpha, minOddsFromProb, playerHitRateK, cardsPerFoulLeague,
    probCardFromFouls, projectLineupTotal,
    // --- narrativo
    generateMatchNarrative,
    narrGoalBandaPiuPlausibile,
    computeTeamLineCrossExtremes,
  };
}
