# Protocollo per gli aggiornamenti di giornata — database Serie A 2026/27

**Scopo.** Aggiornare le partite senza alterare il roster già corretto, attribuire statistiche al giocatore e alla squadra giusti e consegnare risultati controllabili. Questo documento integra `02_DATABASE_SERIE_A_PROTOCOL.md` e `06_AUDIT_CHECKLIST.md`. Si applica al master operativo più recente e, se presenti, ai file derivati e al sito.

## 1. Individuare il file e congelare ciò che non cambia

- Usare la versione più recente ricevuta, registrando nome, data, impronta SHA-256, schede, intestazioni, dimensioni, formule, convalide e righe compilate. Conservare una copia immutata per il confronto finale. Non scegliere una versione precedente per comodità.
- L'attuale master allegato ha i fogli `CALENDAR`, `MATCH_LOG`, `ROSTERS`, `ARBITRI`. Le giornate non ancora disputate sono presenti in `CALENDAR`: distinguere righe programmate e partite concluse. Non scambiare l'estensione del calendario per il numero di match già popolati.
- **Preservare il roster e i ruoli già corretti manualmente:** usare come base il master più recente che contiene le tue correzioni di `Posizione_Analitica`. Nell'aggiornamento di giornata non ricalcolare né sovrascrivere i ruoli in `ROSTERS` o nelle righe storiche di `MATCH_LOG` ricorrendo a ruoli di siti esterni, formazioni, posizioni nominali o versioni precedenti del file. Non riscrivere, riordinare, normalizzare o eliminare righe esistenti senza motivo documentato. Registrare i nuovi trasferimenti accertati come interventi nominativi separati, con fonte, data di efficacia e audit prima/dopo; lasciare invariati i giocatori per i quali non risulta un cambiamento. Non dedurre un trasferimento dai soli minuti, dall'assenza in una lista o da una discordanza tra squadra storica e attuale.
- Non modificare partite già chiuse per far tornare un totale. Una rettifica storica richiede evidenza della stessa statistica, del giocatore e della partita.

## 2. Stabilire la fonte e il significato di ogni valore

Per ogni nuova partita registrare `Match_ID`, giornata, squadre, risultato, fonte e data di acquisizione; distinguere valori per partita, cumulativi stagionali, medie e valori per 90 minuti. Usare la stessa definizione per tutte le giornate e annotare divergenze fra fonti. I tabellini di squadra e le liste giocatori devono riferirsi **alla stessa partita e alla stessa metrica**. Non dedurre quote, minuti o statistiche dalla semplice presenza in rosa; non inserire decimali inventati né trasformare dati non disponibili in zero.

| Campo | Regola di inserimento |
| --- | --- |
| `CALENDAR` | Una riga per `Match_ID`; stato, casa/trasferta, punteggio e corner da fonte verificata. Non marcare conclusa una partita senza risultato verificato. |
| `MATCH_LOG` | Una riga per giocatore e partita, identificata da `Player_ID` + `Match_ID`; `Team_ID` è la squadra **con cui ha giocato quel match**, `Opp_ID` l'avversario, `Venue` H/A dal calendario. Per le nuove righe riportare `Posizione_Analitica` dal `ROSTERS` corrente tramite `Player_ID`, senza inferirla dalla fonte della partita. |
| Minuti e presenza | Minuti effettivi documentati; zero reale distinto da vuoto/non disponibile. Segnalare subito eventi positivi con minuti assenti o zero, minuti impossibili o statistiche attribuite a un omonimo. |
| Tiri, falli, cartellini | Valori individuali della stessa partita; verificare definizione di tiri in porta, falli commessi/subiti, gialli e rossi (compreso il caso della seconda ammonizione). Nessuna correzione per semplice simmetria di aggregati. |
| `ARBITRI` | Un solo nome per `Match_ID` verificato dalla designazione/tabellino. Una stringa di orario (`h. 20.45`, `18:30`), `TBD` o un valore numerico in `Arbitro` è un errore o dato non disponibile, mai il nome di un secondo arbitro. |

## 3. Trasferimenti: non spostare la storia

I conteggi storici di squadra, gli split casa/trasferta, ranking e concessioni per ruolo aggregano `MATCH_LOG.Team_ID` e `Match_ID`. Menu e shortlist attuali usano `ROSTERS.Team_ID_Rif` con `Attivo_SerieA = SI`. Un ceduto `NO` rimane nel log delle partite già giocate; `NO` non equivale a zero minuti. Una riga storica con squadra diversa da quella attuale **non è di per sé un errore**: controllare il trasferimento e la data della partita.

Per un giocatore passato tra club di Serie A, produrre due campioni separati: **con la squadra attuale** (`MATCH_LOG.Team_ID = ROSTERS.Team_ID_Rif`) e **complessivo stagionale** (tutte le sue righe per `Player_ID`). Mostrare presenze/minuti per entrambi e non attribuire automaticamente i dati della vecchia squadra alla tesi pre-gara della nuova. Per esempio, il tiro di Rowe col Bologna in G1 rimane al Bologna, mentre Rowe può figurare tra i giocatori attuali dell'Atalanta. I 13 `NO` già confermati come ceduti non vanno cambiati sulla base delle presenze storiche.

Per ciascun trasferimento accertato aggiornare la squadra attuale e `Attivo_SerieA` della riga corrispondente secondo il criterio «`SI` solo se il giocatore appartiene alla `Squadra_Rif` indicata», preservando `Player_ID` e tutte le righe storiche di `MATCH_LOG`. Se il giocatore passa a un altro club di Serie A, verificare che sia selezionabile nella nuova squadra e non nella vecchia; se esce dalla Serie A, conservarne lo storico e toglierlo dalle shortlist attuali. Controllare i conteggi della squadra vecchia e nuova prima/dopo: i valori delle partite concluse non devono spostarsi.

**Registro dei dubbi per controllo umano.** Durante l'acquisizione segnare ogni omonimia, nome incompleto o variato, ruolo discordante, ID ambiguo, squadra inattesa, trasferimento senza fonte o data certa e stato `SI/NO` incoerente con fonti credibili. Cercare prima nel roster, nel tabellino della partita e in una fonte indipendente accessibile; se l'identità o il trasferimento resta incerto, presentare a Marco un elenco nominativo con `Player_ID`/`Match_ID`, partita, squadra nel file, alternativa possibile, fonti esaminate, motivo del dubbio e decisione richiesta. Lasciare sospesa **la sola attribuzione o modifica interessata** fino al suo riscontro: non scegliere un omonimo per somiglianza del nome, non cambiare `Attivo_SerieA` per ipotesi e non attribuire statistiche a caso. Registrare la risposta o la prova che risolve il dubbio nell'audit.

**Ruoli congelati.** Le correzioni manuali già presenti nel master sono la classificazione operativa per le analisi per ruolo. Una discordanza con SofaScore o altra fonte si segnala nel registro dei dubbi, ma non autorizza a cambiare `Posizione_Analitica`. Se un nuovo giocatore non ha ancora un ruolo nel roster, registrare il caso e chiedere la classificazione prima di inserirlo nelle concessioni per ruolo; non assegnargli un ruolo per analogia. Una futura modifica dei ruoli richiesta esplicitamente da Marco è un'operazione separata: aggiornare per `Player_ID` sia `ROSTERS.Posizione_Analitica` sia tutte le righe collegate di `MATCH_LOG.Posizione_Analitica`, poi verificare gli eventuali derivati e il sito. Non cambiare tiri, falli, minuti, squadra della partita o stato di attività per effetto della sola modifica del ruolo.

## 4. Riconciliare i falli senza fabbricare correzioni

Tre confronti distinti, con esito e fonte registrati:

1. **Lista individuale vs fonte individuale:** stessi giocatori, partita, squadra e metrica. Una differenza identificata e documentata può motivare la modifica di una cella.
2. **Somma individuale vs totale ufficiale della stessa squadra e stessa metrica:** verificare se la copertura e le definizioni coincidono. Nell'attuale `CALENDAR` il totale ufficiale dei falli non è memorizzato: non chiamare “totale ufficiale” la somma derivata da `MATCH_LOG`. Se si propone di aggiungere campi per tiri/falli squadra, inserirli solo da tabellino indipendente e con modifica di schema esplicitata.
3. **Falli commessi da A vs falli subiti da B:** registrare lo scarto, ma **non usarlo da solo come prova di un giocatore mancante o di un errore**. Definizioni e copertura possono divergere. Non aumentare un valore individuale per far quadrare le somme.

Ogni scarto resta classificato come `dato errato documentato`, `differenza di fonte/definizione`, `copertura da verificare` oppure `non risolto`; non descrivere tutte queste categorie come “errori nei falli”. Le immagini ricevute e già controllate non vanno richieste di nuovo.

## 5. Chiedere screenshot soltanto per una verifica che può cambiare il file

Prima esaminare il master, le fonti accessibili e gli screenshot già inviati. Chiedere un'immagine **solo** quando manca la prova necessaria per decidere una correzione specifica. Indicare in un'unica richiesta:

- giornata, partita, squadra e metrica esatta (per esempio `falli commessi → Lecce`, non “screenshot dei falli”);
- giocatore e valore contestato, se già individuati, oppure il totale ufficiale preciso che serve per distinguere due ipotesi;
- cosa è stato verificato, cosa manca e **quale cella o valutazione potrà cambiare** in base alla schermata;
- necessità di scorrere la lista per mostrare tutti i giocatori, solo se pertinente.

Se entrambe le liste coincidono col file, fermarsi: non chiedere altre foto per il solo scarto incrociato. Se la fonte non è ottenibile e lo screenshot non può sciogliere il dubbio, registrare il limite senza inventare un valore. Correggere solo dopo il confronto del dettaglio ricevuto.

## 6. Audit prima della consegna

| Controllo | Esito necessario |
| --- | --- |
| Identità e partite | `Player_ID` di ogni riga in `ROSTERS`; `Match_ID` in `CALENDAR`; nessuna coppia giocatore-partita duplicata; squadra, avversario e H/A coerenti col calendario della partita. Confrontare anche `ARBITRI` per ID e squadre. |
| Omonimie e ruolo | Identità e `Player_ID` verificati sulla fonte; `Posizione_Analitica` di ogni riga `MATCH_LOG`, storica o nuova, uguale a quella del `ROSTERS` corrente per lo stesso ID. Il ruolo della fonte può differire dalla classificazione manuale: registrare il dubbio, non sovrascrivere il master. Caso sentinella: Napoli–Bologna G4, 90 minuti a **Massimo Pessina**, portiere del Bologna, non a Matteo Pessina del Monza. |
| Attività e roster | `ROSTERS` invariato salvo trasferimenti o rettifiche nominative documentati; nessuna modifica accidentale a `Posizione_Analitica` in `ROSTERS` e `MATCH_LOG`. Nessun giocatore disponibile escluso dalle viste attuali a causa di un `NO` errato; nessun ceduto mostrato come disponibile. Le presenze storiche dei `NO` rimangono consultabili; i dubbi su identità, stato e ruolo sono elencati per controllo umano. |
| Minuti e derivati | Eventi con 0 minuti/vuoto segnalati; se esiste un foglio `PLAYERS` o dashboard, totali di presenze/minuti/tiri/falli e indicatori per 90 confrontati su casi nominativi e sull'intera tabella con `MATCH_LOG`. Il problema di Douglas Luiz (minuti nel log, riepilogo a zero/vuoto) deve impedire di dichiarare valido quel derivato. |
| Arbitri | Nessun orario al posto del nome, nessun `Match_ID` duplicato; nomi con fonte. Casi sentinella: Lazio–Genoa G2 **Feliciani** e Napoli–Como G2 **Pairetto**. |
| Statistiche e fonte | Liste individuali verificate; anomalie dei falli classificate come al §4; le cifre pubblicate per squadra e giocatore derivano dalle colonne corrette. |
| Integrità file | Confronto cella per cella con la copia iniziale: ogni differenza deve essere nella lista delle modifiche intenzionali. Controllare formule, riferimenti, formattazione e convalide se presenti; riaprire l'output e verificare i risultati calcolati nel motore Excel compatibile quando si consegnano formule. |
| Pubblicazione | Se si aggiorna il sito, esaminare codice del generatore e `data.json`, rigenerare dalla versione validata e confrontare numeri/filtri pubblicati. Un file Excel corretto non prova che il sito sia stato aggiornato. |

Un audit degli ID dimostra **coerenza interna**, non accuratezza del tabellino esterno né correttezza dei riepiloghi. Registrare entrambi i livelli separatamente. Se una verifica critica fallisce, correggere e ripetere l'audit prima della consegna; se non è risolvibile, non presentare il file come validato per quell'uso.

## 7. Consegna e traccia delle modifiche

Consegnare insieme al file una breve scheda: versione d'ingresso e versione d'uscita; giornate e `Match_ID` aggiunti/corretti; numero e elenco delle celle variate (foglio, riga/ID, prima → dopo, fonte); conferma che ruoli già corretti e `ROSTERS` sono invariati **oppure elenco nominativo dei trasferimenti accertati**; controllo della corrispondenza dei ruoli per `Player_ID` tra `ROSTERS` e `MATCH_LOG`; controlli eseguiti e risultati; dubbi su identità, omonimie, stato e ruolo sottoposti a Marco con esito o ancora aperti; anomalie residue con impatto pratico; screenshot ancora necessari **solo** se soddisfano il §5. Distinguere sempre `verificato sulla fonte`, `coerente internamente`, `dedotto` e `non verificato`.

Non chiamare “definitivo” un master per tutte le analisi se è stato verificato soltanto che il file si apre o che gli ID si collegano. Se un derivato o il sito non è stato controllato, dirlo esplicitamente.
