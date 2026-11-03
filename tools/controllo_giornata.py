#!/usr/bin/env python3
"""Controllo ruoli/lati di una nuova giornata contro i dati già pubblicati.

Uso:
  python3 tools/controllo_giornata.py NUOVO.xlsx [data.json] [--applica]

- data.json = dati pubblicati (default: ./data.json) = riferimento confermato.
- Controlla solo le righe MATCH_LOG di partite NON presenti in data.json.
- Stampa le sole eccezioni. Non modifica mai il file originale.
- --applica: scrive NUOVO_applicato.xlsx con, sulle righe nuove, Posizione_Analitica
  = ruolo ROSTERS e Lato_Campo = lato prevalente (solo per giocatori con lato fisso).
  Le eccezioni da decidere restano nel report.
Il lato fisso vale solo per Esterno offensivo, Terzino, Esterno/Quinto (e portieri=Centro);
centrocampisti, centrali e punte cambiano lato per modulo e non vengono controllati.
Lato prevalente: >=3 presenze nel ruolo e un lato >=70% (come RIFERIMENTO_RUOLI_LATI).
"""
import sys, json, re, collections
import pandas as pd

args = [a for a in sys.argv[1:] if not a.startswith('--')]
APPLICA = '--applica' in sys.argv
if not args:
    sys.exit(__doc__)
xlsx = args[0]
pub = json.load(open(args[1] if len(args) > 1 else 'data.json', encoding='utf-8'))

pub_role = {p['Player_ID']: p['Posizione_Analitica'] for p in pub['rosters']}
pub_name = {p['Player_ID']: p['Nome_completo'] for p in pub['rosters']}
pub_team = {p['Player_ID']: p['Squadra_Rif'] for p in pub['rosters']}
pub_matches = {m['Match_ID'] for m in pub['match_log']}
rows = collections.defaultdict(list)
for r in pub['match_log']:
    rows[r['Player_ID']].append(r)

LARGHI = {'Esterno offensivo', 'Terzino', 'Esterno/Quinto'}

def lato_fisso(pid, role):
    """Lato fisso solo per ruoli larghi (dove il lato conta); portieri = Centro."""
    if role == 'Portiere':
        return 'Centro'
    if role not in LARGHI:
        return None
    rs = [r for r in rows[pid] if r['Posizione_Analitica'] == role and r['Lato_Campo'] in ('Sinistra', 'Destra', 'Centro')]
    n = len(rs)
    if n < 3:
        return None
    l, k = collections.Counter(r['Lato_Campo'] for r in rs).most_common(1)[0]
    return l if k / n >= 0.7 else None

X = pd.read_excel(xlsx, sheet_name=None)
R, M = X['ROSTERS'], X['MATCH_LOG']
new_role = dict(zip(R['Player_ID'], R['Posizione_Analitica']))
new_name = dict(zip(R['Player_ID'], R['Nome_completo']))
new_team = dict(zip(R['Player_ID'], R['Squadra_Rif']))
Mn = M[~M['Match_ID'].isin(pub_matches)]
print(f"Righe MATCH_LOG nuove: {len(Mn)} | partite nuove: {Mn['Match_ID'].nunique()}")

out = []
# 1. giocatori nuovi
nuovi = [p for p in R['Player_ID'] if p not in pub_role]
print(f"\n[1] Giocatori NUOVI in ROSTERS: {len(nuovi)}")
for p in nuovi:
    ok = bool(re.fullmatch(r'P_[0-9A-F]{10}|UCL_[0-9A-F]{12}', p))  # Serie A: P_+10 hex | UCL: UCL_+12 hex
    print(f"   {new_name[p]} | {new_team[p]} | {new_role[p]} | {p}" + ("" if ok else "  <-- ID NON STANDARD"))
# 2. ruoli cambiati in ROSTERS
print("\n[2] Ruoli in ROSTERS diversi dal pubblicato (cambio retroattivo? serve motivo):")
c2 = 0
for p, r in new_role.items():
    if p in pub_role and pub_role[p] != r:
        c2 += 1
        print(f"   {new_name[p]} ({new_team[p]}): {pub_role[p]} -> {r}")
print("   nessuno" if not c2 else f"   totale {c2}")
# 3. ruolo righe nuove != ruolo ROSTERS
print("\n[3] Righe nuove con ruolo diverso da ROSTERS (ruolo giocato quella partita):")
c3 = 0
for _, r in Mn.iterrows():
    rr = new_role.get(r['Player_ID'])
    if rr and r['Posizione_Analitica'] != rr:
        c3 += 1
        print(f"   {new_name[r['Player_ID']]} ({r['Team_ID']}) {r['Match_ID']}: giocato {r['Posizione_Analitica']} | roster {rr}")
print("   nessuna" if not c3 else f"   totale {c3}")
# 4. lato diverso dal prevalente per giocatori fissi
print("\n[4] Righe nuove con lato diverso dal lato fisso del giocatore:")
c4 = 0
for _, r in Mn.iterrows():
    pid = r['Player_ID']
    if pid not in pub_role:
        continue
    f = lato_fisso(pid, pub_role[pid])
    if f and r['Lato_Campo'] != f and r['Posizione_Analitica'] == pub_role[pid]:
        c4 += 1
        print(f"   {new_name[pid]} ({r['Team_ID']}) {r['Match_ID']}: lato {r['Lato_Campo']} | fisso {f}")
print("   nessuna" if not c4 else f"   totale {c4}")
# 5. modifiche a righe già pubblicate
Mo = M[M['Match_ID'].isin(pub_matches)]
pubkey = {(r['Player_ID'], r['Match_ID']): r for r in pub['match_log']}
ch = collections.Counter()
for _, r in Mo.iterrows():
    o = pubkey.get((r['Player_ID'], r['Match_ID']))
    if o is None:
        ch['riga nuova in partita vecchia'] += 1
        continue
    for col in ('Posizione_Analitica', 'Lato_Campo'):
        if o[col] != r[col]:
            ch[col] += 1
print(f"\n[5] Righe VECCHIE modificate: {dict(ch) if ch else 'nessuna'}")

if APPLICA:
    import openpyxl
    wb = openpyxl.load_workbook(xlsx)
    ws = wb['MATCH_LOG']
    h = [c.value for c in ws[1]]
    ip, im, ir, il = h.index('Player_ID'), h.index('Match_ID'), h.index('Posizione_Analitica'), h.index('Lato_Campo')
    nrow = 0
    for row in ws.iter_rows(min_row=2):
        if row[im].value in pub_matches:
            continue
        pid = row[ip].value
        rr = new_role.get(pid)
        if not rr:
            continue
        if row[ir].value == rr:  # solo se il ruolo giocato coincide col roster
            f = lato_fisso(pid, pub_role.get(pid, rr)) if pid in pub_role else None
            if f and row[il].value != f:
                row[il].value = f
                nrow += 1
    dst = xlsx.rsplit('.', 1)[0] + '_applicato.xlsx'
    wb.save(dst)
    print(f"\nScritto {dst}: {nrow} lati riallineati al lato fisso (ruoli non toccati).")
