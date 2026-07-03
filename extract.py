import pandas as pd
import json
import re

EXCLUDE_COLS = ['DATAS', 'CONFRONTOS', 'TIME A', 'TIME B', 'PLACAR', 'Penalti', 'Prorrogação',
                'Unnamed: 7', 'Gols A', 'Gols B', 'RESULTADO', 'Passou']

# Ordem oficial das fases no bolão (nome da aba -> rótulo de exibição)
PHASES = [
    ('ORGANIZAÇÃO',  'Fase de Grupos'),
    ('MATA-MATA 32', '32avos de Final'),
    ('MATA-MATA 16', 'Oitavas de Final'),
    ('MATA-MATA 8',  'Quartas de Final'),
    ('MATA-MATA 4',  'Semifinal'),
    ('MATA-MATA 2',  'Final'),
]

# Prefixo do placeholder de chaveamento usado em cada fase (ex.: "Vencedor Jogo 3")
# aponta para o índice local (1-based) do jogo na fase anterior
PLACEHOLDER_PREFIX = {
    'MATA-MATA 16': 'Vencedor Jogo',
    'MATA-MATA 8':  'Vencedor Oitava',
    'MATA-MATA 4':  'Vencedor Quarta',
    'MATA-MATA 2':  'Vencedor Semifinal',
}
PREV_PHASE = {
    'MATA-MATA 16': 'MATA-MATA 32',
    'MATA-MATA 8':  'MATA-MATA 16',
    'MATA-MATA 4':  'MATA-MATA 8',
    'MATA-MATA 2':  'MATA-MATA 4',
}

excel = pd.ExcelFile('BOLAO.xlsx')
sheets = {name: pd.read_excel('BOLAO.xlsx', sheet_name=name) for name, _ in PHASES}

participants = sorted([c for c in sheets['ORGANIZAÇÃO'].columns
                       if c and 'Unnamed' not in str(c) and c not in EXCLUDE_COLS])

def fmt_date(date):
    if pd.isna(date):
        return ''
    if isinstance(date, str):
        return date.split()[0]
    try:
        return date.strftime('%d/%m')
    except Exception:
        return str(date)

# ---- 1a passada: monta os jogos de cada fase (com indice local 1-based) ----
games_by_sheet = {}   # sheet -> list of game dicts (na ordem da planilha)
all_games = []

for sheet_name, phase_label in PHASES:
    df = sheets[sheet_name]
    local_games = []
    for idx, row in df.iterrows():
        match = row.get('CONFRONTOS', '')
        if not match or pd.isna(match):
            continue

        guesses = {}
        for p in participants:
            g = row.get(p, '')
            guesses[p] = str(g).upper() if pd.notna(g) and g != '' else ''

        resultado = row.get('RESULTADO', '')
        resultado = str(resultado).upper() if pd.notna(resultado) and resultado != '' else ''

        gols_a = row.get('Gols A', '')
        gols_a = str(int(gols_a)) if pd.notna(gols_a) and gols_a != '' else ''

        gols_b = row.get('Gols B', '')
        gols_b = str(int(gols_b)) if pd.notna(gols_b) and gols_b != '' else ''

        penalti = row.get('Penalti', '')
        penalti = str(penalti).strip().upper() if pd.notna(penalti) and penalti != '' else ''

        prorrogacao = row.get('Prorrogação', '')
        prorrogacao = str(prorrogacao).strip().upper() if pd.notna(prorrogacao) and prorrogacao != '' else ''

        times = match.split(' x ')
        timeA = times[0].strip() if len(times) > 0 else ''
        timeB = times[1].strip() if len(times) > 1 else ''

        passou, tipo_decisao = '', ''
        if penalti:
            tipo_decisao = 'Pênaltis'
            passou = timeA if penalti == 'A' else timeB if penalti == 'B' else ''
        elif prorrogacao:
            tipo_decisao = 'Prorrogação'
            passou = timeA if prorrogacao == 'A' else timeB if prorrogacao == 'B' else ''
        elif resultado:
            passou = timeA if resultado == 'A' else timeB if resultado == 'B' else ''

        game_id = len(all_games)
        game = {
            'id': game_id,
            'local_index': len(local_games) + 1,   # posicao dentro da propria fase (1-based)
            'match': match,
            'timeA': timeA,
            'timeB': timeB,
            'date': fmt_date(row.get('DATAS')),
            'resultado': resultado,
            'gols_a': gols_a,
            'gols_b': gols_b,
            'guesses': guesses,
            'sheet': sheet_name,
            'phase_label': phase_label,
            'passou': passou,
            'penalties': penalti != '',
            'overtime': prorrogacao != '',
            'tipo_decisao': tipo_decisao,
        }
        local_games.append(game)
        all_games.append(game)
    games_by_sheet[sheet_name] = local_games

# ---- 2a passada: resolve placeholders "Vencedor Jogo N" com o time real, se ja definido ----
def resolve_placeholder(text, sheet_name):
    prefix = PLACEHOLDER_PREFIX.get(sheet_name)
    if not prefix or not text:
        return text
    m = re.fullmatch(rf'{re.escape(prefix)} (\d+)', text.strip())
    if not m:
        return text
    n = int(m.group(1))
    prev_sheet = PREV_PHASE[sheet_name]
    prev_games = games_by_sheet.get(prev_sheet, [])
    if n <= len(prev_games) and prev_games[n - 1]['passou']:
        return prev_games[n - 1]['passou']
    return text  # ainda nao definido

for sheet_name, _ in PHASES:
    if sheet_name not in PLACEHOLDER_PREFIX:
        continue
    for game in games_by_sheet[sheet_name]:
        rA = resolve_placeholder(game['timeA'], sheet_name)
        rB = resolve_placeholder(game['timeB'], sheet_name)
        game['timeA_resolved'] = rA
        game['timeB_resolved'] = rB
        game['match_resolved'] = f"{rA} x {rB}"
        game['pending'] = (rA == game['timeA'] and rA.startswith(PLACEHOLDER_PREFIX[sheet_name])) or \
                           (rB == game['timeB'] and rB.startswith(PLACEHOLDER_PREFIX[sheet_name]))

for game in all_games:
    if 'timeA_resolved' not in game:
        game['timeA_resolved'] = game['timeA']
        game['timeB_resolved'] = game['timeB']
        game['match_resolved'] = game['match']
        game['pending'] = False

# ---- 3a passada: pontuacao por fase e acumulada, por participante ----
def phase_points(sheet_name, participant):
    pts = 0
    for g in games_by_sheet[sheet_name]:
        if g['resultado'] and (g['guesses'].get(participant, '') == g['resultado']):
            pts += 1
    return pts

participants_detail = []
cumulative = {p: 0 for p in participants}
by_phase_matrix = {p: {} for p in participants}

for sheet_name, phase_label in PHASES:
    for p in participants:
        pp = phase_points(sheet_name, p)
        cumulative[p] += pp
        by_phase_matrix[p][sheet_name] = {'label': phase_label, 'points': pp, 'total_after': cumulative[p]}

for p in participants:
    acertos_total = sum(v['points'] for v in by_phase_matrix[p].values())
    participants_detail.append({
        'name': p,
        'total_points': cumulative[p],
        'acertos': acertos_total,
        'by_phase': by_phase_matrix[p],
    })

ranking = sorted(participants_detail, key=lambda x: (-x['total_points'], x['name']))
for i, r in enumerate(ranking):
    r['position'] = i + 1

results = {
    'participants': participants,
    'phases': [s for s, _ in PHASES],
    'phase_labels': {s: l for s, l in PHASES},
    'games': all_games,
    'ranking': ranking,
}

with open('results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("results.json gerado")
for sheet_name, phase_label in PHASES:
    n = len(games_by_sheet[sheet_name])
    print(f"   {phase_label}: {n} jogo(s)")
print(f"   {len(participants)} participantes")
