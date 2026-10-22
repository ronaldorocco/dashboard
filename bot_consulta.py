"""
Bot Interativo GMBC — Responde consultas no Telegram
Roda via GitHub Actions a cada 5 minutos.

Comandos disponíveis:
  centro              → consulta bairro (texto livre)
  /bairro centro      → consulta bairro
  /tipo furto         → consulta por tipo de crime
  /turno noite        → consulta por turno
  /resumo             → resumo geral
  /bairros            → lista todos os bairros
  /ajuda              → lista de comandos
"""

import sys, json, math, warnings, unicodedata
from datetime import datetime, timezone, timedelta
from collections import Counter
import urllib.request, urllib.error
import os as _os

warnings.filterwarnings('ignore')

BRT = timezone(timedelta(hours=-3))

BOT_TOKEN       = _os.environ.get('BOT_TOKEN',       'SEU_TOKEN_AQUI').strip()
CHAT_ID         = _os.environ.get('CHAT_ID',         '1931364088').strip()
GOOGLE_DRIVE_ID = _os.environ.get('GOOGLE_DRIVE_ID', '1w_4WgORfWrxonI-tL6uKOkoCZJQ9K5VN').strip()

# ── Mapas de normalização ────────────────────────────────────────────────────
DIA_MAP = {
    'SEGUNDA':'Segunda','TERÇA':'Terça','TERCA':'Terça','QUARTA':'Quarta',
    'QUINTA':'Quinta','SEXTA':'Sexta','SABADO':'Sábado','SÁBADO':'Sábado','DOMINGO':'Domingo',
} 
BAIRRO_MAP = {
    'BARRA SUL':'Barra Sul','SÃO J. TADEU':'São J. Tadeu','SAO J. TADEU':'São J. Tadeu',
    'N. ESPERANÇA':'N. Esperança','N. ESPERANCA':'N. Esperança','MUNICIPIOS':'Municípios',
    'NAÇÕES':'Nações','NACOES':'Nações','PONTAL NORTE':'Pontal Norte',
    'VILA REAL':'Vila Real','PIONEIROS':'Pioneiros','ARIRIBA':'Ariribá',
    'CENTRO':'Centro','ESTADOS':'Estados','BARRA':'Barra',
}
MES_PARA_NUM = {
    'JANEIRO':'01','FEVEREIRO':'02','MARCO':'03','MARÇO':'03','ABRIL':'04','MAIO':'05',
    'JUNHO':'06','JULHO':'07','AGOSTO':'08','SETEMBRO':'09','OUTUBRO':'10',
    'NOVEMBRO':'11','DEZEMBRO':'12',
}
MES_NOME = {'01':'Janeiro','02':'Fevereiro','03':'Março','04':'Abril','05':'Maio',
            '06':'Junho','07':'Julho','08':'Agosto','09':'Setembro','10':'Outubro',
            '11':'Novembro','12':'Dezembro'}
TIPO_MAP = {'FURTO':'Furto','ROUBO':'Roubo'}
ITEM_MAP = {
    'VEICULO':'Ve\u00edculo','VEILCULO':'Ve\u00edculo','VEICULO ':'Ve\u00edculo',
    'CELULAR':'Celular','FIA\u00c7\u00c3O EL\u00c9TRICA':'Fia\u00e7\u00e3o El\u00e9trica',
    'FIA\u00c7AO ELETRICA':'Fia\u00e7\u00e3o El\u00e9trica','FIA\u00c7\u00c3O ELETRICA':'Fia\u00e7\u00e3o El\u00e9trica',
    'ELETRODOMESTICO':'Eletrodom\u00e9stico','COMBUSTIVEL':'Combust\u00edvel',
    'JETSKI':'JetSki','ALUMINIO':'Alum\u00ednio','VESTUARIO':'Vestu\u00e1rio',
    'PRODUTOS':'Produtos','REBOQUE':'Reboque','MERCADORIAS':'Mercadorias',
    'OBJETOS':'Objetos','DINHEIRO':'Dinheiro','SCOOTER':'Scooter',
    'BICICLETA':'Bicicleta','ELETR\u00d4NICO':'Eletr\u00f4nico','ELETRONICO':'Eletr\u00f4nico',
    'MOTOCICLETA':'Motocicleta',
}
ITEM_MAP_NORMALIZED = {
    'VEICULO':'Ve\u00edculo','VEILCULO':'Ve\u00edculo',
    'CELULAR':'Celular','FIACAO ELETRICA':'Fia\u00e7\u00e3o El\u00e9trica',
    'ELETRODOMESTICO':'Eletrodom\u00e9stico','COMBUSTIVEL':'Combust\u00edvel',
    'JETSKI':'JetSki','ALUMINIO':'Alum\u00ednio','VESTUARIO':'Vestu\u00e1rio',
    'PRODUTOS':'Produtos','REBOQUE':'Reboque','MERCADORIAS':'Mercadorias',
    'OBJETOS':'Objetos','DINHEIRO':'Dinheiro','SCOOTER':'Scooter',
    'BICICLETA':'Bicicleta','ELETRONICO':'Eletr\u00f4nico',
    'MOTOCICLETA':'Motocicleta',
}
ORDEM_TURNO = ['Madrugada','Manhã','Tarde','Noite']


def norm(v, mapa):
    if not v or str(v).strip().upper() in ('','NAN'): return ''
    u = str(v).strip().upper()
    return mapa.get(u, str(v).strip().title())

def sem_acento(v):
    txt = unicodedata.normalize('NFKD', str(v).lower())
    return ''.join(c for c in txt if not unicodedata.combining(c))

def norm_tipo(v):
    if not v or str(v).strip().upper() in ('','NAN'): return ''
    v = str(v).strip().upper()
    if 'TENTATIVA' in v and 'ROUBO' in v: return 'Tentativa de Roubo'
    if 'TENTATIVA' in v and 'FURTO' in v: return 'Tentativa de Furto'
    if 'ARROMBAMENTO' in v: return 'Arrombamento'
    return TIPO_MAP.get(v, v.title())

def norm_item(v):
    if not v or str(v).strip().upper() in ('','NAN'): return ''
    u = str(v).strip().upper()
    u_norm = sem_acento(u).upper()
    return ITEM_MAP.get(u, ITEM_MAP_NORMALIZED.get(u_norm, str(v).strip().title()))

def clean_text_cell(v):
    if not v or str(v).strip().upper() in ('','NAN'):
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def calcular_turno(hora_val):
    try:
        h = hora_val.hour if hasattr(hora_val,'hour') else int(str(hora_val).strip().split(':')[0])
        if  6<=h<=11: return 'Manhã'
        if 12<=h<=17: return 'Tarde'
        if 18<=h<=23: return 'Noite'
        return 'Madrugada'
    except: return ''


# ── Carregar dados ───────────────────────────────────────────────────────────
def carregar_dados():
    import tempfile, os
    try: import pandas as pd
    except ImportError:
        print("ERRO: pandas nao instalado."); sys.exit(1)

    GITHUB_URL = 'https://github.com/ronaldorocco/relatorio-gmbc/raw/refs/heads/master/secretario.xlsx'

    def _baixar_xlsx(url, nome):
        print(f"  Baixando planilha de {nome}...")
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        tmp.write(data); tmp.close()
        return tmp.name

    DASHBOARD_XLSX_URL = 'https://raw.githubusercontent.com/ronaldorocco/dashboard/main/secretario.xlsx'

    df = None
    # 1) Mesma planilha usada pelo dashboard (fonte principal — garante dados identicos)
    try:
        tmp_path = _baixar_xlsx(DASHBOARD_XLSX_URL, 'repositorio do dashboard')
        df = pd.read_excel(tmp_path, sheet_name='DADOS', engine='openpyxl')
        os.unlink(tmp_path)
        print("  OK: dados carregados do repositorio do dashboard.")
    except Exception as e:
        print(f"  Falha ao baixar do repositorio do dashboard: {e}")

    # 2) Google Drive (fallback)
    if df is None and GOOGLE_DRIVE_ID:
        try:
            tmp_path = _baixar_xlsx(
                f"https://docs.google.com/spreadsheets/d/{GOOGLE_DRIVE_ID}/export?format=xlsx",
                'Google Drive'
            )
            df = pd.read_excel(tmp_path, sheet_name='DADOS', engine='openpyxl')
            os.unlink(tmp_path)
            print("  OK: dados carregados do Google Drive.")
        except Exception as e:
            print(f"  Falha ao baixar do Google Drive: {e}")

    # 3) Arquivo local do deployment (fallback)
    if df is None:
        excel_local = 'secretario.xlsx'
        if os.path.exists(excel_local):
            print(f"  Lendo {excel_local} local...")
            try:
                df = pd.read_excel(excel_local, sheet_name='DADOS', engine='openpyxl')
                print("  OK: dados carregados do arquivo local.")
            except Exception as e:
                print(f"  Falha ao ler local: {e}")

    # 4) GitHub relatorio-gmbc (último recurso)
    if df is None:
        try:
            tmp_path = _baixar_xlsx(GITHUB_URL, 'GitHub')
            df = pd.read_excel(tmp_path, sheet_name='DADOS', engine='openpyxl')
            os.unlink(tmp_path)
            print("  OK: dados carregados do GitHub.")
        except Exception as e:
            print(f"ERRO ao baixar planilha: {e}"); sys.exit(1)

    # Normaliza os cabeçalhos pra maiúsculo — a planilha às vezes é resalva com
    # outra capitalização (ex.: "Tipificacao", "Dia_Semana" em vez de
    # "TIPIFICACAO", "DIA_SEMANA") e o resto do script busca as colunas em maiúsculo.
    df.columns = [str(c).strip().upper() for c in df.columns]
    _RENOMEIA_COLUNAS = {
        "NUMERO DE SÉRIE": "NUMERO_SERIE",
        "NUMERO DE SERIE": "NUMERO_SERIE",
        "BOLETIM_OCORRENCIA": "B.O.",
        "BOLETIM OCORRENCIA": "B.O.",
        "BOLETIM_OCORRÊNCIA": "B.O.",
        "BOLETIM OCORRÊNCIA": "B.O.",
    }
    df = df.rename(columns=_RENOMEIA_COLUNAS)

    df['TIPIFICACAO'] = df['TIPIFICACAO'].apply(norm_tipo)
    df['DIA_SEMANA']  = df['DIA_SEMANA'].apply(lambda v: norm(v, DIA_MAP))
    df['TURNO']       = df.apply(lambda r: calcular_turno(r['HORA']), axis=1)
    df['BAIRRO']      = df['BAIRRO'].apply(lambda v: norm(v, BAIRRO_MAP))
    df['MES_UPPER']   = df['MES'].apply(lambda v: str(v).strip().upper() if v else '')

    df = df[
        df['B.O.'].notna() &
        (df['B.O.'].astype(str).str.strip().str.upper() != 'NAN') &
        (df['B.O.'].astype(str).str.strip() != '') &
        (df['TIPIFICACAO'] != '')
    ].copy()

    def build_data(row):
        mes_num = MES_PARA_NUM.get(row['MES_UPPER'],'')
        if not mes_num: return ''
        try: return f"{int(row['ANO'])}-{mes_num}-{str(int(row['DATA'])).zfill(2)}"
        except: return ''

    df['DATA_STR'] = df.apply(build_data, axis=1)
    df['HORA_STR'] = df['HORA'].apply(lambda x: x.strftime('%H:%M') if hasattr(x,'strftime') else str(x)[:5] if str(x).strip() not in ('','nan') else '')
    df['ENDERECO'] = df['ENDEREÇO'].fillna('').astype(str)
    df['ITEM_STR'] = df['ITEM'].apply(norm_item) if 'ITEM' in df.columns else ''
    df['MARCA_STR'] = df['MARCA_MODELO'].fillna('').astype(str).apply(lambda v: '' if v.strip().lower() in ('','nan') else v.strip()) if 'MARCA_MODELO' in df.columns else ''
    df['IMEI_STR'] = df['IMEI'].apply(clean_text_cell) if 'IMEI' in df.columns else ''
    df['BO_STR'] = df['B.O.'].fillna('').astype(str)

    records = []
    for _, r in df.iterrows():
        if not r['DATA_STR']: continue
        records.append({
            'data':    r['DATA_STR'],
            'ano':     r['DATA_STR'][:4],
            'mes':     MES_NOME.get(r['DATA_STR'][5:7], ''),
            'dia':     r['DIA_SEMANA'],
            'turno':   r['TURNO'],
            'tipo':    r['TIPIFICACAO'],
            'bairro':  r['BAIRRO'],
            'endereco':r['ENDERECO'],
            'hora':    r['HORA_STR'],
            'item':    r.get('ITEM_STR','') if 'ITEM_STR' in r else '',
            'marca':   r.get('MARCA_STR','') if 'MARCA_STR' in r else '',
            'imei':    r.get('IMEI_STR','') if 'IMEI_STR' in r else '',
            'bo':      r['BO_STR'],
        })
    return records


# ── Telegram API ─────────────────────────────────────────────────────────────
def get_updates(offset=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates?timeout=0"
    if offset: url += f"&offset={offset}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            return json.loads(resp.read().decode()).get('result', [])
    except Exception as e:
        print(f"Erro getUpdates: {e}"); return []

def send_message(chat_id, text):
    MAX = 4000  # Telegram limita 4096; deixamos margem
    chunks = []
    lines = text.split('\n')
    current = []
    current_len = 0
    for line in lines:
        # +1 para o '\n'
        if current_len + len(line) + 1 > MAX and current:
            chunks.append('\n'.join(current))
            current = [line]
            current_len = len(line) + 1
        else:
            current.append(line)
            current_len += len(line) + 1
    if current:
        chunks.append('\n'.join(current))

    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    for chunk in chunks:
        # Tenta com Markdown; se falhar, reenvia sem formatação
        for parse_mode in ("Markdown", None):
            payload = {"chat_id": chat_id, "text": chunk}
            if parse_mode:
                payload["parse_mode"] = parse_mode
            data = json.dumps(payload).encode('utf-8')
            req  = urllib.request.Request(url, data=data, headers={'Content-Type':'application/json'})
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    result = json.loads(resp.read().decode())
                    if result.get('ok'):
                        print(f"  Resposta enviada para {chat_id} (parse_mode={parse_mode})")
                        break  # Sucesso — não precisa retry
                    else:
                        print(f"  ERRO Telegram (parse_mode={parse_mode}): {result.get('description')}")
            except Exception as e:
                # HTTP 400 (Markdown inválido, ex.: "_" ou "*" no texto) levanta exceção:
                # continua pro próximo parse_mode (texto puro) em vez de perder a resposta.
                print(f"  ERRO ao enviar (parse_mode={parse_mode}): {e}")
                continue


# ── Consultas ─────────────────────────────────────────────────────────────────
def pct(n, total):
    return f"{round(n/total*100)}%" if total else "0%"

# Deduplica por número de B.O.: um B.O. com vários itens conta como 1
# ocorrência só — mesma regra usada no card "Total de Ocorrências" do
# dashboard (build_dashboard.py/dedupBO), pra bater os números nos dois
# lugares. Registros sem B.O. (não deveria acontecer, carregar_dados já
# filtra) são mantidos, nunca descartados por engano.
def dedup_bo(registros):
    vistos = set()
    out = []
    for r in registros:
        k = (r.get('bo') or '').strip()
        if not k:
            out.append(r)
            continue
        if k in vistos:
            continue
        vistos.add(k)
        out.append(r)
    return out

def top_ruas(registros, n=5):
    """Ranking de logradouros agrupando variações de maiúscula/espaço/acento
    ('Avenida das Flores' / 'Avenida Das Flores') como a mesma rua."""
    grupos = {}
    for r in registros:
        e = (r.get('endereco') or '').strip()
        if not e or e.lower() == 'nan':
            continue
        chave = ' '.join(sem_acento(e).split())
        g = grupos.setdefault(chave, Counter())
        g[' '.join(e.split())] += 1
    ranking = [(g.most_common(1)[0][0], sum(g.values())) for g in grupos.values()]
    ranking.sort(key=lambda x: (-x[1], x[0]))
    return ranking[:n]

def achar_nome(nomes, query):
    """Acha o nome mais parecido com a consulta: exato > começa com > contém.
    Evita 'barra sul' cair em 'Barra' ou 'estaleirinho' em 'Estaleiro'."""
    q = ' '.join(sem_acento(query).split())
    if not q:
        return None
    norm_ = [(n, sem_acento(n)) for n in nomes]
    return (
        next((n for n, sn in norm_ if sn == q), None) or
        next((n for n, sn in norm_ if q in sn.split() or sn.startswith(q)), None) or
        next((n for n, sn in norm_ if q in sn), None) or
        next((n for n, sn in norm_ if len(sn) >= 4 and sn in q), None)
    )

def consultar_bairro(records, query):
    bairros = sorted(set(r['bairro'] for r in records if r['bairro']))
    q = sem_acento(query.strip())
    match = achar_nome(bairros, query)

    if not match:
        lista = '\n'.join(f"• {b}" for b in bairros[:15])
        return f"Bairro *'{query}'* não encontrado.\n\n*Bairros disponíveis:*\n{lista}"

    dados  = dedup_bo([r for r in records if r['bairro'] == match])
    total  = len(dados)
    tipos  = Counter(r['tipo']     for r in dados if r['tipo']).most_common(5)
    turnos = Counter(r['turno']    for r in dados if r['turno']).most_common()
    ruas   = top_ruas(dados, 5)
    horas  = Counter(r['hora'][:2] for r in dados if r['hora'] and len(r['hora'])>=2).most_common(3)
    datas  = sorted(set(r['data']  for r in dados))

    linhas = [
        f"📍 *BAIRRO: {match.upper()}*",
        f"📊 Total de ocorrências: *{total}*",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*🔴 Tipos de crime:*",
        *[f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos],
        "",
        "*🕐 Horários de pico:*",
        *[f"  • {h}h: {n} oc." for h,n in horas],
        "",
        "*🛣️ Ruas mais afetadas:*",
        *[f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas) if r and r!='nan'],
    ]
    return '\n'.join(linhas)

def consultar_tipo(records, query):
    tipos = sorted(set(r['tipo'] for r in records if r['tipo']))
    q = sem_acento(query.strip())
    match = achar_nome(tipos, query)

    if not match:
        lista = '\n'.join(f"• {t}" for t in tipos)
        return f"Tipo *'{query}'* não encontrado.\n\n*Tipos disponíveis:*\n{lista}"

    dados   = dedup_bo([r for r in records if r['tipo'] == match])
    total   = len(dados)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']  for r in dados if r['turno']).most_common()
    ruas    = top_ruas(dados, 3)
    datas   = sorted(set(r['data'] for r in dados))

    linhas = [
        f"🔴 *CRIME: {match.upper()}*",
        f"📊 Total: *{total}* ocorrências",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*📍 Bairros mais afetados:*",
        *[f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos],
        "",
        "*🛣️ Ruas de maior risco:*",
        *[f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas) if r and r!='nan'],
    ]
    return '\n'.join(linhas)

def consultar_turno(records, query):
    turno_map = {'manha':'Manhã','manhã':'Manhã','tarde':'Tarde','noite':'Noite','madrugada':'Madrugada'}
    match = turno_map.get(query.strip().lower())
    if not match:
        return "Turnos disponíveis: *Madrugada*, *Manhã*, *Tarde*, *Noite*"

    dados   = dedup_bo([r for r in records if r['turno'] == match])
    total   = len(dados)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    tipos   = Counter(r['tipo']   for r in dados if r['tipo']).most_common(5)

    linhas = [
        f"⏰ *TURNO: {match.upper()}*",
        f"📊 Total: *{total}* ocorrências",
        "",
        "*📍 Bairros mais afetados:*",
        *[f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)],
        "",
        "*🔴 Crimes mais frequentes:*",
        *[f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)],
    ]
    return '\n'.join(linhas)

def consultar_mes(records, query):
    q = sem_acento(query.strip())
    mes_num = mes_nome = None
    for nome, num in MES_PARA_NUM.items():
        if sem_acento(nome) == q:
            mes_num, mes_nome = num, nome.capitalize()
            break

    if not mes_num:
        return "Meses disponíveis: Janeiro, Fevereiro, Março, Abril, Maio, Junho, Julho, Agosto, Setembro, Outubro, Novembro, Dezembro"

    dados = dedup_bo([r for r in records if r.get('data') and len(r['data']) >= 7 and r['data'][5:7] == mes_num])
    total = len(dados)
    if not total:
        return f"Nenhuma ocorrência encontrada em *{mes_nome}*."

    tipos   = Counter(r['tipo']   for r in dados if r['tipo']).most_common(5)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']  for r in dados if r['turno']).most_common()

    linhas = [
        f"🗓 *MÊS: {mes_nome.upper()}*",
        f"📊 Total: *{total}* ocorrências",
        "",
        "*🔴 Tipos de crime:*",
        *[f"  {i+1}. {t}: *{n}* ({pct(n,total)})" for i,(t,n) in enumerate(tipos)],
        "",
        "*📍 Bairros mais afetados:*",
        *[f"  {i+1}. {b}: *{n}* ({pct(n,total)})" for i,(b,n) in enumerate(bairros)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: *{n}* ({pct(n,total)})" for t,n in turnos],
    ]
    return '\n'.join(linhas)

def consultar_resumo(records):
    records = dedup_bo(records)
    total   = len(records)
    bairros = Counter(r['bairro'] for r in records if r['bairro']).most_common(3)
    tipos   = Counter(r['tipo']   for r in records if r['tipo']).most_common(3)
    turnos  = Counter(r['turno']  for r in records if r['turno']).most_common()
    datas   = sorted(set(r['data'] for r in records))
    now     = datetime.now(BRT)

    linhas = [
        f"📊 *RESUMO GERAL — GMBC*",
        f"🕐 {now.strftime('%d/%m/%Y às %H:%M')}",
        f"Total de ocorrências: *{total}*",
        f"Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
        "",
        "*📍 Top bairros:*",
        *[f"  {i+1}. {b}: {n} oc." for i,(b,n) in enumerate(bairros)],
        "",
        "*🔴 Top crimes:*",
        *[f"  {i+1}. {t}: {n} oc." for i,(t,n) in enumerate(tipos)],
        "",
        "*⏰ Por turno:*",
        *[f"  • {t}: {n} ({pct(n,total)})" for t,n in turnos],
    ]
    return '\n'.join(linhas)

# ── Menu numerado pós-busca ──────────────────────────────────────────────────
ESTADO_BUSCA = {}          # chat_id -> {'query', 'base', 'ts', 'opcoes'}
ESTADO_TTL = timedelta(hours=3)

DIMENSOES = [   # (chave, rótulo)
    ('tipo',   '🔴 Tipificação'),
    ('bairro', '📍 Bairro'),
    ('rua',    '🛣️ Rua'),
    ('turno',  '⏰ Turno'),
    ('dia',    '📆 Dia da semana'),
    ('mes',    '🗓 Mês'),
    ('ano',    '📅 Ano'),
    ('item',   '📦 Item'),
    ('marca',  '🏷️ Marca/Modelo'),
]

def _ranking_dim(base, dim, limite=15):
    """Lista [(rótulo, qtd)] da dimensão pedida, sobre os B.O.s da busca."""
    if dim == 'rua':
        return top_ruas(base, limite)
    if dim == 'mes':
        c = Counter(r['data'][:7] for r in base if r.get('data') and len(r['data']) >= 7)
        return [(f"{MES_NOME.get(k[5:7], k[5:7])}/{k[:4]}", v) for k, v in sorted(c.items())]
    campo = {'tipo': 'tipo', 'bairro': 'bairro', 'turno': 'turno', 'dia': 'dia',
             'ano': 'ano', 'item': 'item', 'marca': 'marca'}[dim]
    # agrupa variações de maiúscula/acento ("Iphone" / "iPhone")
    grupos = {}
    for r in base:
        v = r.get(campo)
        if v and str(v).strip().lower() != 'nan':
            grupos.setdefault(sem_acento(str(v).strip()), Counter())[str(v).strip()] += 1
    ranking = [(g.most_common(1)[0][0], sum(g.values())) for g in grupos.values()]
    ranking.sort(key=lambda x: (-x[1], x[0]))
    return ranking[:limite]

def _opcoes_disponiveis(base):
    """Só oferece dimensões com mais de um valor (ex.: busca por bairro não pergunta o bairro)."""
    return [(k, rot) for k, rot in DIMENSOES if len(_ranking_dim(base, k, 1000)) > 1]

def _cabecalho_busca(est):
    base = est['base']
    datas = sorted(set(r['data'] for r in base if r.get('data')))
    linhas = [f"🔍 *BUSCA: \"{est['query'].upper()}\"*", f"📊 Total: *{len(base)}* ocorrência(s)"]
    if datas:
        linhas.append(f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]}/{datas[0][:4]} — {datas[-1][8:10]}/{datas[-1][5:7]}/{datas[-1][:4]}")
    return '\n'.join(linhas)

def _menu_busca(est):
    opcoes = [('dim:' + k, rot) for k, rot in _opcoes_disponiveis(est['base'])]
    opcoes += [('ultimas', '🕘 Últimas ocorrências'), ('resumo', '📋 Resumo completo')]
    est['opcoes'] = [k for k, _ in opcoes]
    linhas = ["*Escolha uma opção (responda só com o número):*"]
    linhas += [f"  *{i+1}* - {rot}" for i, (_, rot) in enumerate(opcoes)]
    linhas.append("_Ou digite uma nova busca._")
    return '\n'.join(linhas)

def responder_opcao(chat_id, numero):
    """Resposta a um número digitado depois de uma busca. None se não houver busca ativa."""
    est = ESTADO_BUSCA.get(chat_id)
    if not est or datetime.now(BRT) - est['ts'] > ESTADO_TTL:
        ESTADO_BUSCA.pop(chat_id, None)
        return None
    est['ts'] = datetime.now(BRT)
    base = est['base']
    if 'opcoes' not in est:
        _menu_busca(est)
    if not (1 <= numero <= len(est['opcoes'])):
        return f"Opção inválida. Responda com um número de 1 a {len(est['opcoes'])}.\n\n" + _menu_busca(est)

    chave = est['opcoes'][numero - 1]
    linhas = [_cabecalho_busca(est), ""]
    if chave.startswith('dim:'):
        dim = chave[4:]
        rot = dict(DIMENSOES)[dim]
        total = len(base)
        linhas.append(f"*{rot}:*")
        linhas += [f"  {i+1}. {v}: *{q}* ({pct(q, total)})" for i, (v, q) in enumerate(_ranking_dim(base, dim))]
    elif chave == 'ultimas':
        recentes = sorted(base, key=lambda r: (r.get('data', ''), r.get('hora', '')), reverse=True)[:10]
        linhas.append("*🕘 Últimas ocorrências:*")
        for r in recentes:
            d = r.get('data', '')
            linhas.append(f"  • {d[8:10]}/{d[5:7]}/{d[:4]} {r.get('hora','')} — {r.get('tipo','')} — "
                          f"{r.get('endereco','')} ({r.get('bairro','')})")
    else:
        return busca_universal_resumo(est['query'], base) + '\n\n' + _menu_busca(est)
    return '\n'.join(linhas) + '\n\n' + _menu_busca(est)

def busca_universal(records, query, chat_id=None):
    """Busca guiada por campo: cada parte da pergunta é procurada na COLUNA certa.
      • nome de rua ("quinta avenida", "avenida brasil") → só em ENDEREÇO
      • nome de bairro ("centro", "nova esperanca")      → só em BAIRRO
      • dia / mês / ano / turno / tipo de crime          → só na coluna respectiva
      • item (celular, bicicleta...)                     → só em ITEM; marca → MARCA
    As partes são combinadas com E ('furto celular sabado centro'). Só palavras
    que não são nada disso caem num fallback entre ENDEREÇO/BAIRRO e depois
    ITEM/MARCA/IMEI/B.O. Sem ano na pergunta, vale todos os anos."""
    DIAS_SET   = {'segunda', 'terca', 'quarta', 'quinta', 'sexta', 'sabado', 'domingo'}
    MESES_SET  = {sem_acento(m) for m in MES_NOME.values()}
    TURNOS_SET = {'madrugada', 'manha', 'tarde', 'noite'}
    TIPO_WORDS = {'furto', 'roubo', 'arrombamento', 'tentativa'}
    STOPWORDS  = {'de', 'da', 'do', 'das', 'dos', 'em', 'na', 'no', 'nas', 'nos', 'a', 'o', 'e'}
    ABREV      = {'av': 'avenida', 'av.': 'avenida'}

    # "sexta-feira"/"segunda feira" → "sexta"/"segunda"
    q_norm = sem_acento(query.strip()).replace('-feira', '').replace(' feira', '')
    if len(q_norm) < 3 and not q_norm.isdigit():
        return "Digite pelo menos 3 letras para buscar. Ex.: *centro*, *furto celular*, *quinta avenida*."
    tokens_all = [ABREV.get(t, t) for t in q_norm.split() if t]

    def n(r, campo):
        return sem_acento(r.get(campo, '') or '')

    def variacoes_de(tok):
        vs = {tok}
        if tok.endswith('s') and len(tok) > 3:
            vs.add(tok[:-1])      # bicicletas → bicicleta
        if tok.endswith('es') and len(tok) > 4:
            vs.add(tok[:-2])      # veiculos → veiculo
        return vs

    bairros_norm = {n(r, 'bairro') for r in records if r.get('bairro')}
    ends_norm    = {' '.join(n(r, 'endereco').split()) for r in records if r.get('endereco')}
    ends_norm.discard('nan')
    itens_norm   = {n(r, 'item') for r in records if r.get('item')}
    itens_norm  |= {i.rstrip('s') for i in itens_norm}

    restricoes = []          # (descricao, funcao(registro) -> bool)
    usados = [False] * len(tokens_all)

    def consome(i, j):
        for k in range(i, j):
            usados[k] = True

    def livre(i, j):
        return not any(usados[i:j])

    # 1) Bairro com mais de uma palavra ("nova esperanca", "praia dos amores")
    for bn in sorted((b for b in bairros_norm if ' ' in b), key=len, reverse=True):
        w = bn.split()
        for i in range(len(tokens_all) - len(w) + 1):
            if tokens_all[i:i+len(w)] == w and livre(i, i+len(w)):
                restricoes.append(('bairro', lambda r, bn=bn: n(r, 'bairro') == bn))
                consome(i, i+len(w))

    # 2) Nome de rua com 2+ palavras ("quinta avenida", "avenida do estado")
    for tam in range(len(tokens_all), 1, -1):
        for i in range(len(tokens_all) - tam + 1):
            if not livre(i, i+tam):
                continue
            frase = ' '.join(tokens_all[i:i+tam])
            if any((' ' + frase + ' ') in (' ' + e + ' ') for e in ends_norm):
                restricoes.append(('endereco', lambda r, f=frase: f in ' '.join(n(r, 'endereco').split())))
                consome(i, i+tam)

    # 3) Palavras soltas, cada uma na sua coluna
    for i, tok in enumerate(tokens_all):
        if usados[i] or tok in STOPWORDS:
            continue
        vs = variacoes_de(tok)
        if tok.isdigit() and len(tok) == 4:
            restricoes.append(('ano', lambda r, t=tok: r.get('ano') == t))
        elif tok in DIAS_SET:
            restricoes.append(('dia', lambda r, t=tok: n(r, 'dia') == t))
        elif tok in MESES_SET:
            restricoes.append(('mes', lambda r, t=tok: n(r, 'mes') == t))
        elif tok in TURNOS_SET:
            restricoes.append(('turno', lambda r, t=tok: n(r, 'turno') == t))
        elif vs & TIPO_WORDS:
            restricoes.append(('tipo', lambda r, v=vs: any(x in n(r, 'tipo') for x in v)))
        elif vs & bairros_norm:
            restricoes.append(('bairro', lambda r, v=vs: n(r, 'bairro') in v))
        elif vs & itens_norm:
            restricoes.append(('item', lambda r, v=vs: n(r, 'item') in v or n(r, 'item').rstrip('s') in v))
        else:
            # Palavra solta: logradouro e/ou bairro; senão item/marca/IMEI/B.O.
            f_end  = lambda r, v=vs: any(x in n(r, 'endereco') for x in v)
            f_bai  = lambda r, v=vs: any(x in w for x in v for w in n(r, 'bairro').split()) if len(tok) >= 5 else False
            f_item = lambda r, v=vs: any(x in n(r, 'item') for x in v)
            f_marca = lambda r, v=vs: any(x in n(r, 'marca') for x in v)
            f_id   = lambda r, v=vs: len(tok) >= 6 and any(x in n(r, 'imei') or x in n(r, 'bo') for x in v)
            for f in (lambda r, a=f_end, b=f_bai: a(r) or b(r), f_item, f_marca, f_id):
                if any(f(r) for r in records):
                    restricoes.append(('livre', f))
                    break
            else:
                restricoes.append(('livre', lambda r: False))

    if not restricoes:
        restricoes.append(('livre', lambda r: False))

    matches = [r for r in records if all(f(r) for _, f in restricoes)]

    if not matches:
        return (
            f"Nenhuma ocorrência encontrada para *'{query}'*.\n\n"
            "Tente: nome de bairro, tipo de crime, item (bicicleta, celular...), IMEI, turno ou logradouro.\n"
            "Digite /ajuda para ver todos os comandos."
        )

    # Deduplicado por B.O. — 1 boletim com vários itens conta 1 ocorrência.
    base = dedup_bo(matches)
    total_base = len(base)

    # Com chat_id (uso no Telegram): guarda a busca e mostra só o total + menu
    # numerado; cada número abre um detalhamento (ano, mês, bairro, rua...).
    if chat_id is not None:
        ESTADO_BUSCA[chat_id] = {'query': query, 'base': base, 'ts': datetime.now(BRT)}
        return _cabecalho_busca(ESTADO_BUSCA[chat_id]) + '\n\n' + _menu_busca(ESTADO_BUSCA[chat_id])

    return busca_universal_resumo(query, base)


def busca_universal_resumo(query, base):
    """Relatório completo da busca (tipificação, bairros, turno, dia, mês, ruas)."""
    total_base = len(base)
    tipos   = Counter(r['tipo']    for r in base if r['tipo']).most_common(5)
    bairros = Counter(r['bairro']  for r in base if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']   for r in base if r['turno']).most_common()
    ruas    = top_ruas(base, 5)
    dias    = Counter(r['dia']     for r in base if r['dia']).most_common()
    meses   = Counter(r['data'][5:7] for r in base if r.get('data') and len(r['data'])>=7).most_common()
    datas   = sorted(set(r['data'] for r in base))

    linhas = [
        f"🔍 *BUSCA: \"{query.upper()}\"*",
        f"📊 *{total_base}* ocorrência(s) encontrada(s)",
        f"📅 Período: {datas[0][8:10]}/{datas[0][5:7]} — {datas[-1][8:10]}/{datas[-1][5:7]}",
    ]

    linhas[1] = f"📊 Total: *{total_base}* ocorrência(s)"

    if tipos:
        linhas += ["", "*🔴 Tipificação:*"]
        linhas += [f"  {i+1}. {t}: *{n}* ({pct(n,total_base)})" for i,(t,n) in enumerate(tipos)]

    if bairros:
        linhas += ["", "*📍 Bairros:*"]
        linhas += [f"  {i+1}. {b}: *{n}* ({pct(n,total_base)})" for i,(b,n) in enumerate(bairros)]

    if turnos:
        linhas += ["", "*⏰ Turno:*"]
        linhas += [f"  • {t}: *{n}* ({pct(n,total_base)})" for t,n in turnos]

    if dias:
        linhas += ["", "*📆 Dia da semana:*"]
        linhas += [f"  • {d}: *{n}* ({pct(n,total_base)})" for d,n in dias]

    if meses:
        linhas += ["", "*🗓 Mês:*"]
        linhas += [f"  • {MES_NOME.get(m,m)}: *{n}*" for m,n in meses]

    if ruas:
        linhas += ["", "*🛣️ Ruas:*"]
        linhas += [f"  {i+1}. {r}: *{n}* oc." for i,(r,n) in enumerate(ruas)]

    return '\n'.join(linhas)


def top_n(records, campo, n=5):
    c = Counter(r[campo] for r in records if r[campo])
    return c.most_common(n)

def turno_atual():
    h = datetime.now(BRT).hour
    if  6 <= h <= 11: return 'Manhã'
    if 12 <= h <= 17: return 'Tarde'
    if 18 <= h <= 23: return 'Noite'
    return 'Madrugada'

def proxima_turno(t):
    prox = {'Madrugada':'Manhã','Manhã':'Tarde','Tarde':'Noite','Noite':'Madrugada'}
    return prox.get(t, t)

def gerar_analise_diaria(records):
    now        = datetime.now(BRT)
    dia_num    = now.weekday()
    DIAS       = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']
    PLURAL     = ['Segundas','Terças','Quartas','Quintas','Sextas','Sábados','Domingos']
    dia_semana = DIAS[dia_num]
    plural     = PLURAL[dia_num]
    data_atual = now.strftime('%d/%m/%Y')
    hora_atual = now.strftime('%H:%M')

    hist  = dedup_bo([r for r in records if r['dia'] == dia_semana])
    datas = sorted(set(r['data'] for r in hist))
    nDias = len(datas)

    if nDias == 0:
        return (
            f"📊 *ANÁLISE — {plural} — GMBC*\n"
            f"📅 {dia_semana}, {data_atual} às {hora_atual}\n\n"
            f"⚠️ Nenhuma ocorrência registrada em {plural} até o momento.\n\n"
            f"🌐 dashboardgmbc.com.br"
        )

    total      = len(hist)
    media      = round(total / nDias, 1)
    tipos      = top_n(hist, 'tipo', 3)
    bairros    = top_n(hist, 'bairro', 3)
    turnos_c   = {t: sum(1 for r in hist if r['turno'] == t) for t in ORDEM_TURNO}
    top_tipo   = tipos[0]   if tipos   else ('–', 0)
    top_bairro = bairros[0] if bairros else ('–', 0)
    top_turno  = max(turnos_c, key=turnos_c.get)

    tipos_linhas   = '\n'.join(f"  • {t} — {pct(n, total)}" for t, n in tipos)
    bairros_linhas = '\n'.join(f"  • {b} — {n} oc." for b, n in bairros)

    return '\n'.join([
        f"📊 *ANÁLISE — {plural} — GMBC*",
        f"📅 {dia_semana}, {data_atual} · {hora_atual}",
        "",
        f"📌 *{total} ocorrências* em {nDias} {plural}",
        f"📈 Média histórica: *{media} oc./dia*",
        "",
        f"🔴 *Crimes mais frequentes:*",
        tipos_linhas,
        "",
        f"📍 *Bairros com mais ocorrências:*",
        bairros_linhas,
        "",
        f"⏰ *Turno crítico:* {top_turno} ({pct(turnos_c[top_turno], total)})",
        "",
        f"🌐 dashboardgmbc.com.br",
    ])


def gerar_previsao(records):
    now        = datetime.now(BRT)
    dia_num    = now.weekday()
    DIAS       = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']
    PLURAL     = ['Segundas','Terças','Quartas','Quintas','Sextas','Sábados','Domingos']
    dia_semana = DIAS[dia_num]
    plural     = PLURAL[dia_num]
    data_atual = now.strftime('%d/%m/%Y')
    hora_atual = now.strftime('%H:%M')
    t_atual    = turno_atual()
    t_proximo  = proxima_turno(t_atual)

    hist  = dedup_bo([r for r in records if r['dia'] == dia_semana])
    datas = sorted(set(r['data'] for r in hist))
    nDias = len(datas)

    if nDias == 0:
        return (
            f"📈 *PREVISÃO — {plural} — GMBC*\n"
            f"📅 {dia_semana}, {data_atual} às {hora_atual}\n\n"
            f"⚠️ Nenhuma ocorrência registrada em {plural} até o momento.\n\n"
            f"🌐 dashboardgmbc.com.br"
        )

    per_day   = [len([r for r in hist if r['data'] == d]) for d in datas]
    mean      = sum(per_day) / nDias
    std_dev   = math.sqrt(sum((n - mean)**2 for n in per_day) / nDias)
    min_exp   = max(0, round(mean - std_dev))
    max_exp   = round(mean + std_dev)
    turnos_c  = {t: sum(1 for r in hist if r['turno'] == t) for t in ORDEM_TURNO}
    total_t   = len(hist) or 1
    top_turno = max(turnos_c, key=turnos_c.get)
    tipos     = top_n(hist, 'tipo', 3)
    bairros   = top_n(hist, 'bairro', 3)
    max_b     = bairros[0][1] if bairros else 1

    trend_str  = 'Estável'
    trend_icon = '→'
    trend_diff = 0.0
    if len(datas) >= 4:
        mid  = len(datas) // 2
        avg1 = sum(len([r for r in hist if r['data'] == d]) for d in datas[:mid]) / mid
        avg2 = sum(len([r for r in hist if r['data'] == d]) for d in datas[mid:]) / (len(datas) - mid)
        trend_diff = avg2 - avg1
        if trend_diff > 1.5:
            trend_str, trend_icon = 'Crescente', '↗'
        elif trend_diff < -1.5:
            trend_str, trend_icon = 'Decrescente', '↘'

    risk_score = 0
    if mean > 8: risk_score += 2
    elif mean > 4: risk_score += 1
    if trend_icon == '↗': risk_score += 2
    if turnos_c[t_atual] / total_t > 0.35: risk_score += 1
    risk_level = 'ALTO' if risk_score >= 4 else 'MÉDIO' if risk_score >= 2 else 'BAIXO'
    risk_emoji = '🔴' if risk_level == 'ALTO' else '🟡' if risk_level == 'MÉDIO' else '🟢'

    bairros_linhas = []
    for i, (nome, qtd) in enumerate(bairros):
        score = qtd / max_b
        nivel = '🔴' if score > 0.6 else '🟡' if score > 0.3 else '🟢'
        bairros_linhas.append(f"  {nivel} {i+1}. {nome} — {qtd} oc.")

    tipos_linhas = '\n'.join(f"  • {t} — {pct(n, total_t)}" for t, n in tipos)

    linhas = [
        f"📈 *PREVISÃO — {plural} — GMBC*",
        f"📅 {dia_semana}, {data_atual} · {hora_atual}",
        "",
        f"{risk_emoji} *RISCO {risk_level}* — Previsão: *{min_exp}–{max_exp} oc.*",
        f"📊 Tendência: {trend_str} {trend_icon} · Base: {nDias} {plural} · Média: {round(mean,1)}/dia",
    ]
    if trend_icon == '↗':
        linhas.append(f"⚠️ *Atenção:* Tendência *crescente* (+{round(trend_diff,1)} oc./semana)")
    linhas += [
        "",
        f"⏱ *Turno atual:* {t_atual} ({pct(turnos_c[t_atual], total_t)}) ◀",
        f"⏱ *Próximo turno:* {t_proximo} ({pct(turnos_c[t_proximo], total_t)})",
        "",
        f"📍 *Bairros em alerta:*",
        *bairros_linhas,
        "",
        f"⚠️ *Crimes esperados:*",
        tipos_linhas,
        "",
        f"🌐 dashboardgmbc.com.br",
    ]
    return '\n'.join(linhas)


def gerar_relatorio_diario(records):
    now = datetime.now(BRT)
    hora_atual = now.strftime('%H:%M')

    datas_disponiveis = sorted(set(r['data'] for r in records if r['data']), reverse=True)
    if not datas_disponiveis:
        return "⚠️ Nenhuma ocorrência registrada na base."

    hoje = now.strftime('%Y-%m-%d')
    data_ref = hoje if hoje in datas_disponiveis else datas_disponiveis[0]
    registros_dia = [r for r in records if r['data'] == data_ref]

    DIAS = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']

    # Data de hoje para o cabeçalho
    d_hoje = datetime.strptime(hoje, '%Y-%m-%d')
    ano_h, mes_h, dia_h = hoje.split('-')
    dia_semana = DIAS[d_hoje.weekday()]
    label_hoje = f"{dia_h}/{mes_h}/{ano_h}"

    # Data de referência dos dados (pode ser diferente de hoje)
    ano_r, mes_r, dia_r = data_ref.split('-')
    label_data = f"{dia_r}/{mes_r}/{ano_r}"

    # itens fica com as linhas cruas (não deduplicadas): 2 celulares no mesmo
    # B.O. contam como 2 pra "itens mais furtados hoje" — igual ao gráfico de
    # itens do dashboard. O resto (total/tipos/bairros/turnos) é por
    # ocorrência (1 B.O. = 1), deduplicado.
    itens   = Counter(r['item']   for r in registros_dia if r.get('item')).most_common(5)
    registros_dia = dedup_bo(registros_dia)
    total = len(registros_dia)
    tipos   = Counter(r['tipo']   for r in registros_dia if r['tipo']).most_common(8)
    bairros = Counter(r['bairro'] for r in registros_dia if r['bairro']).most_common(5)

    ORDEM_T = ['Madrugada', 'Manhã', 'Tarde', 'Noite']
    turnos_cnt = Counter(r['turno'] for r in registros_dia if r['turno'])
    turnos = [(t, turnos_cnt.get(t, 0)) for t in ORDEM_T]

    # Totais por tipo principal
    furtos    = sum(n for t, n in tipos if 'Furto' in t and 'Tentativa' not in t)
    roubos    = sum(n for t, n in tipos if 'Roubo' in t and 'Tentativa' not in t)
    arrombs   = sum(n for t, n in tipos if 'Arrombamento' in t)
    tent_f    = sum(n for t, n in tipos if 'Tentativa de Furto' in t)
    tent_r    = sum(n for t, n in tipos if 'Tentativa de Roubo' in t)

    def pp(n): return f"{round(n/total*100)}%" if total else "0%"

    linhas = [
        "📋 *RELATÓRIO DO DIA — GMBC*",
        f"📅 {dia_semana}, {label_hoje} | {hora_atual}",
    ]
    if data_ref != hoje:
        linhas.append(f"_⚠️ Sem registros hoje — último dia com dados: {label_data}_")

    linhas += [
        "",
        f"📌 *Total: {total} ocorrências*",
        f"  🔴 Furtos: {furtos} ({pp(furtos)})",
        f"  🔴 Roubos: {roubos} ({pp(roubos)})",
        f"  🟠 Arromb.: {arrombs} ({pp(arrombs)})",
        f"  🟠 T.Furto: {tent_f} ({pp(tent_f)})",
    ]
    if tent_r:
        linhas.append(f"  🟠 T.Roubo: {tent_r} ({pp(tent_r)})")
    linhas += [
        "",
        "🔴 *Tipos de ocorrência:*",
    ]
    for t, n in tipos:
        linhas.append(f"  • {t}: {n} ({pp(n)})")

    linhas += ["", "⏰ *Por turno:*"]
    for t, n in turnos:
        linhas.append(f"  • {t}: {n} ({pp(n)})")

    linhas += ["", "📍 *Bairros afetados:*"]
    for b, n in bairros:
        linhas.append(f"  • {b}: {n} oc. ({pp(n)})")

    if itens:
        linhas += ["", "📦 *Itens furtados/roubados:*"]
        for it, n in itens:
            linhas.append(f"  • {it}: {n} ({pp(n)})")

    linhas += ["", "🌐 dashboardgmbc.com.br"]
    return '\n'.join(linhas)


AJUDA = (
    "*🛡️ Bot GMBC — Consulta de Ocorrências*\n\n"
    "*📊 Resumos rápidos:*\n"
    "  `Analise` → análise histórica do dia da semana\n"
    "  `Previsao` → previsão de risco para hoje\n"
    "  `Relatorio` → boletim do dia (hoje ou último com dados)\n"
    "  `Recarregar` → força atualização dos dados da planilha\n\n"
    "*Busca livre — digite qualquer palavra:*\n"
    "  `bicicleta`, `celular`, `rua 906`, `furto`, `noite`...\n"
    "  O bot mostra o total e um *menu numerado* (ano, mês, bairro, rua, turno...). Responda com o número.\n\n"
    "*Comandos específicos:*\n"
    "  `/bairro centro` · `/tipo furto` · `/turno noite`\n"
    "  `/resumo` · `/bairros` · `/tipos` · `/ajuda`\n\n"
    "🌐 *Relatório completo:* dashboardgmbc.com.br\n"
    "_Resposta em até 5 minutos._"
)


# ── Processar mensagem ────────────────────────────────────────────────────────
def processar(text, chat_id, records):
    t  = text.strip()
    # Em grupos o Telegram manda "/resumo@alertas_gmbc_bot" — tira o "@bot".
    if t.startswith('/'):
        cmd, _, resto = t.partition(' ')
        t = cmd.split('@')[0] + (' ' + resto if resto else '')
    tl = t.lower()
    tl_norm = sem_acento(t)

    if tl in ['/start','/ajuda','/help','ajuda','help']:
        return send_message(chat_id, AJUDA)

    # Resposta numérica ao menu da última busca ("1", "2"...) ou "menu".
    if tl_norm.isdigit() and len(tl_norm) <= 2:
        resp = responder_opcao(chat_id, int(tl_norm))
        return send_message(chat_id, resp or "Nenhuma busca ativa. Digite o que procura, ex.: *celular*, *rua 906*, *centro*.")
    if tl in ('menu', '/menu') and chat_id in ESTADO_BUSCA:
        est = ESTADO_BUSCA[chat_id]
        return send_message(chat_id, _cabecalho_busca(est) + '\n\n' + _menu_busca(est))

    if tl == '/resumo':
        return send_message(chat_id, consultar_resumo(records))

    if tl == '/bairros':
        bairros = sorted(set(r['bairro'] for r in records if r['bairro']))
        return send_message(chat_id, "*📍 Bairros disponíveis:*\n" + '\n'.join(f"• {b}" for b in bairros))

    if tl == '/tipos':
        tipos = sorted(set(r['tipo'] for r in records if r['tipo']))
        return send_message(chat_id, "*🔴 Tipos de crime:*\n" + '\n'.join(f"• {t}" for t in tipos))

    if tl in ['/chatid', '/id']:
        return send_message(chat_id, f"ID deste chat: `{chat_id}`")

    if tl in ('/bairro', '/tipo', '/turno'):
        return send_message(chat_id, f"Use assim: *{tl} <nome>*. Ex.: {tl} " +
                            {'/bairro': 'centro', '/tipo': 'furto', '/turno': 'noite'}[tl])

    if tl.startswith('/bairro '):
        return send_message(chat_id, consultar_bairro(records, t[8:]))

    if tl.startswith('/tipo '):
        return send_message(chat_id, consultar_tipo(records, t[6:]))

    if tl.startswith('/turno '):
        return send_message(chat_id, consultar_turno(records, t[7:]))

    # Detecta automaticamente se é bairro, tipo, turno ou mês conhecido.
    # Prioriza match exato antes de match parcial/token para evitar que um
    # bairro parecido (ex: "Amores" vs "Praia Dos Amores") roube a resposta
    # de um outro que so bate por substring.
    bairros_lista = sorted(set(r['bairro'] for r in records if r['bairro']))
    bairro_match = (
        next((b for b in bairros_lista if tl_norm == sem_acento(b)), None) or
        next((b for b in bairros_lista if len(tl_norm) >= 4 and tl_norm in sem_acento(b).split()
              and tl_norm not in ('praia', 'jardim', 'dos', 'sul')), None)
    )
    if bairro_match:
        return send_message(chat_id, consultar_bairro(records, bairro_match))

    tipos_lista = sorted(set(r['tipo'] for r in records if r['tipo']))
    tipo_match = (
        next((tp for tp in tipos_lista if tl_norm == sem_acento(tp)), None) or
        next((tp for tp in tipos_lista if len(tl_norm) >= 4 and tl_norm in sem_acento(tp).split()
              and tl_norm != 'tentativa'), None)
    )
    if tipo_match:
        return send_message(chat_id, consultar_tipo(records, tipo_match))

    turno_palavras = ['madrugada','manha','manhã','tarde','noite']
    if tl_norm in turno_palavras:
        return send_message(chat_id, consultar_turno(records, t))

    meses_norm = {sem_acento(m) for m in MES_PARA_NUM.keys()}
    if tl_norm in meses_norm:
        return send_message(chat_id, consultar_mes(records, t))

    # Palavras-chave do dashboard
    tl_sa = sem_acento(tl)
    print(f"  [debug] tl='{tl}' tl_sa='{tl_sa}'")
    if tl_sa in ('analise', '/analise'):
        print("  [debug] handler: ANALISE")
        return send_message(chat_id, gerar_analise_diaria(records))
    if tl_sa in ('previsao', '/previsao'):
        print("  [debug] handler: PREVISAO")
        return send_message(chat_id, gerar_previsao(records))
    if tl_sa in ('relatorio', '/relatorio'):
        print("  [debug] handler: RELATORIO")
        try:
            return send_message(chat_id, gerar_relatorio_diario(records))
        except Exception as e:
            print(f"  [debug] ERRO em gerar_relatorio_diario: {e}")
            return send_message(chat_id, f"⚠️ Erro ao gerar relatório: {e}")

    # Busca universal em todos os campos
    return send_message(chat_id, busca_universal(records, t, chat_id))


# ── Modo Railway: long-polling contínuo ───────────────────────────────────────
RELOAD_INTERVAL = 1800  # recarrega a cada 30 minutos

def run_continuo():
    import time
    print("Bot GMBC iniciado (modo Railway - resposta instantanea)")
    records = carregar_dados()
    print(f"  {len(records)} registros carregados.")
    offset = None
    ultima_carga = time.time()
    force_reload = False

    while True:
        # Recarrega dados a cada 30 minutos ou quando forçado
        if force_reload or time.time() - ultima_carga > RELOAD_INTERVAL:
            print("Recarregando dados...")
            records = carregar_dados()
            print(f"  {len(records)} registros carregados.")
            ultima_carga = time.time()
            force_reload = False

        # Long-polling: aguarda até 30s por novas mensagens
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates?timeout=30"
        if offset:
            url += f"&offset={offset}"
        try:
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req, timeout=40) as resp:
                updates = json.loads(resp.read().decode()).get('result', [])
        except Exception as e:
            print(f"Erro polling: {e}")
            time.sleep(5)
            continue

        now_unix = time.time()
        for upd in updates:
            offset = upd['update_id'] + 1
            msg = upd.get('message') or upd.get('edited_message')
            if not msg:
                continue
            if now_unix - msg.get('date', 0) > 300:
                continue
            text = msg.get('text', '').strip()
            if not text:
                continue
            cid  = str(msg['chat']['id'])
            nome = msg['chat'].get('first_name') or msg['chat'].get('title') or cid
            print(f"[{datetime.now(BRT).strftime('%H:%M:%S')}] [{nome}] '{text}'")
            if sem_acento(text.strip().lower()) in ('recarregar', '/recarregar'):
                force_reload = True
                send_message(cid, "♻️ Recarregando dados da planilha...")
            else:
                try:
                    processar(text, cid, records)
                except Exception as e:
                    print(f"  ERRO ao processar '{text}': {e}")
                    send_message(cid, "⚠️ Erro ao processar sua consulta. Tente reformular.")


# ── Modo GitHub Actions: one-shot ─────────────────────────────────────────────
def run_once():
    import time
    print("Bot GMBC - verificando mensagens...")
    print("Carregando dados...")
    records = carregar_dados()
    print(f"  {len(records)} registros carregados.")

    updates = get_updates()
    if not updates:
        print("Nenhuma mensagem pendente.")
        sys.exit(0)

    now_unix = time.time()
    processed = 0
    last_offset = None

    for upd in updates:
        last_offset = upd['update_id'] + 1
        msg = upd.get('message') or upd.get('edited_message')
        if not msg:
            continue
        if now_unix - msg.get('date', 0) > 3600:
            continue
        text = msg.get('text', '').strip()
        if not text:
            continue
        cid  = str(msg['chat']['id'])
        nome = msg['chat'].get('first_name') or msg['chat'].get('title') or cid
        print(f"[{datetime.now(BRT).strftime('%H:%M:%S')}] [{nome}] '{text}'")
        try:
            processar(text, cid, records)
        except Exception as e:
            print(f"  ERRO ao processar '{text}': {e}")
            send_message(cid, "⚠️ Erro ao processar sua consulta. Tente reformular.")
        processed += 1

    if last_offset:
        get_updates(offset=last_offset)
    print(f"\nConcluido. {processed} mensagem(ns) processada(s).")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    # Railway define estas variáveis automaticamente
    is_railway = bool(
        _os.environ.get('RAILWAY_ENVIRONMENT_NAME') or
        _os.environ.get('RAILWAY_PROJECT_NAME') or
        _os.environ.get('RAILWAY_SERVICE_NAME')
    )
    # BOT_RUN_MODE=continuo força o modo long-polling em qualquer host (ex: Easypanel/VPS)
    run_mode = _os.environ.get('BOT_RUN_MODE', '').strip().lower()
    if is_railway or run_mode == 'continuo':
        run_continuo()
    else:
        run_once()
