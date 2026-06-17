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

BOT_TOKEN       = _os.environ.get('BOT_TOKEN',       '8971067969:AAF73XtvvHyhkb_KX0dC3Tny6DQ6DtRdjjM').strip()
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

    excel_local = 'secretario.xlsx'
    if os.path.exists(excel_local):
        print(f"  Lendo {excel_local} local...")
        try:
            df = pd.read_excel(excel_local, sheet_name='DADOS', engine='openpyxl')
        except Exception as e:
            print(f"ERRO ao ler planilha local: {e}"); sys.exit(1)
    else:
        print(f"  Baixando planilha do Google Drive...")
        url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_DRIVE_ID}/export?format=xlsx"
        req = urllib.request.Request(url, headers={'User-Agent':'Mozilla/5.0'})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                data = resp.read()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            tmp.write(data); tmp.close()
            df = pd.read_excel(tmp.name, sheet_name='DADOS', engine='openpyxl')
            os.unlink(tmp.name)
        except Exception as e:
            print(f"ERRO ao baixar planilha: {e}"); sys.exit(1)

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
def close_other_sessions():
    """Encerra sessões de long-polling concorrentes (resolve HTTP 409 Conflict)."""
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates?timeout=0&offset=-1"
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            pass
    except Exception:
        pass

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
                print(f"  ERRO ao enviar (parse_mode={parse_mode}): {e}")
                break


# ── Consultas ─────────────────────────────────────────────────────────────────
def pct(n, total):
    return f"{round(n/total*100)}%" if total else "0%"

def consultar_bairro(records, query):
    bairros = sorted(set(r['bairro'] for r in records if r['bairro']))
    q = sem_acento(query.strip())
    match = next((b for b in bairros if q in sem_acento(b) or sem_acento(b) in q), None)

    if not match:
        lista = '\n'.join(f"• {b}" for b in bairros[:15])
        return f"Bairro *'{query}'* não encontrado.\n\n*Bairros disponíveis:*\n{lista}"

    dados  = [r for r in records if r['bairro'] == match]
    total  = len(dados)
    tipos  = Counter(r['tipo']     for r in dados if r['tipo']).most_common(5)
    turnos = Counter(r['turno']    for r in dados if r['turno']).most_common()
    ruas   = Counter(r['endereco'] for r in dados if r['endereco']).most_common(5)
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
    match = next((t for t in tipos if q in sem_acento(t) or sem_acento(t) in q), None)

    if not match:
        lista = '\n'.join(f"• {t}" for t in tipos)
        return f"Tipo *'{query}'* não encontrado.\n\n*Tipos disponíveis:*\n{lista}"

    dados   = [r for r in records if r['tipo'] == match]
    total   = len(dados)
    bairros = Counter(r['bairro'] for r in dados if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']  for r in dados if r['turno']).most_common()
    ruas    = Counter(r['endereco'] for r in dados if r['endereco']).most_common(3)
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

    dados   = [r for r in records if r['turno'] == match]
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

def consultar_resumo(records):
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

def busca_universal(records, query):
    """Busca em todos os campos: bairro, tipo, item, marca, IMEI, turno, logradouro, dia."""
    q = sem_acento(query.strip())
    campos = ['bairro','tipo','item','marca','imei','turno','endereco','dia','bo']

    # Variações da busca: original, sem 's' final, sem 'es' final
    variacoes = {q}
    if q.endswith('s') and len(q) > 3:
        variacoes.add(q[:-1])     # bicicletas → bicicleta
    if q.endswith('es') and len(q) > 4:
        variacoes.add(q[:-2])     # veiculos → veiculo

    def campo_match(valor):
        v = sem_acento(valor)
        return any(var in v for var in variacoes)

    def item_match_exato(valor):
        v = sem_acento(valor)
        return any(var == v for var in variacoes)

    matches = [r for r in records if any(campo_match(r.get(c,'')) for c in campos)]

    if not matches:
        return (
            f"Nenhuma ocorrência encontrada para *'{query}'*.\n\n"
            "Tente: nome de bairro, tipo de crime, item (bicicleta, celular...), IMEI, turno ou logradouro.\n"
            "Digite /ajuda para ver todos os comandos."
        )

    # Registros onde a busca bateu no campo ITEM (igual ao filtro do dashboard)
    matches_item_exato = [r for r in matches if item_match_exato(r.get('item',''))]
    matches_item = matches_item_exato or [r for r in matches if campo_match(r.get('item',''))]
    total_item = len(matches_item)

    MES_NOME = {'01':'Janeiro','02':'Fevereiro','03':'Março','04':'Abril','05':'Maio',
                '06':'Junho','07':'Julho','08':'Agosto','09':'Setembro','10':'Outubro',
                '11':'Novembro','12':'Dezembro'}

    # Usa registros do item principal para estatísticas (igual ao dashboard)
    base = matches_item if total_item > 0 else matches
    total_base = len(base)

    tipos   = Counter(r['tipo']    for r in base if r['tipo']).most_common(5)
    bairros = Counter(r['bairro']  for r in base if r['bairro']).most_common(5)
    turnos  = Counter(r['turno']   for r in base if r['turno']).most_common()
    ruas    = Counter(r['endereco'] for r in base if r['endereco'] and r['endereco'] != 'nan').most_common(5)
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

    hist  = [r for r in records if r['dia'] == dia_semana]
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

    hist  = [r for r in records if r['dia'] == dia_semana]
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

    ano, mes, dia = data_ref.split('-')
    d = datetime.strptime(data_ref, '%Y-%m-%d')
    DIAS = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']
    dia_semana = DIAS[d.weekday()]
    label_data = f"{dia}/{mes}/{ano}"

    total = len(registros_dia)
    tipos   = Counter(r['tipo']   for r in registros_dia if r['tipo']).most_common(8)
    bairros = Counter(r['bairro'] for r in registros_dia if r['bairro']).most_common(5)
    itens   = Counter(r['item']   for r in registros_dia if r.get('item')).most_common(5)

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
        f"📅 {dia_semana}, {label_data} | {hora_atual}",
    ]
    if data_ref != hoje:
        linhas.append(f"_⚠️ Sem registros hoje — último dia com dados: {label_data}_")

    linhas += [
        "",
        f"📌 *Total: {total} ocorrências*",
        f"  🔴 Furtos: {furtos} ({pp(furtos)})  |  Roubos: {roubos} ({pp(roubos)})",
        f"  🟠 Arromb.: {arrombs} ({pp(arrombs)})  |  T.Furto: {tent_f} ({pp(tent_f)})" + (f"  |  T.Roubo: {tent_r} ({pp(tent_r)})" if tent_r else ""),
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
    "  `Relatorio` → boletim do dia (hoje ou último com dados)\n\n"
    "*Busca livre — digite qualquer palavra:*\n"
    "  `bicicleta`, `celular`, `centro`, `furto`, `noite`...\n\n"
    "*Comandos específicos:*\n"
    "  `/bairro centro` · `/tipo furto` · `/turno noite`\n"
    "  `/resumo` · `/bairros` · `/tipos` · `/ajuda`\n\n"
    "🌐 *Relatório completo:* dashboardgmbc.com.br\n"
    "_Resposta em até 5 minutos._"
)


# ── Processar mensagem ────────────────────────────────────────────────────────
def processar(text, chat_id, records):
    t  = text.strip()
    tl = t.lower()
    tl_norm = sem_acento(t)

    if tl in ['/start','/ajuda','/help','ajuda','help']:
        return send_message(chat_id, AJUDA)

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

    if tl.startswith('/bairro '):
        return send_message(chat_id, consultar_bairro(records, t[8:]))

    if tl.startswith('/tipo '):
        return send_message(chat_id, consultar_tipo(records, t[6:]))

    if tl.startswith('/turno '):
        return send_message(chat_id, consultar_turno(records, t[7:]))

    # Detecta automaticamente se é bairro, tipo ou turno conhecido
    bairros_lista = sorted(set(r['bairro'] for r in records if r['bairro']))
    for b in bairros_lista:
        b_norm = sem_acento(b)
        if tl_norm == b_norm or tl_norm in b_norm.split() or b_norm == tl_norm:
            return send_message(chat_id, consultar_bairro(records, t))

    tipos_lista = sorted(set(r['tipo'] for r in records if r['tipo']))
    for tp in tipos_lista:
        tp_norm = sem_acento(tp)
        if tl_norm == tp_norm or tl_norm in tp_norm:
            return send_message(chat_id, consultar_tipo(records, t))

    turno_palavras = ['madrugada','manha','manhã','tarde','noite']
    if tl_norm in turno_palavras:
        return send_message(chat_id, consultar_turno(records, t))

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
    return send_message(chat_id, busca_universal(records, t))


# ── Modo Railway: long-polling contínuo ───────────────────────────────────────
def run_continuo():
    import time
    print("Bot GMBC iniciado (modo Railway - resposta instantanea)")
    records = carregar_dados()
    print(f"  {len(records)} registros carregados.")
    offset = None
    ultima_carga = time.time()

    while True:
        # Recarrega dados a cada 2 horas
        if time.time() - ultima_carga > 7200:
            print("Recarregando dados...")
            records = carregar_dados()
            print(f"  {len(records)} registros carregados.")
            ultima_carga = time.time()

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
            processar(text, cid, records)


# ── Modo GitHub Actions: one-shot ─────────────────────────────────────────────
def run_once():
    import time
    print("Bot GMBC - verificando mensagens...")
    close_other_sessions()
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
        processar(text, cid, records)
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
    if is_railway:
        run_continuo()
    else:
        run_once()

    # Marca todas as atualizações como lidas
    if last_offset:
        get_updates(offset=last_offset)

    print(f"\nConcluído. {processed} mensagem(ns) processada(s).")
