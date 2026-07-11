import pandas as pd
import json
import os
import re
import hashlib
import urllib.request
import warnings
warnings.filterwarnings('ignore')

# ╔══════════════════════════════════════════════════════╗
# ║         SENHA DE ACESSO AO DASHBOARD                 ║
# ║  Altere SENHA_DASHBOARD para mudar a senha           ║
# ╚══════════════════════════════════════════════════════╝
SENHA_DASHBOARD = "GuardaBC2026"
_senha_hash = hashlib.sha256(SENHA_DASHBOARD.encode('utf-8')).hexdigest()
import base64 as _b64
_senha_b64 = _b64.b64encode(SENHA_DASHBOARD.encode('utf-8')).decode('ascii')

# ── Baixar/embedar bibliotecas JS/CSS ─────────────────────────────────────────
LIBS_DIR = 'libs'
os.makedirs(LIBS_DIR, exist_ok=True)

def fetch_lib(url, filename):
    path = os.path.join(LIBS_DIR, filename)
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            return f.read()
    print(f"Baixando {filename}...")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read()
        try:
            content = raw.decode('utf-8')
        except UnicodeDecodeError:
            content = raw.decode('latin-1')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"  OK: {len(content)//1024} KB")
        return content
    except Exception as e:
        print(f"  FALHA ao baixar {filename}: {e}")
        return ''

print("Carregando bibliotecas (primeira vez faz download)...")
plotly_js    = fetch_lib('https://cdn.plot.ly/plotly-2.26.0.min.js',
                         'plotly.min.js')
leaflet_css  = fetch_lib('https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css',
                         'leaflet.min.css')
leaflet_js   = fetch_lib('https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js',
                         'leaflet.min.js')
cluster_css1 = fetch_lib('https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/MarkerCluster.min.css',
                         'MarkerCluster.min.css')
cluster_css2 = fetch_lib('https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.min.css',
                         'MarkerCluster.Default.min.css')
cluster_js   = fetch_lib('https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.min.js',
                         'leaflet.markercluster.min.js')
jspdf_js     = fetch_lib('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
                         'jspdf.umd.min.js')

# ── Logo da Guarda Municipal (usado nos slides em PDF) ────────────────────────
LOGO_GMBC_B64 = ''
if os.path.exists('logo_gmbc.png'):
    with open('logo_gmbc.png', 'rb') as f:
        LOGO_GMBC_B64 = 'data:image/png;base64,' + _b64.b64encode(f.read()).decode('ascii')

def _embed_js(code):
    if not code:
        return ''
    # Só escapa "</script" (caso insensível) — trocar todo "</" quebra
    # regexes legítimas no código das libs, tipo /</g (visto no pptxgenjs).
    safe = re.sub(r'</script', '<\\/script', code, flags=re.IGNORECASE)
    return '<script>' + safe + '</script>'

def _embed_css(code):
    return f'<style>{code}</style>' if code else ''

# ── Carregar geocache ─────────────────────────────────────────────────────────
GEOCACHE = {}
if os.path.exists('geocache.json'):
    with open('geocache.json', 'r', encoding='utf-8') as f:
        GEOCACHE = json.load(f)
    print(f"Geocache: {sum(1 for v in GEOCACHE.values() if v)} coordenadas carregadas")

# ── Carregar dados de Inteligência Criminal (opcional) ────────────────────────
INTEL_DATA = None
if os.path.exists('bos_grupos.json'):
    with open('bos_grupos.json', 'r', encoding='utf-8') as f:
        INTEL_DATA = json.load(f)
INTEL_JSON = json.dumps(INTEL_DATA, ensure_ascii=False) if INTEL_DATA else 'null'

# ── Carregar e limpar dados ───────────────────────────────────────────────────
df = pd.read_excel('secretario.xlsx', sheet_name='DADOS', engine='openpyxl')

def norm_tipo(v):
    if pd.isna(v): return ''
    v = str(v).strip().upper()
    if not v or v == 'NAN': return ''
    if 'TENTATIVA' in v and 'ROUBO' in v: return 'Tentativa de Roubo'
    if 'TENTATIVA' in v and 'FURTO' in v: return 'Tentativa de Furto'
    if v == 'FURTO': return 'Furto'
    if v == 'ROUBO': return 'Roubo'
    if 'ARROMBAMENTO' in v: return 'Arrombamento'
    return v.title()

ITEM_MAP = {
    'VEICULO':'Veículo','VEILCULO':'Veículo','VEICULO ':'Veículo',
    'CELULAR':'Celular','FIAÇAO ELÉTRICA':'Fiação Elétrica',
    'FIAÇAO ELETRICA':'Fiação Elétrica','FIAÇÃO ELÉTRICA':'Fiação Elétrica',
    'ELETRODOMESTICO':'Eletrodoméstico','COMBUSTIVEL':'Combustível',
    'JETSKI':'JetSki','ALUMINIO':'Alumínio','VESTUARIO':'Vestuário',
    'PRODUTOS':'Produtos','REBOQUE':'Reboque','MERCADORIAS':'Mercadorias',
    'OBJETOS':'Objetos','DINHEIRO':'Dinheiro','SCOOTER':'Scooter',
    'BICICLETA':'Bicicleta','ELETRÔNICO':'Eletrônico','ELETRONICO':'Eletrônico',
    'MOTOCICLETA':'Motocicleta',
}
def norm_item(v):
    if pd.isna(v): return ''
    u = str(v).strip().upper()
    if not u or u == 'NAN': return ''
    return ITEM_MAP.get(u, str(v).strip().title())

BAIRRO_MAP = {
    'BARRA SUL':'Barra Sul','SÃO J. TADEU':'São J. Tadeu',
    'SAO J. TADEU':'São J. Tadeu','N. ESPERANÇA':'N. Esperança',
    'N. ESPERANCA':'N. Esperança','MUNICIPIOS':'Municípios',
    'NAÇÕES':'Nações','NACOES':'Nações','PONTAL NORTE':'Pontal Norte',
    'VILA REAL':'Vila Real','PIONEIROS':'Pioneiros','ARIRIBA':'Ariribá',
    'CENTRO':'Centro','ESTADOS':'Estados','BARRA':'Barra',
}
def norm_bairro(v):
    if pd.isna(v): return ''
    u = str(v).strip().upper()
    if not u or u == 'NAN': return ''
    return BAIRRO_MAP.get(u, str(v).strip().title())

DIA_MAP = {
    'SEGUNDA':'Segunda','SEGUNDA-FEIRA':'Segunda',
    'TERÇA':'Terça','TERCA':'Terça','TERÇA-FEIRA':'Terça','TERCA-FEIRA':'Terça',
    'QUARTA':'Quarta','QUARTA-FEIRA':'Quarta',
    'QUINTA':'Quinta','QUINTA-FEIRA':'Quinta',
    'SEXTA':'Sexta','SEXTA-FEIRA':'Sexta',
    'SABADO':'Sábado','SÁBADO':'Sábado','SÁBADO-FEIRA':'Sábado',
    'DOMINGO':'Domingo',
}
def norm_dia(v):
    if pd.isna(v): return ''
    u = str(v).strip().upper()
    if not u or u == 'NAN': return ''
    return DIA_MAP.get(u, str(v).strip().title())

TURNO_MAP = {
    'MANHA':'Manhã','MANHÃ':'Manhã',
    'TARDE':'Tarde','NOITE':'Noite','MADRUGADA':'Madrugada',
}
def norm_turno(v):
    if pd.isna(v): return ''
    u = str(v).strip().upper()
    if not u or u == 'NAN': return ''
    return TURNO_MAP.get(u, str(v).strip().title())

def calcular_turno(hora_val, turno_fallback=''):
    """Calcula turno pela hora:
       Manhã 06-11h | Tarde 12-17h | Noite 18-23h | Madrugada 00-05h
       Usa turno_fallback se hora inválida/ausente."""
    try:
        if hasattr(hora_val, 'hour'):
            h = hora_val.hour
        elif pd.notna(hora_val) and str(hora_val).strip():
            h = int(str(hora_val).strip().split(':')[0])
        else:
            return norm_turno(turno_fallback)
        if  6 <= h <= 11: return 'Manhã'
        if 12 <= h <= 17: return 'Tarde'
        if 18 <= h <= 23: return 'Noite'
        return 'Madrugada'   # 00-05h
    except Exception:
        return norm_turno(turno_fallback)

MES_MAP = {
    'ABRIL':'Abril','MAIO':'Maio','MARCO':'Março','MARÇO':'Março',
    'JANEIRO':'Janeiro','FEVEREIRO':'Fevereiro','JUNHO':'Junho',
    'JULHO':'Julho','AGOSTO':'Agosto','SETEMBRO':'Setembro',
    'OUTUBRO':'Outubro','NOVEMBRO':'Novembro','DEZEMBRO':'Dezembro',
}
def norm_mes(v):
    if pd.isna(v): return ''
    u = str(v).strip().upper()
    if not u or u == 'NAN': return ''
    return MES_MAP.get(u, str(v).strip().title())

df['TIPIFICACAO'] = df['TIPIFICACAO'].apply(norm_tipo)
df['ITEM']        = df['ITEM'].apply(norm_item)
df['BAIRRO']      = df['BAIRRO'].apply(norm_bairro)
df['DIA_SEMANA']  = df['DIA_SEMANA'].apply(norm_dia)
df['MES']         = df['MES'].apply(norm_mes)
# Turno calculado pela HORA (06-11h=Manhã, 12-17h=Tarde, 18-23h=Noite, 00-05h=Madrugada)
df['TURNO']       = df.apply(lambda r: calcular_turno(r['HORA'], r['TURNO']), axis=1)

# Remover linhas sem B.O. ou sem tipificação (células em branco / linhas vazias)
df = df[
    df['B.O.'].notna() &
    (df['B.O.'].astype(str).str.strip().str.upper() != 'NAN') &
    (df['B.O.'].astype(str).str.strip() != '') &
    (df['TIPIFICACAO'] != '')
].copy()
df['DATA_COMPLETA'] = pd.to_datetime(
    df['ANO'].astype(str) + '-' +
    df['MES'].str.upper().map({
        'JANEIRO':'01','FEVEREIRO':'02','MARÇO':'03','MARCO':'03',
        'ABRIL':'04','MAIO':'05','JUNHO':'06','JULHO':'07',
        'AGOSTO':'08','SETEMBRO':'09','OUTUBRO':'10',
        'NOVEMBRO':'11','DEZEMBRO':'12'
    }) + '-' +
    df['DATA'].astype(str).str.zfill(2),
    format='%Y-%m-%d', errors='coerce'
)
df['DATA_STR'] = df['DATA_COMPLETA'].dt.strftime('%Y-%m-%d')
df['HORA_STR'] = df['HORA'].apply(lambda x: x.strftime('%H:%M') if hasattr(x,'strftime') else str(x)[:5] if pd.notna(x) else '')
df['BO']       = df['B.O.'].fillna('').astype(str)
df['ENDERECO'] = df['ENDEREÇO'].fillna('').astype(str)
df['REF']      = df['PONTO_REFERENCIA'].fillna('').astype(str)
df['MAPA_URL'] = df['MAPA'].fillna('').astype(str)
_desc_col = next((c for c in df.columns if 'DESCRI' in c.upper()), None)
df['DESCRICAO'] = df[_desc_col].fillna('').astype(str).apply(lambda x: '' if x.lower() == 'nan' else x) if _desc_col else pd.Series('', index=df.index)
def clean_text_cell(row, col):
    if col not in df.columns or pd.isna(row[col]):
        return ''
    val = row[col]
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    txt = str(val).strip()
    return '' if txt.lower() in ('', 'nan') else txt

def get_coords(mapa_str):
    geo = GEOCACHE.get(str(mapa_str).strip())
    if geo:
        return geo['lat'], geo['lon']
    return None, None
df['LAT'], df['LON'] = zip(*df['MAPA_URL'].apply(get_coords))

# ── Exportar registros como JSON ──────────────────────────────────────────────
records = []
for _, r in df.iterrows():
    records.append({
        'data': r['DATA_STR'],
        'mes': r['MES'],
        'ano': int(r['ANO']),
        'hora': r['HORA_STR'],
        'turno': r['TURNO'],
        'dia': r['DIA_SEMANA'],
        'bo': r['BO'],
        'tipo': r['TIPIFICACAO'],
        'item': r['ITEM'],
        'descricao': r['DESCRICAO'],
        'marca': str(r['MARCA_MODELO']) if 'MARCA_MODELO' in df.columns and pd.notna(r['MARCA_MODELO']) else '',
        'cor': clean_text_cell(r, 'COR'),
        'detalhes': clean_text_cell(r, 'DETALHES'),
        'imei': clean_text_cell(r, 'IMEI'),
        'placa': clean_text_cell(r, 'PLACA'),
        'numero_serie': clean_text_cell(r, 'NUMERO_SERIE'),
        'recuperado': clean_text_cell(r, 'RECUPERADO'),
        'endereco': r['ENDERECO'],
        'bairro': r['BAIRRO'],
        'ref': r['REF'],
        'mapa': r['MAPA_URL'],
        'lat': r['LAT'] if pd.notna(r['LAT']) else None,
        'lon': r['LON'] if pd.notna(r['LON']) else None,
        'link': str(r['LINK']).strip() if 'LINK' in df.columns and pd.notna(r['LINK']) and str(r['LINK']).strip() not in ('', 'nan') else '',
    })

data_json = json.dumps(records, ensure_ascii=True)
data_json = data_json.replace('</', '<\\/')            # impede </script> fechar a tag

# ── Template HTML ─────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Segurança – Balneário Camboriú 2026</title>
__PLOTLY_JS__
__LEAFLET_CSS__
__LEAFLET_JS__
__CLUSTER_CSS1__
__CLUSTER_CSS2__
__CLUSTER_JS__
__JSPDF_JS__
<style>
:root {{
  --azul:#0078D4; --azul2:#106EBE; --azul-clr:#50B2FF;
  --laranja:#E07B00; --verde:#107C10; --vermelho:#D13438;
  --roxo:#8764B8; --amarelo:#FFB900; --bg:#F3F2F1;
  --sidebar:#1A1A2E; --card:#FFFFFF;
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);display:flex;flex-direction:column;min-height:100vh;}}

/* HEADER */
.header{{
  background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);
  color:white;padding:0;display:flex;flex-direction:column;
  box-shadow:0 2px 8px rgba(0,0,0,.3);
  position:sticky;top:0;z-index:100;
}}
.header-row1{{
  display:flex;align-items:center;gap:10px;
  padding:9px 20px;border-bottom:1px solid rgba(255,255,255,.12);
}}
.header-row2{{
  display:flex;align-items:center;gap:6px;
  padding:5px 20px 6px;flex-wrap:wrap;
  justify-content:flex-end;
}}
.header h1{{font-size:15px;font-weight:700;letter-spacing:.3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;min-width:0;}}
.pesquisa-header{{display:flex;align-items:center;gap:4px;flex-shrink:0;}}
.pesquisa-header input{{
  border:1.5px solid rgba(255,255,255,.35);border-radius:5px;
  background:rgba(255,255,255,.12);color:white;
  padding:4px 10px;font-size:11px;font-family:inherit;outline:none;width:200px;
}}
.pesquisa-header input::placeholder{{color:rgba(255,255,255,.6);}}
.pesquisa-header input:focus{{border-color:rgba(255,255,255,.8);background:rgba(255,255,255,.2);}}
.pesquisa-header input::-webkit-search-cancel-button{{display:none;}}
.btn-pdf-r1{{display:none;}}
.btn-pdf-r2{{display:inline-flex;}}
.header .sub{{font-size:11px;opacity:.8;margin-top:2px;}}
.header-right{{display:flex;gap:7px;align-items:center;flex-shrink:0;}}
.btxt{{}}
.btn-wa-dash{{background:#25D366;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-wa-dash:hover{{background:#1DA851;}}
.badge{{background:rgba(255,255,255,.2);border-radius:4px;padding:3px 10px;font-size:11px;font-weight:600;}}
/* ── LOGIN ── */
.login-overlay{{position:fixed;top:0;left:0;width:100%;height:100%;
  background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);
  z-index:99999;display:flex;align-items:center;justify-content:center;}}
.login-card{{background:white;border-radius:14px;padding:40px 36px;width:min(380px,92vw);
  box-shadow:0 24px 64px rgba(0,0,0,.5);text-align:center;}}
.login-shield{{margin-bottom:10px;text-align:center;}}
.login-title{{font-size:15px;font-weight:800;color:#1A1A2E;margin-bottom:4px;}}
.login-sub{{font-size:11px;color:#666;margin-bottom:24px;line-height:1.6;}}
.login-input-wrap{{position:relative;width:100%;margin-bottom:10px;}}
.login-input{{width:100%;box-sizing:border-box;padding:11px 44px 11px 14px;
  border:2px solid #ddd;border-radius:7px;font-size:14px;
  font-family:inherit;transition:border-color .2s;}}
.login-input:focus{{outline:none;border-color:#333D65;}}
.login-eye{{position:absolute;right:12px;top:50%;transform:translateY(-50%);
  background:none;border:none;cursor:pointer;color:#999;font-size:18px;
  padding:0;line-height:1;user-select:none;}}
.login-eye:hover{{color:#333D65;}}
.login-btn{{width:100%;padding:11px;background:#333D65;color:white;border:none;
  border-radius:7px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;
  transition:background .2s;}}
.login-btn:hover{{background:#222847;}}
.login-error{{color:#D13438;font-size:12px;margin-top:8px;min-height:18px;}}
@keyframes shake{{0%,100%{{transform:translateX(0)}}25%{{transform:translateX(-7px)}}75%{{transform:translateX(7px)}}}}
.login-shake{{animation:shake .35s ease;}}
.btn-reset{{background:#D13438;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-reset:hover{{background:#A4262C;}}
.btn-pdf{{background:#107C10;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-pdf:hover{{background:#0B5A0B;}}
.btn-analise{{background:#5C2D91;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-analise:hover{{background:#3D1A6B;}}
.btn-sair{{background:transparent;color:rgba(255,255,255,.8);border:1px solid rgba(255,255,255,.4);
  border-radius:4px;padding:5px 12px;font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-sair:hover{{background:rgba(255,255,255,.15);color:white;border-color:rgba(255,255,255,.7);}}
.btn-prev{{background:#C05700;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-prev:hover{{background:#8F3F00;}}
.btn-predit{{background:linear-gradient(135deg,#4A0E8F,#7B2FBE);color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-predit:hover{{background:linear-gradient(135deg,#3A0A6F,#5A1F9E);}}
.btn-intelig{{background:linear-gradient(135deg,#1B4332,#2D6A4F);color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-intelig:hover{{background:linear-gradient(135deg,#0D2B1F,#1B4332);}}
.btn-relatorio{{background:#0097A7;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-relatorio:hover{{background:#006978;}}
.btn-resumoia{{background:linear-gradient(135deg,#0078D4,#00B7C3);color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-resumoia:hover{{background:linear-gradient(135deg,#005A9E,#008B96);}}
.btn-slides{{background:linear-gradient(135deg,#8E44AD,#C0392B);color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-slides:hover{{background:linear-gradient(135deg,#6C3483,#922B21);}}
.btn-slides:disabled{{opacity:.6;cursor:wait;}}
.chat-ia-panel{{position:fixed;top:0;right:-380px;width:360px;max-width:92vw;height:100vh;background:white;
  box-shadow:-4px 0 24px rgba(0,0,0,.25);z-index:9999;display:flex;flex-direction:column;
  transition:right .25s ease;font-family:inherit;}}
.chat-ia-panel.aberto{{right:0;}}
.chat-ia-header{{background:linear-gradient(135deg,#0078D4,#00B7C3);color:white;padding:14px 16px;
  font-weight:700;font-size:14px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0;}}
.chat-ia-log{{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px;}}
.chat-msg{{padding:9px 12px;border-radius:10px;font-size:12px;line-height:1.5;max-width:88%;}}
.chat-msg-user{{background:#0078D4;color:white;align-self:flex-end;border-bottom-right-radius:2px;}}
.chat-msg-ia{{background:#F0F0F0;color:#222;align-self:flex-start;border-bottom-left-radius:2px;text-align:justify;}}
.chat-ia-inputbar{{display:flex;gap:8px;padding:12px;border-top:1px solid #EEE;flex-shrink:0;align-items:flex-end;}}
.chat-ia-inputbar textarea{{flex:1;border:1px solid #DDD;border-radius:6px;padding:8px 10px;font-size:12px;
  font-family:inherit;min-width:0;resize:vertical;min-height:64px;max-height:200px;line-height:1.4;}}
.chat-ia-inputbar button{{background:#0078D4;color:white;border:none;border-radius:6px;padding:8px 14px;
  font-size:12px;font-weight:600;cursor:pointer;flex-shrink:0;}}
.chat-ia-inputbar button:hover{{background:#005A9E;}}
#chat-ia-mic{{background:#F0F0F0;color:#333;border:none;border-radius:6px;padding:8px 10px;
  font-size:14px;cursor:pointer;flex-shrink:0;}}
#chat-ia-mic:hover{{background:#E0E0E0;}}
#chat-ia-mic.gravando{{background:#D13438;color:white;animation:pulseGravando 1s infinite;}}
@keyframes pulseGravando{{0%{{opacity:1;}}50%{{opacity:.5;}}100%{{opacity:1;}}}}
#chat-ia-fab{{position:fixed;bottom:24px;right:24px;width:56px;height:56px;border-radius:50%;
  background:linear-gradient(135deg,#0078D4,#00B7C3);color:white;border:none;font-size:24px;
  box-shadow:0 4px 16px rgba(0,0,0,.3);cursor:pointer;z-index:9998;}}
#chat-ia-fab:hover{{background:linear-gradient(135deg,#005A9E,#008B96);}}
/* ── MODAL RELATÓRIO ── */
.rel-date-bar{{background:#f2f4f8;padding:11px 18px;border-bottom:1px solid #ddd;
  display:flex;align-items:center;gap:10px;flex-wrap:wrap;flex-shrink:0;}}
.rel-date-bar label{{font-size:11px;font-weight:600;color:#1A1A2E;}}
.rel-date-input{{padding:5px 9px;border:1px solid #ccc;border-radius:5px;font-size:12px;
  font-family:inherit;background:white;color:#1A1A2E;cursor:pointer;}}
.rel-date-input:focus{{outline:none;border-color:#0078D4;}}
.rel-gerar-btn{{background:#0078D4;color:white;border:none;border-radius:5px;padding:6px 16px;
  font-size:12px;font-weight:700;cursor:pointer;font-family:inherit;}}
.rel-gerar-btn:hover{{background:#005A9E;}}
/* ── RISCO BADGES ── */
.risco-alto{{background:#D13438;color:white;border-radius:4px;padding:2px 8px;font-weight:800;font-size:11px;}}
.risco-medio{{background:#E07B00;color:white;border-radius:4px;padding:2px 8px;font-weight:800;font-size:11px;}}
.risco-baixo{{background:#107C10;color:white;border-radius:4px;padding:2px 8px;font-weight:800;font-size:11px;}}
.risco-bar-wrap{{background:#eee;height:8px;border-radius:4px;width:70px;display:inline-block;vertical-align:middle;}}
.risco-bar-fill{{height:8px;border-radius:4px;}}
/* ── MODAL ANÁLISE DIÁRIA ── */
.analise-overlay{{position:fixed;top:0;left:0;width:100%;height:100%;
  background:rgba(0,0,0,.65);z-index:9999;display:none;align-items:center;justify-content:center;}}
.analise-overlay.ativo{{display:flex;}}
.analise-box{{background:#fff;width:min(840px,96vw);max-height:91vh;border-radius:10px;
  box-shadow:0 12px 48px rgba(0,0,0,.45);display:flex;flex-direction:column;overflow:hidden;}}
.analise-header{{background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);color:white;
  padding:13px 18px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0;}}
.analise-header h2{{margin:0;font-size:14px;font-weight:700;}}
.analise-close{{background:none;border:none;color:white;font-size:22px;cursor:pointer;line-height:1;padding:0 4px;opacity:.8;}}
.analise-close:hover{{opacity:1;}}
.analise-corpo{{padding:18px 20px;overflow-y:auto;flex:1;font-size:12px;}}
.analise-section{{margin-bottom:14px;}}
.analise-section h3{{font-size:11px;font-weight:700;color:#1A1A2E;
  border-bottom:2px solid #0078D4;padding-bottom:3px;margin:0 0 8px;
  text-transform:uppercase;letter-spacing:.5px;}}
.analise-table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:2px;}}
.analise-table th{{background:#1A1A2E;color:white;padding:4px 8px;text-align:left;font-weight:600;}}
.analise-table td{{padding:4px 8px;border-bottom:1px solid #eee;}}
.analise-table tr:nth-child(even){{background:#f7f8fa;}}
.analise-rec-ol{{padding-left:18px;margin:0;}}
.analise-rec-ol li{{margin-bottom:6px;line-height:1.55;}}
.analise-footer{{padding:11px 18px;background:#f2f4f8;border-top:1px solid #ddd;
  display:flex;gap:8px;justify-content:flex-end;flex-shrink:0;}}
.badge-resumo{{display:inline-block;background:#0078D4;color:white;
  border-radius:4px;padding:1px 9px;font-weight:700;margin-left:4px;font-size:13px;}}

/* IMPRESSÃO / PDF */
@media print{{
  *{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
  .sidebar,.sidebar-overlay,.menu-toggle,.header-right,.active-filters,
  .sel-hint,.btn-reset,.btn-pdf,.btn-analise,.badge,.analise-overlay{{display:none!important;}}
  .header{{
    background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%)!important;
    padding:10px 18px!important;box-shadow:none!important;
  }}
  .main-layout{{display:block!important;}}
  .content{{overflow:visible!important;padding:10px!important;}}
  .row,.row2,.row3,.row-full{{display:grid!important;}}
  .chart-card,.kpi-card,.insight-card,.rec-card,.plano-card,.mapa-card{{
    break-inside:avoid;box-shadow:none!important;
    border:1px solid #ddd!important;
  }}
  .kpi-row{{page-break-after:avoid;}}
  #mapa-crime{{height:300px!important;}}
  .section-header{{break-before:auto;}}
  .plano-grid-wide,.plano-grid,.rec-grid,.insight-grid{{
    display:grid!important;
  }}
  body{{font-size:11px;}}
}}

/* LAYOUT */
.main-layout{{display:flex;flex:1;overflow:hidden;}}

/* SIDEBAR */
.sidebar{{
  width:220px;min-width:220px;background:var(--sidebar);color:white;
  padding:16px 12px;overflow-y:auto;flex-shrink:0;
}}
.sidebar h2{{font-size:11px;text-transform:uppercase;letter-spacing:1px;
  color:#aaa;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #333;}}
.filter-group{{margin-bottom:16px;}}
.filter-label{{font-size:11px;font-weight:600;color:var(--azul-clr);
  margin-bottom:6px;display:block;text-transform:uppercase;letter-spacing:.5px;}}
.filter-item{{display:flex;align-items:center;gap:7px;margin-bottom:4px;cursor:pointer;}}
.filter-item input[type=checkbox]{{accent-color:var(--azul);width:13px;height:13px;cursor:pointer;}}
.filter-item label{{font-size:12px;color:#ddd;cursor:pointer;flex:1;}}
.filter-item .cnt{{font-size:10px;color:#888;}}
.filter-search{{
  width:100%;background:#2A2A3E;border:1px solid #444;color:white;
  border-radius:4px;padding:5px 8px;font-size:11px;margin-bottom:6px;
  font-family:inherit;outline:none;
}}
.filter-search:focus{{border-color:var(--azul);}}
.filter-scroll{{max-height:160px;overflow-y:auto;}}
.filter-scroll::-webkit-scrollbar{{width:4px;}}
.filter-scroll::-webkit-scrollbar-track{{background:#1A1A2E;}}
.filter-scroll::-webkit-scrollbar-thumb{{background:#444;border-radius:2px;}}

/* CONTENT */
.content{{flex:1;overflow-y:auto;padding:16px;}}

/* KPI */
.kpi-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px;}}
.kpi-card{{
  background:white;border-radius:8px;padding:12px 14px;
  border-top:4px solid var(--azul);box-shadow:0 2px 6px rgba(0,0,0,.07);
  transition:transform .15s,box-shadow .15s;cursor:default;
}}
.kpi-card:hover{{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.12);}}
.kpi-card.laranja{{border-color:var(--laranja);}}
.kpi-card.vermelho{{border-color:var(--vermelho);}}
.kpi-card.verde{{border-color:var(--verde);}}
.kpi-card.roxo{{border-color:var(--roxo);}}
.kpi-card.marrom{{border-color:#795548;}}
.kpi-card.teal{{border-color:#0097A7;}}
.kpi-card.indigo{{border-color:#3949AB;}}
.kpi-val{{font-size:27px;font-weight:700;color:var(--azul);line-height:1;}}
.kpi-card.laranja .kpi-val{{color:var(--laranja);}}
.kpi-card.vermelho .kpi-val{{color:var(--vermelho);}}
.kpi-card.verde .kpi-val{{color:var(--verde);}}
.kpi-card.roxo .kpi-val{{color:var(--roxo);}}
.kpi-card.marrom .kpi-val{{color:#795548;}}
.kpi-card.teal .kpi-val{{color:#0097A7;}}
.kpi-card.indigo .kpi-val{{color:#3949AB;}}
.kpi-label{{font-size:11px;color:#666;margin-top:5px;font-weight:500;}}
.kpi-sub{{font-size:10px;color:#999;margin-top:2px;}}
.kpi-media{{font-size:10px;color:#555;margin-top:4px;background:#F4F8FF;border-radius:4px;padding:2px 6px;display:inline-block;font-weight:500;}}
.kpi-count{{margin-top:6px;display:flex;align-items:baseline;gap:5px;flex-wrap:wrap;}}
.kpi-count-num{{font-size:24px;font-weight:700;color:#1A1A2E;line-height:1;}}
.kpi-count-unit{{font-size:11px;color:#888;}}
.kpi-count-pct{{font-size:10px;font-weight:600;color:white;background:#888;border-radius:3px;padding:1px 5px;margin-left:2px;}}
.kpi-card.verde .kpi-count-pct{{background:var(--verde);}}
.kpi-card.teal  .kpi-count-pct{{background:#0097A7;}}
.kpi-card.roxo  .kpi-count-pct{{background:var(--roxo);}}
.kpi-card.indigo .kpi-count-pct{{background:#3949AB;}}

/* CHART GRID */
.row{{display:grid;gap:12px;margin-bottom:12px;}}
.row2{{grid-template-columns:1fr 1fr;}}
.row3{{grid-template-columns:1.1fr 2fr 1fr;}}
.row-full{{grid-template-columns:2fr 1fr;}}
.chart-card{{
  background:white;border-radius:8px;padding:14px 12px 6px;
  box-shadow:0 2px 6px rgba(0,0,0,.07);
  transition:box-shadow .15s;
}}
.chart-card:hover{{box-shadow:0 4px 14px rgba(0,0,0,.12);}}
.chart-title{{font-size:12px;font-weight:700;color:#1A1A2E;margin-bottom:6px;
  padding-bottom:6px;border-bottom:2px solid var(--azul);display:flex;
  align-items:center;gap:6px;}}
.chart-title .icon{{font-size:14px;}}

/* TABELA */
.table-wrap{{overflow:auto;max-height:280px;border-radius:6px;}}
.table-wrap::-webkit-scrollbar{{height:4px;width:4px;}}
.table-wrap::-webkit-scrollbar-thumb{{background:#ddd;border-radius:2px;}}
table{{width:100%;border-collapse:collapse;font-size:11px;}}
thead th{{
  background:var(--azul);color:white;padding:7px 8px;
  position:sticky;top:0;font-weight:600;text-align:left;white-space:nowrap;
}}
tbody tr:nth-child(even){{background:#F3F8FE;}}
tbody tr:hover{{background:#DCE9F8;}}
tbody td{{padding:6px 8px;border-bottom:1px solid #F0F0F0;color:#333;white-space:nowrap;}}

/* ATIVO */
.active-filters{{
  background:#EBF3FC;border:1px solid #BDD6F0;border-radius:6px;
  padding:8px 12px;font-size:11px;color:#0078D4;margin-bottom:12px;
  display:none;flex-wrap:wrap;gap:6px;align-items:center;
}}
.active-filters.show{{display:flex;}}
.filter-chip{{
  background:#0078D4;color:white;border-radius:20px;
  padding:2px 10px 2px 8px;font-size:10px;display:flex;align-items:center;gap:4px;cursor:pointer;
}}
.filter-chip:hover{{background:#106EBE;}}
.filter-chip::after{{content:'×';font-size:13px;margin-left:2px;}}

/* Tooltip de seleção */
.sel-hint{{font-size:10px;color:#999;font-style:italic;margin-top:4px;}}

/* ── RESPONSIVO ──────────────────────────────────────────────────────────── */
.menu-toggle{{
  display:none;background:rgba(255,255,255,.15);border:none;
  color:white;font-size:20px;cursor:pointer;padding:6px 10px;
  border-radius:6px;line-height:1;flex-shrink:0;
}}
.sidebar-overlay{{
  display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:50;
}}
.sidebar-overlay.show{{display:block;}}
.sidebar-topbar{{
  display:flex;align-items:center;justify-content:space-between;
  padding-bottom:10px;margin-bottom:12px;
  border-bottom:1px solid #333;
  font-size:12px;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:1px;
}}
.sidebar-close{{
  display:none;background:rgba(255,255,255,.15);border:none;
  color:white;font-size:16px;cursor:pointer;padding:4px 9px;
  border-radius:5px;line-height:1;font-weight:700;
}}
.sidebar-close:hover{{background:rgba(255,255,255,.3);}}

/* Tablet (≤1100px) */
@media(max-width:1100px){{
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:8px;}}
  .row3{{grid-template-columns:1fr 1fr;}}
  .row3 .chart-card:last-child{{grid-column:1/-1;}}
  .insight-grid{{grid-template-columns:repeat(2,1fr);}}
  .tipo-grid{{grid-template-columns:repeat(3,1fr);}}
  .plano-grid{{grid-template-columns:1fr 1fr;}}
  .sidebar{{width:200px;min-width:200px;}}
}}

/* ── MOBILE (≤768px) ──────────────────────────────────────────────────────── */
@media(max-width:768px){{
  body{{font-size:13px;}}

  /* Header row 1: título resumido + pesquisa + sair */
  .header-row1{{padding:7px 10px;gap:8px;flex-wrap:nowrap;}}
  .header h1{{font-size:0;white-space:nowrap;overflow:hidden;flex:0 0 auto;}}
  .header h1::before{{content:'🛡️ GMBC';font-size:13px;font-weight:700;}}
  .pesquisa-header{{flex:1;min-width:0;}}
  .pesquisa-header input{{width:100%;min-width:0;box-sizing:border-box;}}
  .btn-sair{{flex-shrink:0;padding:5px 10px;font-size:11px;}}

  /* Botões linha 1: somente ícones (Limpar e Sair) */
  .btn-reset .btxt,.btn-sair .btxt{{display:none !important;}}
  .btn-pdf-r1{{display:block;}}
  .btn-pdf-r2{{display:none;}}

  /* Header row 2: grid fixo de 4 colunas (garante 4 por linha independente
     do tamanho do texto do botão — flexbox deixava o texto forçar só 2) */
  .header-row2{{
    display:grid;grid-template-columns:repeat(4,1fr);gap:4px;
    padding:3px 6px 4px;
  }}
  .header-row2 button{{padding:5px 1px;font-size:8.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:center;}}

  .badge{{display:none;}}
  .menu-toggle{{display:block;flex-shrink:0;}}
  .sidebar-close{{display:block;}}

  /* Layout principal */
  .main-layout{{flex-direction:column;}}
  .sidebar{{
    position:fixed;top:0;left:-260px;width:260px;min-width:260px;
    height:100%;z-index:200;transition:left .25s ease;
    padding-top:0;box-shadow:4px 0 24px rgba(0,0,0,.35);
  }}
  .sidebar.open{{left:0;}}
  .sidebar-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:199;}}
  .sidebar-overlay.open{{display:block;}}
  .content{{padding:10px;overflow-y:auto;width:100%;}}

  /* KPIs — 2 colunas */
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:8px;}}
  .kpi-val{{font-size:26px;}}
  .kpi-label{{font-size:10px;}}

  /* Grids → coluna única */
  .row,.row2,.row3,.row-full,
  .insight-grid,.rec-grid,
  .plano-grid,.plano-grid-wide{{
    grid-template-columns:1fr !important;
    gap:10px !important;
  }}
  .row3 .chart-card:last-child{{grid-column:auto;}}

  /* Tipificações — 2 colunas */
  .tipo-grid{{grid-template-columns:repeat(2,1fr);gap:6px;}}

  /* Altura dos gráficos */
  .chart-card{{padding:10px;}}
  #chart-bairro,#chart-item,#chart-ruas,#chart-refs{{height:280px !important;}}
  #chart-heatmap,#chart-hora{{height:200px !important;}}
  #chart-linha,#chart-comp{{height:220px !important;}}
  #mapa-crime{{height:300px !important;}}

  /* Tabela */
  .table-wrap{{font-size:10px;overflow-x:auto;}}
  thead th{{padding:5px 6px;font-size:9px;white-space:nowrap;}}
  tbody td{{padding:4px 6px;font-size:10px;}}

  /* Mapa */
  .mapa-legenda{{flex-wrap:wrap;gap:6px;}}
  .mapa-header{{flex-direction:column;align-items:flex-start;gap:8px;}}

  /* Sidebar filtros */
  .filter-scroll{{max-height:140px;}}
  .section-header{{font-size:12px;}}
  .turno-label{{min-width:56px;}}

  /* Modais */
  .analise-box{{width:97vw !important;max-height:90vh;}}
  .analise-header h2{{font-size:14px;}}
  .analise-corpo{{padding:12px !important;}}
}}

/* ── TABLET / NOTEBOOK MÉDIO (769px–1024px) ──────────────────────────────── */
@media(min-width:769px) and (max-width:1024px){{
  .sidebar{{width:190px;min-width:190px;}}
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:8px;}}
  .kpi-val{{font-size:26px;}}
  .row3{{grid-template-columns:1fr 1fr;}}
  .row3 .chart-card:last-child{{grid-column:1/-1;}}
  .insight-grid{{grid-template-columns:repeat(2,1fr);}}
  .plano-grid{{grid-template-columns:1fr 1fr;}}
  .rec-grid{{grid-template-columns:1fr 1fr;}}
  .pesquisa-header input{{width:160px;}}
}}

/* ── TELAS MUITO PEQUENAS (≤480px) ──────────────────────────────────────── */
@media(max-width:480px){{
  .kpi-row{{gap:6px;}}
  .kpi-val{{font-size:22px;}}
  .kpi-label{{font-size:9px;}}
  .content{{padding:8px;}}
  .rec-text{{font-size:11.5px;}}
  .chart-card{{padding:8px;}}
  .section-title{{font-size:12px;}}
  .header-row2 button{{font-size:7.5px;padding:5px 1px;}}
}}

/* INSIGHTS */
.insight-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px;}}
.insight-card{{
  background:white;border-radius:8px;padding:14px 16px;
  box-shadow:0 2px 6px rgba(0,0,0,.07);border-left:4px solid var(--azul);
}}
.insight-card.alerta{{border-color:var(--vermelho);}}
.insight-card.atencao{{border-color:var(--laranja);}}
.insight-card.info{{border-color:var(--roxo);}}
.insight-card.positivo{{border-color:var(--verde);}}
.insight-title{{font-size:11px;font-weight:700;color:#1A1A2E;text-transform:uppercase;
  letter-spacing:.4px;margin-bottom:8px;display:flex;align-items:center;gap:6px;}}
.insight-text{{font-size:12px;color:#444;line-height:1.6;}}
.insight-highlight{{font-weight:700;color:var(--azul);}}
.insight-card.alerta .insight-highlight{{color:var(--vermelho);}}
.insight-card.atencao .insight-highlight{{color:var(--laranja);}}
.insight-card.info .insight-highlight{{color:var(--roxo);}}
.insight-card.positivo .insight-highlight{{color:var(--verde);}}

/* TIPIFICAÇÕES */
.tipo-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:12px;}}
.tipo-card{{
  background:white;border-radius:8px;padding:12px 14px;
  box-shadow:0 2px 6px rgba(0,0,0,.07);text-align:center;
}}
.tipo-icon{{font-size:26px;margin-bottom:6px;}}
.tipo-nome{{font-size:11px;font-weight:700;color:#1A1A2E;margin-bottom:4px;}}
.tipo-desc{{font-size:10px;color:#666;line-height:1.5;}}
.tipo-badge{{
  display:inline-block;border-radius:4px;padding:2px 8px;
  font-size:10px;font-weight:700;color:white;margin-top:6px;
}}

/* RECOMENDAÇÕES */
.rec-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-bottom:12px;}}
.rec-card{{
  background:white;border-radius:10px;
  box-shadow:0 2px 8px rgba(0,0,0,.08);
  overflow:hidden;
}}
.rec-header{{
  display:flex;align-items:center;gap:12px;
  padding:14px 18px 12px;
  background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);
}}
.rec-icon{{font-size:24px;}}
.rec-title{{font-size:12px;font-weight:700;color:white;text-transform:uppercase;letter-spacing:.5px;}}
.rec-body{{padding:0 16px 8px;}}
.rec-item{{
  display:flex;align-items:flex-start;gap:10px;
  padding:9px 0;border-bottom:1px solid #F0F0F0;
}}
.rec-item:last-child{{border-bottom:none;}}
.rec-prio{{
  flex-shrink:0;font-size:9px;font-weight:700;
  border-radius:4px;padding:3px 7px;margin-top:1px;
  text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;
}}
.rec-prio.alta{{background:#FDDCDC;color:#A4262C;border:1px solid #F1BBBC;}}
.rec-prio.media{{background:#FFF4CE;color:#7A5100;border:1px solid #F0D070;}}
.rec-prio.baixa{{background:#DFF6DD;color:#107C10;border:1px solid #9FD89F;}}
.rec-text{{font-size:12.5px;color:#333;line-height:1.55;}}

.section-header{{
  font-size:13px;font-weight:700;color:#1A1A2E;
  margin:18px 0 10px;padding-bottom:6px;
  border-bottom:3px solid var(--azul);
  display:flex;align-items:center;gap:8px;
}}

/* PLANO POLICIAL */
.plano-grid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px;}}
.plano-grid-wide{{display:grid;grid-template-columns:2fr 1fr;gap:12px;margin-bottom:12px;}}
.plano-card{{background:white;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08);}}
.plano-head{{padding:11px 16px;font-size:11px;font-weight:700;color:white;
  text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:8px;}}
.plano-body{{padding:10px 14px 12px;}}
.turno-row{{display:flex;align-items:stretch;gap:0;margin-bottom:8px;border-radius:8px;
  overflow:hidden;border:1px solid #E8E8E8;}}
.turno-label{{
  display:flex;flex-direction:column;justify-content:center;align-items:center;
  min-width:72px;padding:8px 6px;font-size:10px;font-weight:700;color:white;text-align:center;
}}
.turno-label .th{{font-size:13px;margin-bottom:2px;}}
.turno-label .tn{{font-size:9px;opacity:.9;}}
.turno-detalhes{{flex:1;padding:8px 12px;}}
.turno-bairro{{font-size:11px;font-weight:600;color:#1A1A2E;margin-bottom:3px;}}
.turno-info{{font-size:10.5px;color:#555;line-height:1.5;}}
.turno-badge{{
  display:inline-block;font-size:9px;font-weight:700;border-radius:3px;
  padding:2px 6px;margin-right:4px;margin-top:3px;
}}
.tb-efetivo{{background:#EBF3FC;color:#0078D4;}}
.tb-foco{{background:#FFF4CE;color:#7A5100;}}
.tb-alerta{{background:#FDDCDC;color:#A4262C;}}
.posto-item{{
  display:flex;align-items:flex-start;gap:10px;
  padding:8px 0;border-bottom:1px solid #F0F0F0;
}}
.posto-item:last-child{{border-bottom:none;}}
.posto-num{{
  background:#1A1A2E;color:white;border-radius:50%;
  width:22px;height:22px;display:flex;align-items:center;
  justify-content:center;font-size:10px;font-weight:700;flex-shrink:0;margin-top:1px;
}}
.posto-info{{flex:1;}}
.posto-local{{font-size:12px;font-weight:600;color:#1A1A2E;}}
.posto-detalhe{{font-size:10.5px;color:#666;margin-top:2px;line-height:1.4;}}
.risco-bar{{height:6px;border-radius:3px;margin-top:4px;background:#EEE;overflow:hidden;}}
.risco-fill{{height:100%;border-radius:3px;}}
.alerta-box{{
  background:#FFF4F4;border:1px solid #F1BBBC;border-radius:8px;
  padding:10px 14px;margin-bottom:8px;font-size:12px;color:#A4262C;
  display:flex;gap:8px;align-items:flex-start;line-height:1.5;
}}
.ok-box{{
  background:#F0FFF0;border:1px solid #9FD89F;border-radius:8px;
  padding:10px 14px;margin-bottom:8px;font-size:12px;color:#107C10;
  display:flex;gap:8px;align-items:flex-start;line-height:1.5;
}}

/* MAPA */
#mapa-crime{{height:460px;width:100%;border-radius:0 0 8px 8px;z-index:1;}}
.mapa-card{{background:white;border-radius:10px;overflow:hidden;
  box-shadow:0 2px 8px rgba(0,0,0,.08);margin-bottom:12px;}}
.mapa-header{{
  background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);
  padding:12px 18px;display:flex;align-items:center;justify-content:space-between;
}}
.mapa-header-left{{display:flex;align-items:center;gap:10px;}}
.mapa-title{{font-size:12px;font-weight:700;color:white;text-transform:uppercase;letter-spacing:.5px;}}
.mapa-legenda{{display:flex;gap:12px;flex-wrap:wrap;}}
.mapa-leg-item{{display:flex;align-items:center;gap:5px;font-size:10px;color:white;}}
.mapa-leg-dot{{width:10px;height:10px;border-radius:50%;flex-shrink:0;border:2px solid rgba(255,255,255,.5);}}
.mapa-stat{{font-size:11px;color:rgba(255,255,255,.8);}}
/* ── COMPARAÇÃO ANUAL ── */
.comp-ano-btn{{
  border:none;border-radius:20px;padding:5px 16px;font-size:11px;font-weight:700;
  cursor:pointer;font-family:inherit;transition:all .2s;
}}
.comp-ano-btn.ativo{{opacity:1;color:white;box-shadow:0 2px 8px rgba(0,0,0,.22);}}
.comp-kpi-delta{{
  background:white;border-radius:8px;padding:12px 14px;
  box-shadow:0 2px 6px rgba(0,0,0,.07);text-align:center;min-width:0;
}}
.comp-delta-val{{font-size:24px;font-weight:800;margin:4px 0;line-height:1;}}
.comp-delta-label{{font-size:10px;color:#666;margin-top:3px;line-height:1.4;font-weight:600;}}
.delta-up{{color:#D13438;}}
.delta-down{{color:#107C10;}}
.delta-eq{{color:#888;}}
.leaflet-popup-content{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;min-width:200px;}}
.popup-tipo{{display:inline-block;border-radius:4px;padding:2px 8px;
  font-size:10px;font-weight:700;color:white;margin-bottom:6px;}}
.popup-row{{display:flex;gap:6px;margin-top:4px;color:#333;font-size:11px;}}
.popup-label{{color:#888;font-size:10px;min-width:50px;}}
</style>
</head>
<body>

<div id="__err__" style="display:none;position:fixed;top:0;left:0;right:0;background:#c00;color:#fff;padding:16px;font-family:monospace;font-size:13px;z-index:99999;white-space:pre-wrap"></div>
<script>window.onerror=function(m,s,l,c,e){{var d=document.getElementById('__err__');if(d){{d.style.display='block';d.textContent='ERRO JS: '+m+' (linha '+l+')\\n'+(e&&e.stack?e.stack:'');}}return false;}};</script>

<script>
/* ── LOGIN STANDALONE — funciona mesmo se o script principal falhar ── */
(function(){{
  if(sessionStorage.getItem('gmbc_auth')==='1'){{
    document.addEventListener('DOMContentLoaded',function(){{
      var o=document.getElementById('login-overlay');
      if(o)o.style.display='none';
    }});
  }}
}})();
function toggleSenha(){{
  var i=document.getElementById('login-senha');
  var b=document.querySelector('.login-eye');
  if(!i)return;
  i.type=i.type==='password'?'text':'password';
  if(b)b.textContent=i.type==='password'?'👁':'🙈';
}}
function verificarSenha(){{
  var i=document.getElementById('login-senha');
  var e=document.getElementById('login-erro');
  var c=document.getElementById('login-card');
  var v=i?i.value:'';
  if(!v){{if(e)e.textContent='Digite a senha.';return;}}
  try{{
    var enc=btoa(unescape(encodeURIComponent(v)));
    if(enc==='{_senha_b64}'){{
      sessionStorage.setItem('gmbc_auth','1');
      var o=document.getElementById('login-overlay');
      if(o)o.style.display='none';
    }}else{{
      if(e)e.textContent='❌ Senha incorreta. Tente novamente.';
      if(c){{c.classList.remove('login-shake');void c.offsetWidth;c.classList.add('login-shake');}}
      if(i){{i.value='';i.focus();}}
    }}
  }}catch(ex){{if(e)e.textContent='⚠️ Erro. Tente novamente.';}}
}}
function sair(){{
  sessionStorage.removeItem('gmbc_auth');
  var i=document.getElementById('login-senha');
  var e=document.getElementById('login-erro');
  var o=document.getElementById('login-overlay');
  if(i)i.value='';if(e)e.textContent='';
  if(o)o.style.display='flex';
}}
</script>

<div class="sidebar-overlay" id="sidebar-overlay" onclick="toggleSidebar()"></div>

<div class="header">
  <!-- Linha 1: título + pesquisa + sair -->
  <div class="header-row1">
    <button class="menu-toggle" onclick="toggleSidebar()" title="Filtros">☰</button>
    <h1>🛡️ Secretaria de Segurança e Ordem Pública de Balneário Camboriú — GMBC</h1>
    <div class="pesquisa-header">
      <input id="pesquisa-input" type="text" placeholder="🔎 Pesquisar..." autocomplete="off"
        oninput="onPesquisaInput(this.value)"
        onkeydown="if(event.key==='Enter')executarPesquisa()">
      <button id="pesquisa-limpar" onclick="limparPesquisa()" title="Limpar"
        style="display:none;background:rgba(255,255,255,.2);color:white;border:none;border-radius:5px;padding:4px 8px;font-size:11px;cursor:pointer">✕</button>
    </div>
    <button class="btn-reset" onclick="resetFilters()">⟳<span class="btxt"> Limpar</span></button>
    <button class="btn-pdf btn-pdf-r1" onclick="window.print()">🖨️</button>
    <button class="btn-sair" onclick="sair()">🔒<span class="btxt"> Sair</span></button>
  </div>
  <!-- Linha 2: botões de ação -->
  <div class="header-row2">
    <button class="btn-pdf btn-pdf-r2" onclick="window.print()">🖨️<span class="btxt"> PDF</span></button>
    <button class="btn-analise" onclick="analiseDiaria()">📋<span class="btxt"> Análise</span></button>
    <button class="btn-predit" onclick="abrirAnalisePredit()">🔮<span class="btxt"> Preditiva</span></button>
    <button class="btn-intelig" onclick="abrirInteligencia()">🔍<span class="btxt"> Inteligência</span></button>
    <button class="btn-prev" onclick="previsao()">📈<span class="btxt"> Previsão</span></button>
    <button class="btn-relatorio" onclick="relatorioDiario()">📅<span class="btxt"> Relatório</span></button>
    <button class="btn-resumoia" onclick="abrirResumoIA()">🤖<span class="btxt"> Resumo IA</span></button>
    <button id="btn-gerar-slides" class="btn-slides" onclick="gerarSlides()">📊<span class="btxt"> Gerar Slides</span></button>
  </div>
</div>

<div class="main-layout">

  <!-- SIDEBAR -->
  <div class="sidebar" id="sidebar">
    <div class="sidebar-topbar">
      <span>🔎 Filtros</span>
      <button class="sidebar-close" onclick="toggleSidebar()" title="Fechar filtros">✕</button>
    </div>

    <div id="pesquisa-badge" style="display:none;padding:5px 10px;background:#E8F4FD;border-bottom:1px solid #BDE0F7;font-size:10px;color:#1A3A5C;font-weight:600;text-align:center"></div>

    <div class="filter-group">
      <span class="filter-label">Ano</span>
      <div id="filter-ano"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Mês</span>
      <div id="filter-mes"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Turno</span>
      <div id="filter-turno"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Tipificação</span>
      <div id="filter-tipo"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Bairro</span>
      <input class="filter-search" type="text" placeholder="Buscar bairro…" oninput="filterSearch('bairro',this.value)">
      <div class="filter-scroll" id="filter-bairro"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Item</span>
      <input class="filter-search" type="text" placeholder="Buscar item…" oninput="filterSearch('item',this.value)">
      <div class="filter-scroll" id="filter-item"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Dia da Semana</span>
      <div id="filter-dia"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">Logradouro</span>
      <input class="filter-search" type="text" placeholder="Buscar rua…" oninput="filterSearch('logradouro',this.value)">
      <div class="filter-scroll" id="filter-logradouro"></div>
    </div>

    <div class="filter-group">
      <span class="filter-label">IMEI</span>
      <input id="filter-imei-input" class="filter-search" type="text" placeholder="Buscar IMEI..." oninput="filterImei(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Marca / Modelo</span>
      <input id="filter-marca-input" class="filter-search" type="text" placeholder="Buscar marca/modelo..." oninput="filterMarca(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Placa</span>
      <input id="filter-placa-input" class="filter-search" type="text" placeholder="Buscar placa..." oninput="filterPlaca(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Número de Série</span>
      <input id="filter-nserie-input" class="filter-search" type="text" placeholder="Buscar número de série..." oninput="filterNumeroSerie(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Cor</span>
      <input id="filter-cor-input" class="filter-search" type="text" placeholder="Buscar cor..." oninput="filterCor(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Detalhes</span>
      <input id="filter-detalhes-input" class="filter-search" type="text" placeholder="Buscar detalhes..." oninput="filterDetalhes(this.value)">
    </div>

    <div class="filter-group">
      <span class="filter-label">Recuperado</span>
      <div class="filter-scroll" id="filter-recuperado"></div>
    </div>
  </div>

  <!-- CONTENT -->
  <div class="content">

    <!-- Filtros ativos -->
    <div class="active-filters" id="active-filters"></div>

    <!-- KPIs Linha 1: Ocorrências -->
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="kpi-val" id="kpi-total">–</div>
        <div class="kpi-label">Total de Ocorrências</div>
        <div class="kpi-sub" id="kpi-total-sub">2026</div>
        <div class="kpi-media" id="kpi-total-media"></div>
      </div>
      <div class="kpi-card laranja">
        <div class="kpi-val" id="kpi-furtos">–</div>
        <div class="kpi-label">Furtos</div>
        <div class="kpi-sub" id="kpi-furtos-sub"></div>
        <div class="kpi-media" id="kpi-furtos-media"></div>
      </div>
      <div class="kpi-card vermelho">
        <div class="kpi-val" id="kpi-roubos">–</div>
        <div class="kpi-label">Roubos</div>
        <div class="kpi-sub" id="kpi-roubos-sub"></div>
        <div class="kpi-media" id="kpi-roubos-media"></div>
      </div>
      <div class="kpi-card marrom">
        <div class="kpi-val" id="kpi-arrom">–</div>
        <div class="kpi-label">Arrombamentos</div>
        <div class="kpi-sub" id="kpi-arrom-sub"></div>
        <div class="kpi-media" id="kpi-arrom-media"></div>
      </div>
    </div>
    <!-- KPIs Linha 2: Análise -->
    <div class="kpi-row">
      <div class="kpi-card verde">
        <div class="kpi-val" id="kpi-bairro">–</div>
        <div class="kpi-label">Bairro Mais Afetado</div>
        <div class="kpi-count">
          <span class="kpi-count-num" id="kpi-bairro-num">–</span>
          <span class="kpi-count-unit">casos</span>
          <span class="kpi-count-pct" id="kpi-bairro-pct"></span>
        </div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-val" id="kpi-rua" style="font-size:15px;padding-top:4px;line-height:1.3">–</div>
        <div class="kpi-label">Rua Mais Crítica</div>
        <div class="kpi-count">
          <span class="kpi-count-num" id="kpi-rua-num">–</span>
          <span class="kpi-count-unit">casos</span>
          <span class="kpi-count-pct" id="kpi-rua-pct"></span>
        </div>
      </div>
      <div class="kpi-card roxo">
        <div class="kpi-val" id="kpi-turno">–</div>
        <div class="kpi-label">Turno Mais Crítico</div>
        <div class="kpi-count">
          <span class="kpi-count-num" id="kpi-turno-num">–</span>
          <span class="kpi-count-unit">casos</span>
          <span class="kpi-count-pct" id="kpi-turno-pct"></span>
        </div>
      </div>
      <div class="kpi-card indigo">
        <div class="kpi-val" id="kpi-dia">–</div>
        <div class="kpi-label">Dia Mais Crítico</div>
        <div class="kpi-count">
          <span class="kpi-count-num" id="kpi-dia-num">–</span>
          <span class="kpi-count-unit">casos</span>
          <span class="kpi-count-pct" id="kpi-dia-pct"></span>
        </div>
      </div>
    </div>

    <!-- Linha 1: Tipificação | Bairros | Turno -->
    <div class="row row3">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🥧</span> Tipificação</div>
        <div id="chart-tipo" style="height:220px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📍</span> 10 Bairros com Mais Ocorrências</div>
        <div id="chart-bairro" style="height:220px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🕐</span> Turno</div>
        <div id="chart-turno" style="height:220px"></div>
      </div>
    </div>

    <!-- Linha 2: Série temporal | Por Mês -->
    <div class="row row-full">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📈</span> Ocorrências por Dia</div>
        <div id="chart-linha" style="height:190px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📅</span> Por Mês</div>
        <div id="chart-mes" style="height:190px"></div>
      </div>
    </div>

    <!-- Linha 3: Itens | Dia da semana -->
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📦</span> 10 Itens Mais Furtados / Roubados</div>
        <div id="chart-item" style="height:220px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📆</span> Dia da Semana</div>
        <div id="chart-dia" style="height:220px"></div>
      </div>
    </div>

    <!-- Linha 4: Ruas | Pontos de referência -->
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🛣️</span> 10 Ruas com Mais Ocorrências</div>
        <div id="chart-ruas" style="height:340px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📌</span> 10 Pontos de Referência</div>
        <div id="chart-refs" style="height:340px"></div>
      </div>
    </div>

    <!-- Linha 5: Heatmap Dia x Turno | Distribuição por Hora -->
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🌡️</span> Mapa de Calor – Dia da Semana × Turno</div>
        <div id="chart-heatmap" style="height:260px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🕐</span> Distribuição por Hora do Dia</div>
        <div id="chart-hora" style="height:260px"></div>
      </div>
    </div>

    <!-- COMPARAÇÃO ANUAL -->
    <div class="section-header">📊 Comparação Anual
      <span id="comp-anos-label" style="font-size:10px;font-weight:400;color:#888;margin-left:8px"></span>
      <span style="font-size:10px;font-weight:400;color:#aaa;margin-left:4px">(atualiza com os filtros)</span>
    </div>
    <div id="comp-anos-selector" style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center;
      padding:10px 14px;background:#f5f7fa;border-radius:8px;border:1px solid #e0e6f0">
      <span style="font-size:11px;color:#555;font-weight:700;align-self:center">Selecione os anos:</span>
    </div>
    <div class="kpi-row" id="comp-kpi-row" style="display:none;margin-bottom:12px"></div>
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📅</span> Total por Mês — Comparativo Anual</div>
        <div id="comp-chart-mes" style="height:240px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📊</span> Por Tipificação — Comparativo Anual</div>
        <div id="comp-chart-tipo" style="height:240px"></div>
      </div>
    </div>
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📍</span> Por Bairro — Top 10 Comparativo</div>
        <div id="comp-chart-bairro" style="height:300px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🕐</span> Por Turno — Comparativo Anual</div>
        <div id="comp-chart-turno" style="height:300px"></div>
      </div>
    </div>
    <div class="row row2">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📆</span> Dia da Semana — Comparativo Anual</div>
        <div id="comp-chart-dia" style="height:240px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🛣️</span> Top 10 Ruas — Comparativo Anual</div>
        <div id="comp-chart-ruas" style="height:360px"></div>
      </div>
    </div>

    <!-- ANALISTA DE SEGURANÇA -->
    <div class="section-header">👮 Plano de Distribuição Policial <span style="font-size:10px;font-weight:400;color:#888">(gerado automaticamente com base nos dados)</span></div>
    <div id="plano-policial"></div>

    <!-- MAPA -->
    <div class="section-header">🗺️ Mapa de Ocorrências <span style="font-size:10px;font-weight:400;color:#888">(atualiza com os filtros)</span></div>
    <div class="mapa-card">
      <div class="mapa-header">
        <div class="mapa-header-left">
          <span style="font-size:20px">📍</span>
          <div>
            <div class="mapa-title">Distribuição Geográfica – Balneário Camboriú</div>
            <div class="mapa-stat" id="mapa-stat">Clique em um marcador para ver detalhes</div>
          </div>
        </div>
        <div class="mapa-legenda">
          <div class="mapa-leg-item"><div class="mapa-leg-dot" style="background:#0078D4"></div>Furto</div>
          <div class="mapa-leg-item"><div class="mapa-leg-dot" style="background:#50B2FF"></div>Tent. Furto</div>
          <div class="mapa-leg-item"><div class="mapa-leg-dot" style="background:#E07B00"></div>Arrombamento</div>
          <div class="mapa-leg-item"><div class="mapa-leg-dot" style="background:#D13438"></div>Roubo</div>
          <div class="mapa-leg-item"><div class="mapa-leg-dot" style="background:#8764B8"></div>Tent. Roubo</div>
        </div>
      </div>
      <div id="mapa-crime"></div>
    </div>

    <!-- INSIGHTS DINÂMICOS -->
    <div class="section-header">🔍 Análise Automática <span style="font-size:10px;font-weight:400;color:#888">(atualiza com os filtros)</span></div>
    <div class="insight-grid" id="insights-container"></div>

    <!-- RECOMENDAÇÕES -->
    <div class="section-header">💡 Análise e Recomendações do Analista de Segurança
      <span id="diag-header-sub" style="font-size:10px;font-weight:400;color:#888;margin-left:8px"></span>
    </div>

    <!-- Resumo executivo dinâmico -->
    <div id="diagnostico-box" style="background:#FFF4F4;border:1px solid #F1BBBC;border-radius:8px;padding:12px 16px;margin-bottom:14px;font-size:12.5px;color:#333;line-height:1.7"></div>

    <div class="rec-grid" id="rec-container"></div>

    <!-- Tabela -->
    <div class="chart-card" style="margin-bottom:8px">
      <div class="chart-title" style="margin-bottom:8px;justify-content:space-between">
        <div style="display:flex;align-items:center;gap:6px">
          <span class="icon">📋</span> Registros Detalhados
          <span style="font-size:10px;color:#888;font-weight:400;margin-left:4px" id="tabela-count"></span>
        </div>
        <button onclick="exportarTabelaPDF()" style="
          background:#107C10;color:white;border:none;border-radius:4px;
          padding:4px 12px;font-size:11px;font-weight:600;cursor:pointer;
          font-family:inherit;white-space:nowrap;flex-shrink:0">
          🖨️ Exportar PDF
        </button>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Data</th><th>Mês</th><th>Ano</th><th>Hora</th><th>Turno</th><th>Dia</th>
              <th>B.O.</th><th>Tipificação</th><th>Item</th><th>Descrição</th>
              <th>Marca/Modelo</th><th>Endereço</th><th>Bairro</th><th>Referência</th>
              <th>IMEI</th><th>Placa</th><th>Número de Série</th><th>Recuperado</th>
            </tr>
          </thead>
          <tbody id="tabela-body"></tbody>
        </table>
      </div>
      <div class="sel-hint">Clique em qualquer gráfico para filtrar os dados</div>
    </div>

  </div>
</div>

<script type="application/json" id="raw-data">{data_json}</script>
<script>
const LOGO_GMBC = "{LOGO_GMBC_B64}";
// ── DADOS ────────────────────────────────────────────────────────────────────
var RAW = [];
try {{
  RAW = JSON.parse(document.getElementById('raw-data').textContent);
}} catch(e) {{
  var _ed = document.getElementById('__err__');
  if(_ed){{ _ed.style.display='block'; _ed.textContent='ERRO AO CARREGAR DADOS: '+e.message; }}
}}

// ── ESTADO DOS FILTROS ────────────────────────────────────────────────────────
const state = {{
  ano: new Set(),
  mes: new Set(), turno: new Set(), tipo: new Set(),
  bairro: new Set(), item: new Set(), dia: new Set(), logradouro: new Set(),
  recuperado: new Set()
}};
let imeiQ = '', marcaQ = '', placaQ = '', numeroSerieQ = '', corQ = '', detalhesQ = '', pesquisaQ = '';

// ── CORES ─────────────────────────────────────────────────────────────────────
const COLORS = {{
  azul:'#0078D4', azulClr:'#333D65', laranja:'#E07B00',
  verde:'#107C10', vermelho:'#D13438', roxo:'#8764B8', amarelo:'#FFB900',
  roxo2:'#CA5010',
}};
const TIPO_COLORS = {{
  'Furto':COLORS.azul,'Tentativa de Furto':COLORS.azulClr,
  'Arrombamento':COLORS.laranja,'Roubo':COLORS.vermelho,
  'Tentativa de Roubo':COLORS.roxo,
}};
const TURNO_COLORS = {{
  'Manhã':COLORS.amarelo,'Tarde':COLORS.laranja,
  'Noite':COLORS.azul,'Madrugada':COLORS.roxo
}};
const DIA_ORDER = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo'];
const DIA_COLORS = d => ['Sábado','Domingo'].includes(d) ? COLORS.vermelho : COLORS.azul;

const LAYOUT_BASE = {{
  margin:{{l:8,r:8,t:4,b:4}},
  paper_bgcolor:'white', plot_bgcolor:'white',
  font:{{family:'Segoe UI,Arial,sans-serif',size:11,color:'#333'}},
  showlegend:false,
  hoverlabel:{{bgcolor:'#1A1A2E',font:{{color:'white',size:11}}}},
}};
const CONFIG = {{responsive:true,displayModeBar:false,locale:'pt-BR'}};

// ── FILTRAR DADOS ─────────────────────────────────────────────────────────────
function filtered() {{
  return RAW.filter(r =>
    (state.mes.size        === 0 || state.mes.has(r.mes))        &&
    (state.turno.size      === 0 || state.turno.has(r.turno))    &&
    (state.tipo.size       === 0 || state.tipo.has(r.tipo))      &&
    (state.bairro.size     === 0 || state.bairro.has(r.bairro))  &&
    (state.item.size       === 0 || state.item.has(r.item))      &&
    (imeiQ === '' || (r.imei && r.imei.toLowerCase().includes(imeiQ.toLowerCase()))) &&
    (marcaQ === '' || (r.marca && r.marca.toLowerCase().includes(marcaQ.toLowerCase()))) &&
    (corQ === '' || (r.cor && r.cor.toLowerCase().includes(corQ.toLowerCase()))) &&
    (detalhesQ === '' || (r.detalhes && r.detalhes.toLowerCase().includes(detalhesQ.toLowerCase()))) &&
    (placaQ === '' || (r.placa && r.placa.toLowerCase().includes(placaQ.toLowerCase()))) &&
    (numeroSerieQ === '' || (r.numero_serie && r.numero_serie.toLowerCase().includes(numeroSerieQ.toLowerCase()))) &&
    (pesquisaQ === '' || [r.bairro,r.tipo,r.endereco,r.item,r.marca,r.cor,r.detalhes,r.dia,r.turno,r.mes,r.bo].some(v=>v&&String(v).toLowerCase().includes(pesquisaQ.toLowerCase()))) &&
    (state.recuperado.size === 0 || state.recuperado.has(r.recuperado)) &&
    (state.dia.size        === 0 || state.dia.has(r.dia))        &&
    (state.logradouro.size === 0 || state.logradouro.has(r.endereco)) &&
    (state.ano.size        === 0 || state.ano.has(String(r.ano)))
  );
}}

// ── CONTAR ────────────────────────────────────────────────────────────────────
function count(data, key) {{
  const m = {{}};
  data.forEach(r => {{ const v=r[key]; if(v!==null&&v!==undefined&&v!=='') m[v]=(m[v]||0)+1; }});
  return m;
}}
function sortedEntries(obj, desc=true) {{
  return Object.entries(obj).sort((a,b)=>desc?b[1]-a[1]:a[1]-b[1]);
}}
// Deduplica por número de B.O.: um B.O. com 4 itens vira 1 ocorrência.
// Usado em TODOS os gráficos/KPIs exceto renderItem (que conta itens).
function dedupBO(data) {{
  if(!data || data.length === 0) return data;
  const seen = new Set();
  return data.filter(r => {{
    const k = r.bo ? r.bo.trim() : '';
    if(!k) return true;
    if(seen.has(k)) return false;
    seen.add(k); return true;
  }});
}}

// ── PLOTLY HELPERS ────────────────────────────────────────────────────────────
function barH(labels, vals, colors) {{
  return [{{
    type:'bar', orientation:'h',
    x:vals, y:labels,
    marker:{{color:colors, line:{{color:'white',width:1}}}},
    text:vals.map(String), textposition:'outside',
    hovertemplate:'<b>%{{y}}</b><br>%{{x}} casos<extra></extra>',
    cliponaxis:false,
  }}];
}}
function barV(labels, vals, colors) {{
  return [{{
    type:'bar',
    x:labels, y:vals,
    marker:{{color:colors, line:{{color:'white',width:1}}}},
    text:vals.map(String), textposition:'outside',
    hovertemplate:'<b>%{{x}}</b><br>%{{y}} casos<extra></extra>',
    cliponaxis:false,
  }}];
}}

// ── GRÁFICO: TIPIFICAÇÃO (pizza) ──────────────────────────────────────────────
function renderTipo(data) {{
  if(typeof Plotly === 'undefined') return;
  const c = count(data,'tipo');
  const labels = Object.keys(c), vals = Object.values(c);
  const trace = {{
    type:'pie', labels, values:vals,
    marker:{{colors:labels.map(l=>TIPO_COLORS[l]||COLORS.azul),
      line:{{color:'white',width:2}}}},
    textinfo:'percent', textfont:{{size:11,color:'white'}},
    hovertemplate:'<b>%{{label}}</b><br>%{{value}} casos (%{{percent}})<extra></extra>',
    hole:.35, sort:false,
  }};
  const layout = {{...LAYOUT_BASE,
    showlegend:true,
    legend:{{orientation:'v',x:1,y:.5,font:{{size:10}}}},
    margin:{{l:4,r:90,t:4,b:4}},
  }};
  Plotly.react('chart-tipo',[trace],layout,CONFIG);
  document.getElementById('chart-tipo').on('plotly_click', e => {{
    const lab = e.points[0].label;
    toggleFilter('tipo', lab);
  }});
}}

// ── GRÁFICO: BAIRROS ──────────────────────────────────────────────────────────
function renderBairro(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'bairro');
  const entries = sortedEntries(c).slice(0,10); // top 10
  const rawLabels = entries.map(e=>e[0]);
  const labels = rawLabels.map(l=>l.length>22?l.slice(0,20)+'…':l).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:10}},automargin:false}},
    margin:{{l:130,r:30,t:4,b:24}},
  }};
  Plotly.react('chart-bairro', barH(labels,vals,colors), layout, CONFIG);
  document.getElementById('chart-bairro').on('plotly_click', e => {{
    const idx = labels.indexOf(e.points[0].y);
    toggleFilter('bairro', idx>=0 ? rawLabels.slice().reverse()[idx] : e.points[0].y);
  }});
}}

// ── GRÁFICO: TURNO ────────────────────────────────────────────────────────────
function renderTurno(data) {{
  if(typeof Plotly === 'undefined') return;

  const ORDER = ['Manhã','Tarde','Noite','Madrugada'];
  const c = count(data,'turno');
  const labels = ORDER.filter(t=>c[t]);
  const vals   = labels.map(t=>c[t]||0);
  const maxTurno = Math.max(...vals);
  const colors = labels.map((t,i)=>vals[i]===maxTurno?COLORS.vermelho:TURNO_COLORS[t]||COLORS.azul);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{tickfont:{{size:10}},automargin:true}},
    yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}},range:[0,Math.max(...vals)*1.18]}},
    margin:{{l:30,r:10,t:4,b:30}},
  }};
  Plotly.react('chart-turno', barV(labels,vals,colors), layout, CONFIG);
  document.getElementById('chart-turno').on('plotly_click', e => {{
    toggleFilter('turno', e.points[0].x);
  }});
}}

// ── GRÁFICO: LINHA TEMPORAL ───────────────────────────────────────────────────
function renderLinha(data) {{
  if(typeof Plotly === 'undefined') return;

  const anos = [...new Set(data.map(r=>r.ano))].sort();

  if(anos.length <= 1) {{
    // ── Um único ano: últimos 30 dias ─────────────────────────────────────────
    const c = {{}};
    data.forEach(r => {{ if(r.data) c[r.data]=(c[r.data]||0)+1; }});
    const dates = Object.keys(c).sort().slice(-30); // apenas últimos 30 dias
    const vals  = dates.map(d=>c[d]);
    const trace = {{
      type:'scatter', mode:'lines+markers+text',
      x:dates, y:vals,
      line:{{color:COLORS.azul,width:2.5,shape:'spline'}},
      marker:{{color:COLORS.azulClr,size:7,line:{{color:COLORS.azul,width:1.5}}}},
      fill:'tozeroy', fillcolor:'rgba(0,120,212,.1)',
      text:vals.map(String), textposition:'top center', textfont:{{size:9,color:COLORS.azul}},
      hovertemplate:'<b>%{{x|%d/%m/%Y}}</b><br>%{{y}} casos<extra></extra>',
    }};
    const layout = {{...LAYOUT_BASE,
      xaxis:{{tickformat:'%d/%m',tickfont:{{size:9}},dtick:'D2',tickangle:-45,showgrid:false}},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}},range:[0,Math.max(...vals,1)*1.20]}},
      margin:{{l:28,r:12,t:4,b:52}},
    }};
    Plotly.react('chart-linha',[trace],layout,CONFIG);

  }} else {{
    // ── Múltiplos anos: últimos 30 dias do ano mais recente, mesmo período nos demais ──
    const anoRecente = anos[anos.length-1];
    const mmddRecentes = [...new Set(
      data.filter(r=>r.ano===anoRecente&&r.data).map(r=>r.data.slice(5))
    )].sort().slice(-30); // últimos 30 MM-DD do ano mais recente
    const filtroMMDD = new Set(mmddRecentes);

    const traces = anos.map((a,i) => {{
      const c = {{}};
      data.filter(r=>r.ano===a&&r.data&&filtroMMDD.has(r.data.slice(5))).forEach(r => {{
        const key = '2000-' + r.data.slice(5);
        c[key] = (c[key]||0)+1;
      }});
      const dates = Object.keys(c).sort();
      const vals  = dates.map(d=>c[d]);
      const cor   = ANO_PALETTE[i % ANO_PALETTE.length];
      return {{
        type:'scatter', mode:'lines+markers',
        name:String(a),
        x:dates, y:vals,
        line:{{color:cor,width:2.5,shape:'spline'}},
        marker:{{color:cor,size:6,line:{{color:'white',width:1.5}}}},
        fill:'none',
        hovertemplate:'<b>%{{x|%d/%m}} — '+String(a)+'</b><br>%{{y}} casos<extra></extra>',
      }};
    }});
    const allVals = traces.flatMap(t=>t.y);
    const layout = {{...LAYOUT_BASE,
      showlegend:true,
      legend:{{orientation:'h',x:0,y:1.18,font:{{size:11}}}},
      xaxis:{{
        tickformat:'%d/%m', tickfont:{{size:9}},
        dtick:'D3', tickangle:-45, showgrid:false,
        hoverformat:'%d/%m',
      }},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}},range:[0,Math.max(...allVals,1)*1.20]}},
      margin:{{l:28,r:12,t:38,b:52}},
    }};
    Plotly.react('chart-linha',traces,layout,CONFIG);
  }}
}}

// ── GRÁFICO: MÊS ─────────────────────────────────────────────────────────────
function renderMes(data) {{
  if(typeof Plotly === 'undefined') return;

  const ORDER = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const c = count(data,'mes');
  const labels = ORDER.filter(m=>c[m]);
  const vals   = labels.map(m=>c[m]||0);
  const maxMes = Math.max(...vals);
  const colBase = [COLORS.azul, COLORS.laranja, COLORS.verde, COLORS.roxo, COLORS.amarelo, COLORS.azulClr, COLORS.vermelho, COLORS.azul, COLORS.laranja, COLORS.verde, COLORS.roxo, COLORS.amarelo];
  const colors = vals.map((v,i)=>v===maxMes?COLORS.vermelho:colBase[i]||COLORS.azul);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{tickfont:{{size:11}}}},
    yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}},range:[0,Math.max(...vals)*1.18]}},
    margin:{{l:28,r:12,t:4,b:28}},
  }};
  Plotly.react('chart-mes', barV(labels,vals,colors.slice(0,labels.length)), layout, CONFIG);
  document.getElementById('chart-mes').on('plotly_click', e => {{
    toggleFilter('mes', e.points[0].x);
  }});
}}

// ── GRÁFICO: ITENS ────────────────────────────────────────────────────────────
function renderItem(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'item');
  const entries = sortedEntries(c).slice(0,10);
  const labels = entries.map(e=>e[0]).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:10}},automargin:true}},
    margin:{{l:100,r:30,t:4,b:24}},
  }};
  Plotly.react('chart-item', barH(labels,vals,colors), layout, CONFIG);
  document.getElementById('chart-item').on('plotly_click', e => {{
    toggleFilter('item', e.points[0].y);
  }});
}}

// ── GRÁFICO: DIA DA SEMANA ────────────────────────────────────────────────────
function renderDia(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'dia');
  const labels = DIA_ORDER.filter(d=>c[d]);
  const vals   = labels.map(d=>c[d]||0);
  const maxDia = Math.max(...vals);
  const colors = labels.map((_,i)=>vals[i]===maxDia?COLORS.vermelho:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{tickfont:{{size:10}},automargin:true}},
    yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}},range:[0,Math.max(...vals)*1.18]}},
    margin:{{l:30,r:12,t:4,b:40}},
  }};
  Plotly.react('chart-dia', barV(labels,vals,colors), layout, CONFIG);
  document.getElementById('chart-dia').on('plotly_click', e => {{
    toggleFilter('dia', e.points[0].x);
  }});
}}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs(data) {{
  const total   = data.length;
  const furtos  = data.filter(r=>r.tipo==='Furto').length;
  const roubos  = data.filter(r=>r.tipo==='Roubo').length;
  const arrom   = data.filter(r=>r.tipo==='Arrombamento').length;
  const turnos  = count(data,'turno');
  const bairros = count(data,'bairro');
  const dias    = count(data,'dia');
  const enderecos = count(data,'endereco');
  const topTurno  = sortedEntries(turnos)[0]||['–',0];
  const topBairro = sortedEntries(bairros)[0]||['–',0];
  const topDia    = sortedEntries(dias)[0]||['–',0];
  const topRua    = sortedEntries(enderecos).filter(e=>e[0])[0]||['–',0];
  const pFurtos = total ? ((furtos/total)*100).toFixed(1) : '0.0';
  const pRoubos = total ? ((roubos/total)*100).toFixed(1) : '0.0';
  const pArrom  = total ? ((arrom/total)*100).toFixed(1)  : '0.0';

  // Calcular média diária
  const datas = data.map(r=>r.data).filter(Boolean).sort();
  let nDias = 1, mediaLabel = '';
  if(datas.length > 0) {{
    nDias = new Set(datas).size;
    const mediaTotal  = (total/nDias).toFixed(1);
    const mediaFurtos = (furtos/nDias).toFixed(1);
    const mediaRoubos = (roubos/nDias).toFixed(1);
    const mediaArrom  = (arrom/nDias).toFixed(1);
    document.getElementById('kpi-total-media').textContent  = `📅 ${{mediaTotal}}/dia · ${{nDias}} dias`;
    document.getElementById('kpi-furtos-media').textContent = `📅 ${{mediaFurtos}}/dia · ${{nDias}} dias`;
    document.getElementById('kpi-roubos-media').textContent = `📅 ${{mediaRoubos}}/dia · ${{nDias}} dias`;
    document.getElementById('kpi-arrom-media').textContent  = `📅 ${{mediaArrom}}/dia · ${{nDias}} dias`;
  }} else {{
    document.getElementById('kpi-total-media').textContent  = '';
    document.getElementById('kpi-furtos-media').textContent = '';
    document.getElementById('kpi-roubos-media').textContent = '';
    document.getElementById('kpi-arrom-media').textContent  = '';
  }}

  const gt = dedupBO(RAW).length; // total global de ocorrências (não itens)
  const pBairro2 = gt ? ((topBairro[1]/gt)*100).toFixed(1)+'%' : '';
  const pRua2    = gt ? ((topRua[1]/gt)*100).toFixed(1)+'%'    : '';
  const pTurno2  = gt ? ((topTurno[1]/gt)*100).toFixed(1)+'%'  : '';
  const pDia2    = gt ? ((topDia[1]/gt)*100).toFixed(1)+'%'    : '';

  document.getElementById('kpi-total').textContent   = total;
  document.getElementById('kpi-furtos').textContent  = furtos;
  document.getElementById('kpi-roubos').textContent  = roubos;
  document.getElementById('kpi-arrom').textContent   = arrom;
  document.getElementById('kpi-furtos-sub').textContent  = pFurtos + '% do total';
  document.getElementById('kpi-roubos-sub').textContent  = pRoubos + '% do total';
  document.getElementById('kpi-arrom-sub').textContent   = pArrom  + '% do total';

  document.getElementById('kpi-bairro').textContent     = topBairro[0];
  document.getElementById('kpi-bairro-num').textContent  = topBairro[1];
  document.getElementById('kpi-bairro-pct').textContent  = pBairro2;
  document.getElementById('kpi-rua').textContent        = topRua[0];
  document.getElementById('kpi-rua-num').textContent     = topRua[1];
  document.getElementById('kpi-rua-pct').textContent     = pRua2;
  document.getElementById('kpi-turno').textContent      = topTurno[0];
  document.getElementById('kpi-turno-num').textContent   = topTurno[1];
  document.getElementById('kpi-turno-pct').textContent   = pTurno2;
  document.getElementById('kpi-dia').textContent        = topDia[0];
  document.getElementById('kpi-dia-num').textContent     = topDia[1];
  document.getElementById('kpi-dia-pct').textContent     = pDia2;

  _resumoExecutivoDados =
    `total de ${{total}} ocorrências no período filtrado; ` +
    `furtos: ${{furtos}} (${{pFurtos}}%); roubos: ${{roubos}} (${{pRoubos}}%); arrombamentos: ${{arrom}} (${{pArrom}}%); ` +
    `bairro mais afetado: ${{topBairro[0]}} (${{topBairro[1]}} casos, ${{pBairro2}}); ` +
    `rua mais crítica: ${{topRua[0]}} (${{topRua[1]}} casos, ${{pRua2}}); ` +
    `turno mais crítico: ${{topTurno[0]}} (${{topTurno[1]}} casos, ${{pTurno2}}); ` +
    `dia mais crítico: ${{topDia[0]}} (${{topDia[1]}} casos, ${{pDia2}}).`;
}}

// ── TABELA ────────────────────────────────────────────────────────────────────
function renderTabela(data) {{
  const tbody = document.getElementById('tabela-body');
  const show  = [...data].sort((a,b) => {{
    const da = (a.data||'') + ' ' + (a.hora||'');
    const db = (b.data||'') + ' ' + (b.hora||'');
    return db.localeCompare(da);
  }});
  document.getElementById('tabela-count').textContent = `${{data.length}} registros`;
  tbody.innerHTML = show.map(r => `
    <tr>
      <td>${{r.data ? r.data.slice(8)+'/'+r.data.slice(5,7)+'/'+r.data.slice(0,4) : ''}}</td>
      <td>${{r.mes}}</td>
      <td>${{r.ano}}</td>
      <td>${{r.hora}}</td>
      <td>${{r.turno}}</td>
      <td>${{r.dia}}</td>
      <td style="font-size:9px">${{r.link
        ? `<a href="${{r.link}}" target="_blank" rel="noopener"
             title="Abrir PDF do B.O."
             style="color:#0078D4;text-decoration:none;font-weight:600"
             onmouseover="this.style.textDecoration='underline'"
             onmouseout="this.style.textDecoration='none'"
           >${{r.bo}} <span style="font-size:8px">&#128196;</span></a>`
        : r.bo
      }}</td>
      <td><span style="background:${{TIPO_COLORS[r.tipo]||'#888'}};color:white;
        border-radius:3px;padding:1px 5px;font-size:9px">${{r.tipo}}</span></td>
      <td>${{r.item}}</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis">${{r.descricao||''}}</td>
      <td>${{r.marca}}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${{r.endereco}}</td>
      <td>${{r.bairro}}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:10px">${{r.ref}}</td>
      <td style="font-size:10px">${{r.imei}}</td>
      <td style="font-size:10px">${{r.placa}}</td>
      <td style="font-size:10px">${{r.numero_serie||''}}</td>
      <td style="font-size:10px">${{r.recuperado||''}}</td>
    </tr>`).join('');
}}

// ── EXPORTAR TABELA PDF ───────────────────────────────────────────────────────
function exportarTabelaPDF() {{
  const data = filtered();
  if (!data.length) {{
    alert('Nenhum registro para exportar.');
    return;
  }}
  // Calcula min/max das datas dos registros filtrados
  const datas = data.map(r => r.data).filter(Boolean).sort();
  const minData = datas[0] || '';
  const maxData = datas[datas.length - 1] || '';
  document.getElementById('pdf-data-ini').value = minData;
  document.getElementById('pdf-data-fim').value = maxData;
  document.getElementById('pdf-modal').style.display = 'flex';
}}

function fecharPdfModal() {{
  document.getElementById('pdf-modal').style.display = 'none';
}}

function gerarPdfTabela() {{
  const ini = document.getElementById('pdf-data-ini').value;
  const fim = document.getElementById('pdf-data-fim').value;
  if (!ini || !fim) {{ alert('Selecione as datas.'); return; }}
  if (ini > fim) {{ alert('Data inicial nao pode ser maior que a data final.'); return; }}

  const base = filtered();
  const data = base.filter(r => r.data && r.data >= ini && r.data <= fim);

  if (!data.length) {{
    alert('Nenhum registro encontrado no periodo selecionado.');
    return;
  }}

  fecharPdfModal();

  const fmtD = d => d ? d.slice(8)+'/'+d.slice(5,7)+'/'+d.slice(0,4) : '';
  const periodo = ini === fim ? fmtD(ini) : fmtD(ini)+' a '+fmtD(fim);
  const agora = new Date().toLocaleString('pt-BR');

  // Descricao dos filtros ativos
  const chips = [];
  for(const [key,set] of Object.entries(state)) {{
    set.forEach(v => chips.push(v));
  }}
  const filtroDesc = chips.length ? 'Filtros: ' + chips.join(', ') : 'Todos os registros';

  const sorted = [...data].sort((a,b) => {{
    const da = (a.data||'')+' '+(a.hora||'');
    const db = (b.data||'')+' '+(b.hora||'');
    return db.localeCompare(da);
  }});

  const headers = ['Data','Mes','Ano','Hora','Turno','Dia','B.O.','Tipificacao','Item',
    'Descricao','Marca/Modelo','Endereco','Bairro','Referencia','IMEI','Placa','Nº Série','Recuperado'];

  const TIPO_CSS = {{
    'Furto':'background:#0078D4;color:white',
    'Roubo':'background:#D13438;color:white',
    'Arrombamento':'background:#E07B00;color:white',
    'Tentativa de Furto':'background:#50B2FF;color:white',
    'Tentativa de Roubo':'background:#8764B8;color:white',
  }};

  const rowsHTML = sorted.map((r, i) => {{
    const bg = i % 2 === 0 ? '#fff' : '#f5f7fa';
    const tipoStyle = TIPO_CSS[r.tipo] || 'background:#888;color:white';
    return `<tr style="background:${{bg}}">
      <td>${{fmtD(r.data)}}</td>
      <td>${{r.mes||''}}</td>
      <td>${{r.ano||''}}</td>
      <td>${{r.hora||''}}</td>
      <td>${{r.turno||''}}</td>
      <td>${{r.dia||''}}</td>
      <td style="font-size:8.5px">${{r.link
        ? `<a href="${{r.link}}" style="color:#0078D4">${{r.bo}}</a>`
        : (r.bo||'')}}</td>
      <td><span style="border-radius:3px;padding:1px 5px;font-size:8px;${{tipoStyle}}">${{r.tipo||''}}</span></td>
      <td>${{r.item||''}}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${{r.descricao||''}}</td>
      <td>${{r.marca||''}}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${{r.endereco||''}}</td>
      <td>${{r.bairro||''}}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;font-size:8.5px">${{r.ref||''}}</td>
      <td style="font-size:8.5px">${{r.imei||''}}</td>
      <td style="font-size:8.5px">${{r.placa||''}}</td>
      <td style="font-size:8.5px">${{r.numero_serie||''}}</td>
      <td style="font-size:8.5px">${{r.recuperado||''}}</td>
    </tr>`;
  }}).join('');

  const w = window.open('', '_blank');
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
<meta charset="UTF-8">
<title>Registros GMBC - ${{periodo}}</title>
<style>
  body{{font-family:'Segoe UI',Arial,sans-serif;font-size:10px;margin:0;padding:16px;color:#1A1A2E}}
  .logo-bar{{display:flex;align-items:center;gap:12px;border-bottom:3px solid #1A1A2E;padding-bottom:10px;margin-bottom:10px}}
  .logo-bar h1{{font-size:13px;margin:0;color:#1A1A2E;line-height:1.3}}
  .logo-bar .sub{{font-size:10px;color:#555;margin-top:2px}}
  .info-bar{{display:flex;gap:16px;flex-wrap:wrap;background:#f2f4f8;border-radius:5px;
    padding:7px 12px;margin-bottom:12px;font-size:10px;color:#333}}
  .info-item strong{{color:#1A1A2E}}
  table{{width:100%;border-collapse:collapse;font-size:9px}}
  thead tr{{background:#1A1A2E}}
  thead th{{color:white;padding:5px 6px;text-align:left;font-weight:600;white-space:nowrap}}
  tbody td{{padding:4px 6px;border-bottom:1px solid #eee;vertical-align:top}}
  .footer{{margin-top:12px;font-size:8.5px;color:#999;border-top:1px solid #ddd;padding-top:6px;display:flex;justify-content:space-between}}
  @media print{{
    body{{padding:8px}}
    @page{{margin:8mm;size:A4 landscape}}
    .no-print{{display:none}}
  }}
</style>
</head><body>
<div class="logo-bar">
  <div>
    <h1>Secretaria de Seguranca e Ordem Publica de Balneario Camboriu</h1>
    <div class="sub">Guarda Municipal de Balneario Camboriu - GMBC</div>
  </div>
</div>
<div class="info-bar">
  <div class="info-item"><strong>Periodo:</strong> ${{periodo}}</div>
  <div class="info-item"><strong>Total:</strong> ${{data.length}} registro(s)</div>
  <div class="info-item"><strong>${{filtroDesc}}</strong></div>
  <div class="info-item" style="margin-left:auto;color:#888">Gerado em: ${{agora}}</div>
</div>
<table>
  <thead><tr>${{headers.map(h=>`<th>${{h}}</th>`).join('')}}</tr></thead>
  <tbody>${{rowsHTML}}</tbody>
</table>
<div class="footer">
  <span>Dashboard GMBC - Guarda Municipal de Balneario Camboriu</span>
  <span>Total: ${{data.length}} registros no periodo ${{periodo}}</span>
</div>
<script>window.onload=function(){{window.print();}}<\/script>
</body></html>`);
  w.document.close();
}}

// ── DIAGNÓSTICO DINÂMICO ─────────────────────────────────────────────────────
function renderDiagnostico(data) {{
  const total     = data.length;
  const grandTotal= RAW.length;
  const box   = document.getElementById('diagnostico-box');
  const sub   = document.getElementById('diag-header-sub');
  if(!box) return;

  if(total === 0) {{
    box.innerHTML = '<em style="color:#999">Nenhum dado para o filtro selecionado.</em>';
    if(sub) sub.textContent = '';
    return;
  }}

  // Calcular tudo dinamicamente
  const bairros  = count(data,'bairro');
  const turnos   = count(data,'turno');
  const dias     = count(data,'dia');
  const itens    = count(data,'item');
  const enderecos= count(data,'endereco');

  const topBairro = sortedEntries(bairros)[0] || ['–',0];
  const topTurno  = sortedEntries(turnos)[0]  || ['–',0];
  const topDia    = sortedEntries(dias)[0]    || ['–',0];
  const dia2      = sortedEntries(dias)[1]    || ['–',0];
  const dia3      = sortedEntries(dias)[2]    || ['–',0];
  const topItem   = sortedEntries(itens)[0]   || ['–',0];
  const item2     = sortedEntries(itens)[1]   || ['–',0];
  const item3     = sortedEntries(itens)[2]   || ['–',0];
  const topRua    = sortedEntries(enderecos).filter(e=>e[0])[0] || ['–',0];
  const rua2      = sortedEntries(enderecos).filter(e=>e[0])[1] || ['–',0];
  const rua3      = sortedEntries(enderecos).filter(e=>e[0])[2] || ['–',0];

  // Combo bairro x turno mais crítico
  const combos = {{}};
  data.forEach(r=>{{
    const k = r.bairro + ' / ' + r.turno;
    combos[k] = (combos[k]||0)+1;
  }});
  const topCombo = sortedEntries(combos)[0] || ['–',0];

  // Turno mais crítico — usa o campo turno diretamente (igual ao gráfico)
  const turnosPer = count(data,'turno');
  const topPeriodo = sortedEntries(turnosPer)[0] || ['–',0];

  const pBairro = ((topBairro[1]/grandTotal)*100).toFixed(1);
  const pTurno  = ((topTurno[1]/grandTotal)*100).toFixed(1);
  const pItem   = ((topItem[1]/grandTotal)*100).toFixed(1);

  // Meses presentes
  const mesesSet = [...new Set(data.map(r=>r.mes).filter(Boolean))].sort();
  const periodoLabel = mesesSet.length > 1 ? mesesSet.join('–') : (mesesSet[0]||'');
  const anos = [...new Set(data.map(r=>r.ano))].sort().join('/');
  if(sub) sub.textContent = `baseado nos ${{total}} B.O. registrados${{periodoLabel?' · '+periodoLabel:''}}${{anos?' · '+anos:''}}`;

  box.innerHTML = `
    <strong style="color:#A4262C">🚨 Diagnóstico crítico:</strong>
    O bairro <strong>${{topBairro[0]}}</strong> concentra <strong>${{topBairro[1]}} de ${{grandTotal}} ocorrências (${{pBairro}}%)</strong> —
    o principal foco de criminalidade no período analisado${{total<grandTotal?` (filtro ativo: ${{total}} registros)`:''}}.

    O período mais perigoso é a <strong>${{topPeriodo[0]}}</strong> com <strong>${{topPeriodo[1]}} casos</strong>, sendo
    <strong>${{topCombo[0]}} o combo mais crítico (${{topCombo[1]}} casos)</strong>.
    O dia de maior incidência é <strong>${{topDia[0]}} (${{topDia[1]}} casos)</strong>,
    seguido de ${{dia2[0]}} <strong>(${{dia2[1]}})</strong> e ${{dia3[0]}} <strong>(${{dia3[1]}})</strong>.
    O item mais furtado é <strong>${{topItem[0]}} (${{topItem[1]}} casos = ${{pItem}}%)</strong>,
    seguido de ${{item2[0]}} <strong>(${{item2[1]}})</strong> e ${{item3[0]}} <strong>(${{item3[1]}})</strong>.
    As ruas de maior risco são:
    <strong>${{topRua[0]}} (${{topRua[1]}}), ${{rua2[0]}} (${{rua2[1]}}) e ${{rua3[0]}} (${{rua3[1]}})</strong>.
  `;
}}

// ── RECOMENDAÇÕES DINÂMICAS ───────────────────────────────────────────────────
function renderRecomendacoes(data) {{
  const cont = document.getElementById('rec-container');
  if(!cont) return;
  const total = data.length;
  if(total === 0) {{
    cont.innerHTML = '<p style="color:#999;font-size:13px;padding:16px">Nenhum dado para o filtro selecionado.</p>';
    return;
  }}

  const bairroRank = sortedEntries(count(data,'bairro'));
  const itemRank   = sortedEntries(count(data,'item'));
  const ruaRank    = sortedEntries(count(data,'endereco')).filter(e=>e[0]);
  const diaRank    = sortedEntries(count(data,'dia'));
  const turnoRank  = sortedEntries(count(data,'turno'));

  const top1B = bairroRank[0]||['–',0], top2B = bairroRank[1]||['–',0], top3B = bairroRank[2]||['–',0];
  const top1I = itemRank[0]||['–',0],   top2I = itemRank[1]||['–',0],   top3I = itemRank[2]||['–',0];
  const top1R = ruaRank[0]||['–',0],    top2R = ruaRank[1]||['–',0],    top3R = ruaRank[2]||['–',0];
  const top1D = diaRank[0]||['–',0],    top2D = diaRank[1]||['–',0],    top3D = diaRank[2]||['–',0];
  const top1T = turnoRank[0]||['–',0];

  const combos = {{}};
  data.forEach(r=>{{ const k=r.bairro+' / '+r.turno; combos[k]=(combos[k]||0)+1; }});
  const topCombo = sortedEntries(combos)[0]||['–',0];

  // Turno mais crítico — usa campo turno diretamente (igual ao gráfico)
  const topPer = sortedEntries(count(data,'turno'))[0] || ['–',0];

  const gt   = RAW.length;
  const pB1  = ((top1B[1]/gt)*100).toFixed(1);
  const pT1  = ((top1T[1]/gt)*100).toFixed(1);
  const pI1  = ((top1I[1]/gt)*100).toFixed(1);
  const pI2  = ((top2I[1]/gt)*100).toFixed(1);
  const pI3  = ((top3I[1]/gt)*100).toFixed(1);
  const pPer = ((topPer[1]/gt)*100).toFixed(1);
  const top3DTotal   = (top1D[1]||0)+(top2D[1]||0)+(top3D[1]||0);
  const pTop3D       = ((top3DTotal/gt)*100).toFixed(1);
  const itemTop3Tot  = (top1I[1]||0)+(top2I[1]||0)+(top3I[1]||0);
  const pItemTop3    = ((itemTop3Tot/gt)*100).toFixed(1);
  const itens12tot   = (top1I[1]||0)+(top2I[1]||0);

  function ri(prio, text) {{
    const label = prio==='alta'?'Alta':prio==='media'?'Média':'Baixa';
    return `<div class="rec-item"><span class="rec-prio ${{prio}}">${{label}}</span><span class="rec-text">${{text}}</span></div>`;
  }}

  cont.innerHTML = `
    <div class="rec-card">
      <div class="rec-header"><span class="rec-icon">👮</span><span class="rec-title">Policiamento e Distribuição de Guarnições</span></div>
      <div class="rec-body">
        ${{ri('alta',`<strong>Concentrar no mínimo 50% do efetivo em ${{top1B[0]}}</strong> em todos os turnos. Com ${{top1B[1]}} ocorrências (${{pB1}}%), é o bairro prioritário absoluto em ambos os turnos operacionais.`)}}
        ${{ri('alta',`<strong>Reforço no combo ${{topCombo[0]}} — ${{topCombo[1]}} casos.</strong> Manter pelo menos 2 viaturas neste bairro/turno, com ronda intensiva nas ruas ${{top1R[0]}} e ${{top2R[0]}}.`)}}
        ${{ri('alta',`<strong>Reforço nas ${{top1D[0]}}s (${{top1D[1]}} casos)</strong>, ${{top2D[0]}}s (${{top2D[1]}}) e ${{top3D[0]}}s (${{top3D[1]}}) — os 3 dias mais críticos (${{pTop3D}}% do total). Escalar efetivo adicional nestes dias.`)}}
        ${{ri('alta',`<strong>${{top1T[0]}} = ${{top1T[1]}} casos (${{pT1}}%).</strong> Não reduzir efetivo neste turno. Priorizar cobertura em ${{top1B[0]}} e ${{top2B[0]}} durante ${{top1T[0].toLowerCase()}}.`)}}
        ${{ri('media',`<strong>Destinar viatura fixa para ${{top2B[0]}} e ${{top3B[0]}}</strong> — 2º e 3º bairros mais afetados (${{top2B[1]}} e ${{top3B[1]}} casos). Foco no período ${{topPer[0]}}.`)}}
        ${{ri('media',`<strong>Ronda intensiva em ${{top1R[0]}} (${{top1R[1]}}) e ${{top2R[0]}} (${{top2R[1]}})</strong> — principais endereços de ocorrência. Especialmente no período ${{topPer[0]}} (${{topPer[1]}} casos = ${{pPer}}%).`)}}
      </div>
    </div>
    <div class="rec-card">
      <div class="rec-header"><span class="rec-icon">🎯</span><span class="rec-title">Operações Específicas por Item e Rua</span></div>
      <div class="rec-body">
        ${{ri('alta',`<strong>Prioridade nº 1: ${{top1I[0]}}</strong> — ${{top1I[1]}} furtos (${{pI1}}% do total). Operação direcionada para este item nas ruas ${{top1R[0]}} e ${{top2R[0]}}, especialmente no período ${{topPer[0]}}.`)}}
        ${{ri('alta',`<strong>${{top1R[0]}} = rua mais crítica (${{top1R[1]}} casos).</strong> Requer ronda permanente. Seguida de ${{top2R[0]}} (${{top2R[1]}}) e ${{top3R[0]}} (${{top3R[1]}}) — eixo central das ocorrências.`)}}
        ${{ri('alta',`<strong>${{top2I[0]}} (${{top2I[1]}} casos = ${{pI2}}%) e ${{top3I[0]}} (${{top3I[1]}} = ${{pI3}}%).</strong> Juntos com ${{top1I[0]}}, esses 3 itens representam ${{pItemTop3}}% de todos os furtos. Abordagem ostensiva em áreas de lazer.`)}}
        ${{ri('media',`<strong>Fiscalizar pontos de receptação</strong> nas regiões de ${{top2B[0]}} e ${{top3B[0]}}. Itens de alto valor indicam crime de oportunidade. Revista em pontos de venda informal e ferro-velho.`)}}
        ${{ri('media',`<strong>Intensificar abordagens no combo ${{topCombo[0]}}</strong> (${{topCombo[1]}} casos). Este é o pico mais crítico — abordagem ostensiva a suspeitos no período de maior incidência.`)}}
      </div>
    </div>
    <div class="rec-card">
      <div class="rec-header"><span class="rec-icon">📷</span><span class="rec-title">Tecnologia, Infraestrutura e Prevenção</span></div>
      <div class="rec-body">
        ${{ri('alta',`<strong>Câmeras urgentes nas ruas críticas:</strong> ${{top1R[0]}} (${{top1R[1]}} casos), ${{top2R[0]}} (${{top2R[1]}}) e ${{top3R[0]}} (${{top3R[1]}}). Cobertura 24h com gravação mínima de 30 dias. Integração com central de monitoramento.`)}}
        ${{ri('alta',`<strong>Ampliar iluminação pública em ${{top1R[0]}} e ${{top2R[0]}}.</strong> O período ${{topPer[0]}} concentra ${{topPer[1]}} casos (${{pPer}}%). Iluminação adequada é o fator preventivo mais custo-efetivo contra furtos oportunistas.`)}}
        ${{ri('alta',`<strong>Infraestrutura segura para ${{top1I[0]}} e ${{top2I[0]}}</strong> (${{itens12tot}} casos combinados = ${{((itens12tot/gt)*100).toFixed(1)}}%). Pontos seguros monitorados nas áreas de maior incidência em ${{top1B[0]}}.`)}}
        ${{ri('media',`<strong>Câmeras no entorno dos principais pontos de concentração</strong> em ${{top1B[0]}} e ${{top2B[0]}}. Parceria público-privada para ampliação do monitoramento em locais críticos.`)}}
        ${{ri('media',`<strong>Programa de rastreamento</strong> dos itens mais furtados (${{top1I[0]}}, ${{top2I[0]}}, ${{top3I[0]}}). Registro, aplicativo e banco de dados integrado com PM e PC para recuperação dos bens.`)}}
        ${{ri('baixa',`<strong>Central de monitoramento integrada</strong> com câmeras, rádio e GPS das viaturas. Resposta imediata a acionamentos em ${{top1B[0]}}, especialmente no combo ${{topCombo[0]}}.`)}}
      </div>
    </div>
    <div class="rec-card">
      <div class="rec-header"><span class="rec-icon">📊</span><span class="rec-title">Gestão, Inteligência e Comunidade</span></div>
      <div class="rec-body">
        ${{ri('alta',`<strong>Atualização semanal deste dashboard.</strong> Os dados mostram padrões claros (${{top1D[0]}}, ${{top2D[0]}}, ${{topPer[0]}}, ${{top1B[0]}}) que devem ser monitorados continuamente para realocar viaturas conforme mudanças de comportamento criminal.`)}}
        ${{ri('alta',`<strong>CONSEG atuante em ${{top2B[0]}} e ${{top3B[0]}}</strong> (2º e 3º bairros com ${{top2B[1]}} e ${{top3B[1]}} casos). Reuniões mensais com moradores, comerciantes e PM para troca de informações sobre suspeitos e pontos de risco.`)}}
        ${{ri('media',`<strong>Cruzar dados com agenda de eventos locais:</strong> coincidências com picos de ${{top1D[0]}} e ${{top2D[0]}}. Planejar reforço preventivo nos dias e locais de maior movimento.`)}}
        ${{ri('media',`<strong>Integração com Polícia Civil:</strong> cruzar B.O.s para identificar reincidentes. A concentração em ${{top1R[0]}} (${{top1R[1]}}), ${{top2R[0]}} (${{top2R[1]}}) e ${{top3R[0]}} (${{top3R[1]}}) sugere atuação de grupos com horários definidos.`)}}
        ${{ri('media',`<strong>Campanhas direcionadas por bairro:</strong> em ${{top1B[0]}}, foco em ${{top1I[0]}} e ${{top2I[0]}}; em ${{top2B[0]}}, foco em ${{top2I[0]}} e ${{top3I[0]}}. Abordagem específica por perfil de vitimização local.`)}}
        ${{ri('baixa',`<strong>Canal de denúncia anônima 24h</strong> com foco em ${{topCombo[0]}}. Monitoramento contínuo dos padrões de ${{top1D[0]}} e ${{top2D[0]}} para planejamento tático semanal.`)}}
      </div>
    </div>
  `;
}}

// ── INSIGHTS DINÂMICOS ────────────────────────────────────────────────────────
function renderInsights(data) {{
  const total   = data.length;
  if(total === 0) {{
    document.getElementById('insights-container').innerHTML =
      '<p style="color:#999;font-size:13px">Nenhum dado para o filtro selecionado.</p>';
    return;
  }}
  const turnos  = count(data,'turno');
  const bairros = count(data,'bairro');
  const dias    = count(data,'dia');
  const itens   = count(data,'item');
  const tipos   = count(data,'tipo');
  const topTurno  = sortedEntries(turnos)[0]  || ['–',0];
  const topBairro = sortedEntries(bairros)[0] || ['–',0];
  const topDia    = sortedEntries(dias)[0]    || ['–',0];
  const topItem   = sortedEntries(itens)[0]   || ['–',0];
  const furtos  = tipos['Furto']  || 0;
  const roubos  = tipos['Roubo']  || 0;
  const pFurtos = ((furtos/total)*100).toFixed(0);
  const pBairro = ((topBairro[1]/total)*100).toFixed(0);
  const pTurno  = ((topTurno[1]/total)*100).toFixed(0);
  // Calcular nDias igual ao KPI card
  const _dts = data.map(r=>r.data).filter(Boolean);
  const nDias = _dts.length > 0 ? new Set(_dts).size : 1;
  const pDia    = ((topDia[1]/total)*100).toFixed(0);
  const pItem   = ((topItem[1]/total)*100).toFixed(0);

  const cards = [
    {{
      cls:'alerta', icon:'🚨', title:'Área Crítica',
      text:`O bairro <span class="insight-highlight">${{topBairro[0]}}</span> concentra
            <span class="insight-highlight">${{topBairro[1]}} ocorrências (${{pBairro}}%)</span>
            do total. Exige atenção prioritária de policiamento e monitoramento.`
    }},
    {{
      cls:'atencao', icon:'🕐', title:'Horário de Maior Risco',
      text:`O turno da <span class="insight-highlight">${{topTurno[0]}}</span> responde por
            <span class="insight-highlight">${{topTurno[1]}} casos (${{pTurno}}%)</span>.
            Reforçar rondas neste período pode reduzir significativamente os índices.`
    }},
    {{
      cls:'info', icon:'📦', title:'Item Mais Visado',
      text:`<span class="insight-highlight">${{topItem[0]}}</span> é o bem mais furtado:
            <span class="insight-highlight">${{topItem[1]}} casos (${{pItem}}%)</span>.
            Programas de rastreamento e bicicletários seguros são medidas eficazes.`
    }},
    {{
      cls:'alerta', icon:'📅', title:'Dia Mais Perigoso',
      text:`<span class="insight-highlight">${{topDia[0]}}</span> é o dia com mais ocorrências:
            <span class="insight-highlight">${{topDia[1]}} casos (${{pDia}}%)</span>.
            Operações direcionadas neste dia podem impactar os resultados mensais.`
    }},
    {{
      cls:'positivo', icon:'⚖️', title:'Perfil das Ocorrências',
      text:`<span class="insight-highlight">${{pFurtos}}%</span> dos casos são furtos
            (sem violência). Apenas <span class="insight-highlight">${{roubos}} roubos</span>
            foram registrados, o que indica menor nível de violência direta.`
    }},
    {{
      cls:'info', icon:'📊', title:'Volume do Período',
      text:`Foram registradas <span class="insight-highlight">${{total}} ocorrências</span>
            no período analisado. A média é de
            <span class="insight-highlight">${{(total/nDias).toFixed(1)}} casos/dia</span>,
            com variações conforme dia da semana e turno.`
    }},
  ];

  document.getElementById('insights-container').innerHTML = cards.map(c => `
    <div class="insight-card ${{c.cls}}">
      <div class="insight-title">${{c.icon}} ${{c.title}}</div>
      <div class="insight-text">${{c.text}}</div>
    </div>`).join('');
}}

function renderAll() {{
  const data = filtered();
  const ocs  = dedupBO(data); // 1 registo por B.O. (ocorrência), NÃO por item
  renderKPIs(ocs);
  renderTipo(ocs);
  renderBairro(ocs);
  renderTurno(ocs);
  renderLinha(ocs);
  renderMes(ocs);
  renderItem(data);   // único que usa data bruto: conta cada item separado
  renderDia(ocs);
  renderRuas(ocs);
  renderRefs(ocs);
  renderHeatmap(ocs);
  renderHora(ocs);
  renderComparacao(ocs);
  renderPlanoPolicial(ocs);
  renderDiagnostico(ocs);
  renderRecomendacoes(ocs);
  renderInsights(ocs);
  renderMapa(ocs);
  renderTabela(data);   // tabela detalhada mostra todos os itens do B.O.
  renderChips();
  buildSidebar();
}}

// ── CHIPS DE FILTROS ATIVOS ───────────────────────────────────────────────────
const FILTER_LABELS = {{ano:'Ano',mes:'Mês',turno:'Turno',tipo:'Tipo',bairro:'Bairro',item:'Item',dia:'Dia',logradouro:'Logradouro',recuperado:'Recuperado'}};
function renderChips() {{
  const div = document.getElementById('active-filters');
  const chips = [];
  for(const [key,set] of Object.entries(state)) {{
    set.forEach(v => chips.push(
      `<span class="filter-chip" onclick="toggleFilter('${{key}}','${{v}}')">${{FILTER_LABELS[key]}}: ${{v}}</span>`
    ));
  }}
  div.innerHTML = chips.length
    ? '<span style="font-size:11px;font-weight:600">Filtros:</span>' + chips.join('')
    : '';
  div.className = 'active-filters' + (chips.length?' show':'');
}}

// ── TOGGLE FILTRO ─────────────────────────────────────────────────────────────
function toggleFilter(key, val) {{
  if(state[key].has(val)) state[key].delete(val);
  else state[key].add(val);
  renderAll();
  if(window.innerWidth <= 768) toggleSidebar();
}}

// ── RESET ─────────────────────────────────────────────────────────────────────
function resetFilters() {{
  for(const k of Object.keys(state)) state[k].clear();
  imeiQ = ''; marcaQ = ''; placaQ = ''; numeroSerieQ = ''; corQ = ''; detalhesQ = '';
  const imeiInput = document.querySelector('#filter-imei-input');
  if(imeiInput) imeiInput.value = '';
  const marcaInput = document.querySelector('#filter-marca-input');
  if(marcaInput) marcaInput.value = '';
  const placaInput = document.querySelector('#filter-placa-input');
  if(placaInput) placaInput.value = '';
  const nSerieInput = document.querySelector('#filter-nserie-input');
  if(nSerieInput) nSerieInput.value = '';
  const corInput = document.querySelector('#filter-cor-input');
  if(corInput) corInput.value = '';
  const detalhesInput = document.querySelector('#filter-detalhes-input');
  if(detalhesInput) detalhesInput.value = '';
  renderAll();
}}

// ── GERAR PDF ─────────────────────────────────────────────────────────────────
function gerarPDF() {{
  if(mapaInst) mapaInst.invalidateSize();
  setTimeout(() => window.print(), 300);
}}

// ── COMPARTILHAR VIA WHATSAPP ─────────────────────────────────────────────────
function compartilharWA() {{
  const data = filtered();
  const total  = data.length;
  const furtos = data.filter(r=>r.tipo==='Furto').length;
  const roubos = data.filter(r=>r.tipo==='Roubo').length;
  const arrom  = data.filter(r=>r.tipo==='Arrombamento').length;

  const bairros  = count(data,'bairro');
  const turnos   = count(data,'turno');
  const itens    = count(data,'item');
  const enderecos= count(data,'endereco');
  const topBairro= sortedEntries(bairros)[0]||['–',0];
  const topTurno = sortedEntries(turnos)[0]||['–',0];
  const topItem  = sortedEntries(itens)[0]||['–',0];
  const topRua   = sortedEntries(enderecos).filter(e=>e[0])[0]||['–',0];

  const now = new Date();
  const dataBR = now.toLocaleDateString('pt-BR');
  const horaBR = now.toLocaleTimeString('pt-BR',{{hour:'2-digit',minute:'2-digit'}});
  const pct = (n,t) => t ? ' ('+((n/t)*100).toFixed(0)+'%)' : '';

  const filtros = [];
  if(state.mes.size)        filtros.push('Mês: '+[...state.mes].join(', '));
  if(state.turno.size)      filtros.push('Turno: '+[...state.turno].join(', '));
  if(state.tipo.size)       filtros.push('Tipo: '+[...state.tipo].join(', '));
  if(state.bairro.size)     filtros.push('Bairro: '+[...state.bairro].join(', '));
  if(state.item.size)       filtros.push('Item: '+[...state.item].join(', '));
  if(state.dia.size)        filtros.push('Dia: '+[...state.dia].join(', '));
  if(state.logradouro.size) filtros.push('Logradouro: '+[...state.logradouro].join(', '));
  if(imeiQ)        filtros.push('IMEI: '+imeiQ);
  if(marcaQ)       filtros.push('Marca: '+marcaQ);
  if(corQ)         filtros.push('Cor: '+corQ);
  if(detalhesQ)    filtros.push('Detalhes: '+detalhesQ);
  if(placaQ)       filtros.push('Placa: '+placaQ);
  if(numeroSerieQ) filtros.push('Nº Série: '+numeroSerieQ);
  if(state.recuperado.size) filtros.push('Recuperado: '+[...state.recuperado].join(', '));

  const lines = [
    `🛡️ *DASHBOARD GMBC — Resumo Operacional*`,
    `📅 ${{dataBR}} às ${{horaBR}}`,
    filtros.length ? `🔍 Filtros: ${{filtros.join(' | ')}}` : `📊 Todos os registros`,
    ``,
    `*📊 OCORRÊNCIAS: ${{total}}*`,
    `🔵 Furtos: *${{furtos}}*${{pct(furtos,total)}}`,
    `🔴 Roubos: *${{roubos}}*${{pct(roubos,total)}}`,
    `🟠 Arrombamentos: *${{arrom}}*${{pct(arrom,total)}}`,
    ``,
    `*📍 Bairro mais afetado:* ${{topBairro[0]}} — ${{topBairro[1]}} casos`,
    `*⏰ Turno mais crítico:* ${{topTurno[0]}} — ${{topTurno[1]}} casos`,
    `*📦 Item mais furtado:* ${{topItem[0]}} — ${{topItem[1]}} casos`,
    `*🛣️ Rua mais crítica:* ${{topRua[0]}} — ${{topRua[1]}} casos`,
    ``,
    `_Secretaria de Segurança e Ordem Pública — Guarda Municipal BC_`
  ];
  window.open('https://wa.me/?text='+encodeURIComponent(lines.join('\\n')), '_blank');
}}

// ── ANÁLISE DIÁRIA ────────────────────────────────────────────────────────────
let _analiseTextoWA = '';

// ── AUTENTICAÇÃO ──────────────────────────────────────────────────────────────
(function() {{
  if (sessionStorage.getItem('gmbc_auth') === '1') {{
    const ov = document.getElementById('login-overlay');
    if (ov) ov.style.display = 'none';
  }}
}})();

function sair() {{
  sessionStorage.removeItem('gmbc_auth');
  document.getElementById('login-senha').value = '';
  document.getElementById('login-erro').textContent = '';
  document.getElementById('login-overlay').style.display = 'flex';
}}

function toggleSenha() {{
  const inp = document.getElementById('login-senha');
  const btn = document.querySelector('.login-eye');
  if (!inp) return;
  if (inp.type === 'password') {{
    inp.type = 'text';
    if (btn) btn.textContent = '🙈';
  }} else {{
    inp.type = 'password';
    if (btn) btn.textContent = '👁';
  }}
}}

function verificarSenha() {{
  const inp    = document.getElementById('login-senha');
  const erroEl = document.getElementById('login-erro');
  const cardEl = document.getElementById('login-card');
  const input  = inp ? inp.value : '';
  if (!input) {{ if(erroEl) erroEl.textContent = 'Digite a senha.'; return; }}
  try {{
    const encoded = btoa(unescape(encodeURIComponent(input)));
    if (encoded === '{_senha_b64}') {{
      sessionStorage.setItem('gmbc_auth', '1');
      const ov = document.getElementById('login-overlay');
      if (ov) ov.style.display = 'none';
    }} else {{
      if (erroEl) erroEl.textContent = '❌ Senha incorreta. Tente novamente.';
      if (cardEl) {{ cardEl.classList.remove('login-shake'); void cardEl.offsetWidth; cardEl.classList.add('login-shake'); }}
      if (inp) {{ inp.value = ''; inp.focus(); }}
    }}
  }} catch(e) {{
    if (erroEl) erroEl.textContent = '⚠️ Erro inesperado. Tente novamente.';
  }}
}}

function analiseDiaria() {{
  const now = new Date();
  const horaBR    = now.toLocaleTimeString('pt-BR', {{hour:'2-digit', minute:'2-digit'}});
  const nowDayBR  = now.toLocaleDateString('pt-BR', {{weekday:'long'}}).replace(/^\w/, c=>c.toUpperCase());
  const nowDateBR = now.toLocaleDateString('pt-BR');

  // Dia da semana atual
  const DIAS_JS   = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
  const diaSemana = DIAS_JS[now.getDay()];
  const plural    = diaSemana + 's';

  // Todos os registros desse dia da semana (histórico completo)
  const dayData = RAW.filter(r => r.dia === diaSemana);
  const total   = dayData.length;
  const datas   = [...new Set(dayData.map(r => r.data))];
  const nDias   = datas.length;
  const media   = nDias > 0 ? (total/nDias).toFixed(1) : '–';

  // Atualiza título do modal
  const tituloEl = document.getElementById('analise-titulo');
  if (tituloEl) tituloEl.textContent = '📊 Análise de ' + plural + ' — Guarda Municipal BC';

  if (total === 0) {{
    document.getElementById('analise-corpo').innerHTML =
      '<div style="color:#888;text-align:center;padding:40px;font-style:italic">⚠️ Nenhuma ocorrência registrada em ' + plural + ' até o momento.</div>';
    document.getElementById('analise-overlay').classList.add('ativo');
    return;
  }}

  const cntF = (arr, f) => {{
    const m = {{}};
    arr.forEach(r => {{ if(r[f]) m[r[f]] = (m[r[f]]||0)+1; }});
    return Object.entries(m).sort((a,b) => b[1]-a[1]);
  }};
  const pct = n => total ? ((n/total)*100).toFixed(0)+'%' : '0%';
  const med = n => nDias > 0 ? (n/nDias).toFixed(1) : '–';

  const tipos   = cntF(dayData,'tipo');
  const bairros = cntF(dayData,'bairro');
  const ruas    = cntF(dayData,'endereco');
  const turnos  = cntF(dayData,'turno');
  const topTipo   = tipos[0]  || ['–',0];
  const topBairro = bairros[0]|| ['–',0];
  const topTurno  = turnos[0] || ['–',0];
  const ORDEM_TURNO = ['Madrugada','Manhã','Tarde','Noite'];

  // Recomendações baseadas no padrão histórico do dia da semana
  const recs = [];
  if (topBairro[0] !== '–')
    recs.push(`Em ${{plural}}, o bairro <strong>${{topBairro[0]}}</strong> concentra historicamente ${{topBairro[1]}} ocorrência(s) (${{pct(topBairro[1])}}). Reforçar patrulhamento preventivo nessa área.`);
  if ((ruas[0]||['–'])[0] !== '–')
    recs.push(`A <strong>${{ruas[0][0]}}</strong> é o logradouro com maior incidência em ${{plural}}. Considerar ponto fixo de ronda.`);
  const recTurno = {{
    'Madrugada': `O turno da <strong>Madrugada (00h–05h)</strong> é o mais crítico em ${{plural}}. Reforçar efetivo e rondas preventivas nos pontos vulneráveis.`,
    'Manhã':     `O turno <strong>Matutino (06h–11h)</strong> concentra maior índice em ${{plural}}. Ampliar guarnições nos bairros críticos.`,
    'Tarde':     `O período <strong>Vespertino (12h–17h)</strong> é o mais crítico em ${{plural}}. Atenção especial à movimentação urbana.`,
    'Noite':     `O turno <strong>Noturno (18h–23h)</strong> é crítico em ${{plural}}. Coordenar blitz e pontos de controle nas vias principais.`
  }};
  if (recTurno[topTurno[0]]) recs.push(recTurno[topTurno[0]]);
  const recTipo = {{
    'Furto':        `<strong>Furtos</strong> lideram em ${{plural}}. Orientar guarnições: abordagem preventiva próxima a comércio e estacionamentos.`,
    'Roubo':        `<strong>Roubos</strong> são o crime mais frequente em ${{plural}}. Acionar inteligência para identificar padrões e possíveis autores.`,
    'Arrombamento': `<strong>Arrombamentos</strong> predominam em ${{plural}}. Intensificar rondas em imóveis comerciais e residenciais.`
  }};
  if (recTipo[topTipo[0]]) recs.push(recTipo[topTipo[0]]);
  recs.push(`Compartilhar esta análise com todas as guarnições no início do turno e manter comunicação ativa pelo rádio.`);

  const mkRows = (data, cols) => data.map((e,i) =>
    `<tr>${{cols.map(c=>`<td>${{c(e,i)}}</td>`).join('')}}</tr>`).join('');
  const mkTable = (rows, headers) =>
    `<table class="analise-table"><thead><tr>${{headers.map(h=>`<th>${{h}}</th>`).join('')}}</tr></thead><tbody>${{rows}}</tbody></table>`;

  const tiposRows   = mkRows(tipos.slice(0,7),  [(e,i)=>`${{i+1}}. ${{e[0]}}`, e=>e[1], e=>pct(e[1]), e=>med(e[1])]);
  const bairrosRows = mkRows(bairros.slice(0,5), [(e,i)=>`${{i+1}}. ${{e[0]}}`, e=>e[1], e=>pct(e[1])]);
  const ruasRows    = mkRows(ruas.slice(0,5),    [(e,i)=>`${{i+1}}. ${{e[0]}}`, e=>e[1]]);
  const turnosRows  = ORDEM_TURNO.map(t => {{
    const n = (turnos.find(e=>e[0]===t)||[t,0])[1];
    return `<tr><td>${{t}}</td><td>${{n}}</td><td>${{pct(n)}}</td><td>${{med(n)}}/dia</td></tr>`;
  }}).join('');

  const datasFormatadas = datas.sort().map(d => {{
    const pts = d.split('-');
    return pts[2]+'/'+pts[1];
  }}).join(', ');

  const corpo = `
  <div style="text-align:center;background:linear-gradient(135deg,#f0f4ff,#e8f0fe);border-radius:8px;
    padding:12px;margin-bottom:14px;border-left:4px solid #0078D4;">
    <div style="font-weight:800;font-size:15px;color:#1A1A2E">📊 Análise de ${{plural}}</div>
    <div style="color:#555;font-size:11px;margin-top:4px">
      Guarda Municipal de Balneário Camboriú<br>
      📅 Gerado em: <strong>${{nowDayBR}}, ${{nowDateBR}}</strong> às <strong>${{horaBR}}</strong>
    </div>
  </div>
  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:14px;">
    <div style="background:#1A1A2E;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Total histórico</div>
      <div style="font-size:26px;font-weight:800;line-height:1.1">${{total}}</div>
      <div style="font-size:9px;opacity:.65">ocorrências</div>
    </div>
    <div style="background:#0078D4;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">${{plural}} analisados</div>
      <div style="font-size:26px;font-weight:800;line-height:1.1">${{nDias}}</div>
      <div style="font-size:9px;opacity:.65">com registros</div>
    </div>
    <div style="background:#107C10;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Média por ${{diaSemana}}</div>
      <div style="font-size:26px;font-weight:800;line-height:1.1">${{media}}</div>
      <div style="font-size:9px;opacity:.65">oc./dia</div>
    </div>
    <div style="background:#5C2D91;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Turno crítico</div>
      <div style="font-size:14px;font-weight:800;line-height:1.3;margin-top:4px">${{topTurno[0]}}</div>
      <div style="font-size:9px;opacity:.65">${{topTurno[1]}} oc. (${{pct(topTurno[1])}})</div>
    </div>
  </div>
  <div class="analise-section">
    <h3>🔴 Tipos de Ocorrência em ${{plural}}</h3>
    ${{mkTable(tiposRows,['Tipo','Total','%','Média/dia'])}}
  </div>
  <div class="analise-section">
    <h3>📍 Bairros Mais Afetados em ${{plural}}</h3>
    ${{mkTable(bairrosRows,['Bairro','Total','%'])}}
  </div>
  <div class="analise-section">
    <h3>🛣️ Logradouros de Risco em ${{plural}}</h3>
    ${{mkTable(ruasRows,['Logradouro','Total'])}}
  </div>
  <div class="analise-section">
    <h3>⏰ Distribuição por Turno em ${{plural}}</h3>
    ${{mkTable(turnosRows,['Turno','Total','%','Média/dia'])}}
  </div>
  <div class="analise-section">
    <h3>🎯 Recomendações para ${{plural}}</h3>
    <ol class="analise-rec-ol">${{recs.map(r=>`<li>${{r}}</li>`).join('')}}</ol>
  </div>
  <div style="border-top:1px solid #ddd;margin-top:12px;padding-top:8px;color:#888;font-size:10px;text-align:center;">
    Secretaria de Segurança e Ordem Pública de Balneário Camboriú — Guarda Municipal
  </div>`;

  const waLns = [
    `📊 *ANÁLISE DE ${{plural.toUpperCase()}} — GUARDA MUNICIPAL BC*`,
    `📅 ${{nowDayBR}}, ${{nowDateBR}} às ${{horaBR}}`,
    ``,
    `*📋 RESUMO HISTÓRICO DE ${{plural.toUpperCase()}}*`,
    `Total de ocorrências: *${{total}}*`,
    `${{plural}} com dados: *${{nDias}}* | Média: *${{media}}* oc./dia`,
    `Tipo mais frequente: *${{topTipo[0]}}* (${{topTipo[1]}} — ${{pct(topTipo[1])}})`,
    `Bairro mais afetado: *${{topBairro[0]}}* (${{topBairro[1]}} oc.)`,
    `Turno crítico: *${{topTurno[0]}}* (${{topTurno[1]}} oc. — ${{pct(topTurno[1])}})`,
    ``,
    `*🔴 TIPOS EM ${{plural.toUpperCase()}}*`,
    ...tipos.slice(0,6).map(e=>`• ${{e[0]}}: ${{e[1]}} total (${{pct(e[1])}}) | média ${{med(e[1])}}/dia`),
    ``,
    `*📍 BAIRROS EM ${{plural.toUpperCase()}}*`,
    ...bairros.slice(0,4).map((e,i)=>`${{i+1}}. ${{e[0]}}: ${{e[1]}} oc. (${{pct(e[1])}})`),
    ``,
    `*⏰ TURNOS EM ${{plural.toUpperCase()}}*`,
    ...ORDEM_TURNO.map(t=>{{ const n=(turnos.find(e=>e[0]===t)||[t,0])[1]; return `• ${{t}}: ${{n}} (${{pct(n)}}) | média ${{med(n)}}/dia`; }}),
    ``,
    `*🎯 RECOMENDAÇÕES*`,
    ...recs.map((r,i)=>`${{i+1}}. ${{r.replace(/<[^>]+>/g,'')}}`),
    ``,
    `_${{plural}} analisados: ${{datasFormatadas}}_`,
    `_Guarda Municipal de Balneário Camboriú_`,
    `_Secretaria de Segurança e Ordem Pública_`
  ];
  _analiseTextoWA = waLns.join('\\n');

  document.getElementById('analise-corpo').innerHTML = corpo;
  document.getElementById('analise-overlay').classList.add('ativo');
}}

function fecharAnalise() {{
  document.getElementById('analise-overlay').classList.remove('ativo');
}}

// ── EXPLICAR COM IA (OpenAI via backend) ──────────────────────────────────────
let _preditResumoIA = '';
let _resumoExecutivoDados = '';

async function explicarComIA(resumoTexto, boxId, btnId) {{
  const btn = btnId ? document.getElementById(btnId) : null;
  const box = document.getElementById(boxId);
  if (!box || !resumoTexto) return;
  if (btn) {{ btn.disabled = true; btn.textContent = '⏳ Gerando...'; }}
  box.style.display = 'block';
  box.innerHTML = '<span style="color:#888;font-style:italic">Consultando IA...</span>';
  try {{
    const resp = await fetch('/api/explicar', {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify({{resumo: resumoTexto}})
    }});
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.erro || 'Erro desconhecido');
    const texto = data.texto.replace(/</g,'&lt;').replace(/\\n/g,'<br>');
    box.innerHTML = '<div style="line-height:1.6">' + texto + '</div>';
  }} catch (e) {{
    box.innerHTML = '<span style="color:#D13438">Erro ao gerar: ' + e.message + '</span>';
  }} finally {{
    if (btn) {{ btn.disabled = false; btn.textContent = btn.dataset.label; }}
  }}
}}

// ── CHAT IA (perguntas livres sobre os dados reais) ───────────────────────────
function _normalizarTexto(s) {{
  return (s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'');
}}

function montarContextoChat(pergunta) {{
  // Respeita os filtros já ativos na sidebar (ano, mês, bairro, item, etc.)
  // e só então aplica os critérios adicionais identificados na pergunta.
  const dataFiltrada = filtered();
  const dadosU = dedupBO(dataFiltrada);
  const pNorm = _normalizarTexto(pergunta);

  const bairros = [...new Set(dadosU.map(r=>r.bairro).filter(Boolean))];
  const tipos   = [...new Set(dadosU.map(r=>r.tipo).filter(Boolean))];
  const turnos  = [...new Set(dadosU.map(r=>r.turno).filter(Boolean))];
  const itens   = [...new Set(dataFiltrada.map(r=>r.item).filter(Boolean))];

  const acharNaLista = lista => lista.filter(v => pNorm.includes(_normalizarTexto(v)));

  const critBairro = acharNaLista(bairros);
  const critTipo   = acharNaLista(tipos);
  const critTurno  = acharNaLista(turnos);
  const critDia    = acharNaLista(DIA_ORDER);
  const critItem   = acharNaLista(itens);

  const semCriterio = !critBairro.length && !critTipo.length && !critTurno.length && !critDia.length && !critItem.length;
  const semFiltroAtivo = dataFiltrada.length === RAW.length;
  if (semCriterio && semFiltroAtivo) return '';

  let subset = dadosU;
  if (critBairro.length) subset = subset.filter(r => critBairro.includes(r.bairro));
  if (critTipo.length)   subset = subset.filter(r => critTipo.includes(r.tipo));
  if (critTurno.length)  subset = subset.filter(r => critTurno.includes(r.turno));
  if (critDia.length)    subset = subset.filter(r => critDia.includes(r.dia));
  if (critItem.length) {{
    const bosComItem = new Set(dataFiltrada.filter(r => critItem.includes(r.item)).map(r=>r.bo));
    subset = subset.filter(r => bosComItem.has(r.bo));
  }}

  const criteriosTxt = [...critBairro,...critTipo,...critTurno,...critDia,...critItem].join(', ') || 'filtros já ativos na tela';
  if (subset.length === 0) return `Nenhum registro encontrado para os critérios identificados (${{criteriosTxt}}).`;

  const porTipo   = count(subset,'tipo');
  const porBairro = count(subset,'bairro');
  const porTurno  = count(subset,'turno');
  const porDia    = count(subset,'dia');
  const fmtObj = o => Object.entries(o).map(([k,v])=>`${{k}}: ${{v}}`).join(', ');

  const amostra = [...subset].sort((a,b)=>(b.data||'').localeCompare(a.data||'')).slice(0,15);
  const amostraTxt = amostra.map(r =>
    `${{r.data||'?'}} | ${{r.bairro||'?'}} | ${{r.turno||'?'}} | ${{r.tipo||'?'}} | ${{r.endereco||'?'}}`
  ).join('\\n');

  return `Critérios identificados na pergunta: ${{criteriosTxt}}.\n` +
    `Total de registros encontrados: ${{subset.length}}.\n` +
    `Distribuição por tipo: ${{fmtObj(porTipo)}}.\n` +
    `Distribuição por bairro: ${{fmtObj(porBairro)}}.\n` +
    `Distribuição por turno: ${{fmtObj(porTurno)}}.\n` +
    `Distribuição por dia da semana: ${{fmtObj(porDia)}}.\n` +
    `Amostra de até 15 registros mais recentes (data | bairro | turno | tipo | endereço):\n${{amostraTxt}}`;
}}

function toggleChatIA() {{
  document.getElementById('chat-ia-panel').classList.toggle('aberto');
}}

async function enviarPerguntaChat() {{
  const input = document.getElementById('chat-ia-input');
  const pergunta = input.value.trim();
  if (!pergunta) return;
  input.value = '';

  const log = document.getElementById('chat-ia-log');
  const respId = 'chat-resp-' + Date.now();
  log.insertAdjacentHTML('beforeend', `<div class="chat-msg chat-msg-user">${{pergunta.replace(/</g,'&lt;')}}</div>`);
  log.insertAdjacentHTML('beforeend', `<div class="chat-msg chat-msg-ia" id="${{respId}}"><span style="color:#888;font-style:italic">Consultando IA...</span></div>`);
  log.scrollTop = log.scrollHeight;

  const contexto = montarContextoChat(pergunta);
  const alvo = document.getElementById(respId);
  try {{
    const resp = await fetch('/api/chat', {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify({{pergunta, contexto}})
    }});
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.erro || 'Erro desconhecido');
    alvo.innerHTML = data.texto.replace(/</g,'&lt;').replace(/\\n/g,'<br>');
  }} catch (e) {{
    alvo.innerHTML = '<span style="color:#D13438">Erro: ' + e.message + '</span>';
  }}
  log.scrollTop = log.scrollHeight;
}}

// ── ENTRADA POR VOZ (Whisper via backend) ─────────────────────────────────────
let _mediaRecorder = null;
let _audioChunks = [];
let _gravando = false;

async function toggleGravacao() {{
  const btn = document.getElementById('chat-ia-mic');
  if (!_gravando) {{
    try {{
      const stream = await navigator.mediaDevices.getUserMedia({{ audio: true }});
      _audioChunks = [];
      _mediaRecorder = new MediaRecorder(stream);
      _mediaRecorder.ondataavailable = e => {{ if (e.data.size > 0) _audioChunks.push(e.data); }};
      _mediaRecorder.onstop = async () => {{
        stream.getTracks().forEach(t => t.stop());
        const blob = new Blob(_audioChunks, {{ type: 'audio/webm' }});
        await transcreverAudio(blob);
      }};
      _mediaRecorder.start();
      _gravando = true;
      btn.textContent = '⏹';
      btn.classList.add('gravando');
    }} catch (e) {{
      alert('Não foi possível acessar o microfone: ' + e.message);
    }}
  }} else {{
    _mediaRecorder.stop();
    _gravando = false;
    btn.textContent = '🎤';
    btn.classList.remove('gravando');
  }}
}}

async function transcreverAudio(blob) {{
  const input = document.getElementById('chat-ia-input');
  const placeholderOriginal = input.placeholder;
  input.placeholder = 'Transcrevendo áudio...';
  input.disabled = true;
  try {{
    const formData = new FormData();
    formData.append('audio', blob, 'audio.webm');
    const resp = await fetch('/api/transcrever', {{ method: 'POST', body: formData }});
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.erro || 'Erro desconhecido');
    input.value = data.texto;
    input.focus();
  }} catch (e) {{
    alert('Erro ao transcrever: ' + e.message);
  }} finally {{
    input.disabled = false;
    input.placeholder = placeholderOriginal;
  }}
}}

// ── GERAR SLIDES (PDF, abre direto no navegador) ──────────────────────────
async function gerarSlides() {{
  const btn = document.getElementById('btn-gerar-slides');
  const labelOriginal = btn.innerHTML;

  // Abre a aba já no clique (síncrono) pra não ser bloqueada como pop-up.
  const novaAba = window.open('', '_blank');
  if (novaAba) {{
    novaAba.document.write('<title>Gerando apresentação...</title><body style="font-family:sans-serif;padding:40px;text-align:center;color:#555">Gerando apresentação, aguarde...</body>');
  }}

  btn.disabled = true;
  btn.innerHTML = '⏳<span class="btxt"> Gerando...</span>';

  try {{
    const chips = [];
    for (const [key, set] of Object.entries(state)) {{ set.forEach(v => chips.push(v)); }}
    const filtroDesc = chips.length ? chips.join(', ') : 'Todos os registros';
    const agora = new Date();
    const dataHora = agora.toLocaleDateString('pt-BR') + ' às ' + agora.toLocaleTimeString('pt-BR',{{hour:'2-digit',minute:'2-digit'}});

    const NAVY = [51,61,101], GRAY = [168,170,174], LIGHT = [222,224,227];
    const TXT = [70,74,82], SUBTXT = [145,148,155];
    const W = 10, H = 5.63;
    const {{ jsPDF }} = window.jspdf;
    const doc = new jsPDF({{ orientation:'landscape', unit:'in', format:[W,H] }});

    function novoSlide() {{ doc.addPage([W,H], 'landscape'); }}
    function tituloSlide(txt) {{
      doc.setFillColor(...NAVY);
      doc.rect(0, 0, W, 0.72, 'F');
      doc.setFillColor(...GRAY);
      doc.rect(0, 0.72, W, 0.05, 'F');
      doc.setFillColor(...GRAY);
      doc.rect(0, 0.77, 0.07, H-0.77, 'F');
      doc.setTextColor(255,255,255);
      doc.setFont(undefined, 'bold');
      doc.setFontSize(17);
      doc.text(txt, 0.4, 0.36);
      doc.setFont(undefined, 'bold');
      doc.setFontSize(9.5);
      doc.setTextColor(...LIGHT);
      doc.text('FILTRO: ' + filtroDesc.toUpperCase(), 0.4, 0.58);
      doc.setFont(undefined, 'normal');
      if (LOGO_GMBC) {{
        const lh = 0.5, lw = lh * (1442/1612);
        doc.addImage(LOGO_GMBC, 'PNG', W - lw - 0.25, 0.11, lw, lh);
      }}
    }}
    function rodapeSlide(pagina, total) {{
      doc.setFillColor(...GRAY);
      doc.rect(0, H-0.34, W, 0.035, 'F');
      doc.setFont(undefined, 'normal');
      doc.setFontSize(7.5);
      doc.setTextColor(...SUBTXT);
      doc.text('Secretaria de Segurança e Ordem Pública — Guarda Municipal BC', 0.4, H-0.15);
      doc.setFontSize(6.5);
      doc.setTextColor(170,170,178);
      doc.text('Desenvolvido por Ronaldo E. Barbosa - Guarda Municipal', W/2, H-0.15, {{ align:'center' }});
      doc.setFontSize(7.5);
      doc.setTextColor(...SUBTXT);
      doc.text(pagina + ' / ' + total, W-0.4, H-0.15, {{ align:'right' }});
    }}

    // Escreve UMA linha justificada (esquerda E direita), palavra por palavra.
    // Não força linhas curtas (poucas palavras) a esticar até a borda —
    // fica feio um "gap" enorme numa linha de 2-3 palavras.
    function linhaJustificada(linha, x, ly, maxWidth) {{
      const palavras = linha.trim().split(/\s+/).filter(Boolean);
      const larguraPalavras = palavras.reduce((s,p) => s + doc.getTextWidth(p), 0);
      if (palavras.length <= 2 || larguraPalavras < maxWidth*0.55) {{
        doc.text(linha, x, ly);
        return;
      }}
      const espacoPorGap = (maxWidth - larguraPalavras) / (palavras.length - 1);
      let cx = x;
      palavras.forEach((p) => {{
        doc.text(p, cx, ly);
        cx += doc.getTextWidth(p) + espacoPorGap;
      }});
    }}

    // Escreve texto justificado numa única coluna.
    // IMPORTANTE: maxWidth deve ficar abaixo de ~5 in — versões deste jsPDF
    // truncam silenciosamente texto de uma única chamada text() acima de
    // ~5.2 in de largura acumulada (bug da biblioteca, não do nosso código).
    function escreverJustificado(texto, x, y, maxWidth, lineHeight) {{
      const linhas = doc.splitTextToSize(texto, maxWidth);
      linhas.forEach((linha, i) => {{
        const isLast = i === linhas.length - 1;
        const ly = y + i*lineHeight;
        if (isLast) {{ doc.text(linha, x, ly); return; }}
        linhaJustificada(linha, x, ly, maxWidth);
      }});
      return linhas.length;
    }}

    // Escreve texto justificado fluindo em 2 colunas lado a lado, dividindo
    // as linhas de forma equilibrada entre as duas (evita coluna vazia).
    function escreverJustificado2Colunas(texto, x1, x2, y, colW, lineHeight) {{
      const linhas = doc.splitTextToSize(texto, colW);
      const linhasPorColuna = Math.ceil(linhas.length / 2);
      linhas.forEach((linha, i) => {{
        const col = Math.floor(i / linhasPorColuna);
        const linhaNaCol = i % linhasPorColuna;
        const x = col === 0 ? x1 : x2;
        const ly = y + linhaNaCol*lineHeight;
        const isLastDoTexto = i === linhas.length - 1;
        const isUltimaDaCol = linhaNaCol === linhasPorColuna - 1;
        if (isLastDoTexto || isUltimaDaCol) {{ doc.text(linha, x, ly); return; }}
        linhaJustificada(linha, x, ly, colW);
      }});
    }}

    // ── Slide 1: título ──
    doc.setFillColor(...NAVY);
    doc.rect(0, 0, W, H, 'F');
    doc.setFillColor(...GRAY);
    doc.rect(0, 0, 0.14, H, 'F');
    doc.setFillColor(...GRAY);
    doc.rect(0, H-0.14, W, 0.14, 'F');
    if (LOGO_GMBC) {{
      const lh = 1.1, lw = lh * (1442/1612);
      doc.addImage(LOGO_GMBC, 'PNG', (W-lw)/2, 0.45, lw, lh);
    }}
    doc.setTextColor(255,255,255);
    doc.setFont(undefined, 'bold');
    doc.setFontSize(22);
    doc.text('Secretaria de Segurança e Ordem Pública', W/2, 2.0, {{ align:'center' }});
    doc.setFontSize(13);
    doc.setTextColor(195,205,220);
    doc.setFont(undefined, 'normal');
    doc.text('Balneário Camboriú — Guarda Municipal', W/2, 2.5, {{ align:'center' }});
    doc.setTextColor(...LIGHT);
    doc.setFont(undefined, 'bold');
    doc.setFontSize(12);
    doc.text(filtroDesc, W/2, 3.2, {{ align:'center' }});
    doc.setFont(undefined, 'normal');
    doc.setFontSize(9);
    doc.setTextColor(...SUBTXT);
    doc.text('Gerado em ' + dataHora, W/2, 4.6, {{ align:'center' }});
    doc.setFontSize(7);
    doc.setTextColor(120,130,145);
    doc.text('Desenvolvido por Ronaldo E. Barbosa - Guarda Municipal', W/2, H-0.32, {{ align:'center' }});

    // ── Slide 2: KPIs (cards) ──
    const kpiCards = [
      {{ label:'Total de Ocorrências',  value: document.getElementById('kpi-total').textContent }},
      {{ label:'Bairro Mais Afetado',   value: document.getElementById('kpi-bairro').textContent, sub: document.getElementById('kpi-bairro-num').textContent+' casos' }},
      {{ label:'Turno Mais Crítico',    value: document.getElementById('kpi-turno').textContent,  sub: document.getElementById('kpi-turno-num').textContent+' casos' }},
      {{ label:'Dia Mais Crítico',      value: document.getElementById('kpi-dia').textContent,    sub: document.getElementById('kpi-dia-num').textContent+' casos' }},
      {{ label:'Furtos',                value: document.getElementById('kpi-furtos').textContent, sub: document.getElementById('kpi-furtos-sub').textContent }},
      {{ label:'Roubos',                value: document.getElementById('kpi-roubos').textContent, sub: document.getElementById('kpi-roubos-sub').textContent }},
      {{ label:'Arrombamentos',         value: document.getElementById('kpi-arrom').textContent,  sub: document.getElementById('kpi-arrom-sub').textContent }},
    ];
    novoSlide();
    tituloSlide('Indicadores Principais');
    const cardW = 2.15, cardH = 1.42, gapX = 0.2, gapY = 0.2, startX = 0.4, startY = 1.05;
    kpiCards.forEach((c, idx) => {{
      const col = idx % 4, row = Math.floor(idx/4);
      const x = startX + col*(cardW+gapX);
      const y = startY + row*(cardH+gapY);
      doc.setFillColor(...NAVY);
      doc.roundedRect(x, y, cardW, cardH, 0.05, 0.05, 'F');
      doc.setFillColor(...GRAY);
      doc.rect(x, y, cardW, 0.06, 'F');
      doc.setFont(undefined, 'bold');
      doc.setFontSize(17);
      doc.setTextColor(...LIGHT);
      doc.text(String(c.value), x+cardW/2, y+0.65, {{ align:'center' }});
      doc.setFont(undefined, 'normal');
      doc.setFontSize(9);
      doc.setTextColor(...GRAY);
      const labelLinhas = doc.splitTextToSize(c.label, cardW-0.2);
      doc.text(labelLinhas, x+cardW/2, y+0.95, {{ align:'center' }});
      if (c.sub) {{
        doc.setFont(undefined, 'bold');
        doc.setFontSize(7.5);
        doc.setTextColor(...LIGHT);
        doc.text(String(c.sub), x+cardW/2, y+cardH-0.15, {{ align:'center' }});
        doc.setFont(undefined, 'normal');
      }}
    }});

    // ── Gráficos principais: 1 por página, com observação/sugestão da IA ──
    const dataFiltrada = filtered();
    const dadosU = dedupBO(dataFiltrada);
    function topEntradas(obj, n) {{
      return sortedEntries(obj).slice(0, n || 6).map(([k,v]) => k + ' ' + v).join(', ');
    }}
    const graficos = [
      {{ id:'chart-tipo',   titulo:'Tipificação das Ocorrências',      dados:() => topEntradas(count(dadosU,'tipo')) }},
      {{ id:'chart-bairro', titulo:'Bairros com Mais Ocorrências',      dados:() => topEntradas(count(dadosU,'bairro')) }},
      {{ id:'chart-turno',  titulo:'Ocorrências por Turno',             dados:() => topEntradas(count(dadosU,'turno')) }},
      {{ id:'chart-dia',    titulo:'Ocorrências por Dia da Semana',     dados:() => topEntradas(count(dadosU,'dia')) }},
      {{ id:'chart-item',   titulo:'Itens Mais Furtados/Roubados',      dados:() => topEntradas(count(dataFiltrada,'item')) }},
      {{ id:'chart-ruas',   titulo:'Ruas com Mais Ocorrências',         dados:() => topEntradas(count(dadosU,'endereco')) }},
    ];

    let insightsBlocos = [];
    try {{
      const categoriasTxt = graficos.map((g,i) => (i+1) + ') ' + g.titulo + ': ' + g.dados() + '.').join('\\n');
      const resp = await fetch('/api/insights', {{
        method: 'POST',
        headers: {{'Content-Type': 'application/json'}},
        body: JSON.stringify({{categorias: categoriasTxt}})
      }});
      const data = await resp.json();
      if (resp.ok) insightsBlocos = data.texto.split(/###\d+###/).map(b => b.trim()).filter(Boolean);
    }} catch (e) {{ /* segue sem observações/sugestões se a IA falhar */ }}

    function parseInsight(bloco) {{
      if (!bloco) return {{ obs:'', sug:'' }};
      const obsMatch = bloco.match(/Observação:\s*([\s\S]*?)(?=\\s*Sugestão:|$)/i);
      const sugMatch = bloco.match(/Sugestão:\s*([\s\S]*)/i);
      return {{
        obs: obsMatch ? obsMatch[1].trim() : '',
        sug: sugMatch ? sugMatch[1].trim() : '',
      }};
    }}

    for (let i = 0; i < graficos.length; i++) {{
      const g = graficos[i];
      const el = document.getElementById(g.id);
      if (!el || !window.Plotly) continue;
      try {{
        const img = await Plotly.toImage(el, {{ format:'png', width:800, height:450 }});
        novoSlide();
        tituloSlide(g.titulo);

        // ── Coluna esquerda: gráfico ──
        const chartColW = 5.0;
        const imgW = 4.5, imgH = imgW * (450/800);
        const panelY = 0.95, panelH = 4.25;
        const imgX = 0.4 + (chartColW - imgW)/2;
        const imgY = panelY + (panelH - imgH)/2;
        doc.addImage(img, 'PNG', imgX, imgY, imgW, imgH);

        // ── Coluna direita: caixa com observação/sugestão ──
        const panelX = 5.6, panelW = 4.0;
        doc.setFillColor(...NAVY);
        doc.roundedRect(panelX, panelY, panelW, panelH, 0.06, 0.06, 'F');

        const {{ obs, sug }} = parseInsight(insightsBlocos[i]);
        const textX = panelX + 0.3;
        const textW = panelW - 0.6;
        let ty = panelY + 0.45;
        if (obs) {{
          doc.setFillColor(...LIGHT);
          doc.rect(textX-0.18, ty-0.1, 0.045, 0.15, 'F');
          doc.setFont(undefined, 'bold');
          doc.setFontSize(10);
          doc.setTextColor(...LIGHT);
          doc.text('OBSERVAÇÃO', textX, ty);
          doc.setFont(undefined, 'normal');
          doc.setFontSize(10.5);
          doc.setTextColor(...GRAY);
          const nObsLinhas = escreverJustificado(obs, textX, ty+0.22, textW, 0.19);
          ty += 0.22 + 0.19*nObsLinhas + 0.4;
        }}
        if (sug) {{
          doc.setFillColor(...GRAY);
          doc.rect(textX-0.18, ty-0.1, 0.045, 0.15, 'F');
          doc.setFont(undefined, 'bold');
          doc.setFontSize(10);
          doc.setTextColor(...GRAY);
          doc.text('SUGESTÃO', textX, ty);
          doc.setFont(undefined, 'normal');
          doc.setFontSize(10.5);
          doc.setTextColor(...GRAY);
          escreverJustificado(sug, textX, ty+0.22, textW, 0.19);
        }}
      }} catch (e) {{ /* ignora gráfico que falhar */ }}
    }}

    // ── Slide final: Resumo Executivo IA ──
    let resumoTexto = 'Resumo não disponível.';
    try {{
      const resp = await fetch('/api/explicar', {{
        method: 'POST',
        headers: {{'Content-Type': 'application/json'}},
        body: JSON.stringify({{resumo: _resumoExecutivoDados}})
      }});
      const data = await resp.json();
      if (resp.ok) resumoTexto = data.texto;
    }} catch (e) {{ /* mantém texto padrão se a IA falhar */ }}
    novoSlide();
    tituloSlide('Resumo Executivo — IA');
    doc.setFontSize(11);
    doc.setTextColor(...TXT);
    doc.setFont(undefined, 'normal');
    // Coluna única, mas com largura abaixo do limite de ~5.2in que trunca
    // texto silenciosamente no jsPDF (bug da biblioteca, ver escreverJustificado).
    // 2 colunas de 4.2in cada, preenchendo a largura toda do slide sem
    // ultrapassar o limite seguro de largura por linha (~5.2in, ver acima).
    escreverJustificado2Colunas(resumoTexto, 0.6, 5.2, 1.1, 4.2, 0.2);

    // ── Slide de encerramento ──
    novoSlide();
    doc.setFillColor(...NAVY);
    doc.rect(0, 0, W, H, 'F');
    doc.setFillColor(...GRAY);
    doc.rect(0, 0, 0.14, H, 'F');
    doc.setFillColor(...GRAY);
    doc.rect(0, H-0.14, W, 0.14, 'F');
    if (LOGO_GMBC) {{
      const lh = 2.6, lw = lh * (1442/1612);
      doc.addImage(LOGO_GMBC, 'PNG', (W-lw)/2, 0.65, lw, lh);
    }}
    doc.setTextColor(255,255,255);
    doc.setFont(undefined, 'bold');
    doc.setFontSize(18);
    doc.text('Apresentação Concluída', W/2, 3.7, {{ align:'center' }});
    doc.setFont(undefined, 'normal');
    doc.setFontSize(11);
    doc.setTextColor(...LIGHT);
    doc.text('Secretaria de Segurança e Ordem Pública — Guarda Municipal BC', W/2, 4.1, {{ align:'center' }});

    // ── Rodapé (numeração + identificação) em todas as páginas, exceto a capa ──
    const totalPaginas = doc.internal.getNumberOfPages();
    for (let p = 2; p <= totalPaginas; p++) {{
      doc.setPage(p);
      rodapeSlide(p, totalPaginas);
    }}

    const blobUrl = doc.output('bloburl');
    if (novaAba) {{
      novaAba.location.href = blobUrl;
    }} else {{
      window.open(blobUrl, '_blank');
    }}
  }} catch (e) {{
    if (novaAba) novaAba.close();
    alert('Erro ao gerar apresentação: ' + e.message);
  }} finally {{
    btn.disabled = false;
    btn.innerHTML = labelOriginal;
  }}
}}

function abrirResumoIA() {{
  const agora = new Date();
  document.getElementById('resumoia-meta').textContent =
    'Gerado em: ' + agora.toLocaleDateString('pt-BR') + ' às ' +
    agora.toLocaleTimeString('pt-BR',{{hour:'2-digit',minute:'2-digit'}});
  document.getElementById('resumoia-overlay').classList.add('ativo');
  explicarComIA(_resumoExecutivoDados, 'resumoia-resultado', null);
}}
function fecharResumoIA() {{
  document.getElementById('resumoia-overlay').classList.remove('ativo');
}}
function imprimirResumoIA() {{ window.print(); }}

// ── PREVISÃO DE RISCO ─────────────────────────────────────────────────────────
let _prevTextoWA = '';
function previsao() {{
  const now      = new Date();
  const horaBR   = now.toLocaleTimeString('pt-BR', {{hour:'2-digit',minute:'2-digit'}});
  const nowDayBR = now.toLocaleDateString('pt-BR', {{weekday:'long'}}).replace(/^\w/,c=>c.toUpperCase());
  const nowDateBR= now.toLocaleDateString('pt-BR');
  const DIAS_JS  = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
  const diaSemana= DIAS_JS[now.getDay()];
  const plural   = diaSemana + 's';
  const tituloEl = document.getElementById('prev-titulo');
  if (tituloEl) tituloEl.textContent = '📈 Previsão de Risco — ' + plural + ' — Guarda Municipal BC';

  const hist  = dedupBO(RAW).filter(r => r.dia === diaSemana);
  const datas = [...new Set(hist.map(r=>r.data))].sort();
  const nDias = datas.length;

  if (nDias === 0) {{
    document.getElementById('prev-corpo').innerHTML =
      '<div style="color:#888;text-align:center;padding:40px;font-style:italic">⚠️ Nenhuma ocorrência registrada em ' + plural + ' até o momento.</div>';
    document.getElementById('prev-overlay').classList.add('ativo');
    return;
  }}

  // Estatísticas por dia
  const perDay  = datas.map(d => hist.filter(r=>r.data===d).length);
  const mean    = perDay.reduce((s,n)=>s+n,0) / nDias;
  const stdDev  = Math.sqrt(perDay.reduce((s,n)=>s+Math.pow(n-mean,2),0)/nDias);
  const minExp  = Math.max(0, Math.round(mean-stdDev));
  const maxExp  = Math.round(mean+stdDev);

  // Turno atual e próximo
  const h = now.getHours();
  const turnoAtual   = h>=6&&h<=11?'Manhã':h>=12&&h<=17?'Tarde':h>=18&&h<=23?'Noite':'Madrugada';
  const PROX_TURNO   = {{Madrugada:'Manhã',Manhã:'Tarde',Tarde:'Noite',Noite:'Madrugada'}};
  const turnoProximo = PROX_TURNO[turnoAtual];
  const ORDEM_TURNO  = ['Madrugada','Manhã','Tarde','Noite'];
  const turnoCnt     = {{}};
  ORDEM_TURNO.forEach(t => turnoCnt[t] = hist.filter(r=>r.turno===t).length);
  const totalTurnos  = hist.length || 1;
  const topTurno     = ORDEM_TURNO.reduce((b,t)=>turnoCnt[t]>turnoCnt[b]?t:b, ORDEM_TURNO[0]);

  // Contagens
  const cntF = (arr,f) => {{
    const m={{}};
    arr.forEach(r=>{{ if(r[f]) m[r[f]]=(m[r[f]]||0)+1; }});
    return Object.entries(m).sort((a,b)=>b[1]-a[1]);
  }};
  const bairros = cntF(hist,'bairro');
  const tipos   = cntF(hist,'tipo');
  const ruas    = cntF(hist,'endereco');
  const maxB    = bairros[0]?bairros[0][1]:1;

  // Tendência: comparar metade inicial vs. final
  let trendStr='Estável', trendColor='#555', trendIcon='→', trendDiff=0;
  if (datas.length >= 4) {{
    const mid  = Math.floor(datas.length/2);
    const avg1 = datas.slice(0,mid).reduce((s,d)=>s+hist.filter(r=>r.data===d).length,0)/mid;
    const avg2 = datas.slice(mid).reduce((s,d)=>s+hist.filter(r=>r.data===d).length,0)/(datas.length-mid);
    trendDiff  = avg2-avg1;
    if (trendDiff>1.5)       {{ trendStr='Crescente';  trendColor='#D13438'; trendIcon='↗'; }}
    else if (trendDiff<-1.5) {{ trendStr='Decrescente';trendColor='#107C10'; trendIcon='↘'; }}
  }} else {{
    trendStr='Poucos dados'; trendColor='#888'; trendIcon='–';
  }}

  // Nível de risco geral
  let riskScore = 0;
  if (mean > 8) riskScore+=2; else if (mean>4) riskScore+=1;
  if (trendIcon==='↗') riskScore+=2;
  if (turnoCnt[turnoAtual]/totalTurnos > 0.35) riskScore+=1;
  const riskLevel = riskScore>=4?'ALTO':riskScore>=2?'MÉDIO':'BAIXO';
  const riskCl    = riskLevel==='ALTO'?'risco-alto':riskLevel==='MÉDIO'?'risco-medio':'risco-baixo';
  const riskBg    = riskLevel==='ALTO'?'#fff0f0':riskLevel==='MÉDIO'?'#fffbe6':'#f0fff4';
  const riskColor = riskLevel==='ALTO'?'#D13438':riskLevel==='MÉDIO'?'#E07B00':'#107C10';

  // HTML dos turnos com barra de risco
  const turnoRowsHtml = ORDEM_TURNO.map(t => {{
    const n   = turnoCnt[t]||0;
    const pct = Math.round(n/totalTurnos*100);
    const isAtual  = t===turnoAtual;
    const isProx   = t===turnoProximo;
    const barColor = isAtual?'#E07B00':isProx?'#5C2D91':'#0078D4';
    const rowBg    = isAtual?'background:#fff3cd;font-weight:700':isProx?'background:#f5f0ff':'';
    return `<tr style="${{rowBg}}">
      <td>${{isAtual?'▶ ':isProx?'⏱ ':''}}${{t}}</td>
      <td style="text-align:center">${{n}}</td>
      <td style="text-align:center">${{pct}}%</td>
      <td><div class="risco-bar-wrap"><div class="risco-bar-fill"
        style="background:${{barColor}};width:${{pct}}%"></div></div></td>
      <td style="font-size:10px">${{isAtual?'🔴 AGORA':isProx?'⏳ Próximo':''}}</td>
    </tr>`;
  }}).join('');

  // HTML dos bairros com badge de risco
  const bairroRowsHtml = bairros.slice(0,8).map((e,i) => {{
    const score = e[1]/maxB;
    const nivel = score>0.6?'ALTO':score>0.3?'MÉDIO':'BAIXO';
    const cl    = score>0.6?'risco-alto':score>0.3?'risco-medio':'risco-baixo';
    return `<tr>
      <td>${{i+1}}. ${{e[0]}}</td>
      <td style="text-align:center">${{e[1]}}</td>
      <td><span class="${{cl}}">${{nivel}}</span></td>
    </tr>`;
  }}).join('');

  // HTML dos tipos
  const tipoRowsHtml = tipos.slice(0,6).map((e,i) => {{
    const pct = hist.length?Math.round(e[1]/hist.length*100):0;
    const bar = `<div class="risco-bar-wrap"><div class="risco-bar-fill"
      style="background:#D13438;width:${{pct}}%"></div></div>`;
    return `<tr><td>${{i+1}}. ${{e[0]}}</td><td>${{e[1]}}</td><td>${{pct}}%</td><td>${{bar}}</td></tr>`;
  }}).join('');

  // Tendência por data (mini-histórico)
  const tendRowsHtml = datas.map(d => {{
    const n   = hist.filter(r=>r.data===d).length;
    const pts = d.split('-');
    const fmt = pts[2]+'/'+pts[1];
    const pct = maxExp>0?Math.round(n/maxExp*100):0;
    const cl  = n>mean+stdDev?'risco-alto':n>mean?'risco-medio':'risco-baixo';
    return `<tr>
      <td>${{fmt}}</td>
      <td style="text-align:center">${{n}}</td>
      <td><div class="risco-bar-wrap" style="width:100px">
        <div class="risco-bar-fill" style="background:#0078D4;width:${{pct}}%"></div>
      </div></td>
      <td><span class="${{cl}}" style="font-size:9px">${{n>mean+stdDev?'ACIMA':n<mean-stdDev?'ABAIXO':'NORMAL'}}</span></td>
    </tr>`;
  }}).join('');

  // Orientações preventivas
  const ori = [];
  if (bairros[0]&&bairros[0][0]!=='–')
    ori.push(`Reforçar guarnição no <strong>${{bairros[0][0]}}</strong> — bairro historicamente mais afetado em ${{plural}}.`);
  if (ruas[0]&&ruas[0][0]!=='–')
    ori.push(`Estabelecer ronda na <strong>${{ruas[0][0]}}</strong> — logradouro de maior risco em ${{plural}}.`);
  const oriTurno={{
    Madrugada:`Atenção especial ao turno da <strong>Madrugada (00h–05h)</strong> — concentra ${{Math.round(turnoCnt.Madrugada/totalTurnos*100)}}% das ocorrências em ${{plural}}.`,
    'Manhã':  `Turno <strong>Matutino (06h–11h)</strong> é crítico em ${{plural}} (${{Math.round(turnoCnt['Manhã']/totalTurnos*100)}}%). Ampliar presença nos bairros de risco.`,
    Tarde:    `Período <strong>Vespertino (12h–17h)</strong> de alto risco em ${{plural}} (${{Math.round(turnoCnt.Tarde/totalTurnos*100)}}%). Atenção à movimentação urbana.`,
    Noite:    `Turno <strong>Noturno (18h–23h)</strong> concentra ${{Math.round(turnoCnt.Noite/totalTurnos*100)}}% dos crimes em ${{plural}}. Coordenar blitz nas vias principais.`
  }};
  if (oriTurno[topTurno]) ori.push(oriTurno[topTurno]);
  if (tipos[0]) ori.push(`Crime mais frequente em ${{plural}}: <strong>${{tipos[0][0]}}</strong>. Orientar guarnições com abordagem preventiva específica.`);
  if (trendIcon==='↗') ori.push(`⚠️ <strong>Tendência crescente detectada</strong> em ${{plural}} (+${{trendDiff.toFixed(1)}} oc./semana). Considerar reforço de efetivo.`);
  ori.push(`Compartilhar esta previsão com todas as guarnições antes do início do turno.`);

  const mkTable = (rows,headers) =>
    `<table class="analise-table"><thead><tr>${{headers.map(h=>`<th>${{h}}</th>`).join('')}}</tr></thead><tbody>${{rows}}</tbody></table>`;

  const corpo = `
  <div style="background:${{riskBg}};border:2px solid ${{riskColor}};border-radius:8px;
    padding:14px 16px;margin-bottom:14px;">
    <div style="font-size:10px;color:#555;font-weight:700;text-transform:uppercase;margin-bottom:8px">
      Previsão para hoje — ${{nowDayBR}}, ${{nowDateBR}} | Gerado às ${{horaBR}}
    </div>
    <div style="display:grid;grid-template-columns:auto 1fr 1fr 1fr;gap:16px;align-items:center">
      <div style="text-align:center">
        <div style="font-size:9px;color:#555;text-transform:uppercase;margin-bottom:4px">Risco geral</div>
        <span class="${{riskCl}}" style="font-size:16px;padding:5px 14px">${{riskLevel}}</span>
      </div>
      <div style="text-align:center">
        <div style="font-size:9px;color:#555;text-transform:uppercase">Previsão</div>
        <div style="font-size:20px;font-weight:800;color:#1A1A2E">${{minExp}}–${{maxExp}}</div>
        <div style="font-size:9px;color:#888">ocorrências</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:9px;color:#555;text-transform:uppercase">Tendência</div>
        <div style="font-size:18px;font-weight:800;color:${{trendColor}}">${{trendIcon}}</div>
        <div style="font-size:10px;color:${{trendColor}};font-weight:600">${{trendStr}}</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:9px;color:#555;text-transform:uppercase">Base histórica</div>
        <div style="font-size:20px;font-weight:800;color:#1A1A2E">${{nDias}}</div>
        <div style="font-size:9px;color:#888">${{plural}} analisados</div>
      </div>
    </div>
  </div>

  <div class="analise-section">
    <h3>⏰ Risco por Turno em ${{plural}} — Turno atual: <span style="color:#E07B00">${{turnoAtual}}</span></h3>
    ${{mkTable(turnoRowsHtml,['Turno','Total','%','Intensidade',''])}}
  </div>

  <div style="display:grid;grid-template-columns:1.3fr 1fr;gap:12px">
    <div class="analise-section">
      <h3>📍 Bairros em Alerta em ${{plural}}</h3>
      ${{mkTable(bairroRowsHtml,['Bairro','Ocorrências','Risco'])}}
    </div>
    <div class="analise-section">
      <h3>🔴 Crimes Mais Prováveis</h3>
      ${{mkTable(tipoRowsHtml,['Tipo','Total','%',''])}}
    </div>
  </div>

  <div class="analise-section">
    <h3>📊 Histórico de ${{plural}} — Ocorrências por dia</h3>
    ${{mkTable(tendRowsHtml,['Data','Qtd','Volume','Status'])}}
    <div style="font-size:10px;color:#888;margin-top:6px">
      Média: <strong>${{mean.toFixed(1)}}</strong> oc./dia &nbsp;|&nbsp;
      Intervalo esperado hoje: <strong>${{minExp}}–${{maxExp}}</strong> ocorrências
    </div>
  </div>

  <div class="analise-section">
    <h3>🎯 Orientações Preventivas para Hoje</h3>
    <ol class="analise-rec-ol">${{ori.map(o=>`<li>${{o}}</li>`).join('')}}</ol>
  </div>

  <div style="border-top:1px solid #ddd;margin-top:12px;padding-top:8px;color:#888;font-size:10px;text-align:center">
    Secretaria de Segurança e Ordem Pública de Balneário Camboriú — Guarda Municipal
  </div>`;

  const waLnsPrev = [
    `📈 *PREVISÃO DE RISCO — ${{plural.toUpperCase()}} — GUARDA MUNICIPAL BC*`,
    `📅 ${{nowDayBR}}, ${{nowDateBR}} | Gerado às ${{horaBR}}`,
    ``,
    `*🚨 NÍVEL DE RISCO GERAL: ${{riskLevel}}*`,
    `Previsão: *${{minExp}}–${{maxExp}}* ocorrências esperadas`,
    `Tendência: *${{trendStr}} ${{trendIcon}}*`,
    `Base histórica: *${{nDias}}* ${{plural}} analisados`,
    ``,
    `*⏰ RISCO POR TURNO*`,
    ...ORDEM_TURNO.map(t=>{{
      const n=turnoCnt[t]||0;
      const pct=Math.round(n/totalTurnos*100);
      const mark=t===turnoAtual?' ◀ AGORA':t===turnoProximo?' ⏳ Próximo':'';
      return `• ${{t}}: ${{n}} oc. (${{pct}}%)${{mark}}`;
    }}),
    ``,
    `*📍 BAIRROS EM ALERTA*`,
    ...bairros.slice(0,5).map((e,i)=>{{
      const score=e[1]/maxB;
      const nivel=score>0.6?'🔴 ALTO':score>0.3?'🟡 MÉDIO':'🟢 BAIXO';
      return `${{i+1}}. ${{e[0]}}: ${{e[1]}} oc. — ${{nivel}}`;
    }}),
    ``,
    `*🔴 CRIMES MAIS PROVÁVEIS*`,
    ...tipos.slice(0,5).map((e,i)=>{{
      const pct=hist.length?Math.round(e[1]/hist.length*100):0;
      return `${{i+1}}. ${{e[0]}}: ${{e[1]}} (${{pct}}%)`;
    }}),
    ``,
    `*🎯 ORIENTAÇÕES PREVENTIVAS*`,
    ...ori.map((o,i)=>`${{i+1}}. ${{o.replace(/<[^>]+>/g,'')}}`),
    ``,
    `_Guarda Municipal de Balneário Camboriú_`,
    `_Secretaria de Segurança e Ordem Pública_`
  ];
  _prevTextoWA = waLnsPrev.join('\\n');

  document.getElementById('prev-corpo').innerHTML = corpo;
  document.getElementById('prev-overlay').classList.add('ativo');
}}

function fecharPrevisao() {{
  document.getElementById('prev-overlay').classList.remove('ativo');
}}

// ── PESQUISA GERAL ────────────────────────────────────────────────────────────
const INTEL_DATA = {INTEL_JSON};
const RELATO_INDEX = INTEL_DATA ? INTEL_DATA.relato_index || [] : [];

let _pesquisaTimer = null;
function onPesquisaInput(val) {{
  clearTimeout(_pesquisaTimer);
  _pesquisaTimer = setTimeout(()=>{{ if(val.trim().length >= 2) executarPesquisa(); }}, 400);
}}
function limparPesquisa() {{
  pesquisaQ = '';
  document.getElementById('pesquisa-input').value = '';
  document.getElementById('pesquisa-limpar').style.display = 'none';
  const badge = document.getElementById('pesquisa-badge');
  if(badge) {{ badge.style.display='none'; badge.textContent=''; }}
  renderAll(filtered());
}}
function fecharPesquisa() {{
  document.getElementById('pesquisa-overlay').classList.remove('ativo');
}}
function executarPesquisa() {{
  const termo = document.getElementById('pesquisa-input').value.trim();
  if(!termo) return;
  pesquisaQ = termo;
  document.getElementById('pesquisa-limpar').style.display = '';

  const termoLow = termo.toLowerCase();

  // ── 1. Resultados no dashboard (dados estruturados) ───────────────────────
  const dashResult = filtered();
  const ocsDash    = dedupBO(dashResult);

  // ── 2. Resultados nos relatos dos B.O.s ───────────────────────────────────
  const boResult = RELATO_INDEX.filter(b => {{
    const texto = [b.relato,b.local,b.bairro,b.tipo,b.numero].join(' ').toLowerCase();
    return texto.includes(termoLow);
  }});

  // ── Atualiza gráficos ─────────────────────────────────────────────────────
  renderAll(dashResult);

  // ── Monta conteúdo do modal ───────────────────────────────────────────────
  function highlight(txt, q) {{
    if(!txt||!q) return txt||'';
    const re = new RegExp('('+q.replace(/[.*+?^${{}}()|[\\]\\\\]/g,'\\\\$&')+')','gi');
    return txt.replace(re,'<mark style="background:#FFE066;padding:0 2px;border-radius:2px">$1</mark>');
  }}

  const dashHtml = ocsDash.length > 0 ? `
    <div style="font-size:11px;color:#2E6DA4;font-weight:600;margin-bottom:6px">
      ✅ ${{ocsDash.length}} ocorrência(s) encontrada(s) nos dados do dashboard — gráficos atualizados
    </div>` : `
    <div style="font-size:11px;color:#888;margin-bottom:6px">
      ℹ️ Nenhuma correspondência nos dados estruturados do dashboard
    </div>`;

  const boHtml = boResult.length > 0 ? `
    <div style="font-size:12px;font-weight:700;color:#1A3A5C;border-bottom:2px solid #2E6DA4;padding-bottom:5px;margin:14px 0 10px">
      📄 ${{boResult.length}} B.O.(s) com menção nos relatos
    </div>
    ${{boResult.map(b=>{{
      const trecho = b.relato ? highlight(b.relato.substring(0,400), termo) : '—';
      return `<div style="border:1px solid #E0E0E0;border-radius:8px;margin-bottom:10px;overflow:hidden">
        <div style="background:#F0F6FF;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:6px">
          <span style="font-size:11px;font-weight:700;color:#1A3A5C">${{b.numero}}</span>
          <span style="font-size:10px;color:#555">${{b.data}} &nbsp;•&nbsp; ${{b.hora||'?'}} — ${{b.turno}}</span>
          <span style="font-size:10px;color:#555">${{b.local||b.bairro||'—'}}</span>
          ${{b.tipo?`<span style="background:#E07B00;color:white;padding:1px 7px;border-radius:4px;font-size:10px;font-weight:700">${{b.tipo}}</span>`:''}}
        </div>
        <div style="padding:8px 12px;font-size:11px;color:#333;line-height:1.6;background:white">
          ${{trecho}}${{b.relato&&b.relato.length>=400?'<span style="color:#aaa"> [...]</span>':''}}
        </div>
      </div>`;
    }}).join('')}}` : `
    <div style="font-size:11px;color:#888;margin-top:10px">
      ℹ️ Nenhum B.O. com menção a "<strong>${{termo}}</strong>" nos relatos
      ${{RELATO_INDEX.length===0?'<br><span style="color:#E07B00">⚠️ Execute INTELIGENCIA.bat para indexar os B.O.s</span>':''}}
    </div>`;

  document.getElementById('pesquisa-subtitulo').textContent =
    `"${{termo}}" — ${{ocsDash.length}} no dashboard · ${{boResult.length}} nos B.O.s`;
  document.getElementById('pesquisa-corpo').innerHTML = dashHtml + boHtml;
  document.getElementById('pesquisa-overlay').classList.add('ativo');
  // Badge na sidebar
  const badge = document.getElementById('pesquisa-badge');
  if(badge) {{
    badge.style.display = '';
    badge.textContent = `✅ ${{ocsDash.length}} no dash · ${{boResult.length}} nos B.O.s`;
  }}
}}

// ── INTELIGÊNCIA CRIMINAL ─────────────────────────────────────────────────────

function fecharInteligencia() {{
  document.getElementById('intel-overlay').classList.remove('ativo');
}}
function imprimirInteligencia() {{
  const c = document.getElementById('intel-corpo').innerHTML;
  const w = window.open('','_blank','width=1000,height=760');
  w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Inteligência Criminal — GMBC</title>
    <style>body{{font-family:'Segoe UI',Arial,sans-serif;padding:20px;font-size:11px}}
    table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #ddd;padding:5px 8px}}
    th{{background:#1B4332;color:white}} .badge{{border-radius:4px;padding:2px 7px;font-weight:700;font-size:10px}}
    </style></head><body>${{c}}</body></html>`);
  w.document.close(); w.print();
}}
function abrirInteligencia() {{
  if(!INTEL_DATA) {{
    alert('Dados de inteligência não disponíveis.\\n\\nExecute no seu computador:\\n1. INTELIGENCIA.bat (duplo clique)\\n2. atualizar.ps1 (duplo clique)\\n\\nIsso processa os B.O.s e publica os dados no dashboard.');
    return;
  }}
  const d = INTEL_DATA;
  const RISCO_COR = {{'CRÍTICO':'#D13438','ALTO':'#E07B00','MÉDIO':'#E6A817'}};
  const CRIME_COR = {{'Roubo':'#D13438','Furto':'#E07B00'}};
  function badge(t,c){{ return `<span class="badge" style="background:${{c}};color:white">${{t}}</span>`; }}
  function corCrime(t){{ for(const[k,v] of Object.entries(CRIME_COR)) if(t&&t.includes(k)) return v; return '#555'; }}

  // KPIs
  const kpis = `
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:16px">
    <div style="background:#F0FFF4;border-left:4px solid #2D6A4F;border-radius:8px;padding:12px 14px">
      <div style="font-size:24px;font-weight:700;color:#2D6A4F">${{d.total_bos}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">B.O.s analisados</div>
    </div>
    <div style="background:#FFF5F5;border-left:4px solid #D13438;border-radius:8px;padding:12px 14px">
      <div style="font-size:24px;font-weight:700;color:#D13438">${{d.total_grupos}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Grupos com padrão similar</div>
    </div>
    <div style="background:#FFF8EC;border-left:4px solid #E07B00;border-radius:8px;padding:12px 14px">
      <div style="font-size:24px;font-weight:700;color:#E07B00">${{d.total_vinculados}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Ocorrências vinculadas</div>
    </div>
    <div style="background:#F8F8FF;border-left:4px solid #555;border-radius:8px;padding:12px 14px">
      <div style="font-size:24px;font-weight:700;color:#555">${{d.total_bos - d.total_vinculados}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Sem vínculo identificado</div>
    </div>
  </div>`;

  // Bairros + Turnos
  const maxB = Math.max(...Object.values(d.bairros_freq),1);
  const bairrosBar = Object.entries(d.bairros_freq).map(([b,c])=>{{
    const pct=Math.round(c/maxB*100);
    return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
      <span style="width:110px;font-size:10px;text-align:right;color:#333;flex-shrink:0">${{b}}</span>
      <div style="flex:1;background:#EEE;border-radius:3px;height:16px">
        <div style="width:${{pct}}%;background:#2D6A4F;height:16px;border-radius:3px"></div>
      </div>
      <span style="width:24px;font-size:10px;font-weight:700;color:#2D6A4F">${{c}}</span>
    </div>`;
  }}).join('');
  const TCOR={{'Madrugada':'#7B2FBE','Manhã':'#0078D4','Tarde':'#E07B00','Noite':'#1A1A2E'}};
  const turnosHtml = Object.entries(d.turnos_freq).sort((a,b)=>b[1]-a[1]).map(([t,c])=>
    `<div style="display:flex;justify-content:space-between;padding:5px 8px;font-size:11px;border-bottom:1px solid #F0F0F0">
      <span style="color:${{TCOR[t]||'#333'}};font-weight:600">${{t}}</span>
      <span style="font-weight:700">${{c}}</span>
    </div>`).join('');

  // Grupos
  const gruposHtml = d.grupos.map(g=>{{
    const cor = RISCO_COR[g.risco]||'#555';
    const padrao = [
      g.n_suspeitos>0?`👤 ${{g.n_suspeitos}} suspeito(s)`:'',
      g.armado?'🔫 Armado(s)':'',
      g.veiculo?`🏍️ Fuga: ${{g.veiculo}}`:'',
      ...(g.mo_tags||[]).slice(0,3).map(t=>`• ${{t}}`),
    ].filter(Boolean).join(' &nbsp; ');
    const linhas = g.crimes.map((c,i)=>{{
      const rzs = (c.razoes||[]).join(' · ')||'referência';
      return `<tr style="background:${{i%2?'#FAFAFA':'white'}}">
        <td style="padding:4px 7px;font-size:10px;font-weight:600;color:#1B4332;white-space:nowrap">${{c.numero||c.arquivo||'—'}}</td>
        <td style="padding:4px 7px;font-size:10px;white-space:nowrap">${{c.data}}</td>
        <td style="padding:4px 7px;font-size:10px;white-space:nowrap">${{c.hora||'?'}} — ${{c.turno}}</td>
        <td style="padding:4px 7px;font-size:10px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{c.local||'—'}}</td>
        <td style="padding:4px 7px">${{c.tipo?badge(c.tipo,corCrime(c.tipo)):'—'}}</td>
      </tr>`;
    }}).join('');
    const dInicio = (g.datas||[])[0]||'';
    const dFim    = (g.datas||[]).slice(-1)[0]||'';
    return `
    <div style="border:1px solid #E0E0E0;border-radius:8px;margin-bottom:16px;overflow:hidden">
      <div style="background:${{cor}};padding:10px 14px;display:flex;justify-content:space-between;align-items:center">
        <span style="color:white;font-size:13px;font-weight:700">Grupo #${{g.id}} — ${{g.n}} ocorrência(s) &nbsp;
          <span style="background:rgba(0,0,0,.2);padding:2px 8px;border-radius:4px;font-size:10px">${{g.risco}}</span>
        </span>
        <span style="color:rgba(255,255,255,.85);font-size:10px">${{dInicio}}${{dFim&&dFim!==dInicio?' → '+dFim:''}}</span>
      </div>
      <div style="padding:10px 14px;background:#FAFAFA;font-size:11px;line-height:1.8">
        ${{padrao?'<div><strong>🧩 Padrão:</strong> &nbsp;'+padrao+'</div>':''}}
        ${{g.objetos&&g.objetos.length?'<div><strong>📦 Objetos:</strong> '+g.objetos.join(' · ')+'</div>':''}}
        ${{g.bairros&&g.bairros.length?'<div><strong>📍 Bairros:</strong> '+g.bairros.join(', ')+'</div>':''}}
      </div>
      <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#F0F0F0">
          <th style="padding:5px 7px;text-align:left;font-size:9px">Nº B.O.</th>
          <th style="padding:5px 7px;text-align:left;font-size:9px">Data</th>
          <th style="padding:5px 7px;text-align:left;font-size:9px">Hora/Turno</th>
          <th style="padding:5px 7px;text-align:left;font-size:9px">Local</th>
          <th style="padding:5px 7px;text-align:left;font-size:9px">Crime</th>
        </tr></thead>
        <tbody>${{linhas}}</tbody>
      </table>
      </div>
    </div>`;
  }}).join('');

  document.getElementById('intel-corpo').innerHTML = `
  <div style="font-size:10px;color:#888;margin-bottom:14px">
    Gerado em: ${{d.gerado_em}} &nbsp;•&nbsp; Base: ${{d.total_bos}} B.O.s processados
    &nbsp;•&nbsp; <span style="color:#2D6A4F;font-weight:700">Análise de Padrões de MO / Suspeitos</span>
  </div>
  <div style="background:#FFF8EC;border-left:4px solid #E07B00;padding:8px 12px;border-radius:4px;font-size:11px;color:#555;margin-bottom:14px">
    ⚠️ As vinculações são baseadas em similaridade de padrões — não constituem prova de autoria.
    Ferramenta de apoio ao planejamento operacional e investigativo.
  </div>
  ${{kpis}}
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px">
    <div>
      <div style="font-size:12px;font-weight:700;border-bottom:2px solid #2D6A4F;padding-bottom:4px;margin-bottom:8px">📍 Bairros com Mais Ocorrências</div>
      ${{bairrosBar}}
    </div>
    <div>
      <div style="font-size:12px;font-weight:700;border-bottom:2px solid #2D6A4F;padding-bottom:4px;margin-bottom:8px">⏰ Distribuição por Turno</div>
      ${{turnosHtml}}
    </div>
  </div>
  <div style="font-size:12px;font-weight:700;border-bottom:2px solid #2D6A4F;padding-bottom:4px;margin-bottom:12px">🔗 Grupos com Padrão Similar</div>
  ${{gruposHtml}}`;

  document.getElementById('intel-overlay').classList.add('ativo');
}}

// ── ANÁLISE PREDITIVA CRIMINAL ────────────────────────────────────────────────
function fecharAnalisePredit() {{
  document.getElementById('predit-overlay').classList.remove('ativo');
}}
function imprimirAnalisePredit() {{
  const corpo = document.getElementById('predit-corpo').innerHTML;
  const w = window.open('','_blank','width=1000,height=760');
  w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8">
    <title>Análise Preditiva — Guarda Municipal BC</title>
    <style>body{{font-family:'Segoe UI',Arial,sans-serif;padding:24px;font-size:12px;}}
    table{{border-collapse:collapse;width:100%;}} th,td{{border:1px solid #ddd;padding:6px 10px;}}
    th{{background:#2D0B6E;color:white;}} .badge{{border-radius:4px;padding:2px 7px;font-weight:700;font-size:11px;}}
    </style></head><body>${{corpo}}</body></html>`);
  w.document.close(); w.print();
}}
function abrirAnalisePredit() {{
  const RAWU = dedupBO(RAW);
  const agora = new Date();
  const dataHoje = agora.toLocaleDateString('pt-BR');

  // ── Helpers ────────────────────────────────────────────────────────────────
  const DIAS_ORD = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo'];
  const TURNS    = ['Manhã','Tarde','Noite','Madrugada'];
  const MESES_PT = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

  function linReg(xs, ys) {{
    const n=xs.length; if(n<2) return {{m:0,b:ys[0]||0,predict:()=>ys[0]||0}};
    const mx=xs.reduce((a,b)=>a+b,0)/n, my=ys.reduce((a,b)=>a+b,0)/n;
    const num=xs.reduce((s,x,i)=>s+(x-mx)*(ys[i]-my),0);
    const den=xs.reduce((s,x)=>s+(x-mx)**2,0);
    const m=den?num/den:0, b=my-m*mx;
    return {{m,b,predict:x=>Math.max(0,m*x+b)}};
  }}
  function addMon(ym,n){{
    let [y,m]=ym.split('-').map(Number); m+=n;
    while(m>12){{m-=12;y++;}} while(m<1){{m+=12;y--;}}
    return y+'-'+(m<10?'0'+m:m);
  }}
  function ymLbl(ym){{
    const [y,mo]=ym.split('-');
    return MESES_PT[parseInt(mo)-1]+'/'+y.slice(2);
  }}
  function riscoBadge(r){{
    const cfg={{
      'Crítico':['#D13438','🔴'],
      'Alto':   ['#D13438','🔴'],
      'Médio':  ['#E6A817','🟡'],
      'Baixo':  ['#107C10','🟢'],
    }}; const [cor,ic]=cfg[r]||['#888','⚪'];
    return `<span class="badge" style="background:${{cor}};color:white">${{ic}} ${{r}}</span>`;
  }}

  // ── 1. Agregação mensal global ──────────────────────────────────────────────
  const monGlobal={{}};
  RAWU.forEach(r=>{{ if(r.data){{ const k=r.data.slice(0,7); monGlobal[k]=(monGlobal[k]||0)+1; }} }});
  const mesesGlobal=Object.keys(monGlobal).sort();
  const contsGlobal=mesesGlobal.map(m=>monGlobal[m]);
  const regG=linReg(mesesGlobal.map((_,i)=>i), contsGlobal);
  const ultimoMesChave=mesesGlobal[mesesGlobal.length-1]||'';
  const prev30 = Math.round(regG.predict(mesesGlobal.length));
  const prev60 = Math.round(regG.predict(mesesGlobal.length+1));
  const prev90 = Math.round(regG.predict(mesesGlobal.length+2));
  const tendPct = contsGlobal.length>=2 ? ((regG.m/(contsGlobal.reduce((a,b)=>a+b,0)/contsGlobal.length))*100).toFixed(1) : '0';
  const tendSeta = regG.m>0.5?'▲ Crescente':regG.m<-0.5?'▼ Decrescente':'→ Estável';
  const tendCor  = regG.m>0.5?'#D13438':regG.m<-0.5?'#107C10':'#555';

  // ── 2. Risco por bairro ────────────────────────────────────────────────────
  const ultMesesBairro = mesesGlobal.filter(m=>m.startsWith('2026')).slice(-5);
  const bairroSet=[...new Set(RAWU.map(r=>r.bairro).filter(Boolean))];
  const bairroStats=bairroSet.map(b=>{{
    const recs=RAWU.filter(r=>r.bairro===b);
    const mb={{}};
    recs.forEach(r=>{{ if(r.data){{const k=r.data.slice(0,7);mb[k]=(mb[k]||0)+1;}} }});
    const ms=Object.keys(mb).sort(), cs=ms.map(m=>mb[m]);
    const med=cs.length?cs.reduce((a,v)=>a+v,0)/cs.length:0;
    const std=cs.length>1?Math.sqrt(cs.reduce((s,v)=>s+(v-med)**2,0)/cs.length):0;
    const recente=mb[ultimoMesChave]||0;
    const z=std>0?(recente-med)/std:0;
    const pctVsMed=med>0?Math.round((recente-med)/med*100):0;
    const tipoFreq={{}};
    recs.forEach(r=>{{ if(r.tipo) tipoFreq[r.tipo]=(tipoFreq[r.tipo]||0)+1; }});
    const topTipo=Object.entries(tipoFreq).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—';
    let risco='Baixo';
    if(z>=1.5) risco='Crítico'; else if(z>=0.5) risco='Alto'; else if(z>=-0.5) risco='Médio';
    const monthCounts=ultMesesBairro.map(m=>mb[m]||0);
    return {{b,total:recs.length,recente,med:med.toFixed(1),z,risco,pctVsMed,topTipo,monthCounts}};
  }}).sort((a,b)=>b.total-a.total).slice(0,10);

  // ── 3. Período crítico (dia × turno) ──────────────────────────────────────
  const dtMap={{}};
  RAWU.forEach(r=>{{ if(r.dia&&r.turno){{ const k=r.dia+'|'+r.turno; dtMap[k]=(dtMap[k]||0)+1; }} }});
  const topPeriodo=Object.entries(dtMap).sort((a,b)=>b[1]-a[1])[0]||['—|—',0];
  const [tpDia,tpTurno]=topPeriodo[0].split('|');
  const totalOcs=RAWU.length||1;

  // ── 4. Alertas automáticos ─────────────────────────────────────────────────
  const now30 = new Date(agora); now30.setDate(now30.getDate()-30);
  const prev30d= new Date(agora); prev30d.setDate(prev30d.getDate()-60);
  const fmt=d=>d.toISOString().slice(0,10);
  const rec30=RAWU.filter(r=>r.data&&r.data>=fmt(now30));
  const ant30=RAWU.filter(r=>r.data&&r.data>=fmt(prev30d)&&r.data<fmt(now30));

  const alertas=[];
  function checkAlerta(label, valRec, valAnt){{
    if(valAnt===0) return;
    const delta=((valRec-valAnt)/valAnt*100).toFixed(0);
    if(delta>25) alertas.push({{ic:'🔴',cor:'#D13438',msg:`${{label}}: aumento de ${{delta}}% nos últimos 30 dias vs período anterior`,nivel:'Crítico'}});
    else if(delta>10) alertas.push({{ic:'🟠',cor:'#E07B00',msg:`${{label}}: aumento de ${{delta}}% nos últimos 30 dias`,nivel:'Alto'}});
    else if(delta<-15) alertas.push({{ic:'🟢',cor:'#107C10',msg:`${{label}}: queda de ${{Math.abs(delta)}}% — tendência positiva`,nivel:'Baixo'}});
  }}
  const tipos=[...new Set(RAWU.map(r=>r.tipo).filter(Boolean))];
  tipos.forEach(t=>checkAlerta('Tipo: '+t, rec30.filter(r=>r.tipo===t).length, ant30.filter(r=>r.tipo===t).length));
  bairroStats.slice(0,5).forEach(bs=>checkAlerta('Bairro: '+bs.b, rec30.filter(r=>r.bairro===bs.b).length, ant30.filter(r=>r.bairro===bs.b).length));
  checkAlerta('Total geral', rec30.length, ant30.length);
  if(alertas.length===0) alertas.push({{ic:'🟢',cor:'#107C10',msg:'Nenhuma variação crítica detectada no período recente',nivel:'Normal'}});

  // ── 5. Recomendações ───────────────────────────────────────────────────────
  const recs=[];
  const bCrit=bairroStats.filter(b=>b.risco==='Crítico'||b.risco==='Alto');
  if(bCrit.length) recs.push(`🚔 Reforçar policiamento ostensivo no(s) bairro(s): <b>${{bCrit.slice(0,3).map(b=>b.b).join(', ')}}</b>`);
  recs.push(`⏰ Ampliar rondas no turno <b>${{tpTurno}}</b> — período de maior concentração de ocorrências`);
  recs.push(`📅 Atenção especial nas <b>${{tpDia}}s</b> — dia com maior histórico de ocorrências`);
  const topTipoGlobal=tipos.map(t=>([t,RAWU.filter(r=>r.tipo===t).length])).sort((a,b)=>b[1]-a[1])[0];
  if(topTipoGlobal) recs.push(`🔍 Estratégia específica para combate a <b>${{topTipoGlobal[0]}}</b> — crime mais frequente (${{topTipoGlobal[1]}} ocorrências)`);
  if(regG.m>0.5) recs.push(`📈 Tendência crescente identificada — considerar reforço de efetivo preventivo`);
  recs.push(`🤝 Integrar ações com Polícia Militar, Polícia Civil e agentes de trânsito nas áreas críticas`);

  // ── Resumo para IA ────────────────────────────────────────────────────────
  const bairrosCriticosTxt = bCrit.slice(0,3).map(b=>`${{b.b}} (${{b.recente}} casos recentes, risco ${{b.risco}})`).join('; ');
  _preditResumoIA =
    `total de ${{totalOcs}} ocorrências analisadas; ` +
    `tendência geral: ${{tendSeta}} (${{tendPct}}% de variação média mensal); ` +
    `bairros de maior risco: ${{bairrosCriticosTxt || 'nenhum em nível crítico/alto'}}; ` +
    `período mais crítico: ${{tpDia}} no turno ${{tpTurno}}; ` +
    `tipo de ocorrência mais frequente: ${{topTipoGlobal ? topTipoGlobal[0]+' ('+topTipoGlobal[1]+' casos)' : 'não informado'}}; ` +
    `previsão para os próximos 30 dias: ~${{prev30}} ocorrências; ` +
    `alertas automáticos: ${{alertas.map(a=>a.msg).join('; ')}}.`;

  // ── 6. Projeção visual (barras CSS) ────────────────────────────────────────
  const ultMeses=mesesGlobal.slice(-4);
  const ultConts=ultMeses.map(m=>monGlobal[m]);
  const projMeses=[addMon(ultimoMesChave||'',1),addMon(ultimoMesChave||'',2),addMon(ultimoMesChave||'',3)];
  const projConts=[prev30,prev60,prev90];
  const maxBar=Math.max(...ultConts,...projConts,1);

  function barRow(lbl, val, cor, isProj){{
    const pct=Math.round(val/maxBar*100);
    const barStyle=isProj
      ?'opacity:.85;background:repeating-linear-gradient(45deg,'+cor+','+cor+' 4px,rgba(255,255,255,.35) 4px,rgba(255,255,255,.35) 8px)'
      :'background:'+cor;
    return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px">
      <span style="width:52px;font-size:10px;color:#555;text-align:right;flex-shrink:0">${{lbl}}</span>
      <div style="flex:1;background:#EBEBEB;border-radius:4px;height:22px;min-width:0">
        <div style="width:${{pct}}%;height:22px;border-radius:4px;${{barStyle}}"></div>
      </div>
      <span style="width:90px;flex-shrink:0;font-size:10px;font-weight:700;color:${{isProj?'#7B2FBE':'#0063B1'}}">${{val}}${{isProj?' proj.':''}}</span>
    </div>`;
  }}

  // ── Heatmap dia × turno (HTML) ─────────────────────────────────────────────
  const maxDT=Math.max(...DIAS_ORD.flatMap(d=>TURNS.map(t=>dtMap[d+'|'+t]||0)),1);
  function heatCell(val){{
    const p=val/maxDT;
    const bg=p>=0.7?'#D13438':p>=0.45?'#E07B00':p>=0.2?'#FFD966':'#E8F4FD';
    const fg=p>=0.45?'white':'#333';
    const pctTot=Math.round(val/totalOcs*100);
    return `<td style="text-align:center;padding:5px 4px;background:${{bg}};color:${{fg}};font-size:10px;font-weight:${{val>0?'700':'400'}};border-radius:3px">${{val>0?val+' <span style=\\'font-size:8px;opacity:.8\\'>'+pctTot+'%</span>':'—'}}</td>`;
  }}
  const heatRows=DIAS_ORD.map(d=>`<tr><td style="padding:5px 10px;font-size:10px;font-weight:600;white-space:nowrap">${{d}}</td>
    ${{TURNS.map(t=>heatCell(dtMap[d+'|'+t]||0)).join('')}}</tr>`).join('');

  // ── Monta HTML do modal ────────────────────────────────────────────────────
  document.getElementById('predit-corpo').innerHTML = `
  <div style="font-size:10px;color:#888;margin-bottom:14px">
    Gerado em: ${{dataHoje}} às ${{agora.toLocaleTimeString('pt-BR',{{hour:'2-digit',minute:'2-digit'}})}}
    &nbsp;•&nbsp; Base: ${{RAWU.length}} ocorrências (deduplicadas por B.O.)
    &nbsp;•&nbsp; <span style="color:#7B2FBE;font-weight:700">IA • Regressão Linear + Análise Estatística</span>
  </div>

  <!-- Explicar com IA generativa -->
  <div style="margin-bottom:16px">
    <button id="btn-explicar-ia" class="btn-analise" data-label="🤖 Explicar com IA"
      onclick="explicarComIA(_preditResumoIA,'predit-ia-resultado','btn-explicar-ia')" style="background:#0078D4">🤖 Explicar com IA</button>
    <div id="predit-ia-resultado" style="display:none;margin-top:10px;background:#F0F6FC;border-left:4px solid #0078D4;
      border-radius:8px;padding:12px 14px;font-size:12px;text-align:justify"></div>
  </div>

  <!-- KPIs -->
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:18px">
    <div style="background:#F5F0FF;border-left:4px solid #7B2FBE;border-radius:8px;padding:12px 14px">
      <div style="font-size:22px;font-weight:700;color:#7B2FBE">~${{prev30}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Previsão próximos 30 dias</div>
      <div style="font-size:10px;color:#888">60d: ~${{prev60}} &nbsp;•&nbsp; 90d: ~${{prev90}}</div>
    </div>
    <div style="background:#FFF5F5;border-left:4px solid #D13438;border-radius:8px;padding:12px 14px">
      <div style="font-size:16px;font-weight:700;color:#D13438;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{bairroStats[0]?.b||'—'}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Bairro mais crítico</div>
      <div style="font-size:10px;color:#888">${{bairroStats[0]?.total||0}} ocorrências históricas</div>
    </div>
    <div style="background:#FFF8EC;border-left:4px solid #E07B00;border-radius:8px;padding:12px 14px">
      <div style="font-size:16px;font-weight:700;color:#E07B00">${{tpDia}} / ${{tpTurno}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Período de maior risco</div>
      <div style="font-size:10px;color:#888">${{topPeriodo[1]}} ocorrências históricas (${{Math.round(topPeriodo[1]/totalOcs*100)}}% do total)</div>
    </div>
    <div style="background:#F0FFF0;border-left:4px solid ${{tendCor}};border-radius:8px;padding:12px 14px">
      <div style="font-size:18px;font-weight:700;color:${{tendCor}}">${{tendSeta}}</div>
      <div style="font-size:11px;color:#555;margin-top:2px">Tendência geral</div>
      <div style="font-size:10px;color:#888">${{Math.abs(tendPct)}}% ao mês</div>
    </div>
  </div>

  <!-- Alertas -->
  <div style="margin-bottom:18px">
    <div style="font-size:12px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #7B2FBE;padding-bottom:5px;margin-bottom:10px">⚠️ Alertas Automáticos</div>
    ${{alertas.slice(0,6).map(a=>`
    <div style="display:flex;align-items:center;gap:8px;padding:8px 10px;background:#FAFAFA;border-left:3px solid ${{a.cor}};border-radius:4px;margin-bottom:5px;font-size:11px">
      <span style="font-size:14px">${{a.ic}}</span>
      <span style="flex:1">${{a.msg}}</span>
      <span class="badge" style="background:${{a.cor}};color:white;white-space:nowrap">${{a.nivel}}</span>
    </div>`).join('')}}
  </div>

  <!-- Tabela bairros largura total com evolução mensal -->
  <div style="margin-bottom:18px">
    <div style="font-size:12px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #7B2FBE;padding-bottom:5px;margin-bottom:8px">📊 Classificação de Risco por Bairro — Evolução Mensal</div>
    <div style="overflow-x:auto">
    <table style="width:100%;border-collapse:collapse;font-size:10px;min-width:560px">
      <thead><tr style="background:#F0EAF8">
        <th style="padding:5px 8px;text-align:left">Bairro</th>
        <th style="padding:5px 8px;text-align:center">Risco</th>
        ${{ultMesesBairro.map((m,i)=>`<th style="padding:5px 6px;text-align:center;color:#555;white-space:nowrap">${{ymLbl(m)}}</th>`).join('')}}
        <th style="padding:5px 8px;text-align:center">Total</th>
      </tr></thead>
      <tbody>
      ${{bairroStats.map((bs,i)=>{{
        const mCells=bs.monthCounts.map((v,j)=>{{
          if(j===0) return `<td style="padding:4px 6px;text-align:center;color:#555">${{v||'—'}}</td>`;
          const prev=bs.monthCounts[j-1];
          const delta=v-prev;
          const cor=delta>0?'#D13438':delta<0?'#107C10':'#888';
          const seta=delta>0?'▲':delta<0?'▼':'=';
          return `<td style="padding:4px 6px;text-align:center;font-weight:${{delta!==0?'700':'400'}};color:${{cor}}">${{v}} <span style="font-size:9px">${{seta}}</span></td>`;
        }}).join('');
        return `<tr style="background:${{i%2?'#FAFAFA':'white'}}">
          <td style="padding:4px 8px;font-weight:600;white-space:nowrap">${{bs.b}}</td>
          <td style="padding:4px 8px;text-align:center">${{riscoBadge(bs.risco)}}</td>
          ${{mCells}}
          <td style="padding:4px 8px;text-align:center;font-weight:700">${{bs.total}}</td>
        </tr>`;
      }}).join('')}}
      </tbody>
    </table>
    </div>
    <div style="margin-top:5px;font-size:9px;color:#999">
      ▲ aumento vs mês anterior &nbsp;•&nbsp; ▼ queda vs mês anterior &nbsp;•&nbsp; 1ª coluna sem comparativo
    </div>
  </div>

  <!-- Heatmap + Projeção lado a lado -->
  <div style="display:grid;grid-template-columns:auto 1fr;gap:20px;margin-bottom:18px;align-items:start">

    <!-- Heatmap -->
    <div>
      <div style="font-size:12px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #7B2FBE;padding-bottom:5px;margin-bottom:8px">🗓️ Mapa de Risco — Dia × Turno</div>
      <table style="border-collapse:separate;border-spacing:3px;font-size:10px">
        <thead><tr>
          <th style="padding:4px;text-align:left;font-size:10px"></th>
          ${{TURNS.map(t=>`<th style="padding:4px 8px;text-align:center;font-size:10px;color:#555">${{t}}</th>`).join('')}}
        </tr></thead>
        <tbody>${{heatRows}}</tbody>
      </table>
      <div style="margin-top:6px;font-size:9px;color:#777">
        🟥 Alto &nbsp; 🟧 Médio-alto &nbsp; 🟨 Médio &nbsp; 🟦 Baixo
      </div>
    </div>

    <!-- Projeção -->
    <div>
      <div style="font-size:12px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #7B2FBE;padding-bottom:5px;margin-bottom:8px">📈 Projeção — Histórico + Previsão 90 Dias</div>
      <div style="font-size:10px;color:#888;margin-bottom:8px">Histórico recente (azul) vs projeção IA/regressão linear (roxo)</div>
      ${{ultMeses.map((m,i)=>barRow(ymLbl(m), ultConts[i], '#0078D4', false)).join('')}}
      ${{projMeses.map((m,i)=>barRow(ymLbl(m), projConts[i], '#7B2FBE', true)).join('')}}
    </div>
  </div>

  <!-- Recomendações -->
  <div>
    <div style="font-size:12px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #7B2FBE;padding-bottom:5px;margin-bottom:10px">🎯 Recomendações Operacionais (IA)</div>
    ${{recs.map(r=>`<div style="padding:8px 12px;background:#F9F5FF;border-left:3px solid #7B2FBE;border-radius:4px;margin-bottom:6px;font-size:11px;line-height:1.5">${{r}}</div>`).join('')}}
    <div style="margin-top:12px;font-size:10px;color:#aaa;text-align:center;font-style:italic">
      ⚠️ A análise preditiva estima riscos em locais e períodos com base em padrões históricos.
      Não identifica indivíduos nem prevê quem cometerá crimes. Ferramenta de apoio ao planejamento operacional.
    </div>
  </div>`;

  document.getElementById('predit-overlay').classList.add('ativo');
}}

function imprimirPrevisao() {{
  const corpo = document.getElementById('prev-corpo').innerHTML;
  const titulo = document.getElementById('prev-titulo').textContent;
  const w = window.open('','_blank','width=860,height=720');
  w.document.write(`<!DOCTYPE html><html><head>
    <meta charset="utf-8">
    <title>Previsão de Risco — Guarda Municipal BC</title>
    <style>
      *{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
      body{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#222;padding:24px;margin:0;}}
      h2{{color:#1A1A2E;font-size:15px;border-bottom:3px solid #C05700;padding-bottom:6px;margin-bottom:14px;}}
      h3{{font-size:11px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #0078D4;
          padding-bottom:3px;margin:14px 0 8px;text-transform:uppercase;letter-spacing:.5px;}}
      table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:4px;}}
      th{{background:#1A1A2E!important;color:white!important;padding:4px 8px;text-align:left;}}
      td{{padding:4px 8px;border-bottom:1px solid #eee;}}
      tr:nth-child(even){{background:#f7f8fa!important;}}
      ol{{padding-left:18px;margin:0;}} li{{margin-bottom:6px;line-height:1.55;}}
      .risco-alto{{background:#D13438!important;color:white!important;border-radius:3px;padding:1px 6px;font-size:9px;font-weight:700;}}
      .risco-medio{{background:#E07B00!important;color:white!important;border-radius:3px;padding:1px 6px;font-size:9px;font-weight:700;}}
      .risco-baixo{{background:#107C10!important;color:white!important;border-radius:3px;padding:1px 6px;font-size:9px;font-weight:700;}}
      .risco-bar-wrap{{background:#eee!important;height:8px;border-radius:4px;width:70px;display:inline-block;vertical-align:middle;}}
      .risco-bar-fill{{height:8px;border-radius:4px;}}
    </style>
  </head><body>
    <h2>${{titulo}}</h2>
    ${{corpo}}
  </body></html>`);
  w.document.close(); w.focus();
  setTimeout(()=>w.print(), 500);
}}

function imprimirAnalise() {{
  const corpo = document.getElementById('analise-corpo').innerHTML;
  const w = window.open('','_blank','width=860,height=720');
  w.document.write(`<!DOCTYPE html><html><head>
    <meta charset="utf-8">
    <title>Análise Diária — Guarda Municipal BC</title>
    <style>
      *{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important;}}
      body{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#222;padding:24px;margin:0;}}
      h2{{color:#1A1A2E;font-size:16px;border-bottom:3px solid #0078D4;padding-bottom:6px;margin-bottom:16px;}}
      h3{{font-size:11px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #0078D4;
          padding-bottom:3px;margin:14px 0 8px;text-transform:uppercase;letter-spacing:.5px;}}
      table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:4px;}}
      th{{background:#1A1A2E!important;color:white!important;padding:4px 8px;text-align:left;font-weight:600;}}
      td{{padding:4px 8px;border-bottom:1px solid #eee;}}
      tr:nth-child(even){{background:#f7f8fa!important;}}
      ol{{padding-left:18px;margin:0;}} li{{margin-bottom:6px;line-height:1.55;}}
      .badge-resumo{{display:inline-block;background:#0078D4!important;color:white!important;
        border-radius:4px;padding:1px 9px;font-weight:700;margin-left:4px;font-size:13px;}}
      [style*="background:#1A1A2E"]{{background:#1A1A2E!important;color:white!important;}}
      [style*="background:#0078D4"]{{background:#0078D4!important;color:white!important;}}
      [style*="background:#107C10"]{{background:#107C10!important;color:white!important;}}
      [style*="background:#5C2D91"]{{background:#5C2D91!important;color:white!important;}}
      [style*="background:linear-gradient"]{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
    </style>
  </head><body>
    <h2 id="analise-titulo">📊 Análise por Dia da Semana — Guarda Municipal BC</h2>
    ${{corpo}}
  </body></html>`);
  w.document.close();
  w.focus();
  setTimeout(()=>w.print(), 500);
}}

function enviarWhatsApp() {{
  if(!_analiseTextoWA) return;
  window.open('https://wa.me/?text=' + encodeURIComponent(_analiseTextoWA), '_blank');
}}

function enviarWhatsAppPrevisao() {{
  if(!_prevTextoWA) return;
  window.open('https://wa.me/?text=' + encodeURIComponent(_prevTextoWA), '_blank');
}}

// ── RELATÓRIO DIÁRIO ──────────────────────────────────────────────────────────
let _relTextoWA = '';

function relatorioDiario() {{
  const now = new Date();
  const today = `${{now.getFullYear()}}-${{String(now.getMonth()+1).padStart(2,'0')}}-${{String(now.getDate()).padStart(2,'0')}}`;
  document.getElementById('rel-data-ini').value = today;
  document.getElementById('rel-data-fim').value = today;
  document.getElementById('rel-titulo').textContent = '📅 Relatório de Ocorrências — Guarda Municipal BC';
  document.getElementById('rel-corpo').innerHTML =
    '<div style="color:#888;text-align:center;padding:40px;font-style:italic">Selecione o período e clique em <strong>Gerar Relatório</strong>.</div>';
  document.getElementById('rel-overlay').classList.add('ativo');
}}

function fecharRelatorio() {{
  document.getElementById('rel-overlay').classList.remove('ativo');
}}

function gerarRelatorio() {{
  const ini = document.getElementById('rel-data-ini').value;
  const fim = document.getElementById('rel-data-fim').value;
  const corpo = document.getElementById('rel-corpo');
  if (!ini || !fim) {{
    corpo.innerHTML = '<div style="color:#D13438;text-align:center;padding:20px;font-weight:600">⚠️ Selecione as datas de início e fim.</div>';
    return;
  }}
  if (ini > fim) {{
    corpo.innerHTML = '<div style="color:#D13438;text-align:center;padding:20px;font-weight:600">⚠️ A data inicial não pode ser maior que a data final.</div>';
    return;
  }}

  const data = RAW.filter(r => r.data && r.data >= ini && r.data <= fim);

  const now = new Date();
  const horaBR    = now.toLocaleTimeString('pt-BR', {{hour:'2-digit',minute:'2-digit'}});
  const nowDateBR = now.toLocaleDateString('pt-BR');
  const fmtDate   = d => d.split('-').reverse().join('/');
  const periodoLabel = ini === fim ? fmtDate(ini) : `${{fmtDate(ini)}} a ${{fmtDate(fim)}}`;

  document.getElementById('rel-titulo').textContent = `📅 Relatório — ${{periodoLabel}} — Guarda Municipal BC`;

  if (data.length === 0) {{
    corpo.innerHTML = `<div style="color:#888;text-align:center;padding:40px;font-style:italic">⚠️ Nenhuma ocorrência registrada no período <strong>${{periodoLabel}}</strong>.</div>`;
    _relTextoWA = '';
    return;
  }}

  const cntF = (arr, f) => {{
    const m = {{}};
    arr.forEach(r => {{ if(r[f]) m[r[f]] = (m[r[f]]||0)+1; }});
    return Object.entries(m).sort((a,b) => b[1]-a[1]);
  }};

  const total      = data.length;
  const furtos     = data.filter(r=>r.tipo==='Furto').length;
  const roubos     = data.filter(r=>r.tipo==='Roubo').length;
  const arrom      = data.filter(r=>r.tipo==='Arrombamento').length;
  const tentFurto  = data.filter(r=>r.tipo==='Tentativa de Furto').length;
  const tentRoubo  = data.filter(r=>r.tipo==='Tentativa de Roubo').length;

  const tipos   = cntF(data,'tipo');
  const bairros = cntF(data,'bairro');
  const turnos  = cntF(data,'turno');
  const itens   = cntF(data,'item');
  const ruas    = cntF(data,'endereco').filter(e=>e[0]);

  const pct = n => total ? ((n/total)*100).toFixed(0)+'%' : '0%';

  const mkTable = (rows, headers) =>
    `<table class="analise-table"><thead><tr>${{headers.map(h=>`<th>${{h}}</th>`).join('')}}</tr></thead><tbody>${{rows}}</tbody></table>`;

  const tiposRows = tipos.map((e,i) =>
    `<tr><td>${{i+1}}. ${{e[0]}}</td><td style="text-align:center">${{e[1]}}</td><td style="text-align:center">${{pct(e[1])}}</td></tr>`
  ).join('');

  const bairrosRows = bairros.slice(0,10).map((e,i) =>
    `<tr><td>${{i+1}}. ${{e[0]}}</td><td style="text-align:center">${{e[1]}}</td><td style="text-align:center">${{pct(e[1])}}</td></tr>`
  ).join('');

  const ORDEM_TURNO = ['Madrugada','Manhã','Tarde','Noite'];
  const turnoMap = {{}};
  turnos.forEach(e => turnoMap[e[0]] = e[1]);
  const turnosRows = ORDEM_TURNO.map(t =>
    `<tr><td>${{t}}</td><td style="text-align:center">${{turnoMap[t]||0}}</td><td style="text-align:center">${{pct(turnoMap[t]||0)}}</td></tr>`
  ).join('');

  const itensRows = itens.slice(0,10).map((e,i) =>
    `<tr><td>${{i+1}}. ${{e[0]}}</td><td style="text-align:center">${{e[1]}}</td><td style="text-align:center">${{pct(e[1])}}</td></tr>`
  ).join('');

  const ruasRows = ruas.slice(0,10).map((e,i) =>
    `<tr><td>${{i+1}}. ${{e[0]}}</td><td style="text-align:center">${{e[1]}}</td><td style="text-align:center">${{pct(e[1])}}</td></tr>`
  ).join('');

  const sorted = [...data].sort((a,b) =>
    ((a.data||'')+' '+(a.hora||'')).localeCompare((b.data||'')+' '+(b.hora||''))
  );
  const boRows = sorted.map(r => `
    <tr>
      <td>${{r.data ? r.data.slice(8)+'/'+r.data.slice(5,7)+'/'+r.data.slice(0,4) : ''}}</td>
      <td>${{r.hora}}</td>
      <td>${{r.turno}}</td>
      <td style="font-size:9px">${{r.bo}}</td>
      <td><span style="background:${{TIPO_COLORS[r.tipo]||'#888'}};color:white;border-radius:3px;padding:1px 5px;font-size:9px">${{r.tipo}}</span></td>
      <td>${{r.item}}</td>
      <td>${{r.bairro}}</td>
      <td style="font-size:10px;max-width:180px;overflow:hidden;text-overflow:ellipsis">${{r.endereco}}</td>
    </tr>`).join('');

  const html = `
  <div style="text-align:center;background:linear-gradient(135deg,#f0f4ff,#e8f0fe);border-radius:8px;
    padding:12px;margin-bottom:14px;border-left:4px solid #0078D4;">
    <div style="font-weight:800;font-size:15px;color:#1A1A2E">📅 Relatório de Ocorrências</div>
    <div style="color:#555;font-size:11px;margin-top:4px">
      Guarda Municipal de Balneário Camboriú<br>
      Período: <strong>${{periodoLabel}}</strong> &nbsp;|&nbsp; Gerado em: <strong>${{nowDateBR}}</strong> às <strong>${{horaBR}}</strong>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:repeat(${{4+(tentFurto>0?1:0)+(tentRoubo>0?1:0)}},1fr);gap:8px;margin-bottom:14px;">
    <div style="background:#1A1A2E;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Total</div>
      <div style="font-size:28px;font-weight:800;line-height:1.1">${{total}}</div>
      <div style="font-size:9px;opacity:.65">ocorrências</div>
    </div>
    <div style="background:#E07B00;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Furtos</div>
      <div style="font-size:28px;font-weight:800;line-height:1.1">${{furtos}}</div>
      <div style="font-size:9px;opacity:.65">${{pct(furtos)}}</div>
    </div>
    <div style="background:#D13438;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Roubos</div>
      <div style="font-size:28px;font-weight:800;line-height:1.1">${{roubos}}</div>
      <div style="font-size:9px;opacity:.65">${{pct(roubos)}}</div>
    </div>
    <div style="background:#795548;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Arrombamentos</div>
      <div style="font-size:22px;font-weight:800;line-height:1.3;margin-top:2px">${{arrom}}</div>
      <div style="font-size:9px;opacity:.65">${{pct(arrom)}}</div>
    </div>
    ${{tentFurto > 0 ? `
    <div style="background:#0078D4;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Tent. Furto</div>
      <div style="font-size:28px;font-weight:800;line-height:1.1">${{tentFurto}}</div>
      <div style="font-size:9px;opacity:.65">${{pct(tentFurto)}}</div>
    </div>` : ''}}
    ${{tentRoubo > 0 ? `
    <div style="background:#8764B8;color:white;border-radius:6px;padding:10px;text-align:center;">
      <div style="font-size:9px;opacity:.75;text-transform:uppercase">Tent. Roubo</div>
      <div style="font-size:28px;font-weight:800;line-height:1.1">${{tentRoubo}}</div>
      <div style="font-size:9px;opacity:.65">${{pct(tentRoubo)}}</div>
    </div>` : ''}}
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;">
    <div class="analise-section">
      <h3>🔴 Tipos de Ocorrência</h3>
      ${{mkTable(tiposRows, ['Tipo','Total','%'])}}
    </div>
    <div class="analise-section">
      <h3>⏰ Distribuição por Turno</h3>
      ${{mkTable(turnosRows, ['Turno','Total','%'])}}
    </div>
  </div>

  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;">
    <div class="analise-section">
      <h3>📍 Bairros Afetados</h3>
      ${{mkTable(bairrosRows, ['Bairro','Total','%'])}}
    </div>
    <div class="analise-section">
      <h3>📦 Itens Furtados / Roubados</h3>
      ${{mkTable(itensRows, ['Item','Total','%'])}}
    </div>
  </div>

  <div class="analise-section" style="margin-bottom:14px">
    <h3>🛣️ Logradouros com Mais Ocorrências</h3>
    ${{mkTable(ruasRows, ['Logradouro','Total','%'])}}
  </div>

  <div class="analise-section">
    <h3>📋 Lista Detalhada — ${{total}} Ocorrência(s)</h3>
    <div style="overflow-x:auto">
      <table class="analise-table">
        <thead><tr>
          <th>Data</th><th>Hora</th><th>Turno</th><th>B.O.</th>
          <th>Tipo</th><th>Item</th><th>Bairro</th><th>Endereço</th>
        </tr></thead>
        <tbody>${{boRows}}</tbody>
      </table>
    </div>
  </div>

  <div style="border-top:1px solid #ddd;margin-top:12px;padding-top:8px;color:#888;font-size:10px;text-align:center;">
    Secretaria de Segurança e Ordem Pública de Balneário Camboriú — Guarda Municipal
  </div>`;

  corpo.innerHTML = html;

  const waLines = [
    `📅 *RELATÓRIO DE OCORRÊNCIAS — GUARDA MUNICIPAL BC*`,
    `Período: *${{periodoLabel}}*`,
    `🕐 Gerado em: ${{nowDateBR}} às ${{horaBR}}`,
    ``,
    `*📊 RESUMO*`,
    `Total: *${{total}}* ocorrências`,
    `Furtos: *${{furtos}}* (${{pct(furtos)}})`,
    `Roubos: *${{roubos}}* (${{pct(roubos)}})`,
    `Arrombamentos: *${{arrom}}* (${{pct(arrom)}})`,
    ``,
    `*🔴 TIPOS*`,
    ...tipos.map((e,i) => `${{i+1}}. ${{e[0]}}: ${{e[1]}} (${{pct(e[1])}})`),
    ``,
    `*📍 BAIRROS*`,
    ...bairros.slice(0,5).map((e,i) => `${{i+1}}. ${{e[0]}}: ${{e[1]}} (${{pct(e[1])}})`),
    ``,
    `*⏰ TURNOS*`,
    ...ORDEM_TURNO.map(t => `• ${{t}}: ${{turnoMap[t]||0}} (${{pct(turnoMap[t]||0)}})`),
    ``,
    `*📦 ITENS*`,
    ...itens.slice(0,5).map((e,i) => `${{i+1}}. ${{e[0]}}: ${{e[1]}} (${{pct(e[1])}})`),
    ``,
    `*🛣️ LOGRADOUROS*`,
    ...ruas.slice(0,5).map((e,i) => `${{i+1}}. ${{e[0]}}: ${{e[1]}}`),
    ``,
    `_Guarda Municipal de Balneário Camboriú_`,
    `_Secretaria de Segurança e Ordem Pública_`
  ];
  _relTextoWA = waLines.join('\\n');
}}

function imprimirRelatorio() {{
  const corpoEl = document.getElementById('rel-corpo');
  if (!_relTextoWA) {{ alert('Gere o relatório primeiro.'); return; }}
  const titulo = document.getElementById('rel-titulo').textContent;
  const w = window.open('','_blank','width=900,height=720');
  w.document.write(`<!DOCTYPE html><html><head>
    <meta charset="utf-8">
    <title>${{titulo}}</title>
    <style>
      *{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
      body{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#222;padding:24px;margin:0;}}
      h2{{color:#1A1A2E;font-size:15px;border-bottom:3px solid #0097A7;padding-bottom:6px;margin-bottom:14px;}}
      h3{{font-size:11px;font-weight:700;color:#1A1A2E;border-bottom:2px solid #0078D4;
          padding-bottom:3px;margin:14px 0 8px;text-transform:uppercase;letter-spacing:.5px;}}
      table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:4px;}}
      th{{background:#1A1A2E!important;color:white!important;padding:4px 8px;text-align:left;}}
      td{{padding:4px 8px;border-bottom:1px solid #eee;}}
      tr:nth-child(even){{background:#f7f8fa!important;}}
      [style*="background:#1A1A2E"]{{background:#1A1A2E!important;color:white!important;}}
      [style*="background:#E07B00"]{{background:#E07B00!important;color:white!important;}}
      [style*="background:#D13438"]{{background:#D13438!important;color:white!important;}}
      [style*="background:#795548"]{{background:#795548!important;color:white!important;}}
      [style*="background:#0078D4"]{{background:#0078D4!important;color:white!important;}}
    </style>
  </head><body>
    <h2>${{titulo}}</h2>
    ${{corpoEl.innerHTML}}
  </body></html>`);
  w.document.close(); w.focus();
  setTimeout(()=>w.print(), 500);
}}

function enviarWhatsAppRelatorio() {{
  if (!_relTextoWA) {{ alert('Gere o relatório primeiro.'); return; }}
  window.open('https://wa.me/?text=' + encodeURIComponent(_relTextoWA), '_blank');
}}

// ── SIDEBAR DINÂMICA ──────────────────────────────────────────────────────────
function buildCheckboxes(containerId, key, values, counts) {{
  const container = document.getElementById(containerId);
  if(!container) return;
  const currentSearch = container.dataset.search||'';
  const items = values.filter(v=>v.toLowerCase().includes(currentSearch.toLowerCase()));
  container.innerHTML = items.map(v => `
    <div class="filter-item">
      <input type="checkbox" id="cb-${{key}}-${{v}}" ${{state[key].has(v)?'checked':''}}>
      <label for="cb-${{key}}-${{v}}">${{v}}</label>
      <span class="cnt">${{counts[v]||0}}</span>
    </div>`).join('');
  container.querySelectorAll('input[type=checkbox]').forEach(cb => {{
    const v = cb.id.replace(`cb-${{key}}-`,'');
    cb.addEventListener('change', () => toggleFilter(key, v));
  }});
}}

function filterSearch(key, q) {{
  const map = {{bairro:'filter-bairro',item:'filter-item',logradouro:'filter-logradouro'}};
  const el = document.getElementById(map[key]);
  if(el) el.dataset.search = q;
  buildSidebar();
}}

function filterImei(q) {{
  imeiQ = q.trim();
  renderAll();
}}

function filterMarca(q) {{
  marcaQ = q.trim();
  renderAll();
}}

function filterPlaca(q) {{
  placaQ = q.trim();
  renderAll();
}}

function filterNumeroSerie(q) {{
  numeroSerieQ = q.trim();
  renderAll();
}}

function filterCor(q) {{
  corQ = q.trim();
  renderAll();
}}

function filterDetalhes(q) {{
  detalhesQ = q.trim();
  renderAll();
}}


function buildSidebar() {{
  const RAWU = dedupBO(RAW); // RAW deduplicado por B.O. (ocorrências)
  const ocCounts = k => count(RAWU, k); // conta ocorrências (não itens)

  const anos        = [...new Set(RAWU.map(r=>String(r.ano)))].sort((a,b)=>a-b);
  const anoCnts     = {{}}; RAWU.forEach(r=>{{ const k=String(r.ano); anoCnts[k]=(anoCnts[k]||0)+1; }});
  const meses       = [...new Set(RAWU.map(r=>r.mes))].sort();
  const turnos      = ['Manhã','Tarde','Noite','Madrugada'];
  const tipos       = [...new Set(RAWU.map(r=>r.tipo))].sort();
  const bairros     = sortedEntries(count(RAWU,'bairro')).map(e=>e[0]);
  const itens       = sortedEntries(count(RAW,'item')).map(e=>e[0]); // itens: usa RAW completo
  const dias        = DIA_ORDER.filter(d=>RAWU.some(r=>r.dia===d));
  const logradouros = [...new Set(RAWU.map(r=>r.endereco).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  const recuperados = sortedEntries(count(RAWU,'recuperado')).map(e=>e[0]).filter(Boolean);

  buildCheckboxes('filter-ano',        'ano',        anos,        anoCnts);
  buildCheckboxes('filter-mes',        'mes',        meses,       ocCounts('mes'));
  buildCheckboxes('filter-turno',      'turno',      turnos,      ocCounts('turno'));
  buildCheckboxes('filter-tipo',       'tipo',        tipos,      ocCounts('tipo'));
  buildCheckboxes('filter-bairro',     'bairro',     bairros,     ocCounts('bairro'));
  buildCheckboxes('filter-item',       'item',        itens,      count(RAW,'item')); // itens: conta cada item
  buildCheckboxes('filter-dia',        'dia',         dias,       ocCounts('dia'));
  buildCheckboxes('filter-logradouro', 'logradouro', logradouros, count(RAWU,'endereco'));
  buildCheckboxes('filter-recuperado', 'recuperado', recuperados, ocCounts('recuperado'));
}}

// ── RUAS ─────────────────────────────────────────────────────────────────────
function renderRuas(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'endereco');
  const entries = sortedEntries(c).filter(e=>e[0]&&e[0]!=='').slice(0,10); // top 10
  const labels = entries.map(e=>e[0]).map(l=>l.length>30?l.slice(0,28)+'…':l).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:i>labels.length-4?COLORS.laranja:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:9}},automargin:false}},
    margin:{{l:190,r:36,t:4,b:24}},
  }};
  Plotly.react('chart-ruas', barH(labels,vals,colors), layout, CONFIG);
}}

// ── PONTOS DE REFERÊNCIA ──────────────────────────────────────────────────────
function renderRefs(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'ref');
  const entries = sortedEntries(c).filter(e=>e[0]&&e[0]!=='').slice(0,10); // top 10
  const labels = entries.map(e=>e[0]).map(l=>l.length>32?l.slice(0,30)+'…':l).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:i>labels.length-4?COLORS.laranja:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:9}},automargin:false}},
    margin:{{l:210,r:36,t:4,b:24}},
  }};
  Plotly.react('chart-refs', barH(labels,vals,colors), layout, CONFIG);
}}

// ── HEATMAP DIA x TURNO ───────────────────────────────────────────────────────
function renderHeatmap(data) {{
  if(typeof Plotly === 'undefined') return;

  const TURNOS = ['Manhã','Tarde','Noite','Madrugada'];
  const DIAS   = ['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo'];
  const matrix = TURNOS.map(t => DIAS.map(d => data.filter(r=>r.turno===t&&r.dia===d).length));
  const trace = {{
    type:'heatmap',
    z:matrix, x:DIAS, y:TURNOS,
    colorscale:[
      [0,'#EBF3FC'],[0.25,'#50B2FF'],[0.55,'#0078D4'],
      [0.8,'#E07B00'],[1,'#D13438']
    ],
    showscale:true,
    colorbar:{{thickness:12,len:.8,tickfont:{{size:9}}}},
    hovertemplate:'<b>%{{y}} – %{{x}}</b><br>%{{z}} ocorrências<extra></extra>',
    text:matrix.map(row=>row.map(v=>v||'')),
    texttemplate:'%{{text}}',
    textfont:{{size:12,color:'white'}},
  }};
  const layout = {{...LAYOUT_BASE,
    margin:{{l:80,r:50,t:8,b:60}},
    xaxis:{{tickfont:{{size:10}},side:'bottom'}},
    yaxis:{{tickfont:{{size:10}},automargin:true}},
  }};
  Plotly.react('chart-heatmap',[trace],layout,CONFIG);
}}

// ── HORA DO DIA ───────────────────────────────────────────────────────────────
function renderHora(data) {{
  if(typeof Plotly === 'undefined') return;

  const hrs = Array(24).fill(0);
  data.forEach(r => {{
    if(r.hora) {{
      const h = parseInt(r.hora.split(':')[0]);
      if(!isNaN(h)) hrs[h]++;
    }}
  }});
  const labels = hrs.map((_,i)=>`${{String(i).padStart(2,'0')}}h`);
  const maxHr = Math.max(...hrs);
  const colors = hrs.map((v,i) => {{
    if(v===maxHr && v>0) return COLORS.vermelho;
    if(i>=22||i<6)       return COLORS.roxo;
    if(i>=18)            return COLORS.azul;
    if(i>=12)            return COLORS.laranja;
    return COLORS.amarelo;
  }});
  const trace = [{{
    type:'bar', x:labels, y:hrs,
    marker:{{color:colors,line:{{color:'white',width:0.5}}}},
    text:hrs.map(v=>v>0?String(v):''),
    textposition:'outside',
    textfont:{{size:9,color:'#333'}},
    cliponaxis:false,
    hovertemplate:'<b>%{{x}}</b><br>%{{y}} casos<extra></extra>',
  }}];
  const layout = {{...LAYOUT_BASE,
    xaxis:{{tickfont:{{size:8.5}},tickangle:-45}},
    yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}},range:[0,Math.max(...hrs,1)*1.20]}},
    margin:{{l:30,r:10,t:8,b:50}},
    shapes:[
      {{type:'rect',x0:'18h',x1:'24h',y0:0,y1:1,xref:'x',yref:'paper',
        fillcolor:'rgba(0,120,212,.06)',line:{{width:0}}}},
      {{type:'rect',x0:'00h',x1:'06h',y0:0,y1:1,xref:'x',yref:'paper',
        fillcolor:'rgba(135,100,184,.08)',line:{{width:0}}}},
    ],
    annotations:[
      {{x:'21h',y:1,xref:'x',yref:'paper',text:'Noite',showarrow:false,
        font:{{size:9,color:COLORS.azul}},xanchor:'center',yanchor:'top'}},
      {{x:'03h',y:1,xref:'x',yref:'paper',text:'Madrugada',showarrow:false,
        font:{{size:9,color:COLORS.roxo}},xanchor:'center',yanchor:'top'}},
    ],
  }};
  Plotly.react('chart-hora',trace,layout,CONFIG);
}}

// ── COMPARAÇÃO ANUAL ─────────────────────────────────────────────────────────
const ANO_PALETTE = ['#0078D4','#E07B00','#107C10','#8764B8','#D13438','#0097A7'];
let compAnosAtivos = new Set();

function _anoColor(ano, anosAll) {{
  const i = anosAll.indexOf(ano);
  return ANO_PALETTE[i >= 0 ? i % ANO_PALETTE.length : 0];
}}

function initCompAnos(anosAll) {{
  // Primeira chamada: ativa todos os anos disponíveis
  if (compAnosAtivos.size === 0) anosAll.forEach(a => compAnosAtivos.add(a));
  const sel = document.getElementById('comp-anos-selector');
  if (!sel) return;
  const spanTitle = '<span style="font-size:11px;color:#555;font-weight:700;align-self:center">Selecione os anos:</span>';
  const btns = anosAll.map((a, i) => {{
    const cor   = ANO_PALETTE[i % ANO_PALETTE.length];
    const ativo = compAnosAtivos.has(a);
    return `<button class="comp-ano-btn${{ativo?' ativo':''}}"
      style="background:${{ativo?cor:'#e8ecf0'}};color:${{ativo?'white':cor}};border:2px solid ${{cor}}"
      onclick="toggleCompAno(${{a}})">${{a}}</button>`;
  }}).join('');
  sel.innerHTML = spanTitle + btns;
  const label = document.getElementById('comp-anos-label');
  if (label) {{
    const ativos = [...compAnosAtivos].sort();
    label.textContent = ativos.length > 1 ? '— ' + ativos.join(' vs ') : ativos.length === 1 ? '— ' + ativos[0] : '';
  }}
}}

function toggleCompAno(ano) {{
  if (compAnosAtivos.has(ano)) {{
    if (compAnosAtivos.size > 1) compAnosAtivos.delete(ano);
  }} else {{
    compAnosAtivos.add(ano);
  }}
  renderComparacao(dedupBO(filtered()));
}}

function renderComparacao(data) {{
  if (typeof Plotly === 'undefined') return;
  const anosAll = [...new Set(RAW.map(r => r.ano))].sort();
  if (anosAll.length === 0) return;
  initCompAnos(anosAll);
  const vis = anosAll.filter(a => compAnosAtivos.has(a));
  if (vis.length === 0) return;

  const LCOMP = {{
    ...LAYOUT_BASE,
    barmode:'group',
    showlegend:true,
    legend:{{orientation:'h',x:0,y:1.18,font:{{size:11}}}},
  }};

  // ── KPI Deltas (2 anos mais recentes visíveis) ────────────────────────────
  const kpiRow = document.getElementById('comp-kpi-row');
  if (kpiRow && vis.length >= 2) {{
    const aA = vis[vis.length-2], aB = vis[vis.length-1];
    const dA = data.filter(r => r.ano === aA), dB = data.filter(r => r.ano === aB);
    const corB = _anoColor(aB, anosAll);
    function mkDelta(vA, vB, label) {{
      const diff = vB - vA;
      const pct  = vA > 0 ? ((diff/vA)*100).toFixed(0) : (diff > 0 ? '+∞' : '0');
      const seta = diff > 0 ? '▲' : diff < 0 ? '▼' : '=';
      const cls  = diff > 0 ? 'delta-up' : diff < 0 ? 'delta-down' : 'delta-eq';
      const col  = diff > 0 ? '#D13438' : diff < 0 ? '#107C10' : '#888';
      return `<div class="comp-kpi-delta" style="border-top:4px solid ${{corB}}">
        <div style="font-size:9px;color:#aaa;margin-bottom:2px;font-weight:600;text-transform:uppercase">${{aA}} → ${{aB}}</div>
        <div class="comp-delta-val ${{cls}}">${{seta}} ${{diff>0?'+':''}}${{diff}}</div>
        <div style="font-size:12px;font-weight:800;color:${{col}}">${{diff>0?'+':''}}${{pct}}%</div>
        <div class="comp-delta-label">${{label}}</div>
        <div style="font-size:9px;color:#888;margin-top:2px">${{aA}}: ${{vA}} &nbsp;|&nbsp; ${{aB}}: ${{vB}}</div>
      </div>`;
    }}
    kpiRow.style.display = 'grid';
    kpiRow.innerHTML = [
      mkDelta(dA.length,                                      dB.length,                                      'Total de Ocorrências'),
      mkDelta(dA.filter(r=>r.tipo==='Furto').length,          dB.filter(r=>r.tipo==='Furto').length,          'Furtos'),
      mkDelta(dA.filter(r=>r.tipo==='Roubo').length,          dB.filter(r=>r.tipo==='Roubo').length,          'Roubos'),
      mkDelta(dA.filter(r=>r.tipo==='Arrombamento').length,   dB.filter(r=>r.tipo==='Arrombamento').length,   'Arrombamentos'),
    ].join('');
  }} else if (kpiRow) {{
    kpiRow.style.display = 'none';
  }}

  // helper: conta por campo dentro de um subset de data para um ano
  const cntAnoCampo = (a, campo, val) => data.filter(r => r.ano===a && r[campo]===val).length;

  // ── Por Mês ────────────────────────────────────────────────────────────────
  const MES_ORD = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
    'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const mesesP = MES_ORD.filter(m => vis.some(a => data.some(r=>r.ano===a&&r.mes===m)));
  Plotly.react('comp-chart-mes',
    vis.map(a => ({{
      type:'bar', name:String(a),
      x:mesesP,
      y:mesesP.map(m => cntAnoCampo(a,'mes',m)),
      marker:{{color:_anoColor(a,anosAll)}},
      text:mesesP.map(m=>{{ const v=cntAnoCampo(a,'mes',m); return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{x}} %{{data.name}}</b><br>%{{y}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{tickfont:{{size:10}}}},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      margin:{{l:30,r:12,t:44,b:36}},
    }}, CONFIG);

  // ── Por Tipificação ────────────────────────────────────────────────────────
  const TIPOS_ORD = ['Furto','Roubo','Arrombamento','Tentativa de Furto','Tentativa de Roubo'];
  const tiposP = TIPOS_ORD.filter(t => vis.some(a => data.some(r=>r.ano===a&&r.tipo===t)));
  Plotly.react('comp-chart-tipo',
    vis.map(a => ({{
      type:'bar', name:String(a),
      x:tiposP,
      y:tiposP.map(t => cntAnoCampo(a,'tipo',t)),
      marker:{{color:_anoColor(a,anosAll)}},
      text:tiposP.map(t=>{{ const v=cntAnoCampo(a,'tipo',t); return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{x}} %{{data.name}}</b><br>%{{y}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{tickfont:{{size:10}},tickangle:-15}},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      margin:{{l:30,r:12,t:44,b:60}},
    }}, CONFIG);

  // ── Por Bairro (top 10 combinado) ─────────────────────────────────────────
  const top10B = sortedEntries(count(data,'bairro')).filter(e=>e[0]).slice(0,10).map(e=>e[0]).reverse();
  Plotly.react('comp-chart-bairro',
    vis.map(a => ({{
      type:'bar', orientation:'h', name:String(a),
      y:top10B,
      x:top10B.map(b => cntAnoCampo(a,'bairro',b)),
      marker:{{color:_anoColor(a,anosAll)}},
      text:top10B.map(b=>{{ const v=cntAnoCampo(a,'bairro',b); return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{y}} %{{data.name}}</b><br>%{{x}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      yaxis:{{tickfont:{{size:10}},automargin:true}},
      margin:{{l:80,r:40,t:44,b:30}},
    }}, CONFIG);

  // ── Por Turno ─────────────────────────────────────────────────────────────
  const TURN_ORD = ['Manhã','Tarde','Noite','Madrugada'];
  Plotly.react('comp-chart-turno',
    vis.map(a => ({{
      type:'bar', name:String(a),
      x:TURN_ORD,
      y:TURN_ORD.map(t => cntAnoCampo(a,'turno',t)),
      marker:{{color:_anoColor(a,anosAll)}},
      text:TURN_ORD.map(t=>{{ const v=cntAnoCampo(a,'turno',t); return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{x}} %{{data.name}}</b><br>%{{y}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{tickfont:{{size:10}}}},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      margin:{{l:30,r:12,t:44,b:36}},
    }}, CONFIG);

  // ── Dia da Semana ──────────────────────────────────────────────────────────
  const diasP = DIA_ORDER.filter(d => vis.some(a => data.some(r=>r.ano===a&&r.dia===d)));
  Plotly.react('comp-chart-dia',
    vis.map(a => ({{
      type:'bar', name:String(a),
      x:diasP,
      y:diasP.map(d => cntAnoCampo(a,'dia',d)),
      marker:{{color:_anoColor(a,anosAll)}},
      text:diasP.map(d=>{{ const v=cntAnoCampo(a,'dia',d); return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{x}} %{{data.name}}</b><br>%{{y}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{tickfont:{{size:10}}}},
      yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      margin:{{l:30,r:12,t:44,b:36}},
    }}, CONFIG);

  // ── Top 10 Ruas ────────────────────────────────────────────────────────────
  const top10R = sortedEntries(count(data,'endereco')).filter(e=>e[0]).slice(0,10).map(e=>e[0]).reverse();
  Plotly.react('comp-chart-ruas',
    vis.map(a => ({{
      type:'bar', orientation:'h', name:String(a),
      y:top10R,
      x:top10R.map(rua => data.filter(r=>r.ano===a&&r.endereco===rua).length),
      marker:{{color:_anoColor(a,anosAll)}},
      text:top10R.map(rua=>{{ const v=data.filter(r=>r.ano===a&&r.endereco===rua).length; return v>0?String(v):''; }}),
      textposition:'outside', cliponaxis:false,
      hovertemplate:'<b>%{{y}} %{{data.name}}</b><br>%{{x}} ocorrências<extra></extra>',
    }})),
    {{...LCOMP,
      xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}}}},
      yaxis:{{tickfont:{{size:10}},automargin:true}},
      margin:{{l:160,r:40,t:44,b:30}},
    }}, CONFIG);
}}

// ── PLANO POLICIAL DINÂMICO ───────────────────────────────────────────────────
function renderPlanoPolicial(data) {{
  const total = data.length;
  if(total===0){{ document.getElementById('plano-policial').innerHTML=''; return; }}

  // ── Separar os dois turnos operacionais (usa campo turno, igual ao gráfico) ──
  // Turno A: Matutino (Manhã 06-11h) + Vespertino (Tarde 12-17h) → 8 viaturas
  // Turno B: Noturno (Noite 18-23h) + Madrugada (00-05h) → 7 viaturas
  const dadosA = data.filter(r=>r.turno==='Manhã'||r.turno==='Tarde');
  const dadosB = data.filter(r=>r.turno==='Noite'||r.turno==='Madrugada');

  const bairrosA = count(dadosA,'bairro');
  const bairrosB = count(dadosB,'bairro');
  const enderecos = count(data,'endereco');
  const refs      = count(data,'ref');
  const dias      = count(data,'dia');
  const horas     = {{}};
  data.forEach(r=>{{if(r.hora){{const h=parseInt(r.hora);if(!isNaN(h))horas[h]=(horas[h]||0)+1;}}}});

  const topRuas    = sortedEntries(enderecos).filter(e=>e[0]).slice(0,5);
  const topRefs    = sortedEntries(refs).filter(e=>e[0]).slice(0,5);
  const topDias    = sortedEntries(dias).slice(0,3);
  const topBairros = sortedEntries(count(data,'bairro')).filter(e=>e[0]).slice(0,5);
  const horasPico  = sortedEntries(horas).slice(0,3).map(e=>`${{String(e[0]).padStart(2,'0')}}h`).join(', ');
  const maxRua = topRuas[0]?topRuas[0][1]:1;
  const maxRef = topRefs[0]?topRefs[0][1]:1;

  // ── Regiões geográficas de Balneário Camboriú ─────────────────────────────
  const REGIOES = [
    {{ nome:'Área Central\\n(Centro · Barra Sul · Barra Norte)',
       bairros:['Centro','Barra Sul','Barra Norte'],
       icon:'🏙️' }},
    {{ nome:'Zona Sul\\n(N. Esperança · Barra · São J. Tadeu\\nEstaleiro · Taquaras · Laranjeiras)',
       bairros:['N. Esperança','Barra','São J. Tadeu','Estaleiro','Estaleirinho','Taquaras','Taquarinhas','Laranjeiras'],
       icon:'🌊' }},
    {{ nome:'Zona Norte\\n(Nações · Ariribá · Pioneiros\\nPraia dos Amores)',
       bairros:['Nações','Ariribá','Pioneiros','Praia dos Amores'],
       icon:'⬆️' }},
    {{ nome:'Zona Oeste\\n(Municípios · Vila Real · Iate Clube)',
       bairros:['Municípios','Vila Real','Iate Clube'],
       icon:'⬅️' }},
    {{ nome:'Zona Leste\\n(Estados · Várzea do Ranchinho\\n3ª Av. · 4ª Av.)',
       bairros:['Estados','Várzea do Ranchinho','3ª Avenida','4ª Avenida'],
       icon:'➡️' }},
  ];

  function distribuirPorRegiao(bairroCount, totalViaturas) {{
    // Soma crimes por região
    const regioes = REGIOES.map(r => {{
      const crimes = r.bairros.reduce((s,b) => s+(bairroCount[b]||0), 0);
      return {{ ...r, crimes }};
    }}).filter(r => r.crimes > 0);

    if(!regioes.length) return [];
    const totalCrimes = regioes.reduce((s,r)=>s+r.crimes,0);

    // Distribuição proporcional com mínimo 1 por região
    let alocado = regioes.map(r => ({{
      ...r,
      viaturas: Math.max(1, Math.floor((r.crimes/totalCrimes)*totalViaturas)),
    }}));

    // Ajustar diferença
    let diff = alocado.reduce((s,r)=>s+r.viaturas,0) - totalViaturas;
    // Remover do menor (exceto mínimo 1)
    let i = alocado.length-1;
    while(diff>0){{ if(alocado[i].viaturas>1){{alocado[i].viaturas--;diff--;}} i=((i-1)+alocado.length)%alocado.length; }}
    // Adicionar ao maior
    i=0;
    while(diff<0){{ alocado[i].viaturas++;diff++;i=(i+1)%alocado.length; }}

    return alocado;
  }}

  const distA = distribuirPorRegiao(bairrosA, 8);
  const distB = distribuirPorRegiao(bairrosB, 7);
  const totalA = distA.reduce((s,e)=>s+e.viaturas,0);
  const totalB = distB.reduce((s,e)=>s+e.viaturas,0);

  // ── HTML de um turno ──────────────────────────────────────────────────────
  function turnoHtml(dist, totalV, label, horario, icon, bgGrad, crimes) {{
    const maxCrimes = Math.max(...dist.map(x=>x.crimes), 1);
    const CORES = ['#D13438','#E07B00','#0078D4','#50B2FF','#8764B8'];
    const rows = dist.map((e,i)=>{{
      const barW = Math.round((e.crimes/maxCrimes)*100);
      const cor  = CORES[i]||'#888';
      const vArr = Array(e.viaturas).fill('🚔').join(' ');
      const nomeFormatado = e.nome.replace(/\\n/g,'<br><span style="font-size:9.5px;color:#888;font-weight:400">');
      return `
      <div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid #F2F2F2">
        <div style="min-width:130px;font-size:11px;font-weight:600;color:#1A1A2E;line-height:1.4">
          ${{nomeFormatado}}</span>
        </div>
        <div style="flex:1">
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;flex-wrap:wrap">
            <span style="font-size:14px;letter-spacing:2px">${{vArr}}</span>
            <span style="font-size:10px;font-weight:700;color:white;background:${{cor}};
              border-radius:3px;padding:2px 7px">${{e.viaturas}} viatura${{e.viaturas>1?'s':''}}</span>
            <span style="font-size:10px;color:#888">${{e.crimes}} ocorrência${{e.crimes>1?'s':''}}</span>
          </div>
          <div style="background:#EEE;border-radius:3px;height:6px;overflow:hidden">
            <div style="width:${{barW}}%;height:100%;background:${{cor}};border-radius:3px"></div>
          </div>
        </div>
      </div>`;
    }}).join('');
    const somaV = dist.reduce((s,e)=>s+e.viaturas,0);
    const checks = `<div style="display:flex;gap:16px;padding:8px 0 2px;font-size:11px;color:#555;flex-wrap:wrap">
      <span>✅ Total alocado: <strong>${{somaV}} / ${{totalV}}</strong></span>
      <span>📍 ${{dist.length}} regiões cobertas</span>
      <span>🔴 ${{crimes}} ocorrências no turno</span>
    </div>`;
    return `
    <div class="plano-card">
      <div class="plano-head" style="background:${{bgGrad}}">
        <div>
          <div style="font-size:14px;font-weight:700;color:white">${{icon}} ${{label}}</div>
          <div style="font-size:11px;color:rgba(255,255,255,.85);margin-top:3px">
            ${{horario}} &nbsp;·&nbsp; <strong style="color:white">${{totalV}} viaturas disponíveis</strong>
            &nbsp;·&nbsp; ${{crimes}} crimes no período
          </div>
        </div>
      </div>
      <div class="plano-body">${{rows}}${{checks}}</div>
    </div>`;
  }}

  const cardA = turnoHtml(distA,8,'Turno A','07h às 19h','☀️',
    'linear-gradient(135deg,#7A5100,#E07B00)', dadosA.length);
  const cardB = turnoHtml(distB,7,'Turno B','19h às 07h','🌙',
    'linear-gradient(135deg,#1A1A2E,#0078D4)', dadosB.length);

  const postosHtml = topRefs.map((e,i)=>{{
    const pct=Math.round((e[1]/maxRef)*100);
    const cor=i===0?COLORS.vermelho:i<2?COLORS.laranja:COLORS.azul;
    return `<div class="posto-item">
      <div class="posto-num">${{i+1}}</div>
      <div class="posto-info">
        <div class="posto-local">${{e[0]||'–'}}</div>
        <div class="posto-detalhe">${{e[1]}} ocorrência${{e[1]>1?'s':''}} registrada${{e[1]>1?'s':''}} neste ponto</div>
        <div class="risco-bar"><div class="risco-fill" style="width:${{pct}}%;background:${{cor}}"></div></div>
      </div></div>`;
  }}).join('');

  const ruasHtml = topRuas.map((e,i)=>{{
    const pct=Math.round((e[1]/maxRua)*100);
    const cor=i===0?COLORS.vermelho:i<2?COLORS.laranja:COLORS.azul;
    return `<div class="posto-item">
      <div class="posto-num">${{i+1}}</div>
      <div class="posto-info">
        <div class="posto-local">${{e[0]||'–'}}</div>
        <div class="posto-detalhe">${{e[1]}} ocorrência${{e[1]>1?'s':''}} · ronda intensiva recomendada</div>
        <div class="risco-bar"><div class="risco-fill" style="width:${{pct}}%;background:${{cor}}"></div></div>
      </div></div>`;
  }}).join('');

  const diasTexto = topDias.map(e=>`<strong>${{e[0]}}</strong> (${{e[1]}} casos)`).join(', ');
  const bairrosTexto = topBairros.slice(0,3).map(e=>e[0]).join(' › ');

  document.getElementById('plano-policial').innerHTML = `
  <div class="alerta-box">
    🚨 <span><strong>Zona de Máximo Risco:</strong> Bairro
    <strong>${{topBairros[0]?topBairros[0][0]:'–'}}</strong>.
    Dias críticos: ${{diasTexto}}. Horários de pico: ${{horasPico}}.</span>
  </div>
  <div class="plano-grid-wide">
    ${{cardA}} ${{cardB}}
  </div>
  <div class="plano-grid">
    <div class="plano-card">
      <div class="plano-head" style="background:linear-gradient(135deg,#A4262C,#D13438)">
        📌 Postos Fixos Prioritários
      </div>
      <div class="plano-body">${{postosHtml}}</div>
    </div>
    <div class="plano-card">
      <div class="plano-head" style="background:linear-gradient(135deg,#004E8C,#0078D4)">
        🛣️ Ruas para Ronda Intensiva
      </div>
      <div class="plano-body">${{ruasHtml}}</div>
    </div>
    <div class="plano-card">
      <div class="plano-head" style="background:linear-gradient(135deg,#004B1C,#107C10)">
        ✅ Diretrizes Operacionais
      </div>
      <div class="plano-body" style="padding-top:10px">
        <div class="ok-box">✔ <strong>Turno A (07h–19h):</strong> concentrar ${{distA[0]?distA[0].viaturas:3}} viaturas na ${{distA[0]?distA[0].nome.split('\\n')[0]:'Área Central'}} — maior volume diurno.</div>
        <div class="ok-box">✔ <strong>Turno B (19h–07h):</strong> concentrar ${{distB[0]?distB[0].viaturas:3}} viaturas na ${{distB[0]?distB[0].nome.split('\\n')[0]:'Área Central'}} — pico noturno e madrugada.</div>
        <div class="ok-box">✔ Criar operação toda <strong>${{topDias[0]?topDias[0][0]:'Sexta'}}</strong> a partir das ${{horasPico.split(',')[0]||'18h'}} com reforço de efetivo.</div>
        <div class="alerta-box">⚠ ${{dadosB.length}} dos ${{total}} crimes (${{((dadosB.length/total)*100).toFixed(0)}}%) ocorrem no Turno B. Não reduzir o efetivo noturno.</div>
      </div>
    </div>
  </div>`;
}}

// ── MAPA LEAFLET ─────────────────────────────────────────────────────────────
const TIPO_COLORS_MAP = {{
  'Furto':'#0078D4','Tentativa de Furto':'#50B2FF',
  'Arrombamento':'#E07B00','Roubo':'#D13438','Tentativa de Roubo':'#8764B8',
}};

let mapaInst = null;
let clusterLayer = null;

function initMapa() {{
  if(mapaInst || typeof L === 'undefined') return;
  try {{
    mapaInst = L.map('mapa-crime', {{
      center:[-26.993, -48.635], zoom:13,
      zoomControl:true, attributionControl:true,
    }});
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      attribution:'© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
      maxZoom:19,
    }}).addTo(mapaInst);
  }} catch(e) {{ console.error('Mapa init error:', e); mapaInst = null; }}
}}

function renderMapa(data) {{
  initMapa();
  if(!mapaInst) return;
  if(clusterLayer) {{ mapaInst.removeLayer(clusterLayer); }}

  clusterLayer = L.markerClusterGroup({{
    maxClusterRadius:40,
    iconCreateFunction: function(cluster) {{
      const c = cluster.getChildCount();
      const size = c < 5 ? 32 : c < 15 ? 40 : 50;
      return L.divIcon({{
        html:`<div style="background:#1A1A2E;color:white;border-radius:50%;
          width:${{size}}px;height:${{size}}px;display:flex;align-items:center;
          justify-content:center;font-weight:700;font-size:${{size>40?13:11}}px;
          border:3px solid white;box-shadow:0 2px 6px rgba(0,0,0,.4)">${{c}}</div>`,
        className:'',iconSize:[size,size],iconAnchor:[size/2,size/2],
      }});
    }},
  }});

  let comCoords = 0;
  data.forEach(r => {{
    if(!r.lat || !r.lon) return;
    comCoords++;
    const cor = TIPO_COLORS_MAP[r.tipo] || '#666';
    const marker = L.circleMarker([r.lat, r.lon], {{
      radius:8, fillColor:cor, color:'white',
      weight:2, opacity:1, fillOpacity:0.9,
    }});

    const dataFmt = r.data ? r.data.slice(8)+'/'+r.data.slice(5,7)+'/'+r.data.slice(0,4) : '–';
    marker.bindPopup(`
      <div>
        <span class="popup-tipo" style="background:${{cor}}">${{r.tipo}}</span>
        <div class="popup-row"><span class="popup-label">Data:</span> ${{dataFmt}} ${{r.hora}} (${{r.turno}})</div>
        <div class="popup-row"><span class="popup-label">Item:</span> ${{r.item}}${{r.marca?' – '+r.marca:''}}</div>
        <div class="popup-row"><span class="popup-label">Bairro:</span> ${{r.bairro}}</div>
        <div class="popup-row"><span class="popup-label">Endereço:</span> ${{r.endereco}}</div>
        <div class="popup-row"><span class="popup-label">Ref.:</span> ${{r.ref}}</div>
        <div class="popup-row" style="margin-top:6px;font-size:10px;color:#888">B.O.: ${{r.bo}}</div>
      </div>`, {{maxWidth:280}});
    clusterLayer.addLayer(marker);
  }});

  mapaInst.addLayer(clusterLayer);
  document.getElementById('mapa-stat').textContent =
    `${{comCoords}} de ${{data.length}} ocorrências com localização mapeada`;
}}

// ── SIDEBAR TOGGLE (MOBILE) ───────────────────────────────────────────────────
function toggleSidebar() {{
  const sb  = document.querySelector('.sidebar');
  const ov  = document.getElementById('sidebar-overlay');
  const open = sb.classList.toggle('open');
  ov.classList.toggle('show', open);
}}
window.addEventListener('resize', () => {{
  if(window.innerWidth > 768) {{
    document.querySelector('.sidebar').classList.remove('open');
    document.getElementById('sidebar-overlay').classList.remove('show');
  }}
  if(mapaInst) setTimeout(()=>mapaInst.invalidateSize(),200);
}});

// ── INIT ──────────────────────────────────────────────────────────────────────
function init() {{
  try {{ renderAll(); }} catch(e) {{ console.error('renderAll error:', e); }}
  setTimeout(() => {{ if(mapaInst) mapaInst.invalidateSize(); }}, 500);
}}

if(document.readyState === 'complete') {{ init(); }}
else {{ window.addEventListener('load', init); }}
</script>
<footer style="text-align:center;padding:12px 24px;font-size:11px;color:rgba(255,255,255,.75);
  background:linear-gradient(135deg,#1A1A2E 0%,#0078D4 100%);
  box-shadow:0 -2px 8px rgba(0,0,0,.3);">
  Desenvolvido por <strong style="color:white">Ronaldo Eliseu Barbosa</strong> — Guarda Municipal
</footer>
<!-- ── LOGIN ── -->
<div class="login-overlay" id="login-overlay">
  <div class="login-card" id="login-card">
    <div class="login-shield"><img src="data:image/png;base64,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" style="width:120px;height:120px;border-radius:50%;display:block;margin:0 auto 6px;box-shadow:0 4px 16px rgba(0,0,0,0.25)"></div>
    <div class="login-title">Secretaria de Segurança e Ordem Pública</div>
    <div class="login-sub">Guarda Municipal de Balneário Camboriú<br>
      <strong>Acesso restrito — uso interno</strong></div>
    <div class="login-input-wrap">
      <input type="password" class="login-input" id="login-senha"
        placeholder="Digite a senha de acesso"
        onkeydown="if(event.key==='Enter')verificarSenha()">
      <button type="button" class="login-eye" onclick="toggleSenha()" tabindex="-1" title="Mostrar/ocultar senha">👁</button>
    </div>
    <button class="login-btn" onclick="verificarSenha()">🔐 Entrar</button>
    <div class="login-error" id="login-erro"></div>
    <div style="margin-top:20px;text-align:center;font-size:10px;color:#aaa;line-height:1.5">
      Desenvolvido por <strong style="color:#888">Ronaldo E. Barbosa</strong> — Guarda Municipal
    </div>
  </div>
</div>
<!-- ── MODAL PESQUISA ── -->
<div class="analise-overlay" id="pesquisa-overlay" onclick="if(event.target===this)fecharPesquisa()">
  <div class="analise-box" style="width:min(860px,97vw);max-height:93vh">
    <div class="analise-header" style="background:linear-gradient(135deg,#1A3A5C 0%,#2E6DA4 100%)">
      <div>
        <h2>🔎 Pesquisa Geral</h2>
        <div style="font-size:10px;opacity:.85;margin-top:2px" id="pesquisa-subtitulo">Resultados da pesquisa</div>
      </div>
      <button class="analise-close" onclick="fecharPesquisa()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="pesquisa-corpo" style="padding:18px 20px;overflow-y:auto"></div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharPesquisa()">✕ Fechar</button>
    </div>
  </div>
</div>

<!-- ── MODAL INTELIGÊNCIA CRIMINAL ── -->
<div class="analise-overlay" id="intel-overlay" onclick="if(event.target===this)fecharInteligencia()">
  <div class="analise-box" style="width:min(1000px,97vw);max-height:93vh">
    <div class="analise-header" style="background:linear-gradient(135deg,#1B4332 0%,#2D6A4F 100%)">
      <div>
        <h2>🔍 Inteligência Criminal — Análise de Padrões por B.O.</h2>
        <div style="font-size:10px;opacity:.85;margin-top:2px">Secretaria de Segurança e Ordem Pública — Balneário Camboriú</div>
      </div>
      <button class="analise-close" onclick="fecharInteligencia()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="intel-corpo" style="padding:18px 20px;overflow-y:auto"></div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharInteligencia()">✕ Fechar</button>
      <button class="btn-pdf" onclick="imprimirInteligencia()">🖨️ Salvar PDF</button>
    </div>
  </div>
</div>

<!-- ── MODAL ANÁLISE PREDITIVA ── -->
<div class="analise-overlay" id="predit-overlay" onclick="if(event.target===this)fecharAnalisePredit()">
  <div class="analise-box" style="width:min(980px,97vw);max-height:93vh">
    <div class="analise-header" style="background:linear-gradient(135deg,#2D0B6E 0%,#7B2FBE 100%)">
      <div>
        <h2>🔮 Análise Preditiva Criminal — Inteligência Artificial</h2>
        <div style="font-size:10px;opacity:.85;margin-top:2px">Secretaria de Segurança e Ordem Pública — Balneário Camboriú</div>
      </div>
      <button class="analise-close" onclick="fecharAnalisePredit()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="predit-corpo" style="padding:18px 20px;overflow-y:auto"></div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharAnalisePredit()">✕ Fechar</button>
      <button class="btn-pdf" onclick="imprimirAnalisePredit()">🖨️ Salvar PDF</button>
    </div>
  </div>
</div>
<!-- ── MODAL RESUMO EXECUTIVO IA ── -->
<div class="analise-overlay" id="resumoia-overlay" onclick="if(event.target===this)fecharResumoIA()">
  <div class="analise-box" style="width:min(700px,95vw)">
    <div class="analise-header" style="background:linear-gradient(135deg,#0078D4 0%,#00B7C3 100%)">
      <div>
        <h2>🤖 Resumo Executivo — Inteligência Artificial</h2>
        <div style="font-size:10px;opacity:.85;margin-top:2px">Secretaria de Segurança e Ordem Pública — Balneário Camboriú</div>
      </div>
      <button class="analise-close" onclick="fecharResumoIA()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="resumoia-corpo" style="padding:18px 20px;overflow-y:auto">
      <div style="font-size:10px;color:#888;margin-bottom:14px" id="resumoia-meta"></div>
      <div id="resumoia-resultado" style="background:#F0F6FC;border-left:4px solid #0078D4;
        border-radius:8px;padding:14px 16px;font-size:13px;text-align:justify"></div>
    </div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharResumoIA()">✕ Fechar</button>
      <button id="btn-resumoia-refazer" class="btn-analise" data-label="🔄 Gerar novamente"
        onclick="explicarComIA(_resumoExecutivoDados,'resumoia-resultado','btn-resumoia-refazer')" style="background:#0078D4">🔄 Gerar novamente</button>
      <button class="btn-pdf" onclick="imprimirResumoIA()">🖨️ Salvar PDF</button>
    </div>
  </div>
</div>

<!-- ── CHAT IA FLUTUANTE ── -->
<button id="chat-ia-fab" onclick="toggleChatIA()" title="Pergunte à IA">💬</button>
<div id="chat-ia-panel" class="chat-ia-panel">
  <div class="chat-ia-header">
    <span>💬 Pergunte à IA</span>
    <button onclick="toggleChatIA()" title="Fechar" style="background:none;border:none;color:white;font-size:18px;cursor:pointer">✕</button>
  </div>
  <div id="chat-ia-log" class="chat-ia-log">
    <div class="chat-msg chat-msg-ia">Olá! Pergunte sobre as ocorrências — ex: "quantos furtos aconteceram no Centro?" ou "o que acontece mais aos sábados de madrugada?"</div>
  </div>
  <div class="chat-ia-inputbar">
    <button id="chat-ia-mic" type="button" onclick="toggleGravacao()" title="Falar">🎤</button>
    <textarea id="chat-ia-input" rows="3" placeholder="Digite ou cole sua pergunta... (Enter envia, Shift+Enter quebra linha)"
      onkeydown="if(event.key==='Enter' && !event.shiftKey){{ event.preventDefault(); enviarPerguntaChat(); }}"></textarea>
    <button onclick="enviarPerguntaChat()">Enviar</button>
  </div>
</div>
<!-- ── MODAL PREVISÃO ── -->
<div class="analise-overlay" id="prev-overlay" onclick="if(event.target===this)fecharPrevisao()">
  <div class="analise-box">
    <div class="analise-header">
      <h2 id="prev-titulo">📈 Previsão de Risco — Guarda Municipal BC</h2>
      <button class="analise-close" onclick="fecharPrevisao()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="prev-corpo"></div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharPrevisao()">✕ Fechar</button>
      <button class="btn-pdf" onclick="imprimirPrevisao()">🖨️ Salvar PDF</button>
      <button class="btn-analise" onclick="enviarWhatsAppPrevisao()">📱 Enviar WhatsApp</button>
    </div>
  </div>
</div>
<!-- ── MODAL RELATÓRIO DIÁRIO ── -->
<div class="analise-overlay" id="rel-overlay" onclick="if(event.target===this)fecharRelatorio()">
  <div class="analise-box" style="width:min(960px,97vw)">
    <div class="analise-header">
      <h2 id="rel-titulo">📅 Relatório de Ocorrências — Guarda Municipal BC</h2>
      <button class="analise-close" onclick="fecharRelatorio()" title="Fechar">✕</button>
    </div>
    <div class="rel-date-bar">
      <label>Data Inicial:</label>
      <input type="date" id="rel-data-ini" class="rel-date-input">
      <label>Data Final:</label>
      <input type="date" id="rel-data-fim" class="rel-date-input">
      <button class="rel-gerar-btn" onclick="gerarRelatorio()">🔍 Gerar Relatório</button>
    </div>
    <div class="analise-corpo" id="rel-corpo">
      <div style="color:#888;text-align:center;padding:40px;font-style:italic">
        Selecione o período e clique em <strong>Gerar Relatório</strong>.
      </div>
    </div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharRelatorio()">✕ Fechar</button>
      <button class="btn-pdf" onclick="imprimirRelatorio()">🖨️ Salvar PDF</button>
      <button class="btn-analise" onclick="enviarWhatsAppRelatorio()">📱 Enviar WhatsApp</button>
    </div>
  </div>
</div>
<!-- ── MODAL EXPORTAR PDF TABELA ── -->
<div id="pdf-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);
  z-index:9999;align-items:center;justify-content:center;">
  <div style="background:white;border-radius:10px;box-shadow:0 12px 48px rgba(0,0,0,.4);
    width:min(440px,94vw);overflow:hidden;">
    <div style="background:linear-gradient(135deg,#1A1A2E 0%,#107C10 100%);
      padding:14px 18px;display:flex;justify-content:space-between;align-items:center;">
      <span style="color:white;font-size:14px;font-weight:700">🖨️ Exportar PDF — Registros</span>
      <button onclick="fecharPdfModal()" style="background:none;border:none;color:white;
        font-size:22px;cursor:pointer;line-height:1;opacity:.8;padding:0 4px">✕</button>
    </div>
    <div style="padding:22px 24px;">
      <p style="margin:0 0 16px;font-size:12px;color:#555;line-height:1.6">
        Selecione o periodo para filtrar os registros antes de gerar o PDF.<br>
        Os filtros da barra lateral (bairro, tipo, item etc.) tambem sao aplicados.
      </p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:20px">
        <div>
          <label style="font-size:11px;font-weight:600;color:#1A1A2E;display:block;margin-bottom:5px">
            Data Inicial
          </label>
          <input type="date" id="pdf-data-ini" style="width:100%;box-sizing:border-box;
            padding:8px 10px;border:2px solid #ddd;border-radius:6px;font-size:13px;
            font-family:inherit;cursor:pointer">
        </div>
        <div>
          <label style="font-size:11px;font-weight:600;color:#1A1A2E;display:block;margin-bottom:5px">
            Data Final
          </label>
          <input type="date" id="pdf-data-fim" style="width:100%;box-sizing:border-box;
            padding:8px 10px;border:2px solid #ddd;border-radius:6px;font-size:13px;
            font-family:inherit;cursor:pointer">
        </div>
      </div>
      <div style="display:flex;gap:10px;justify-content:flex-end">
        <button onclick="fecharPdfModal()" style="padding:9px 20px;border:1px solid #ccc;
          border-radius:6px;background:white;font-size:12px;cursor:pointer;font-family:inherit;
          font-weight:600;color:#555">Cancelar</button>
        <button onclick="gerarPdfTabela()" style="padding:9px 20px;border:none;
          border-radius:6px;background:#107C10;color:white;font-size:12px;cursor:pointer;
          font-family:inherit;font-weight:700">🖨️ Gerar PDF</button>
      </div>
    </div>
  </div>
</div>

<!-- ── MODAL ANÁLISE DIÁRIA ── -->
<div class="analise-overlay" id="analise-overlay" onclick="if(event.target===this)fecharAnalise()">
  <div class="analise-box">
    <div class="analise-header">
      <h2 id="analise-titulo">📊 Análise por Dia da Semana — Guarda Municipal BC</h2>
      <button class="analise-close" onclick="fecharAnalise()" title="Fechar">✕</button>
    </div>
    <div class="analise-corpo" id="analise-corpo"></div>
    <div class="analise-footer">
      <button class="btn-reset" onclick="fecharAnalise()">✕ Fechar</button>
      <button class="btn-pdf" onclick="imprimirAnalise()">🖨️ Salvar PDF</button>
      <button class="btn-analise" onclick="enviarWhatsApp()">📱 Enviar WhatsApp</button>
    </div>
  </div>
</div>
</body>
</html>"""

# ── Substituir placeholders pelas bibliotecas embutidas ───────────────────────
html = html.replace('__PLOTLY_JS__',    _embed_js(plotly_js))
html = html.replace('__LEAFLET_CSS__',  _embed_css(leaflet_css))
html = html.replace('__LEAFLET_JS__',   _embed_js(leaflet_js))
html = html.replace('__CLUSTER_CSS1__', _embed_css(cluster_css1))
html = html.replace('__CLUSTER_CSS2__', _embed_css(cluster_css2))
html = html.replace('__CLUSTER_JS__',   _embed_js(cluster_js))
html = html.replace('__JSPDF_JS__',     _embed_js(jspdf_js))

with open('dashboard_interativo.html', 'w', encoding='utf-8') as f:
    f.write(html)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Dashboard interativo salvo: dashboard_interativo.html + index.html")
print(f"Tamanho: {len(html)//1024} KB")
