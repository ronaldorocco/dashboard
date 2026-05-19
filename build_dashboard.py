import pandas as pd
import json
import os
import urllib.request
import warnings
warnings.filterwarnings('ignore')

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

def _embed_js(code):
    if not code:
        return ''
    return '<script>' + code.replace('</', '<\\/') + '</script>'

def _embed_css(code):
    return f'<style>{code}</style>' if code else ''

# ── Carregar geocache ─────────────────────────────────────────────────────────
GEOCACHE = {}
if os.path.exists('geocache.json'):
    with open('geocache.json', 'r', encoding='utf-8') as f:
        GEOCACHE = json.load(f)
    print(f"Geocache: {sum(1 for v in GEOCACHE.values() if v)} coordenadas carregadas")

# ── Carregar e limpar dados ───────────────────────────────────────────────────
df = pd.read_excel('secretario.xlsx', sheet_name='DADOS', engine='openpyxl')

def norm_tipo(v):
    v = str(v).strip().upper()
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
    u = str(v).strip().upper()
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
    u = str(v).strip().upper()
    return BAIRRO_MAP.get(u, str(v).strip().title())

DIA_MAP = {
    'SEGUNDA':'Segunda','TERÇA':'Terça','TERCA':'Terça','QUARTA':'Quarta',
    'QUINTA':'Quinta','SEXTA':'Sexta','SABADO':'Sábado','SÁBADO':'Sábado',
    'DOMINGO':'Domingo',
}
def norm_dia(v):
    u = str(v).strip().upper()
    return DIA_MAP.get(u, str(v).strip().title())

TURNO_MAP = {
    'MANHA':'Manhã','MANHÃ':'Manhã',
    'TARDE':'Tarde','NOITE':'Noite','MADRUGADA':'Madrugada',
}
def norm_turno(v):
    return TURNO_MAP.get(str(v).strip().upper(), str(v).strip().title())

MES_MAP = {
    'ABRIL':'Abril','MAIO':'Maio','MARCO':'Março','MARÇO':'Março',
    'JANEIRO':'Janeiro','FEVEREIRO':'Fevereiro','JUNHO':'Junho',
    'JULHO':'Julho','AGOSTO':'Agosto','SETEMBRO':'Setembro',
    'OUTUBRO':'Outubro','NOVEMBRO':'Novembro','DEZEMBRO':'Dezembro',
}
def norm_mes(v):
    return MES_MAP.get(str(v).strip().upper(), str(v).strip().title())

df['TIPIFICACAO'] = df['TIPIFICACAO'].apply(norm_tipo)
df['ITEM']        = df['ITEM'].apply(norm_item)
df['BAIRRO']      = df['BAIRRO'].apply(norm_bairro)
df['DIA_SEMANA']  = df['DIA_SEMANA'].apply(norm_dia)
df['MES']         = df['MES'].apply(norm_mes)
df['TURNO']       = df['TURNO'].apply(norm_turno)
df['DATA_COMPLETA'] = pd.to_datetime(
    df['ANO'].astype(str) + '-' +
    df['MES'].str.upper().map({'ABRIL':'04','MAIO':'05'}) + '-' +
    df['DATA'].astype(str).str.zfill(2),
    format='%Y-%m-%d', errors='coerce'
)
df['DATA_STR'] = df['DATA_COMPLETA'].dt.strftime('%Y-%m-%d')
df['HORA_STR'] = df['HORA'].apply(lambda x: x.strftime('%H:%M') if hasattr(x,'strftime') else str(x)[:5] if pd.notna(x) else '')
df['BO']       = df['B.O.'].fillna('').astype(str)
df['ENDERECO'] = df['ENDEREÇO'].fillna('').astype(str)
df['REF']      = df['PONTO_REFERENCIA'].fillna('').astype(str)
df['MAPA_URL'] = df['MAPA'].fillna('').astype(str)
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
        'marca': str(r['MARCA_MODELO']) if pd.notna(r['MARCA_MODELO']) else '',
        'qnt': str(r['QNT']) if pd.notna(r['QNT']) else '1',
        'endereco': r['ENDERECO'],
        'bairro': r['BAIRRO'],
        'ref': r['REF'],
        'mapa': r['MAPA_URL'],
        'lat': r['LAT'] if pd.notna(r['LAT']) else None,
        'lon': r['LON'] if pd.notna(r['LON']) else None,
    })

data_json = json.dumps(records, ensure_ascii=False)
data_json = data_json.replace('</', '<\\/')            # impede </script> fechar a tag
data_json = data_json.replace(' ', '\\u2028').replace(' ', '\\u2029')  # separadores de linha JS

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
  color:white;padding:14px 24px;display:flex;align-items:center;
  justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.3);z-index:10;
}}
.header h1{{font-size:17px;font-weight:700;letter-spacing:.3px;}}
.header .sub{{font-size:11px;opacity:.8;margin-top:2px;}}
.header-right{{display:flex;gap:10px;align-items:center;}}
.badge{{background:rgba(255,255,255,.2);border-radius:4px;padding:3px 10px;font-size:11px;font-weight:600;}}
.btn-reset{{background:#D13438;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-reset:hover{{background:#A4262C;}}
.btn-pdf{{background:#107C10;color:white;border:none;border-radius:4px;padding:5px 12px;
  font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}}
.btn-pdf:hover{{background:#0B5A0B;}}

/* IMPRESSÃO / PDF */
@media print{{
  *{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
  .sidebar,.sidebar-overlay,.menu-toggle,.header-right,.active-filters,
  .sel-hint,.btn-reset,.btn-pdf,.badge{{display:none!important;}}
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
.kpi-row{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:14px;}}
.kpi-card{{
  background:white;border-radius:8px;padding:14px 16px;
  border-top:4px solid var(--azul);box-shadow:0 2px 6px rgba(0,0,0,.07);
  transition:transform .15s,box-shadow .15s;cursor:default;
}}
.kpi-card:hover{{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.12);}}
.kpi-card.laranja{{border-color:var(--laranja);}}
.kpi-card.vermelho{{border-color:var(--vermelho);}}
.kpi-card.verde{{border-color:var(--verde);}}
.kpi-card.roxo{{border-color:var(--roxo);}}
.kpi-val{{font-size:30px;font-weight:700;color:var(--azul);line-height:1;}}
.kpi-card.laranja .kpi-val{{color:var(--laranja);}}
.kpi-card.vermelho .kpi-val{{color:var(--vermelho);}}
.kpi-card.verde .kpi-val{{color:var(--verde);}}
.kpi-card.roxo .kpi-val{{color:var(--roxo);}}
.kpi-label{{font-size:11px;color:#666;margin-top:5px;font-weight:500;}}
.kpi-sub{{font-size:10px;color:#999;margin-top:2px;}}
.kpi-media{{font-size:10px;color:#555;margin-top:4px;background:#F4F8FF;border-radius:4px;padding:2px 6px;display:inline-block;font-weight:500;}}

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
  .kpi-row{{grid-template-columns:repeat(3,1fr);}}
  .row3{{grid-template-columns:1fr 1fr;}}
  .row3 .chart-card:last-child{{grid-column:1/-1;}}
  .insight-grid{{grid-template-columns:repeat(2,1fr);}}
  .tipo-grid{{grid-template-columns:repeat(3,1fr);}}
  .plano-grid{{grid-template-columns:1fr 1fr;}}
  .sidebar{{width:200px;min-width:200px;}}
}}

/* Mobile (≤768px) */
@media(max-width:768px){{
  body{{font-size:13px;}}

  /* Header */
  .header{{padding:10px 14px;flex-wrap:wrap;gap:8px;}}
  .header h1{{font-size:14px;}}
  .header .sub{{font-size:10px;}}
  .header-right{{gap:6px;}}
  .badge{{display:none;}}
  .menu-toggle{{display:block;}}
  .sidebar-close{{display:block;}}

  /* Layout */
  .main-layout{{flex-direction:column;}}
  .sidebar{{
    position:fixed;top:0;left:-240px;width:240px;min-width:240px;
    height:100%;z-index:100;transition:left .3s ease;
    padding-top:56px;box-shadow:4px 0 20px rgba(0,0,0,.3);
  }}
  .sidebar.open{{left:0;}}
  .content{{padding:10px;overflow-y:auto;}}

  /* KPIs */
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:8px;}}
  .kpi-val{{font-size:24px;}}
  .kpi-label{{font-size:10px;}}

  /* Grids → coluna única */
  .row2,.row3,.row-full,
  .insight-grid,.rec-grid,
  .plano-grid,.plano-grid-wide{{
    grid-template-columns:1fr !important;
  }}
  .row3 .chart-card:last-child{{grid-column:auto;}}

  /* Tipificações */
  .tipo-grid{{grid-template-columns:repeat(2,1fr);}}

  /* Altura dos gráficos */
  #chart-bairro,#chart-item,#chart-ruas,#chart-refs{{height:300px !important;}}
  #chart-heatmap,#chart-hora{{height:220px !important;}}
  #mapa-crime{{height:320px !important;}}

  /* Tabela */
  .table-wrap{{font-size:10px;}}
  thead th{{padding:5px 6px;font-size:9px;}}
  tbody td{{padding:4px 6px;font-size:10px;}}

  /* Mapa legenda */
  .mapa-legenda{{flex-wrap:wrap;gap:6px;}}
  .mapa-header{{flex-direction:column;align-items:flex-start;gap:8px;}}

  /* Sidebar filtros */
  .filter-scroll{{max-height:130px;}}
  .section-header{{font-size:12px;}}

  /* Turno rows */
  .turno-label{{min-width:56px;}}
}}

/* Notebook / Desktop médio (769px–1024px) */
@media(min-width:769px) and (max-width:1024px){{
  .sidebar{{width:190px;min-width:190px;}}
  .kpi-row{{grid-template-columns:repeat(3,1fr);gap:10px;}}
  .kpi-val{{font-size:26px;}}
  .row3{{grid-template-columns:1fr 1fr;}}
  .row3 .chart-card:last-child{{grid-column:1/-1;}}
  .insight-grid{{grid-template-columns:repeat(2,1fr);}}
  .plano-grid{{grid-template-columns:1fr 1fr;}}
  .rec-grid{{grid-template-columns:1fr 1fr;}}
}}

/* Telas muito pequenas (≤480px) */
@media(max-width:480px){{
  .header h1{{font-size:12px;}}
  .kpi-row{{grid-template-columns:repeat(2,1fr);gap:6px;}}
  .kpi-val{{font-size:20px;}}
  .kpi-label{{font-size:9px;}}
  .content{{padding:8px;}}
  .rec-text{{font-size:11.5px;}}
  .btn-pdf,.btn-reset{{font-size:10px;padding:4px 8px;}}
  .badge{{font-size:10px;padding:2px 7px;}}
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
.leaflet-popup-content{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;min-width:200px;}}
.popup-tipo{{display:inline-block;border-radius:4px;padding:2px 8px;
  font-size:10px;font-weight:700;color:white;margin-bottom:6px;}}
.popup-row{{display:flex;gap:6px;margin-top:4px;color:#333;font-size:11px;}}
.popup-label{{color:#888;font-size:10px;min-width:50px;}}
</style>
</head>
<body>

<div id="__err__" style="display:none;position:fixed;top:0;left:0;right:0;background:#c00;color:#fff;padding:16px;font-family:monospace;font-size:13px;z-index:99999;white-space:pre-wrap"></div>
<script>window.onerror=function(m,s,l,c,e){{var d=document.getElementById('__err__');if(d){{d.style.display='block';d.textContent='ERRO JS: '+m+' (linha '+l+')\n'+(e&&e.stack?e.stack:'');}}return false;}};</script>
<div class="sidebar-overlay" id="sidebar-overlay" onclick="toggleSidebar()"></div>

<div class="header">
  <div style="display:flex;align-items:center;gap:10px">
    <button class="menu-toggle" onclick="toggleSidebar()" title="Filtros">☰</button>
    <div>
      <h1>🛡️ Dashboard de Segurança Pública – Balneário Camboriú</h1>
      <div class="sub">Boletins de Ocorrência · 2026 · 147 registros</div>
    </div>
  </div>
  <div class="header-right">
    <span class="badge" id="badge-total">147 ocorrências</span>
    <button class="btn-reset" onclick="resetFilters()">⟳ Limpar Filtros</button>
    <button class="btn-pdf" onclick="gerarPDF()">🖨️ Gerar PDF</button>
  </div>
</div>

<div class="main-layout">

  <!-- SIDEBAR -->
  <div class="sidebar" id="sidebar">
    <div class="sidebar-topbar">
      <span>🔎 Filtros</span>
      <button class="sidebar-close" onclick="toggleSidebar()" title="Fechar filtros">✕</button>
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
  </div>

  <!-- CONTENT -->
  <div class="content">

    <!-- Filtros ativos -->
    <div class="active-filters" id="active-filters"></div>

    <!-- KPIs -->
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="kpi-val" id="kpi-total">147</div>
        <div class="kpi-label">Total de Ocorrências</div>
        <div class="kpi-sub" id="kpi-total-sub">2026</div>
        <div class="kpi-media" id="kpi-total-media"></div>
      </div>
      <div class="kpi-card laranja">
        <div class="kpi-val" id="kpi-furtos">133</div>
        <div class="kpi-label">Furtos</div>
        <div class="kpi-sub" id="kpi-furtos-sub">90,5% do total</div>
        <div class="kpi-media" id="kpi-furtos-media"></div>
      </div>
      <div class="kpi-card vermelho">
        <div class="kpi-val" id="kpi-roubos">4</div>
        <div class="kpi-label">Roubos</div>
        <div class="kpi-sub" id="kpi-roubos-sub">2,7% do total</div>
      </div>
      <div class="kpi-card roxo">
        <div class="kpi-val" id="kpi-turno">Noite</div>
        <div class="kpi-label">Turno Mais Crítico</div>
        <div class="kpi-sub" id="kpi-turno-sub"></div>
      </div>
      <div class="kpi-card verde">
        <div class="kpi-val" id="kpi-bairro">Centro</div>
        <div class="kpi-label">Bairro Mais Afetado</div>
        <div class="kpi-sub" id="kpi-bairro-sub"></div>
      </div>
    </div>

    <!-- Linha 1: Tipificação | Bairros | Turno -->
    <div class="row row3">
      <div class="chart-card">
        <div class="chart-title"><span class="icon">🥧</span> Tipificação</div>
        <div id="chart-tipo" style="height:220px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📍</span> Ocorrências por Bairro</div>
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
        <div class="chart-title"><span class="icon">📦</span> Top Itens Furtados / Roubados</div>
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
        <div class="chart-title"><span class="icon">🛣️</span> Top 15 Ruas com Mais Ocorrências</div>
        <div id="chart-ruas" style="height:340px"></div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><span class="icon">📌</span> Top 15 Pontos de Referência</div>
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
      <div class="chart-title" style="margin-bottom:8px">
        <span class="icon">📋</span> Registros Detalhados
        <span style="font-size:10px;color:#888;font-weight:400;margin-left:8px" id="tabela-count"></span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Data</th><th>Hora</th><th>Turno</th><th>Dia</th>
              <th>B.O.</th><th>Tipificação</th><th>Item</th>
              <th>Marca/Modelo</th><th>Qtd</th><th>Bairro</th><th>Endereço</th><th>Referência</th>
            </tr>
          </thead>
          <tbody id="tabela-body"></tbody>
        </table>
      </div>
      <div class="sel-hint">Clique em qualquer gráfico para filtrar os dados</div>
    </div>

  </div>
</div>

<script>
// ── DADOS ────────────────────────────────────────────────────────────────────
const RAW = {data_json};

// ── ESTADO DOS FILTROS ────────────────────────────────────────────────────────
const state = {{
  mes: new Set(), turno: new Set(), tipo: new Set(),
  bairro: new Set(), item: new Set(), dia: new Set()
}};

// ── CORES ─────────────────────────────────────────────────────────────────────
const COLORS = {{
  azul:'#0078D4', azulClr:'#50B2FF', laranja:'#E07B00',
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
    (state.mes.size   === 0 || state.mes.has(r.mes))   &&
    (state.turno.size === 0 || state.turno.has(r.turno))&&
    (state.tipo.size  === 0 || state.tipo.has(r.tipo))  &&
    (state.bairro.size=== 0 || state.bairro.has(r.bairro))&&
    (state.item.size  === 0 || state.item.has(r.item))  &&
    (state.dia.size   === 0 || state.dia.has(r.dia))
  );
}}

// ── CONTAR ────────────────────────────────────────────────────────────────────
function count(data, key) {{
  const m = {{}};
  data.forEach(r => {{ const v=r[key]; m[v]=(m[v]||0)+1; }});
  return m;
}}
function sortedEntries(obj, desc=true) {{
  return Object.entries(obj).sort((a,b)=>desc?b[1]-a[1]:a[1]-b[1]);
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
  const entries = sortedEntries(c);
  const labels = entries.map(e=>e[0]).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:10}},automargin:true}},
    margin:{{l:80,r:30,t:4,b:24}},
  }};
  Plotly.react('chart-bairro', barH(labels,vals,colors), layout, CONFIG);
  document.getElementById('chart-bairro').on('plotly_click', e => {{
    toggleFilter('bairro', e.points[0].y);
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

  const c = {{}};
  data.forEach(r => {{ if(r.data) c[r.data]=(c[r.data]||0)+1; }});
  const dates = Object.keys(c).sort();
  const vals  = dates.map(d=>c[d]);
  const trace = {{
    type:'scatter', mode:'lines+markers+text',
    x:dates, y:vals,
    line:{{color:COLORS.azul,width:2.5,shape:'spline'}},
    marker:{{color:COLORS.azulClr,size:7,line:{{color:COLORS.azul,width:1.5}}}},
    fill:'tozeroy', fillcolor:'rgba(0,120,212,.1)',
    text:vals.map(String), textposition:'top center', textfont:{{size:9,color:COLORS.azul}},
    hovertemplate:'<b>%{{x}}</b><br>%{{y}} casos<extra></extra>',
  }};
  const layout = {{...LAYOUT_BASE,
    xaxis:{{
      tickformat:'%d/%m', tickfont:{{size:9}},
      dtick:'D2',showgrid:false,
    }},
    yaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:9}},range:[0,Math.max(...vals,1)*1.20]}},
    margin:{{l:28,r:12,t:4,b:36}},
  }};
  Plotly.react('chart-linha',[trace],layout,CONFIG);
}}

// ── GRÁFICO: MÊS ─────────────────────────────────────────────────────────────
function renderMes(data) {{
  if(typeof Plotly === 'undefined') return;

  const ORDER = ['Abril','Maio'];
  const c = count(data,'mes');
  const labels = ORDER.filter(m=>c[m]);
  const vals   = labels.map(m=>c[m]||0);
  const maxMes = Math.max(...vals);
  const colBase = [COLORS.azul, COLORS.laranja];
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
  const turnos  = count(data,'turno');
  const bairros = count(data,'bairro');
  const topTurno  = sortedEntries(turnos)[0]||['–',0];
  const topBairro = sortedEntries(bairros)[0]||['–',0];
  const pFurtos = total ? ((furtos/total)*100).toFixed(1) : '0.0';
  const pRoubos = total ? ((roubos/total)*100).toFixed(1) : '0.0';

  // Calcular média diária
  const datas = data.map(r=>r.data).filter(Boolean).sort();
  let nDias = 1, mediaLabel = '';
  if(datas.length > 0) {{
    const d0 = new Date(datas[0]), d1 = new Date(datas[datas.length-1]);
    nDias = Math.max(1, Math.round((d1-d0)/86400000)+1);
    const mediaTotal  = (total/nDias).toFixed(1);
    const mediaFurtos = (furtos/nDias).toFixed(1);
    document.getElementById('kpi-total-media').textContent  = `📅 ${{mediaTotal}}/dia · ${{nDias}} dias`;
    document.getElementById('kpi-furtos-media').textContent = `📅 ${{mediaFurtos}}/dia`;
  }} else {{
    document.getElementById('kpi-total-media').textContent  = '';
    document.getElementById('kpi-furtos-media').textContent = '';
  }}

  document.getElementById('kpi-total').textContent   = total;
  document.getElementById('kpi-furtos').textContent  = furtos;
  document.getElementById('kpi-roubos').textContent  = roubos;
  document.getElementById('kpi-turno').textContent   = topTurno[0];
  document.getElementById('kpi-bairro').textContent  = topBairro[0];
  document.getElementById('kpi-furtos-sub').textContent  = pFurtos + '% do total';
  document.getElementById('kpi-roubos-sub').textContent  = pRoubos + '% do total';
  document.getElementById('kpi-turno-sub').textContent   = topTurno[1] + ' casos';
  document.getElementById('kpi-bairro-sub').textContent  = topBairro[1] + ' casos';
  document.getElementById('badge-total').textContent = total + ' ocorrências';
}}

// ── TABELA ────────────────────────────────────────────────────────────────────
function renderTabela(data) {{
  const tbody = document.getElementById('tabela-body');
  const show  = data.slice(0,200);
  document.getElementById('tabela-count').textContent =
    `${{data.length}} registros${{data.length>200?' (mostrando primeiros 200)':''}}`;
  tbody.innerHTML = show.map(r => `
    <tr>
      <td>${{r.data ? r.data.slice(8)+'/'+(r.mes==='Abril'?'04':'05')+'/'+r.ano : ''}}</td>
      <td>${{r.hora}}</td>
      <td>${{r.turno}}</td>
      <td>${{r.dia}}</td>
      <td style="font-size:9px">${{r.bo}}</td>
      <td><span style="background:${{TIPO_COLORS[r.tipo]||'#888'}};color:white;
        border-radius:3px;padding:1px 5px;font-size:9px">${{r.tipo}}</span></td>
      <td>${{r.item}}</td>
      <td>${{r.marca}}</td>
      <td style="text-align:center">${{r.qnt}}</td>
      <td>${{r.bairro}}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis">${{r.endereco}}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:10px">${{r.ref}}</td>
    </tr>`).join('');
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
            <span class="insight-highlight">${{(total/16).toFixed(1)}} casos/dia</span>,
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
  renderKPIs(data);
  renderTipo(data);
  renderBairro(data);
  renderTurno(data);
  renderLinha(data);
  renderMes(data);
  renderItem(data);
  renderDia(data);
  renderRuas(data);
  renderRefs(data);
  renderHeatmap(data);
  renderHora(data);
  renderPlanoPolicial(data);
  renderDiagnostico(data);
  renderRecomendacoes(data);
  renderInsights(data);
  renderMapa(data);
  renderTabela(data);
  renderChips();
  buildSidebar();
}}

// ── CHIPS DE FILTROS ATIVOS ───────────────────────────────────────────────────
const FILTER_LABELS = {{mes:'Mês',turno:'Turno',tipo:'Tipo',bairro:'Bairro',item:'Item',dia:'Dia'}};
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
  renderAll();
}}

// ── GERAR PDF ─────────────────────────────────────────────────────────────────
function gerarPDF() {{
  // Garante que o mapa está dimensionado corretamente antes de imprimir
  if(mapaInst) mapaInst.invalidateSize();
  // Pequena pausa para os gráficos renderizarem completamente
  setTimeout(() => window.print(), 300);
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
  const map = {{bairro:'filter-bairro',item:'filter-item'}};
  const el = document.getElementById(map[key]);
  if(el) el.dataset.search = q;
  buildSidebar();
}}

function buildSidebar() {{
  const all = filtered();
  const allCounts = k => count(RAW,k);

  const meses  = [...new Set(RAW.map(r=>r.mes))].sort();
  const turnos = ['Manhã','Tarde','Noite','Madrugada'];
  const tipos  = [...new Set(RAW.map(r=>r.tipo))].sort();
  const bairros= sortedEntries(count(RAW,'bairro')).map(e=>e[0]);
  const itens  = sortedEntries(count(RAW,'item')).map(e=>e[0]);
  const dias   = DIA_ORDER.filter(d=>RAW.some(r=>r.dia===d));

  buildCheckboxes('filter-mes',   'mes',   meses,  allCounts('mes'));
  buildCheckboxes('filter-turno', 'turno', turnos, allCounts('turno'));
  buildCheckboxes('filter-tipo',  'tipo',  tipos,  allCounts('tipo'));
  buildCheckboxes('filter-bairro','bairro',bairros,allCounts('bairro'));
  buildCheckboxes('filter-item',  'item',  itens,  allCounts('item'));
  buildCheckboxes('filter-dia',   'dia',   dias,   allCounts('dia'));
}}

// ── RUAS ─────────────────────────────────────────────────────────────────────
function renderRuas(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'endereco');
  const entries = sortedEntries(c).filter(e=>e[0]&&e[0]!=='').slice(0,15);
  const labels = entries.map(e=>e[0]).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:i>labels.length-4?COLORS.laranja:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:10}},automargin:true}},
    margin:{{l:160,r:36,t:4,b:24}},
  }};
  Plotly.react('chart-ruas', barH(labels,vals,colors), layout, CONFIG);
}}

// ── PONTOS DE REFERÊNCIA ──────────────────────────────────────────────────────
function renderRefs(data) {{
  if(typeof Plotly === 'undefined') return;

  const c = count(data,'ref');
  const entries = sortedEntries(c).filter(e=>e[0]&&e[0]!=='').slice(0,15);
  const labels = entries.map(e=>e[0]).map(l=>l.length>35?l.slice(0,33)+'…':l).reverse();
  const vals   = entries.map(e=>e[1]).reverse();
  const colors = labels.map((_,i)=>i===labels.length-1?COLORS.vermelho:i>labels.length-4?COLORS.laranja:COLORS.azulClr);
  const layout = {{...LAYOUT_BASE,
    xaxis:{{showgrid:true,gridcolor:'#F0F0F0',zeroline:false,tickfont:{{size:10}}}},
    yaxis:{{tickfont:{{size:9.5}},automargin:true}},
    margin:{{l:220,r:36,t:4,b:24}},
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

// ── PLANO POLICIAL DINÂMICO ───────────────────────────────────────────────────
function renderPlanoPolicial(data) {{
  const total = data.length;
  if(total===0){{ document.getElementById('plano-policial').innerHTML=''; return; }}

  // ── Separar os dois turnos operacionais (usa campo turno, igual ao gráfico) ──
  // Turno A: Manhã + Tarde → 8 viaturas
  // Turno B: Noite + Madrugada → 7 viaturas
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
</body>
</html>"""

# ── Substituir placeholders pelas bibliotecas embutidas ───────────────────────
html = html.replace('__PLOTLY_JS__',    _embed_js(plotly_js))
html = html.replace('__LEAFLET_CSS__',  _embed_css(leaflet_css))
html = html.replace('__LEAFLET_JS__',   _embed_js(leaflet_js))
html = html.replace('__CLUSTER_CSS1__', _embed_css(cluster_css1))
html = html.replace('__CLUSTER_CSS2__', _embed_css(cluster_css2))
html = html.replace('__CLUSTER_JS__',   _embed_js(cluster_js))

with open('dashboard_interativo.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Dashboard interativo salvo: dashboard_interativo.html")
print(f"Tamanho: {len(html)//1024} KB")
