#!/usr/bin/env python3
"""
Analisador de B.O.s — Inteligência Criminal
Guarda Municipal de Balneário Camboriú
Cruza padrões de MO, suspeitos e localização entre B.O.s
"""

import os
import re
import json
from pathlib import Path
from datetime import datetime, timedelta
import pdfplumber

# ── Configurações ──────────────────────────────────────────────────────────────
PASTA_BOS   = r"C:\Users\rebar\Downloads\B.O. Furtos"
SAIDA_HTML  = r"C:\Users\rebar\Downloads\Secretario\inteligencia_criminal.html"
CACHE_JSON  = r"C:\Users\rebar\Downloads\Secretario\bos_cache.json"
GRUPOS_JSON = r"C:\Users\rebar\Downloads\Secretario\bos_grupos.json"

# ── Extração de texto ──────────────────────────────────────────────────────────
def extrair_texto(caminho):
    try:
        with pdfplumber.open(caminho) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except:
        return ""

def buscar_campo(texto, rotulo, fim=None):
    padrao = rf"{re.escape(rotulo)}\s*\n(.+?)(?=\n[A-ZÁÉÍÓÚÃÕÇ]|\Z)" if fim is None \
             else rf"{re.escape(rotulo)}\s*\n(.+?)(?={re.escape(fim)})"
    m = re.search(padrao, texto, re.DOTALL)
    return m.group(1).strip() if m else ""

# ── Parser de cada B.O. ────────────────────────────────────────────────────────
def parse_bo(caminho):
    texto = extrair_texto(caminho)
    if not texto:
        return None

    bo = {
        "arquivo":      os.path.basename(str(caminho)),
        "caminho":      str(caminho),
        "numero":       "",
        "data_fato":    "",
        "hora_fato":    "",
        "hora_int":     -1,
        "local":        "",
        "bairro":       "",
        "tipo_crime":   "",
        "relato":       "",
        "n_suspeitos":  0,
        "armado":       False,
        "veiculo":      "",
        "objetos":      [],
        "vitimas":      [],
        "mo_tags":      [],
    }

    # Número do BO
    m = re.search(r'Registro\s*\n([\w\-./]+)', texto)
    if m: bo["numero"] = m.group(1).strip()

    # Data do Fato
    m = re.search(r'Data do Fato\s*\n(\d{2}/\d{2}/\d{4})', texto)
    if m: bo["data_fato"] = m.group(1).strip()

    # Hora do Fato
    m = re.search(r'Hora do Fato\s*\n(\d{2}):(\d{2})', texto)
    if m:
        bo["hora_fato"] = f"{m.group(1)}:{m.group(2)}"
        bo["hora_int"]  = int(m.group(1))

    # Local
    m = re.search(r'Local do Fato\s*\n.*?\n(.+?)\n\d{5}', texto, re.DOTALL)
    if m: bo["local"] = m.group(1).strip()

    # Bairro (extrai do endereço "Rua X, N - Bairro")
    m = re.search(r'- ([A-ZÁÉÍÓÚÃÕÇa-záéíóúãõç][^\n\-]+?)\s*\n\d{5}', texto)
    if m: bo["bairro"] = m.group(1).strip()

    # Tipo de crime
    m = re.search(r'Fatos Comunicados\s*\n(.+?)\n', texto)
    if m: bo["tipo_crime"] = m.group(1).strip()

    # Relato Individual (texto livre — núcleo do MO)
    m = re.search(r'Relato Individual\s*\n(.*?)(?=Condi|ATEND|OBJETO|PROVID|$)', texto, re.DOTALL)
    if m: bo["relato"] = m.group(1).strip()

    relato = bo["relato"].lower()

    # ── Número de suspeitos ───────────────────────────────────────────────────
    for pat in [
        r'(\d+)\s+indiv[ií]duo', r'(\d+)\s+homens?', r'(\d+)\s+pessoas?',
        r'(\d+)\s+sujeitos?',    r'(\d+)\s+autores?', r'(\d+)\s+elementos?',
    ]:
        m2 = re.search(pat, relato)
        if m2:
            bo["n_suspeitos"] = int(m2.group(1))
            break
    if bo["n_suspeitos"] == 0:
        if any(p in relato for p in ["autor", "suspeito", "indivíduo", "elemento"]):
            bo["n_suspeitos"] = 1

    # ── Armamento ─────────────────────────────────────────────────────────────
    bo["armado"] = any(p in relato for p in [
        "armado", "arma de fogo", "revólver", "pistola", "espingarda",
        "faca", "simulacro", "garrucha", "arma branca"
    ])

    # ── Veículo de fuga ───────────────────────────────────────────────────────
    for v, label in [
        ("motocicleta","motocicleta"), ("moto ","moto"), ("moto,","moto"),
        ("carro","carro"),             ("veículo","veículo"),
        ("bicicleta","bicicleta"),     ("van","van"),
        ("a pé","a pé"),
    ]:
        if v in relato:
            bo["veiculo"] = label
            break

    # ── Tags de MO (padrões de comportamento) ─────────────────────────────────
    MO_DICT = {
        "perdeu":           "expressão 'perdeu'",
        "anunciou o assalto": "anúncio verbal",
        "rendeu":           "vítima rendida",
        "encapuzado":       "encapuzado",
        "capacete":         "capacete (não identificado)",
        "arrastão":         "arrastão",
        "distraiu":         "distração da vítima",
        "pix":              "solicitou PIX",
        "golpe":            "golpe/estelionato",
        "arrombou":         "arrombamento",
        "quebrou o vidro":  "quebra de vidro",
        "dormia":           "vítima dormia",
        "estacionado":      "veículo estacionado",
        "bolsa":            "furto de bolsa",
        "mochila":          "furto de mochila",
        "celular":          "celular furtado/roubado",
        "notebook":         "notebook furtado/roubado",
        "carteira":         "carteira furtada/roubada",
        "joias":            "joias furtadas",
        "bicicleta":        "bicicleta furtada",
    }
    for chave, tag in MO_DICT.items():
        if chave in relato and tag not in bo["mo_tags"]:
            bo["mo_tags"].append(tag)

    # ── Objetos subtraídos ────────────────────────────────────────────────────
    for m2 in re.finditer(r'Esp[eé]cie\s*\n(.+?)\n', texto):
        obj = m2.group(1).strip()
        if obj and obj not in bo["objetos"]:
            bo["objetos"].append(obj)

    # ── Marcas de objetos ─────────────────────────────────────────────────────
    marcas = []
    for m2 in re.finditer(r'Marca\s*\n(.+?)\n', texto):
        marca = m2.group(1).strip()
        if marca and marca.lower() not in ["não informado", "nao informado"]:
            marcas.append(marca)
    if marcas:
        bo["objetos_marcas"] = marcas

    # ── Vítimas ───────────────────────────────────────────────────────────────
    for m2 in re.finditer(r'Nome\s*\n([A-ZÁÉÍÓÚÃÕÇ][A-ZÁÉÍÓÚÃÕÇ\s]{3,})\nParticipa', texto):
        nome = m2.group(1).strip()
        if nome and nome not in bo["vitimas"]:
            bo["vitimas"].append(nome)

    # Turno
    h = bo["hora_int"]
    if   6 <= h < 12: bo["turno"] = "Manhã"
    elif 12 <= h < 18: bo["turno"] = "Tarde"
    elif 18 <= h < 24: bo["turno"] = "Noite"
    elif 0 <= h < 6:   bo["turno"] = "Madrugada"
    else:              bo["turno"] = "Não informado"

    return bo

# ── Similaridade entre dois B.O.s ─────────────────────────────────────────────
def similaridade(a, b):
    score = 0
    razoes = []

    if a["tipo_crime"] and a["tipo_crime"] == b["tipo_crime"]:
        score += 15; razoes.append(f"crime: {a['tipo_crime']}")

    if a["n_suspeitos"] > 0 and a["n_suspeitos"] == b["n_suspeitos"]:
        score += 25; razoes.append(f"{a['n_suspeitos']} suspeito(s)")

    if a["armado"] and b["armado"]:
        score += 15; razoes.append("ambos armados")

    if a["veiculo"] and a["veiculo"] == b["veiculo"]:
        score += 20; razoes.append(f"fuga: {a['veiculo']}")

    comuns = set(a["mo_tags"]) & set(b["mo_tags"])
    if comuns:
        score += min(len(comuns) * 8, 25)
        razoes.append("MO: " + ", ".join(list(comuns)[:3]))

    if a["bairro"] and a["bairro"] == b["bairro"]:
        score += 10; razoes.append(f"bairro: {a['bairro']}")

    obj_comuns = set(a["objetos"]) & set(b["objetos"])
    if obj_comuns:
        score += 10; razoes.append("objeto: " + ", ".join(list(obj_comuns)[:2]))

    return score, razoes

# ── Agrupamento por MO similar ─────────────────────────────────────────────────
def agrupar(bos, limiar=55):
    grupos = []
    associado = [False] * len(bos)

    for i in range(len(bos)):
        if associado[i]:
            continue
        grupo = [(bos[i], [], 100)]
        for j in range(i + 1, len(bos)):
            if associado[j]:
                continue
            sc, rz = similaridade(bos[i], bos[j])
            if sc >= limiar:
                grupo.append((bos[j], rz, sc))
                associado[j] = True
        if len(grupo) >= 2:
            associado[i] = True
            grupos.append(grupo)

    return sorted(grupos, key=lambda g: -len(g))

# ── HTML do relatório ──────────────────────────────────────────────────────────
COR_RISCO = {"Roubo": "#D13438", "Furto": "#E07B00", "default": "#555"}

def badge(texto, cor):
    return f'<span style="background:{cor};color:white;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700">{texto}</span>'

def cor_crime(tipo):
    for k, v in COR_RISCO.items():
        if k.lower() in tipo.lower():
            return v
    return COR_RISCO["default"]

def gerar_html(grupos, todos_bos, ts):
    total_vinculados = sum(len(g) for g in grupos)
    n_grupos = len(grupos)

    # Estatísticas globais
    bairros_freq = {}
    horas_freq = {}
    for bo in todos_bos:
        if bo["bairro"]: bairros_freq[bo["bairro"]] = bairros_freq.get(bo["bairro"], 0) + 1
        if bo["turno"]:  horas_freq[bo["turno"]]    = horas_freq.get(bo["turno"], 0) + 1

    top_bairros = sorted(bairros_freq.items(), key=lambda x: -x[1])[:6]
    top_turnos  = sorted(horas_freq.items(), key=lambda x: -x[1])

    # Cards dos grupos
    cards_html = ""
    for idx, grupo in enumerate(grupos, 1):
        bo_ref = grupo[0][0]
        n = len(grupo)
        risco_nivel = "CRÍTICO" if n >= 5 else "ALTO" if n >= 3 else "MÉDIO"
        risco_cor   = "#D13438" if n >= 5 else "#E07B00" if n >= 3 else "#E6A817"

        # Coletar datas do grupo
        datas = sorted(set(x[0]["data_fato"] for x in grupo if x[0]["data_fato"]))
        bairros_grupo = list(set(x[0]["bairro"] for x in grupo if x[0]["bairro"]))
        objetos_grupo = list(set(o for x in grupo for o in x[0]["objetos"]))[:5]
        mo_tags_grupo = list(set(t for x in grupo for t in x[0]["mo_tags"]))[:6]

        # Tabela de ocorrências
        linhas = ""
        for bo, razoes, sc in grupo:
            razoes_txt = " · ".join(razoes[:4]) if razoes else "referência"
            linhas += f"""
            <tr>
              <td style="padding:5px 8px;font-size:10px;font-family:monospace;color:#555">{bo["numero"] or bo["arquivo"]}</td>
              <td style="padding:5px 8px;font-size:10px">{bo["data_fato"]}</td>
              <td style="padding:5px 8px;font-size:10px">{bo["hora_fato"] or "?"} — {bo["turno"]}</td>
              <td style="padding:5px 8px;font-size:10px">{bo["local"][:40] if bo["local"] else bo["bairro"]}</td>
              <td style="padding:5px 8px;font-size:10px">{badge(bo["tipo_crime"], cor_crime(bo["tipo_crime"])) if bo["tipo_crime"] else "—"}</td>
              <td style="padding:5px 8px;font-size:10px;color:#555">{razoes_txt}</td>
            </tr>"""

        # Padrão identificado
        padrao_itens = []
        if bo_ref["n_suspeitos"] > 0: padrao_itens.append(f"👤 {bo_ref['n_suspeitos']} suspeito(s)")
        if bo_ref["armado"]:          padrao_itens.append("🔫 Armado(s)")
        if bo_ref["veiculo"]:         padrao_itens.append(f"🏍️ Fuga: {bo_ref['veiculo']}")
        if mo_tags_grupo:             padrao_itens.extend([f"• {t}" for t in mo_tags_grupo[:4]])

        padrao_html = " &nbsp;&nbsp; ".join(padrao_itens) if padrao_itens else "Padrão identificado por análise combinada"

        cards_html += f"""
        <div style="border:1px solid #E0E0E0;border-radius:10px;margin-bottom:20px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.07)">
          <div style="background:{risco_cor};padding:12px 16px;display:flex;justify-content:space-between;align-items:center">
            <div>
              <span style="color:white;font-size:14px;font-weight:700">Grupo #{idx} — {n} ocorrência(s) vinculadas</span>
              <span style="color:rgba(255,255,255,.85);font-size:11px;margin-left:12px">{badge(risco_nivel, 'rgba(0,0,0,.25)')}</span>
            </div>
            <div style="color:rgba(255,255,255,.9);font-size:11px;text-align:right">
              {datas[0] if datas else ''} → {datas[-1] if len(datas)>1 else ''}
            </div>
          </div>
          <div style="padding:14px 16px;background:#FAFAFA">
            <div style="margin-bottom:10px;font-size:11px;color:#333;line-height:1.8">
              <strong>🧩 Padrão:</strong> &nbsp; {padrao_html}
            </div>
            {'<div style="margin-bottom:10px;font-size:11px;color:#333"><strong>📦 Objetos:</strong> ' + ' · '.join(objetos_grupo) + '</div>' if objetos_grupo else ''}
            {'<div style="margin-bottom:10px;font-size:11px;color:#333"><strong>📍 Bairros:</strong> ' + ', '.join(bairros_grupo) + '</div>' if bairros_grupo else ''}
          </div>
          <div style="overflow-x:auto">
          <table style="width:100%;border-collapse:collapse;font-size:11px">
            <thead>
              <tr style="background:#F0F0F0">
                <th style="padding:6px 8px;text-align:left;font-size:10px">Nº B.O.</th>
                <th style="padding:6px 8px;text-align:left;font-size:10px">Data</th>
                <th style="padding:6px 8px;text-align:left;font-size:10px">Hora / Turno</th>
                <th style="padding:6px 8px;text-align:left;font-size:10px">Local</th>
                <th style="padding:6px 8px;text-align:left;font-size:10px">Crime</th>
                <th style="padding:6px 8px;text-align:left;font-size:10px">Similaridade</th>
              </tr>
            </thead>
            <tbody>{linhas}</tbody>
          </table>
          </div>
        </div>"""

    # Top bairros (barras)
    max_b = top_bairros[0][1] if top_bairros else 1
    bairros_bar = ""
    for b, c in top_bairros:
        pct = round(c / max_b * 100)
        bairros_bar += f"""
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:5px">
          <span style="width:120px;font-size:10px;color:#333;text-align:right;flex-shrink:0">{b}</span>
          <div style="flex:1;background:#EEE;border-radius:3px;height:18px">
            <div style="width:{pct}%;background:#D13438;height:18px;border-radius:3px"></div>
          </div>
          <span style="width:30px;font-size:10px;font-weight:700;color:#D13438">{c}</span>
        </div>"""

    # Turnos
    turno_html = ""
    CORES_T = {"Madrugada":"#7B2FBE","Manhã":"#0078D4","Tarde":"#E07B00","Noite":"#1A1A2E"}
    for t, c in top_turnos:
        turno_html += f'<div style="display:flex;justify-content:space-between;padding:4px 8px;font-size:11px;border-bottom:1px solid #F0F0F0"><span style="color:{CORES_T.get(t,"#333")};font-weight:600">{t}</span><span style="font-weight:700">{c}</span></div>'

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Inteligência Criminal — GMBC</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'Segoe UI',Arial,sans-serif; background:#F4F4F8; color:#1A1A2E; }}
  .topo {{ background:linear-gradient(135deg,#1A1A2E 0%,#2D0B6E 100%); color:white; padding:20px 32px; }}
  .topo h1 {{ font-size:20px; font-weight:700; }}
  .topo small {{ font-size:11px; opacity:.75; }}
  .container {{ max-width:1100px; margin:24px auto; padding:0 16px; }}
  .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin-bottom:24px; }}
  .kpi {{ background:white; border-radius:10px; padding:16px; box-shadow:0 2px 6px rgba(0,0,0,.06); }}
  .kpi .val {{ font-size:28px; font-weight:700; }}
  .kpi .lbl {{ font-size:11px; color:#777; margin-top:3px; }}
  .secao {{ background:white; border-radius:10px; padding:20px; margin-bottom:20px; box-shadow:0 2px 6px rgba(0,0,0,.06); }}
  .secao h2 {{ font-size:14px; font-weight:700; border-bottom:2px solid #2D0B6E; padding-bottom:6px; margin-bottom:14px; }}
  .aviso {{ background:#FFF8EC; border-left:4px solid #E07B00; padding:10px 14px; border-radius:4px; font-size:11px; color:#555; margin-bottom:20px; }}
  @media print {{ body{{background:white}} .topo{{-webkit-print-color-adjust:exact}} }}
</style>
</head>
<body>
<div class="topo">
  <h1>🔍 Inteligência Criminal — Análise de Padrões por B.O.</h1>
  <small>Guarda Municipal de Balneário Camboriú &nbsp;•&nbsp; Gerado em: {ts} &nbsp;•&nbsp; Base: {len(todos_bos)} B.O.s processados</small>
</div>
<div class="container">

  <div class="aviso">
    ⚠️ Este relatório é de uso exclusivo da Guarda Municipal. As vinculações são baseadas em similaridade estatística de padrões —
    não constituem prova de autoria. Use como apoio ao planejamento operacional e investigativo.
  </div>

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi"><div class="val">{len(todos_bos)}</div><div class="lbl">B.O.s analisados</div></div>
    <div class="kpi"><div class="val" style="color:#D13438">{n_grupos}</div><div class="lbl">Grupos com padrão similar</div></div>
    <div class="kpi"><div class="val" style="color:#E07B00">{total_vinculados}</div><div class="lbl">Ocorrências vinculadas</div></div>
    <div class="kpi"><div class="val" style="color:#7B2FBE">{len(todos_bos)-total_vinculados}</div><div class="lbl">Sem vínculo identificado</div></div>
  </div>

  <!-- Distribuição -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px">
    <div class="secao" style="margin:0">
      <h2>📍 Bairros com Mais Ocorrências</h2>
      {bairros_bar}
    </div>
    <div class="secao" style="margin:0">
      <h2>⏰ Distribuição por Turno</h2>
      {turno_html}
    </div>
  </div>

  <!-- Grupos -->
  <div class="secao">
    <h2>🔗 Grupos de Crimes com Padrão Similar (MO / Suspeitos)</h2>
    <p style="font-size:11px;color:#888;margin-bottom:14px">
      Agrupados por: mesmo número de suspeitos · mesmo veículo · mesmo MO · mesmos objetos · mesmo bairro.
      Grupos com ≥ 3 ocorrências merecem atenção prioritária.
    </p>
    {cards_html if cards_html else '<p style="color:#888;font-size:12px">Nenhum grupo com padrão similar encontrado com o limiar atual.</p>'}
  </div>

  <div style="text-align:center;font-size:10px;color:#aaa;padding:16px 0">
    Gerado por: Sistema de Inteligência Criminal — GMBC &nbsp;•&nbsp; {ts}
    <br>Secretaria de Segurança e Ordem Pública — Balneário Camboriú
  </div>
</div>
</body>
</html>"""

# ── Principal ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  ANALISADOR DE B.O.s — GUARDA MUNICIPAL BC")
    print("=" * 60)

    # Carrega cache se existir
    cache = {}
    if os.path.exists(CACHE_JSON):
        try:
            with open(CACHE_JSON, encoding="utf-8") as f:
                cache = {item["arquivo"]: item for item in json.load(f)}
            print(f"  Cache carregado: {len(cache)} B.O.s")
        except:
            pass

    # Coleta todos os PDFs
    pdfs = list(Path(PASTA_BOS).rglob("*.pdf"))
    print(f"  PDFs encontrados: {len(pdfs)}")

    bos = []
    novos = 0
    for i, pdf in enumerate(pdfs):
        nome = pdf.name
        if nome in cache:
            bos.append(cache[nome])
        else:
            print(f"  [{i+1}/{len(pdfs)}] Processando: {nome}", end="\r")
            bo = parse_bo(pdf)
            if bo:
                bos.append(bo)
                cache[nome] = bo
                novos += 1

    print(f"\n  Processados: {len(bos)} B.O.s ({novos} novos)")

    # Salva cache
    with open(CACHE_JSON, "w", encoding="utf-8") as f:
        json.dump(list(cache.values()), f, ensure_ascii=False, indent=2)

    # Agrupa por MO similar
    print("  Analisando padrões...")
    grupos = agrupar(bos, limiar=55)
    print(f"  Grupos identificados: {len(grupos)}")
    for g in grupos[:5]:
        print(f"    Grupo de {len(g)} crimes — ex: {g[0][0]['tipo_crime']} / {g[0][0]['bairro']}")

    # Gera HTML
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = gerar_html(grupos, bos, ts)
    with open(SAIDA_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    # Exporta JSON resumido para o dashboard (sem dados pessoais)
    bairros_freq = {}
    turnos_freq  = {}
    for bo in bos:
        if bo.get("bairro"): bairros_freq[bo["bairro"]] = bairros_freq.get(bo["bairro"], 0) + 1
        if bo.get("turno"):  turnos_freq[bo["turno"]]   = turnos_freq.get(bo["turno"], 0) + 1

    grupos_export = []
    for idx, grupo in enumerate(grupos, 1):
        bo_ref = grupo[0][0]
        n = len(grupo)
        datas   = sorted(set(x[0]["data_fato"] for x in grupo if x[0]["data_fato"]))
        bairros = list(set(x[0]["bairro"] for x in grupo if x[0]["bairro"]))
        objetos = list(set(o for x in grupo for o in x[0]["objetos"]))[:5]
        mo_tags = list(set(t for x in grupo for t in x[0]["mo_tags"]))[:6]
        crimes  = [{
            "numero":  x[0]["arquivo"].replace(".pdf",""),
            "data":    x[0]["data_fato"],
            "hora":    x[0]["hora_fato"],
            "turno":   x[0]["turno"],
            "local":   x[0]["local"][:50] if x[0]["local"] else x[0]["bairro"],
            "tipo":    x[0]["tipo_crime"],
            "razoes":  x[1][:4],
        } for x in grupo]
        risco = "CRÍTICO" if n >= 5 else "ALTO" if n >= 3 else "MÉDIO"
        grupos_export.append({
            "id": idx, "n": n, "risco": risco,
            "tipo_ref":    bo_ref["tipo_crime"],
            "n_suspeitos": bo_ref["n_suspeitos"],
            "armado":      bo_ref["armado"],
            "veiculo":     bo_ref["veiculo"],
            "datas":       datas,
            "bairros":     bairros,
            "objetos":     objetos,
            "mo_tags":     mo_tags,
            "crimes":      crimes,
        })

    total_vinculados = sum(len(g) for g in grupos)
    # Índice de busca nos relatos (para pesquisa livre no dashboard)
    relato_index = []
    for bo in bos:
        if not bo.get("relato"):
            continue
        relato_index.append({
            "numero": bo["arquivo"].replace(".pdf", ""),
            "data":   bo.get("data_fato", ""),
            "hora":   bo.get("hora_fato", ""),
            "turno":  bo.get("turno", ""),
            "local":  bo.get("local", ""),
            "bairro": bo.get("bairro", ""),
            "tipo":   bo.get("tipo_crime", ""),
            "relato": bo.get("relato", "")[:800],
        })

    dados_export = {
        "gerado_em":        ts,
        "total_bos":        len(bos),
        "total_grupos":     len(grupos),
        "total_vinculados": total_vinculados,
        "bairros_freq":     dict(sorted(bairros_freq.items(), key=lambda x: -x[1])[:8]),
        "turnos_freq":      turnos_freq,
        "grupos":           grupos_export,
        "relato_index":     relato_index,
    }
    with open(GRUPOS_JSON, "w", encoding="utf-8") as f:
        json.dump(dados_export, f, ensure_ascii=False, indent=2)
    print(f"  JSON dashboard salvo: {GRUPOS_JSON}")

    print(f"\n  Relatório salvo: {SAIDA_HTML}")
    print("=" * 60)

    # Abre no navegador
    import webbrowser
    webbrowser.open(SAIDA_HTML)

if __name__ == "__main__":
    main()
