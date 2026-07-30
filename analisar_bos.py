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

def _calcular_idade(data_nasc, data_fato):
    try:
        d1 = datetime.strptime(data_nasc, "%d/%m/%Y")
        d2 = datetime.strptime(data_fato, "%d/%m/%Y")
        idade = d2.year - d1.year - ((d2.month, d2.day) < (d1.month, d1.day))
        return idade if 0 <= idade <= 110 else None
    except ValueError:
        return None

# ── Extração de autores (ENVOLVIDOS com participação "Autor") ─────────────────
# Usada só para alimentar contagens agregadas no chat de Inteligência Criminal
# (inteligencia_criminal.html, local/gitignored) — nome, CPF, RG, telefone e
# endereço completo nunca são incluídos no retorno.
def extrair_autores(texto, data_fato):
    m = re.search(r'ENVOLVIDOS\s*\n(.*?)(?=\nATENDENTES\b|\nOBJETOS\b|\Z)', texto, re.DOTALL)
    if not m:
        return []
    bloco_geral = "\n" + m.group(1)
    inicios = [mm.start() for mm in re.finditer(r'\nNome\s*\n', bloco_geral)]

    autores = []
    for i, ini in enumerate(inicios):
        fim = inicios[i + 1] if i + 1 < len(inicios) else len(bloco_geral)
        bloco = bloco_geral[ini:fim]
        participacoes = buscar_campo(bloco, "Participações")
        if "autor" not in participacoes.lower():
            continue

        naturalidade = buscar_campo(bloco, "Naturalidade")
        cidade_nat, uf_nat = naturalidade, ""
        if " - " in naturalidade:
            cidade_nat, uf_nat = [p.strip() for p in naturalidade.rsplit(" - ", 1)]

        idade = _calcular_idade(buscar_campo(bloco, "Data de Nascimento"), data_fato)

        bairro_resid, cidade_resid = "", ""
        me = re.search(
            r'Endere[çc]o\(s\)\s*\n.*?\n.+?-\s*([^\n]+?)\s*\n\d{5}-?\d{3}\s+([A-ZÁÉÍÓÚÃÕÇ\s]+?)\s*-\s*[A-Z]{2}',
            bloco, re.DOTALL,
        )
        if me:
            bairro_resid = me.group(1).strip().title()
            cidade_resid = me.group(2).strip().title()

        crime_m = re.search(r'Autor\s*\(([^)]+)\)', participacoes)

        autores.append({
            "naturalidade_cidade": cidade_nat.strip().title() if cidade_nat.strip() else "Não Informado",
            "naturalidade_uf":     uf_nat.strip().upper(),
            "sexo":                buscar_campo(bloco, "Sexo") or "Não Informado",
            "estado_civil":        buscar_campo(bloco, "Estado Civil") or "Não Informado",
            "profissao":           buscar_campo(bloco, "Profissão") or "Não Informado",
            "crime":               crime_m.group(1).strip() if crime_m else "",
            "idade":               idade,
            "menor":               idade is not None and idade < 18,
            "bairro_residencia":   bairro_resid or "Não Informado",
            "cidade_residencia":   cidade_resid or "Não Informado",
        })
    return autores

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
        "autores":      [],
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

    # Autores (ENVOLVIDOS com participação "Autor") — dados demográficos
    # agregados para o chat de Inteligência Criminal
    bo["autores"] = extrair_autores(texto, bo["data_fato"])

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

def gerar_html(grupos, todos_bos, ts, autores):
    total_vinculados = sum(len(g) for g in grupos)
    n_grupos = len(grupos)
    # Só campos demográficos agregáveis viajam para o JS da página — nunca
    # nome, CPF, RG, telefone ou endereço completo (ver extrair_autores()).
    autores_json = json.dumps(autores, ensure_ascii=False)

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
  .chat-ia-panel {{ position:fixed; top:0; right:-380px; width:360px; max-width:92vw; height:100vh; background:white;
    box-shadow:-4px 0 24px rgba(0,0,0,.25); z-index:9999; display:flex; flex-direction:column;
    transition:right .25s ease; font-family:inherit; }}
  .chat-ia-panel.aberto {{ right:0; }}
  .chat-ia-header {{ background:linear-gradient(135deg,#2D0B6E,#7B2FBE); color:white; padding:14px 16px;
    font-weight:700; font-size:14px; display:flex; justify-content:space-between; align-items:center; flex-shrink:0; }}
  .chat-ia-log {{ flex:1; overflow-y:auto; padding:14px; display:flex; flex-direction:column; gap:10px; }}
  .chat-msg {{ padding:9px 12px; border-radius:10px; font-size:12px; line-height:1.5; max-width:88%; }}
  .chat-msg-user {{ background:#2D0B6E; color:white; align-self:flex-end; border-bottom-right-radius:2px; }}
  .chat-msg-ia {{ background:#F0F0F0; color:#222; align-self:flex-start; border-bottom-left-radius:2px; text-align:justify; white-space:pre-line; }}
  .chat-ia-inputbar {{ display:flex; gap:8px; padding:12px; border-top:1px solid #EEE; flex-shrink:0; align-items:flex-end; }}
  .chat-ia-inputbar textarea {{ flex:1; border:1px solid #DDD; border-radius:6px; padding:8px 10px; font-size:12px;
    font-family:inherit; min-width:0; resize:vertical; min-height:64px; max-height:200px; line-height:1.4; }}
  .chat-ia-inputbar button {{ background:#2D0B6E; color:white; border:none; border-radius:6px; padding:8px 14px;
    font-size:12px; font-weight:600; cursor:pointer; flex-shrink:0; }}
  .chat-ia-inputbar button:hover {{ background:#210853; }}
  #chat-ia-fab {{ position:fixed; bottom:24px; right:24px; width:56px; height:56px; border-radius:50%;
    background:linear-gradient(135deg,#2D0B6E,#7B2FBE); color:white; border:none; font-size:24px;
    box-shadow:0 4px 16px rgba(0,0,0,.3); cursor:pointer; z-index:9998; }}
  #chat-ia-fab:hover {{ background:linear-gradient(135deg,#210853,#6425a0); }}
</style>
</head>
<body>
<div class="topo">
  <h1>🔍 Inteligência Criminal — Análise de Padrões por B.O.</h1>
  <small>Guarda Municipal de Balneário Camboriú &nbsp;•&nbsp; Gerado em: {ts} &nbsp;•&nbsp; Base: {len(todos_bos)} B.O.s processados &nbsp;•&nbsp; {len(autores)} autores com perfil extraído</small>
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
    <div class="kpi"><div class="val" style="color:#2D0B6E">{len(autores)}</div><div class="lbl">Autores com perfil extraído</div></div>
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

<!-- ── CHAT IA FLUTUANTE — Perfil dos Autores (só contagens agregadas, sem PII) ── -->
<button id="chat-ia-fab" onclick="toggleChatIA()" title="Fale com a Ana">💬</button>
<div id="chat-ia-panel" class="chat-ia-panel">
  <div class="chat-ia-header">
    <span>💬 Ana — Assistente Virtual</span>
    <button onclick="toggleChatIA()" title="Fechar" style="background:none;border:none;color:white;font-size:18px;cursor:pointer">✕</button>
  </div>
  <div id="chat-ia-log" class="chat-ia-log">
    <div class="chat-msg chat-msg-ia">Olá, eu sou a Ana, a assistente virtual da Guarda Municipal de Balneário Camboriú. Qual o seu nome?</div>
  </div>
  <div class="chat-ia-inputbar">
    <textarea id="chat-ia-input" rows="3" placeholder="Digite sua pergunta... (Enter envia, Shift+Enter quebra linha)"
      onkeydown="if(event.key==='Enter' && !event.shiftKey){{ event.preventDefault(); enviarPerguntaChat(); }}"></textarea>
    <button onclick="enviarPerguntaChat()">Enviar</button>
  </div>
</div>
<script>
// Só campos demográficos agregáveis — nunca nome, CPF, RG, telefone ou
// endereço completo (ver extrair_autores() em analisar_bos.py). Toda a
// contagem roda localmente no navegador; nenhum dado é enviado para fora.
const AUTORES = {autores_json};

function _normalizarTexto(s) {{
  return (s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'');
}}

// Alguns PDFs têm letras acentuadas que o extrator não decodifica (viram
// '�'). Para não perder o match de nomes como "Balneário Camboriú",
// tratamos esse caractere como um coringa de 1 posição na comparação.
function _correspondeValor(valorArmazenado, textoNormalizado) {{
  const base = _normalizarTexto(valorArmazenado).replace(/[.*+?^${{}}()|[\]\\\\]/g, '\\\\$&');
  const padrao = base.replace(/\\uFFFD/g, '.');
  try {{
    return new RegExp(padrao).test(textoNormalizado);
  }} catch (e) {{
    return textoNormalizado.includes(_normalizarTexto(valorArmazenado));
  }}
}}

function _topEntradas(lista, campo, n) {{
  const c = {{}};
  lista.forEach(a => {{ const v = a[campo]; if (v && v !== 'Não Informado') c[v] = (c[v]||0) + 1; }});
  return Object.entries(c).sort((x,y) => y[1]-x[1]).slice(0, n || 5);
}}

function _fmtTop(pares) {{
  return pares.length ? pares.map(([k,v]) => `${{k}} (${{v}})`).join(', ') : 'sem dados suficientes';
}}

function responderPerguntaAutores(pergunta) {{
  if (!AUTORES.length) {{
    return 'Nenhum autor com perfil identificado nos B.O.s processados até o momento.';
  }}
  const pNorm = _normalizarTexto(pergunta);

  const listaUnica = campo => [...new Set(AUTORES.map(a => a[campo]).filter(v => v && v !== 'Não Informado'))];
  const categorias = [
    {{ campo:'naturalidade_cidade', valores:listaUnica('naturalidade_cidade') }},
    {{ campo:'cidade_residencia',   valores:listaUnica('cidade_residencia') }},
    {{ campo:'bairro_residencia',   valores:listaUnica('bairro_residencia') }},
    {{ campo:'profissao',          valores:listaUnica('profissao') }},
    {{ campo:'estado_civil',       valores:listaUnica('estado_civil') }},
    {{ campo:'sexo',               valores:listaUnica('sexo') }},
    {{ campo:'crime',              valores:listaUnica('crime') }},
  ];

  let subset = AUTORES;
  const criteriosEncontrados = [];
  categorias.forEach(({{campo, valores}}) => {{
    const achados = valores.filter(v => _correspondeValor(v, pNorm));
    if (achados.length) {{
      subset = subset.filter(a => achados.includes(a[campo]));
      criteriosEncontrados.push(...achados);
    }}
  }});
  const pedeMenor = /\\bmenor(es)?\\b|adolescente/.test(pNorm);
  if (pedeMenor) {{
    subset = subset.filter(a => a.menor);
    criteriosEncontrados.push('menor de idade');
  }}

  const pedeRanking = /mais comum|mais frequente|qual.*profiss|qual.*naturalidade|ranking|top\\s*\\d*/.test(pNorm);
  if (!criteriosEncontrados.length && pedeRanking) {{
    return `📊 Perfil dos ${{AUTORES.length}} autores identificados nos B.O.s:\\n` +
      `• Naturalidade mais comum: ${{_fmtTop(_topEntradas(AUTORES,'naturalidade_cidade',3))}}\\n` +
      `• Profissão mais comum: ${{_fmtTop(_topEntradas(AUTORES,'profissao',3))}}\\n` +
      `• Estado civil mais comum: ${{_fmtTop(_topEntradas(AUTORES,'estado_civil',3))}}\\n` +
      `• Sexo: ${{_fmtTop(_topEntradas(AUTORES,'sexo',2))}}\\n` +
      `• Menores de idade: ${{AUTORES.filter(a=>a.menor).length}} de ${{AUTORES.length}}.`;
  }}

  if (!criteriosEncontrados.length) {{
    return `Não identifiquei um critério conhecido na pergunta (cidade, profissão, estado civil, sexo ou "menor de idade").\\n` +
      `Total de autores identificados: ${{AUTORES.length}}.\\n` +
      `Exemplos: "quantos autores são naturais de Curitiba?", "quantos são de Balneário Camboriú?", "qual a profissão mais comum entre os autores?", "quantos são menores de idade?".`;
  }}

  const criteriosTxt = [...new Set(criteriosEncontrados)].join(', ');
  if (!subset.length) {{
    return `Nenhum autor identificado corresponde a "${{criteriosTxt}}" (de um total de ${{AUTORES.length}} autores com perfil extraído).`;
  }}

  return `📊 ${{subset.length}} de ${{AUTORES.length}} autores identificados correspondem a "${{criteriosTxt}}".\\n` +
    `• Sexo: ${{_fmtTop(_topEntradas(subset,'sexo',2))}}\\n` +
    `• Profissão: ${{_fmtTop(_topEntradas(subset,'profissao',4))}}\\n` +
    `• Estado civil: ${{_fmtTop(_topEntradas(subset,'estado_civil',3))}}\\n` +
    `• Naturalidade: ${{_fmtTop(_topEntradas(subset,'naturalidade_cidade',4))}}\\n` +
    `• Menores de idade: ${{subset.filter(a=>a.menor).length}}.`;
}}

function toggleChatIA() {{
  document.getElementById('chat-ia-panel').classList.toggle('aberto');
}}

// ── Roteiro de atendimento da Ana (nome, saudação, LGPD, despedida) ───────────
// Tudo aqui roda localmente no navegador, sem chamar nenhuma IA externa —
// combina com o resto deste painel, que já é 100% offline/local.
let _nomeUsuario = null;

function _pareceDespedida(pNorm) {{
  return /\\bobrigad[oa]\\b|\\bvaleu\\b|\\btchau\\b|ate mais|ate logo|\\bso isso\\b|e so isso|foi tudo/.test(pNorm);
}}

function _pedeDadoPessoal(pNorm) {{
  return /\\bcpf\\b|\\brg\\b|\\btelefone\\b|nome (do|da|dos|das) autor|quem (e|é) o autor|identificar o autor|endereco (do|da) autor/.test(pNorm);
}}

function _pareceConsulta(pNorm) {{
  return /quant[oa]s?|\\bqual\\b|\\bquem\\b|profiss|natural|estado civil|\\bautor(es)?\\b|\\bcrime\\b|\\bmenor(es)?\\b|adolescente|\\bcidade\\b|\\bbairro\\b/.test(pNorm);
}}

function gerarRespostaAna(pergunta) {{
  const pNorm = _normalizarTexto(pergunta);

  if (_pareceDespedida(pNorm)) {{
    return 'Fico feliz em ter lhe ajudado! Se precisar de algo mais, é só chamar.';
  }}

  if (_pedeDadoPessoal(pNorm)) {{
    return 'Não posso compartilhar nome, CPF, RG ou qualquer outro dado que identifique uma pessoa — isso é protegido pela Lei Geral de Proteção de Dados (LGPD). Posso te passar estatísticas agregadas, como naturalidade, profissão ou estado civil mais comuns, se ajudar.';
  }}

  if (!_nomeUsuario) {{
    if (_pareceConsulta(pNorm)) {{
      const resposta = responderPerguntaAutores(pergunta);
      return `Ok! Deixa eu buscar isso no sistema...\\n\\n${{resposta}}\\n\\nA propósito, qual o seu nome?`;
    }}
    const primeiraPalavra = pergunta.trim().split(/\\s+/)[0].replace(/[^A-Za-zÀ-ÿ]/g, '');
    if (primeiraPalavra) {{
      _nomeUsuario = primeiraPalavra.charAt(0).toUpperCase() + primeiraPalavra.slice(1).toLowerCase();
      return `Oi, ${{_nomeUsuario}}! Em que posso lhe ajudar hoje?`;
    }}
  }}

  const resposta = responderPerguntaAutores(pergunta);
  return `Ok! Deixa eu buscar isso no sistema...\\n\\n${{resposta}}`;
}}

function enviarPerguntaChat() {{
  const input = document.getElementById('chat-ia-input');
  const pergunta = input.value.trim();
  if (!pergunta) return;
  input.value = '';
  const log = document.getElementById('chat-ia-log');
  log.insertAdjacentHTML('beforeend', `<div class="chat-msg chat-msg-user">${{pergunta.replace(/</g,'&lt;')}}</div>`);
  const resposta = gerarRespostaAna(pergunta);
  log.insertAdjacentHTML('beforeend', `<div class="chat-msg chat-msg-ia">${{resposta.replace(/</g,'&lt;')}}</div>`);
  log.scrollTop = log.scrollHeight;
}}
</script>
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
        if nome in cache and "autores" in cache[nome]:
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
    autores_todos = [a for bo in bos for a in bo.get("autores", [])]
    html = gerar_html(grupos, bos, ts, autores_todos)
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
        # Perfil agregado dos autores (sem nome, CPF, RG, telefone ou
        # endereço completo — ver extrair_autores()) para o chat "Pergunte
        # à IA" do dashboard público.
        "autores":          autores_todos,
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
