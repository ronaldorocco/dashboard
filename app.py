import os

import requests
from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
app = Flask(__name__)

SYSTEM_PROMPT = (
    "Você é um analista de segurança pública auxiliando a Guarda Municipal "
    "de Balneário Camboriú. Você recebe dados estatísticos já processados "
    "sobre ocorrências criminais. Escreva um resumo executivo em português, "
    "tom formal e objetivo, de no máximo 200 palavras, destacando: área e "
    "horário de maior risco, tendência (crescente/decrescente) e uma "
    "recomendação operacional prioritária. Nunca invente números que não "
    "estejam nos dados fornecidos; se um dado não for informado, não o "
    "mencione."
)

CHAT_SYSTEM_PROMPT = (
    "Você é um assistente analítico para a equipe da Guarda Municipal de "
    "Balneário Camboriú, respondendo perguntas sobre ocorrências criminais "
    "registradas no dashboard. Você receberá, junto com a pergunta do "
    "usuário, um conjunto de registros reais já filtrados pelo sistema com "
    "base na pergunta. Responda somente com base nesses registros "
    "fornecidos — nunca use conhecimento externo sobre criminalidade, nunca "
    "invente números, datas ou locais que não estejam nos dados recebidos. "
    "Se os dados fornecidos forem insuficientes ou vazios, diga isso "
    "claramente e sugira ao usuário refinar a pergunta (por exemplo, "
    "especificar bairro, tipo de ocorrência, turno ou dia da semana). "
    "Responda de forma direta, técnica e em português, citando números e "
    "percentuais concretos quando disponíveis. Não dê opiniões pessoais "
    "nem conselhos fora do escopo de segurança pública."
)


def _chamar_openai(system_prompt, user_content, max_tokens=500):
    resp = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": OPENAI_MODEL,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0.3,
            "max_tokens": max_tokens,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "dashboard_interativo.html")


@app.route("/api/explicar", methods=["POST"])
def explicar():
    if not OPENAI_API_KEY:
        return jsonify({"erro": "OPENAI_API_KEY não configurada no servidor"}), 500

    dados = request.get_json(silent=True) or {}
    resumo_dados = dados.get("resumo", "").strip()
    if not resumo_dados:
        return jsonify({"erro": "Campo 'resumo' é obrigatório"}), 400

    try:
        texto = _chamar_openai(SYSTEM_PROMPT, resumo_dados, max_tokens=500)
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao consultar a IA: {exc}"}), 502

    return jsonify({"texto": texto})


@app.route("/api/chat", methods=["POST"])
def chat():
    if not OPENAI_API_KEY:
        return jsonify({"erro": "OPENAI_API_KEY não configurada no servidor"}), 500

    dados = request.get_json(silent=True) or {}
    pergunta = dados.get("pergunta", "").strip()
    contexto = dados.get("contexto", "").strip()
    if not pergunta:
        return jsonify({"erro": "Campo 'pergunta' é obrigatório"}), 400

    user_content = (
        f"Pergunta do usuário: {pergunta}\n\n"
        "Registros/estatísticas encontrados no banco de dados que "
        "correspondem a essa pergunta:\n"
        f"{contexto or '(nenhum registro encontrado para os critérios identificados na pergunta)'}"
    )

    try:
        texto = _chamar_openai(CHAT_SYSTEM_PROMPT, user_content, max_tokens=600)
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao consultar a IA: {exc}"}), 502

    return jsonify({"texto": texto})


@app.route("/<path:filename>")
def static_files(filename):
    # Arquivos sensíveis (ex: inteligencia_criminal.html) ficam fora do
    # Git/imagem Docker e são montados via volume só na VPS.
    if os.path.isfile(os.path.join(DATA_DIR, filename)):
        return send_from_directory(DATA_DIR, filename)
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))
