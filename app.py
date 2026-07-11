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
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": OPENAI_MODEL,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": resumo_dados},
                ],
                "temperature": 0.3,
                "max_tokens": 500,
            },
            timeout=30,
        )
        resp.raise_for_status()
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao consultar a IA: {exc}"}), 502

    texto = resp.json()["choices"][0]["message"]["content"]
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
