import os

import requests
from flask import Flask, Response, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY")
ELEVENLABS_VOICE_ID = os.environ.get("ELEVENLABS_VOICE_ID", "21m00Tcm4TlvDq8ikWAM")
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
        texto = _chamar_openai(CHAT_SYSTEM_PROMPT, user_content, max_tokens=1500)
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao consultar a IA: {exc}"}), 502

    return jsonify({"texto": texto})


INSIGHTS_SYSTEM_PROMPT = (
    "Você é um analista de segurança pública auxiliando a Guarda Municipal "
    "de Balneário Camboriú. Você recebe várias categorias numeradas, cada "
    "uma com dados estatísticos já processados sobre ocorrências "
    "criminais. Para CADA categoria, gere uma observação e uma sugestão "
    "prática, respondendo estritamente neste formato, sem nada antes ou "
    "depois:\n"
    "###N###\n"
    "Observação: <1 a 2 frases descrevendo o padrão nos dados>\n"
    "Sugestão: <1 frase objetiva com recomendação operacional prática>\n"
    "Repita esse bloco para cada categoria recebida, na mesma ordem e "
    "numeração. Nunca invente números que não estejam nos dados "
    "fornecidos. Seja específico, citando os valores concretos recebidos."
)


@app.route("/api/insights", methods=["POST"])
def insights():
    if not OPENAI_API_KEY:
        return jsonify({"erro": "OPENAI_API_KEY não configurada no servidor"}), 500

    dados = request.get_json(silent=True) or {}
    categorias = dados.get("categorias", "").strip()
    if not categorias:
        return jsonify({"erro": "Campo 'categorias' é obrigatório"}), 400

    try:
        texto = _chamar_openai(INSIGHTS_SYSTEM_PROMPT, categorias, max_tokens=1200)
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao consultar a IA: {exc}"}), 502

    return jsonify({"texto": texto})


@app.route("/api/transcrever", methods=["POST"])
def transcrever():
    if not OPENAI_API_KEY:
        return jsonify({"erro": "OPENAI_API_KEY não configurada no servidor"}), 500

    audio_file = request.files.get("audio")
    if not audio_file:
        return jsonify({"erro": "Arquivo de áudio é obrigatório"}), 400

    try:
        resp = requests.post(
            "https://api.openai.com/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
            files={
                "file": (
                    audio_file.filename or "audio.webm",
                    audio_file.stream,
                    audio_file.mimetype or "audio/webm",
                )
            },
            data={"model": "whisper-1", "language": "pt"},
            timeout=30,
        )
        resp.raise_for_status()
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao transcrever áudio: {exc}"}), 502

    texto = resp.json().get("text", "")
    return jsonify({"texto": texto})


@app.route("/api/falar", methods=["POST"])
def falar():
    if not ELEVENLABS_API_KEY:
        return jsonify({"erro": "ELEVENLABS_API_KEY não configurada no servidor"}), 500

    dados = request.get_json(silent=True) or {}
    texto = dados.get("texto", "").strip()
    if not texto:
        return jsonify({"erro": "Campo 'texto' é obrigatório"}), 400

    try:
        resp = requests.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}",
            headers={
                "xi-api-key": ELEVENLABS_API_KEY,
                "Content-Type": "application/json",
                "Accept": "audio/mpeg",
            },
            json={
                "text": texto,
                "model_id": "eleven_multilingual_v2",
                "voice_settings": {"stability": 0.5, "similarity_boost": 0.75},
            },
            timeout=30,
        )
        resp.raise_for_status()
    except requests.RequestException as exc:
        return jsonify({"erro": f"Falha ao gerar áudio: {exc}"}), 502

    return Response(resp.content, mimetype="audio/mpeg")


@app.route("/<path:filename>")
def static_files(filename):
    # Arquivos sensíveis (ex: inteligencia_criminal.html) ficam fora do
    # Git/imagem Docker e são montados via volume só na VPS.
    if os.path.isfile(os.path.join(DATA_DIR, filename)):
        return send_from_directory(DATA_DIR, filename)
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))
