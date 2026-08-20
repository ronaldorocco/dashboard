import os

import requests
from flask import Flask, Response, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY")
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")
ELEVENLABS_VOICE_ID = os.environ.get("ELEVENLABS_VOICE_ID", "21m00Tcm4TlvDq8ikWAM")
# Turbo é o modelo de menor latência da ElevenLabs (troca um pouco de
# qualidade por velocidade) — bom pra conversa em tempo real como a da Ana.
ELEVENLABS_MODEL_ID = os.environ.get("ELEVENLABS_MODEL_ID", "eleven_turbo_v2_5")
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

PROMPT_ANA_PATH = os.path.join(BASE_DIR, "prompt_ana.txt")
_CHAT_SYSTEM_PROMPT_PADRAO = (
    "Você é a Ana, atendente virtual da Secretaria de Segurança e Ordem "
    "Pública da Guarda Municipal de Balneário Camboriú, conversando pelo "
    "chat do dashboard. Seja sempre empática, educada, prestativa e "
    "objetiva. Responda com base apenas nos dados fornecidos, nunca "
    "invente números."
)


def carregar_prompt_ana():
    # Lido do arquivo prompt_ana.txt (editável direto, sem precisar mexer
    # neste .py) — se o arquivo não existir por algum motivo, cai num
    # prompt mínimo padrão em vez de quebrar o chat.
    try:
        with open(PROMPT_ANA_PATH, encoding="utf-8") as f:
            texto = f.read().strip()
        return texto or _CHAT_SYSTEM_PROMPT_PADRAO
    except OSError:
        return _CHAT_SYSTEM_PROMPT_PADRAO


def _chamar_openai(system_prompt, user_content, max_tokens=500, historico=None):
    mensagens = [{"role": "system", "content": system_prompt}]
    if historico:
        mensagens.extend(historico)
    mensagens.append({"role": "user", "content": user_content})

    resp = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": OPENAI_MODEL,
            "messages": mensagens,
            "temperature": 0.3,
            "max_tokens": max_tokens,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def _limpar_historico(bruto):
    """Mantém só as últimas trocas, validando formato — evita abuso do
    endpoint público com payloads gigantes ou papéis inválidos."""
    if not isinstance(bruto, list):
        return []
    historico = []
    for msg in bruto[-20:]:
        if (
            isinstance(msg, dict)
            and msg.get("role") in ("user", "assistant")
            and isinstance(msg.get("content"), str)
        ):
            historico.append({"role": msg["role"], "content": msg["content"][:4000]})
    return historico


def _servir_dashboard(caminho):
    # A chave do Google Maps nunca fica no HTML commitado (repo é público) —
    # o build gera um placeholder e a gente troca aqui, na hora de servir,
    # pela variável de ambiente do servidor.
    with open(caminho, encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__GOOGLE_MAPS_API_KEY__", GOOGLE_MAPS_API_KEY)
    return Response(html, mimetype="text/html")


@app.route("/")
def index():
    return _servir_dashboard(os.path.join(BASE_DIR, "dashboard_interativo.html"))


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
    historico = _limpar_historico(dados.get("historico"))
    if not pergunta:
        return jsonify({"erro": "Campo 'pergunta' é obrigatório"}), 400

    user_content = (
        f"Pergunta do usuário: {pergunta}\n\n"
        "Registros/estatísticas encontrados no banco de dados que "
        "correspondem a essa pergunta:\n"
        f"{contexto or '(nenhum registro encontrado para os critérios identificados na pergunta)'}"
    )

    try:
        texto = _chamar_openai(carregar_prompt_ana(), user_content, max_tokens=600, historico=historico)
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
                "model_id": ELEVENLABS_MODEL_ID,
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
    if filename in ("dashboard_interativo.html", "index.html"):
        return _servir_dashboard(os.path.join(BASE_DIR, filename))
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))
