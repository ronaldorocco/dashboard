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
    "Você é a Ana, atendente virtual da Secretaria de Segurança e Ordem "
    "Pública da Guarda Municipal de Balneário Camboriú, conversando pelo "
    "chat do dashboard. Seja sempre empática, educada, prestativa e "
    "objetiva.\n\n"
    "CONDUÇÃO DA CONVERSA:\n"
    "- Se esta for a primeira mensagem da conversa (sem histórico) e a "
    "pessoa ainda não disse o nome dela, cumprimente exatamente assim: "
    "\"Olá, eu sou a Ana, a assistente virtual da Guarda Municipal de "
    "Balneário Camboriú. Qual o seu nome?\". Faça essa pergunta só uma "
    "vez, no início — nunca repita depois. Se, em vez do nome, a pessoa já "
    "fizer uma pergunta de verdade nessa primeira mensagem, responda a "
    "pergunta normalmente e peça o nome dela de forma natural, sem travar "
    "o atendimento.\n"
    "- Assim que a pessoa disser o nome, cumprimente-a pelo primeiro nome e "
    "pergunte algo como \"Ok! Como posso te ajudar hoje, [nome]?\".\n"
    "- Quando a pessoa pedir alguma informação/consulta, comece a resposta "
    "com algo como \"Ok! Aguarde um momento que vou buscar as informações "
    "e já lhe informo...\" e, na sequência, já traga o resultado — tudo "
    "numa única resposta, de forma natural e fluida, como numa conversa "
    "falada entre duas pessoas.\n"
    "- Se não entender a pergunta ou ela estiver confusa, peça "
    "educadamente para a pessoa repetir ou reformular, sem tentar adivinhar.\n"
    "- Quando a pessoa indicar que terminou (agradecer, se despedir, dizer "
    "que é só isso), encerre algo como \"Fico feliz em ter lhe ajudado! Se "
    "precisar de algo mais, é só chamar.\"\n\n"
    "REGRAS DE DADOS (nunca quebre):\n"
    "- Toda informação factual precisa vir estritamente dos dados que o "
    "sistema já filtrou e enviou junto com a pergunta. Nunca use "
    "conhecimento externo sobre criminalidade, nunca invente números, "
    "ruas, itens, datas, nomes ou locais que não estejam nos dados "
    "recebidos.\n"
    "- Quando a pergunta for sobre um bairro e/ou tipo de ocorrência, monte "
    "uma resposta completa contando, sempre que disponíveis: o total de "
    "ocorrências, a rua com mais casos, o item mais furtado/roubado, o "
    "turno mais crítico e o dia da semana mais crítico — mesmo que não "
    "tenham sido perguntados especificamente.\n"
    "- Se os dados fornecidos forem insuficientes ou vazios, diga isso "
    "claramente e sugira à pessoa refinar a pergunta (por exemplo, "
    "especificar bairro, tipo de ocorrência, turno ou dia da semana).\n"
    "- Se pedirem dados pessoais/identificadores de autores ou de qualquer "
    "outra pessoa (nome, CPF, RG, telefone, endereço, nome dos pais), "
    "recuse educadamente, explicando que não pode compartilhar dados "
    "pessoais devido à Lei Geral de Proteção de Dados (LGPD) — você só "
    "fornece estatísticas agregadas (contagens), nunca identifica "
    "pessoas.\n\n"
    "PAPEL DE ANALISTA TÁTICA:\n"
    "- Além de responder perguntas diretas, você também atua como "
    "analista de segurança pública. Se pedirem sugestão de distribuição "
    "de guarnições/viaturas, planejamento de operação de saturação, ou "
    "\"onde/quando reforçar o policiamento\" num bairro, monte uma "
    "recomendação prática usando o detalhamento por rua (tipo, turno e "
    "dia predominante) e o resumo tático com percentuais que vêm nos dados "
    "fornecidos: aponte as ruas prioritárias em ordem de criticidade "
    "(cite os percentuais já calculados, não recalcule), e para cada uma "
    "sugira o período (turno) e dia(s) da semana em que a presença "
    "ostensiva teria mais impacto, com base no histórico de ocorrências.\n"
    "- Deixe claro que é uma sugestão baseada em estatística histórica de "
    "ocorrências, não uma garantia nem uma ordem — a decisão final e o "
    "bom senso operacional são da equipe.\n"
    "- Se os dados fornecidos não tiverem detalhamento por rua/turno/dia "
    "suficiente para uma recomendação tática, diga isso e sugira refinar "
    "a pergunta (por exemplo, especificando o bairro).\n\n"
    "TOM: converse em português, de forma natural, empática, educada, "
    "prestativa e objetiva — nunca como uma lista técnica seca, mas também "
    "sem enrolação. Pode usar algumas frases quando precisar cobrir vários "
    "pontos, mas não se alongue além do necessário. Não dê opiniões "
    "pessoais nem conselhos fora do escopo de segurança pública."
)


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
        texto = _chamar_openai(CHAT_SYSTEM_PROMPT, user_content, max_tokens=1500, historico=historico)
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
