import os

from flask import Flask, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
app = Flask(__name__)


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "dashboard_interativo.html")


@app.route("/<path:filename>")
def static_files(filename):
    # Arquivos sensíveis (ex: inteligencia_criminal.html) ficam fora do
    # Git/imagem Docker e são montados via volume só na VPS.
    if os.path.isfile(os.path.join(DATA_DIR, filename)):
        return send_from_directory(DATA_DIR, filename)
    return send_from_directory(BASE_DIR, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 3000)))
