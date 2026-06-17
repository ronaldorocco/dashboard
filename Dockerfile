FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY bot_consulta.py .
COPY secretario.xlsx .
CMD ["python", "bot_consulta.py"]
