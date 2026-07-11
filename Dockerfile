FROM python:3.11-slim

WORKDIR /app

COPY requirements-web.txt .
RUN pip install --no-cache-dir -r requirements-web.txt

COPY app.py .
COPY dashboard_interativo.html index.html ./

VOLUME ["/app/data"]
EXPOSE 3000
CMD ["gunicorn", "-b", "0.0.0.0:3000", "app:app"]
