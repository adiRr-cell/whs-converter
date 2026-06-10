FROM python:3.11-slim

# Instala LibreOffice e poppler
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .

ENV PORT=5000
EXPOSE 5000
CMD gunicorn app:app --timeout 180 --bind 0.0.0.0:$PORT
