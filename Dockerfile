FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu-core \
        fonts-noto-core \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs assets

CMD gunicorn -w 1 -b 0.0.0.0:${PORT:-5000} --timeout 180 app:app
