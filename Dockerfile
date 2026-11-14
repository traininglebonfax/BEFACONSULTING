FROM python:3.12-slim

# LibreOffice headless : Calc pour le recalcul du classeur (engine/recalc.py),
# Writer pour la conversion Word -> PDF du Business Plan BEFA (befa/pdf.py).
# Liberation Sans remplace Arial a l'identique (memes metriques).
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        libreoffice-writer \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["uvicorn", "backend.app:app", "--host", "0.0.0.0", "--port", "8000"]
