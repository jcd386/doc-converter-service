FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice libreoffice-writer libreoffice-calc libreoffice-impress \
    default-jre fonts-liberation fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/doc_converter_app.py .

RUN useradd -r -s /bin/false appuser && \
    mkdir -p /tmp/libreoffice && chown appuser:appuser /tmp/libreoffice
USER appuser

EXPOSE 8000
CMD ["uvicorn", "doc_converter_app:app", "--host", "0.0.0.0", "--port", "8000"]
