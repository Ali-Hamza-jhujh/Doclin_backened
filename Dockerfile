FROM python:3.11-slim AS builder

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential libgl1 libglib2.0-0 libsm6 libxext6 \
    libxrender-dev poppler-utils tesseract-ocr \
    tesseract-ocr-ara tesseract-ocr-urd \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /install
COPY requirements.txt .
RUN pip install --prefix=/install/deps -r requirements.txt

# ── Runtime ───────────────────────────────────────────────────────────────────
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 libglib2.0-0 libsm6 libxext6 libxrender-dev \
    poppler-utils tesseract-ocr tesseract-ocr-ara tesseract-ocr-urd \
    && rm -rf /var/lib/apt/lists/*

COPY --from=builder /install/deps /usr/local
WORKDIR /app
COPY . .
RUN mkdir -p /app/outputs /app/uploads

# Pre-warm Docling models so first request is instant
RUN python -c "from docling.document_converter import DocumentConverter; DocumentConverter()" \
    || echo "Pre-warm skipped"

EXPOSE 8000

# 4 parallel workers — tune --workers to match your CPU cores
CMD ["gunicorn", "main:app", \
     "--worker-class", "uvicorn.workers.UvicornWorker", \
     "--workers", "4", \
     "--bind", "0.0.0.0:8000", \
     "--timeout", "300", \
     "--log-level", "info", \
     "--access-logfile", "-"]