FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

# LibreOffice must exist in the final runtime image, not only in a temporary build environment.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-liberation \
    fonts-dejavu-core \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && python -m playwright install --with-deps chromium

COPY . .

ENV PATH="/usr/lib/libreoffice/program:/usr/bin:${PATH}"

CMD ["sh", "-c", "command -v libreoffice && libreoffice --version && uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}"]
