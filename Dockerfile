FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    AUDIVERIS_BIN=/usr/bin/audiveris

WORKDIR /app

# System dependencies for OMR + runtime.
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
      ca-certificates \
      curl \
      jq \
      ghostscript \
      tesseract-ocr \
      libasound2 \
      libfreetype6 \
      libxi6 \
      libxrender1 \
      libxtst6; \
    rm -rf /var/lib/apt/lists/*

# Install latest Audiveris Linux .deb from GitHub releases.
RUN set -eux; \
    AUDIVERIS_DEB_URL="$(curl -fsSL https://api.github.com/repos/Audiveris/audiveris/releases/latest | jq -r '.assets[] | select(.name | test("(?i)linux.*\\.deb$")) | .browser_download_url' | head -n 1)"; \
    test -n "${AUDIVERIS_DEB_URL}"; \
    curl -fsSL "${AUDIVERIS_DEB_URL}" -o /tmp/audiveris.deb; \
    apt-get update; \
    apt-get install -y --no-install-recommends /tmp/audiveris.deb; \
    rm -f /tmp/audiveris.deb; \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--workers", "2", "--threads", "4", "--timeout", "120"]
