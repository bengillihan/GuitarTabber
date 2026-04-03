FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
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
    AUDIVERIS_JSON="$(curl -fsSL https://api.github.com/repos/Audiveris/audiveris/releases/latest)"; \
    AUDIVERIS_DEB_URL="$(printf '%s' "${AUDIVERIS_JSON}" | jq -r '[.assets[]? | select(.name | test("ubuntu24\\.04.*x86_64\\.deb$")) | .browser_download_url][0] // empty')"; \
    if [ -z "${AUDIVERIS_DEB_URL}" ]; then \
      AUDIVERIS_DEB_URL="$(printf '%s' "${AUDIVERIS_JSON}" | jq -r '[.assets[]? | select(.name | test("ubuntu22\\.04.*x86_64\\.deb$")) | .browser_download_url][0] // empty')"; \
    fi; \
    if [ -z "${AUDIVERIS_DEB_URL}" ]; then \
      AUDIVERIS_DEB_URL="$(printf '%s' "${AUDIVERIS_JSON}" | jq -r '[.assets[]? | select(.name | test("\\.deb$")) | .browser_download_url][0] // empty')"; \
    fi; \
    test -n "${AUDIVERIS_DEB_URL}"; \
    curl -fsSL "${AUDIVERIS_DEB_URL}" -o /tmp/audiveris.deb; \
    # Headless containers can fail Audiveris post-install desktop integration.
    # Temporarily no-op xdg-desktop-menu so package configuration succeeds.
    printf '#!/bin/sh\nexit 0\n' > /usr/local/bin/xdg-desktop-menu; \
    chmod +x /usr/local/bin/xdg-desktop-menu; \
    apt-get update; \
    apt-get install -y --no-install-recommends /tmp/audiveris.deb; \
    rm -f /usr/local/bin/xdg-desktop-menu; \
    test -x /usr/bin/audiveris; \
    rm -f /tmp/audiveris.deb; \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--workers", "2", "--threads", "4", "--timeout", "120"]
