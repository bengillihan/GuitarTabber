FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    AUDIVERIS_BIN=/usr/local/bin/audiveris

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
    # Extract package payload directly to avoid post-install desktop scripts in headless builds.
    dpkg-deb -x /tmp/audiveris.deb /tmp/audiveris-root; \
    cp -a /tmp/audiveris-root/. /; \
    rm -rf /tmp/audiveris-root; \
    if [ -x /opt/audiveris/bin/Audiveris ]; then \
      ln -sf /opt/audiveris/bin/Audiveris /usr/local/bin/audiveris; \
    elif [ -x /opt/Audiveris/bin/Audiveris ]; then \
      ln -sf /opt/Audiveris/bin/Audiveris /usr/local/bin/audiveris; \
    fi; \
    if [ ! -x /usr/local/bin/audiveris ]; then \
      AUDIVERIS_CANDIDATE="$(find / -type f \( -name audiveris -o -name Audiveris \) 2>/dev/null | head -n 1)"; \
      test -n "${AUDIVERIS_CANDIDATE}"; \
      ln -sf "${AUDIVERIS_CANDIDATE}" /usr/local/bin/audiveris; \
    fi; \
    ln -sf /usr/local/bin/audiveris /usr/bin/audiveris; \
    test -x /usr/bin/audiveris; \
    rm -f /tmp/audiveris.deb; \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--workers", "2", "--threads", "4", "--timeout", "120"]
