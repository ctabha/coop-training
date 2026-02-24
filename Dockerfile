RUN apt-get update \
 && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-core \
    libreoffice-common \
    fonts-dejavu \
    fonts-noto-core \
    fonts-noto-extra \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    ca-certificates \
 && rm -rf /var/lib/apt/lists/*
