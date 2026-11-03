# Use Python 3.11 slim as base
FROM python:3.11-slim

# Install all Chromium dependencies manually
# (avoids playwright install-deps which fails on newer Debian due to
# missing font packages: ttf-unifont and ttf-ubuntu-font-family)
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    ca-certificates \
    fonts-liberation \
    fonts-unifont \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libc6 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libexpat1 \
    libfontconfig1 \
    libgbm1 \
    libglib2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libstdc++6 \
    libx11-6 \
    libx11-xcb1 \
    libxcb1 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxi6 \
    libxrandr2 \
    libxrender1 \
    libxss1 \
    libxtst6 \
    lsb-release \
    xdg-utils \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium browser only
# (NOT using playwright install-deps — we installed deps manually above)
RUN playwright install chromium

# Copy all backend files
COPY . .

# Create necessary directories
RUN mkdir -p uploads downloads output

# Expose port
EXPOSE 5001

# Start the Flask server
CMD ["python", "server.py"]
