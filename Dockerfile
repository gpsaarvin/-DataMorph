# ── DataMorph — PDF to Excel Converter ──
# Dockerfile for Render / container deployment
# Uses Node 20 with system libs for canvas + Tesseract OCR

FROM node:20-slim

# Install system dependencies for 'canvas' (Cairo, Pango, libjpeg, libgif, librsvg)
# and Tesseract OCR language data
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    pkg-config \
    python3 \
    libcairo2-dev \
    libjpeg-dev \
    libpango1.0-dev \
    libgif-dev \
    librsvg2-dev \
    libpixman-1-dev \
    tesseract-ocr \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files first for better Docker layer caching
COPY package.json package-lock.json ./

# Install Node dependencies
RUN npm ci --omit=dev

# Copy application code
COPY server.js ./
COPY public/ ./public/
COPY eng.traineddata ./

# Create required directories
RUN mkdir -p uploads outputs debug

# Expose port
EXPOSE 8080

# Set environment
ENV PORT=8080
ENV NODE_ENV=production

# Start the application
CMD ["node", "server.js"]
