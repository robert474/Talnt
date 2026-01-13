# Talnt Document Generator
# Python 3.11 + Node.js 20 for resume formatting and DOCX generation

FROM python:3.11-slim

# Install Node.js 20
RUN apt-get update && apt-get install -y \
    curl \
    gnupg \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files first for better caching
COPY package*.json ./
RUN npm install --production

# Copy Python requirements and install
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create required directories
RUN mkdir -p input output rfq/uploads rfq/output

# Expose port
EXPOSE 5050

# Set environment variables
ENV PORT=5050
ENV FLASK_DEBUG=false
ENV RAILWAY_ENVIRONMENT=true

# Run the Flask app with gunicorn for production
CMD ["gunicorn", "--bind", "0.0.0.0:5050", "--workers", "2", "--timeout", "120", "rfq.app:app"]
