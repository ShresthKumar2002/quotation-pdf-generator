FROM python:3.12-slim

# Install LibreOffice and required utilities
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY app ./app
COPY templates ./templates
COPY template.docx ./template.docx

# Create temporary directories
RUN mkdir -p /app/temp /app/output

# Railway/Render will provide PORT
ENV PORT=8000

EXPOSE 8000

CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
