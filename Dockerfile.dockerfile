# ---- Base Python ligera
FROM python:3.10-slim

# ---- Paquetes nativos que necesitan WeasyPrint/CairoSVG
RUN apt-get update && apt-get install -y --no-install-recommends \
    libcairo2 libcairo2-dev \
    libpango-1.0-0 libpangocairo-1.0-0 libpangoft2-1.0-0 \
    libgdk-pixbuf2.0-0 \
    libffi-dev shared-mime-info \
    fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# ---- App
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Evita el buffering de logs
ENV PYTHONUNBUFFERED=1

# Render te pasa $PORT; usa 8000 localmente si no existe
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
