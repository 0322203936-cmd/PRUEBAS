# Usar imagen ligera de Python 3.10
FROM python:3.10-slim

# Evitar escritura de bytecodes y buffer de stdout
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
# Timeout extendido para dar tiempo a Tesseract en CPU lenta
ENV GUNICORN_CMD_ARGS="--timeout 120 --workers 1"

# Configurar directorio de trabajo
WORKDIR /app

# Instalar dependencias del sistema operativo (Tesseract)
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-spa \
    && rm -rf /var/lib/apt/lists/*

# Copiar el archivo de requerimientos
COPY requirements.txt .

# Instalar requerimientos de Python
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copiar el código fuente
COPY . .

# Comando para iniciar la aplicación (Render asigna el puerto con $PORT)
CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-10000} --timeout 120 --workers 1"]
