# EduDocente-Studio: Entorno Contenerizado para Producción Escolar
FROM python:3.11-slim

WORKDIR /app

# Instalar dependencias del sistema mínimas
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copiar requerimientos e instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copiar el código fuente de la aplicación
COPY . .

# Exponer el puerto del servidor web
EXPOSE 8080

# Variables de entorno predeterminadas
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

# Iniciar servidor web institucional sin abrir navegador
CMD ["python", "app.py", "--no-browser", "--port", "8080"]
