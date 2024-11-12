# Utilizar la versión oficial de Python como imagen base
FROM python:3.10-slim

# Instalar dependencias del sistema y Chromium
RUN apt-get update && apt-get install -y \
    wget \
    curl \
    unzip \
    chromium \
    && rm -rf /var/lib/apt/lists/*

# Establecer el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiar el archivo de requerimientos y el script al contenedor
COPY requirements.txt requirements.txt
COPY validar_series.py validar_series.py

# Instalar las dependencias de Python, incluyendo chromedriver-autoinstaller
RUN pip install --no-cache-dir -r requirements.txt

# Establecer las variables de entorno necesarias para Selenium y Chromium
ENV CHROME_BIN=/usr/bin/chromium
ENV CHROME_DRIVER=/usr/local/bin/chromedriver

# Ejecutar el script cuando se inicie el contenedor
ENTRYPOINT ["python", "validar_series.py"]