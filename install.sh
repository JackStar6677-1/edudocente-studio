#!/usr/bin/env bash
# EduDocente-Studio: Script de instalación rápida para Linux y macOS

set -e

echo "===================================================================="
echo "       🎓 INSTALADOR AUTOMÁTICO DE EDUDOCENTE-STUDIO (UNIX)         "
echo "===================================================================="

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 no está instalado en este sistema."
    echo "Por favor instálalo utilizando tu gestor de paquetes (apt, brew, pacman, etc.)."
    exit 1
fi

echo "[1/3] Versión detectada: $(python3 --version)"

echo "[2/3] Instalando dependencias de Python..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

echo "[3/3] Configurando variables de entorno..."
if [ ! -f .env ] && [ -f .env.example ]; then
    cp .env.example .env
    echo " - Archivo .env creado a partir de .env.example"
fi

echo "===================================================================="
echo "      ¡INSTALACIÓN COMPLETADA!                                      "
echo "===================================================================="
echo "Para iniciar la aplicación web ejecuta:"
echo "  python3 app.py"
echo "===================================================================="
