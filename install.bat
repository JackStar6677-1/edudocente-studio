@echo off
title EduDocente-Studio — Instalador Automatizado
chcp 65001 > nul
echo ====================================================================
echo        🎓 INSTALADOR AUTOMÁTICO DE EDUDOCENTE-STUDIO
echo ====================================================================
echo.
echo [1/3] Verificando version de Python...
python --version
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta disponible en el PATH del sistema.
    pause
    exit /b 1
)

echo.
echo [2/3] Instalando dependencias de librerias oficiales (python-docx, openpyxl, matplotlib)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Ocurrio un inconveniente instalando las dependencias.
    pause
    exit /b 1
)

echo.
echo [3/3] Comprobando entorno de ejecucion...
if not exist .env (
    if exist .env.example (
        copy .env.example .env > nul
        echo  - Archivo .env inicializado desde .env.example
    )
)

echo.
echo ====================================================================
echo       ¡INSTALACIÓN COMPLETADA CON ÉXITO!
echo ====================================================================
echo Para abrir la interfaz grafica, haz doble clic en 'start.bat'
echo o ejecuta: python app.py
echo.
pause
