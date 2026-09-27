@echo off
title EduDocente-Studio — Generador de Evaluaciones con IA
chcp 65001 > nul
echo ====================================================================
echo           🎓 INICIANDO EDUDOCENTE-STUDIO (INTERFAZ WEB)
echo ====================================================================
echo.
echo Verificando instalacion de Python...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se detecto Python instalado en el sistema.
    echo Por favor, instala Python 3.10 o superior desde https://www.python.org/
    pause
    exit /b 1
)

echo Iniciando servidor web local y abriendo interfaz grafica...
python app.py
pause
