@echo off
chcp 65001 >nul
title Demo - Generador de Convenios EAFIT

echo.
echo  =====================================================
echo   GENERADOR DE CONVENIOS - TALENTO EAFIT
echo   Reto 2 - Motor de Procesos - Beca IA EAFIT 2025
echo  =====================================================
echo.

:: Verificar que Python este instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado.
    echo Descargarlo en: https://www.python.org/downloads/
    echo Marcar la opcion "Add Python to PATH" al instalar.
    pause
    exit /b 1
)

:: Ir a la carpeta del demo
cd /d "%~dp0"

:: Instalar dependencias si hace falta
echo Verificando dependencias...
pip install -r requirements.txt -q
echo.

:: Abrir el navegador y arrancar la app
echo Iniciando la aplicacion...
echo El navegador se abrira automaticamente.
echo Para cerrar: presionar Ctrl+C en esta ventana.
echo.

python -m streamlit run app.py --server.headless false
pause
