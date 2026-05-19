@echo off
title Extracto de Pagos de Impuestos SAT
echo.
echo  ============================================
echo   Extracto de Pagos de Impuestos SAT
echo   http://127.0.0.1:5050
echo  ============================================
echo.
echo  Iniciando servidor...
echo  Abre tu navegador en: http://127.0.0.1:5050
echo.
start "" "http://127.0.0.1:5050"
python "%~dp0app.py"
pause
