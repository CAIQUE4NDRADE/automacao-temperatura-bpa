@echo off
cd /d "%~dp0"
title Automacao de Temperatura
set PY=py -3.13
%PY% --version >nul 2>&1
if errorlevel 1 set PY=python
del /q "%~dp0monitor.lock" 2>nul
%PY% "%~dp0monitor_temperatura.py"
echo.
echo ===== O programa fechou. Se apareceu erro acima, tire um print. =====
pause
