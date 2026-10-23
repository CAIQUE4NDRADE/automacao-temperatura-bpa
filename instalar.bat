@echo off
set PY=py -3.13
%PY% --version >nul 2>&1
if errorlevel 1 set PY=python
%PY% -m pip install -r "%~dp0requirements.txt"
if not exist "%~dp0config.json" copy "%~dp0config.example.json" "%~dp0config.json"
echo.
echo [OK] Dependencias instaladas. Edite o config.json e rode iniciar.bat
pause
