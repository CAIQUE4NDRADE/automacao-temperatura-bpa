@echo off
echo.
echo  ============================================
echo   INSTALADOR - AUTOMACAO TEMPERATURA BPA
echo  ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale Python 3.11+ em https://python.org
    pause
    exit /b 1
)
echo [OK] Python encontrado.

echo.
echo Instalando dependencias Python...
python -m pip install pdfplumber selenium schedule pywin32 openpyxl webdriver-manager
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.

mkdir "C:\Users\cs338104\Downloads\Atualização de Temperatura\PDFs" 2>nul
mkdir "C:\Users\cs338104\Downloads\Atualização de Temperatura\NewSAT" 2>nul
mkdir "C:\Users\cs338104\Downloads\Atualização de Temperatura\Logs" 2>nul
echo [OK] Pastas criadas.

echo.
echo Criando tarefa agendada no Windows...
schtasks /delete /tn "AutomacaoTemperaturaBPA" /f >nul 2>&1
schtasks /create /tn "AutomacaoTemperaturaBPA" /tr "python \"%~dp0monitor_temperatura.py\"" /sc ONSTART /ru "%USERNAME%" /f
if errorlevel 1 (
    echo [AVISO] Tarefa agendada nao criada. Use iniciar.bat manualmente.
) else (
    echo [OK] Tarefa agendada criada. Inicia automaticamente com o Windows.
)

echo.
echo ============================================
echo  INSTALACAO CONCLUIDA!
echo.
echo  Execute iniciar.bat para rodar agora.
echo  Logs em: Atualização de Temperatura\Logs\
echo ============================================
echo.
pause
