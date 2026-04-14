@echo off
:: Inicia a automação em background (sem janela preta permanente)
echo Iniciando Automação Temperatura BPA...
start "AutomacaoTemperatura" /min python "%~dp0monitor_temperatura.py"
echo.
echo [OK] Automacao iniciada em background.
echo      Verifique os logs em: Logs\temperatura_AAAAMMDD.log
echo.
timeout /t 3 >nul
