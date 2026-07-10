@echo off
REM ============================================================
REM  Robo de reservas Risotolandia
REM  Executa o script PowerShell que le a planilha e envia as
REM  reservas no site da Solvis. Usa apenas o PowerShell nativo
REM  do Windows (nada e baixado ou instalado).
REM ============================================================
setlocal
cd /d "%~dp0"
REM Remove a marca "arquivo baixado da internet" do script, se existir.
REM Sem isso o Windows recusa o script com o erro "nao esta assinado
REM digitalmente". Desbloquear um arquivo proprio e uma acao legitima.
powershell.exe -NoProfile -Command "Unblock-File -LiteralPath '%~dp0robo_risotolandia.ps1'" >nul 2>&1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0robo_risotolandia.ps1" %*
echo.
pause
