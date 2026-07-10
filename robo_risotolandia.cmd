@echo off
REM ============================================================
REM  Robo de reservas Risotolandia
REM  Executa o script PowerShell que le a planilha e envia as
REM  reservas no site da Solvis. Usa apenas o PowerShell nativo
REM  do Windows (nada e baixado ou instalado).
REM ============================================================
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0robo_risotolandia.ps1" %*
echo.
pause
