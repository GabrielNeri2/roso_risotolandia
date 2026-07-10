@echo off
REM ============================================================
REM  MODO TESTE - percorre o formulario inteiro mas NAO clica
REM  em "Finalizar". Nenhuma reserva e criada. Use para conferir
REM  se a planilha e o site estao combinando antes de rodar de
REM  verdade.
REM ============================================================
setlocal
cd /d "%~dp0"
REM Remove a marca "arquivo baixado da internet" do script, se existir.
powershell.exe -NoProfile -Command "Unblock-File -LiteralPath '%~dp0robo_risotolandia.ps1'" >nul 2>&1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0robo_risotolandia.ps1" -SomenteTeste
echo.
pause
