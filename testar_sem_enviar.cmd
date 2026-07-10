@echo off
REM ============================================================
REM  MODO TESTE - percorre o formulario inteiro mas NAO clica
REM  em "Finalizar". Nenhuma reserva e criada. Use para conferir
REM  se a planilha e o site estao combinando antes de rodar de
REM  verdade.
REM ============================================================
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0robo_risotolandia.ps1" -SomenteTeste
echo.
pause
