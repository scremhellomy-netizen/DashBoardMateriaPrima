@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo Atualizacao da planilha, JSON e GitHub Pages
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0ATUALIZAR_MATERIA_PRIMA.ps1"
set "RESULTADO=%ERRORLEVEL%"
echo.
if not "%RESULTADO%"=="0" (echo FALHA: consulte erro_atualizacao.log nesta pasta.) else (echo Processo finalizado.)
pause
exit /b %RESULTADO%
