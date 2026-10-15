@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Atualizacao da planilha, JSON e GitHub Pages
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ATUALIZAR_MATERIA_PRIMA.ps1"
echo.
if errorlevel 1 (echo FALHA: confira as mensagens acima.) else (echo Processo finalizado.)
pause
