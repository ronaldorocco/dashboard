@echo off
chcp 65001 >nul
echo.
echo ================================================
echo   ATUALIZAR PROMPT DA ANA - GUARDA MUNICIPAL BC
echo ================================================
echo.
echo  Publicando o conteudo de prompt_ana.txt...
echo.

cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1" "Atualizar prompt da Ana"

echo.
echo ================================================
echo   Concluido! Em 1-2 minutos o novo prompt entra
echo   em vigor em https://dashboardgmbc.com.br
echo ================================================
echo.
pause
