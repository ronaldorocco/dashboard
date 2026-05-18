@echo off
cd /d "%~dp0"
echo Atualizando dashboard...
"C:\Program Files\Python39\python.exe" build_dashboard.py
if %errorlevel% == 0 (
    echo Dashboard atualizado com sucesso!
    start dashboard_interativo.html
) else (
    echo ERRO ao executar o script.
    pause
)
