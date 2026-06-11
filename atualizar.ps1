# Atualizar Dashboard - Guarda Municipal BC
# Uso: .\atualizar.ps1
# Ou com mensagem: .\atualizar.ps1 "sua mensagem"

param([string]$mensagem = "")

$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Dashboard Seguranca BC - Atualizando..." -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# 1. Gerar HTML
Write-Host ""
Write-Host "[1/5] Gerando dashboard HTML..." -ForegroundColor Yellow
python build_dashboard.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO ao gerar o dashboard. Verifique o Python." -ForegroundColor Red
    pause
    exit 1
}
Write-Host "      OK" -ForegroundColor Green

# 2. Copiar para index.html
Write-Host ""
Write-Host "[2/5] Copiando para index.html..." -ForegroundColor Yellow
Copy-Item dashboard_interativo.html index.html -Force
Write-Host "      OK" -ForegroundColor Green

# 3. Commit
Write-Host ""
Write-Host "[3/5] Fazendo commit..." -ForegroundColor Yellow
git add secretario.xlsx dashboard_interativo.html index.html build_dashboard.py geocache.json

if ($mensagem -eq "") {
    $data = Get-Date -Format "dd/MM/yyyy HH:mm"
    $mensagem = "Atualizar dashboard - $data [skip ci]"
}
git commit -m $mensagem
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Nada novo para commitar (ou ja commitado)." -ForegroundColor DarkYellow
} else {
    Write-Host "      OK - $mensagem" -ForegroundColor Green
}

# 4. Push para GitHub
Write-Host ""
Write-Host "[4/5] Enviando para GitHub..." -ForegroundColor Yellow
git push origin master:main
git push origin master
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO ao fazer push. Verifique a conexao com o GitHub." -ForegroundColor Red
    pause
    exit 1
}
Write-Host "      OK" -ForegroundColor Green

# 5. Concluido
Write-Host ""
Write-Host "[5/5] Deploy enviado para a Vercel..." -ForegroundColor Yellow
Write-Host "      A Vercel atualiza automaticamente em ~1 minuto." -ForegroundColor Gray

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Concluido! Acesse em ~1 minuto:" -ForegroundColor Cyan
Write-Host "  https://dashboardgmbc.com.br" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
pause
