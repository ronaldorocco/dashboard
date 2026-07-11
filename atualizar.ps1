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
Write-Host "[1/6] Gerando dashboard HTML..." -ForegroundColor Yellow
python build_dashboard.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO ao gerar o dashboard. Verifique o Python." -ForegroundColor Red
    pause
    exit 1
}
Write-Host "      OK" -ForegroundColor Green

# 2. Copiar para index.html
Write-Host ""
Write-Host "[2/6] Copiando para index.html..." -ForegroundColor Yellow
Copy-Item dashboard_interativo.html index.html -Force
Write-Host "      OK" -ForegroundColor Green

# 3. Commit
Write-Host ""
Write-Host "[3/6] Fazendo commit..." -ForegroundColor Yellow
git config gc.auto 0
git config gc.autoPackLimit 0
git config maintenance.auto false
git -c gc.auto=0 -c gc.autoPackLimit=0 -c maintenance.auto=false add secretario.xlsx dashboard_interativo.html index.html build_dashboard.py geocache.json

if ($mensagem -eq "") {
    $data = Get-Date -Format "dd/MM/yyyy HH:mm"
    $mensagem = "Atualizar dashboard - $data [skip ci]"
}
git -c gc.auto=0 -c gc.autoPackLimit=0 -c maintenance.auto=false commit -m $mensagem
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Nada novo para commitar (ou ja commitado)." -ForegroundColor DarkYellow
} else {
    Write-Host "      OK - $mensagem" -ForegroundColor Green
}

# 4. Push para GitHub
Write-Host ""
Write-Host "[4/6] Enviando para GitHub..." -ForegroundColor Yellow
git -c gc.auto=0 -c gc.autoPackLimit=0 -c maintenance.auto=false push origin master:main
git -c gc.auto=0 -c gc.autoPackLimit=0 -c maintenance.auto=false push origin master
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO ao fazer push. Verifique a conexao com o GitHub." -ForegroundColor Red
    pause
    exit 1
}
Write-Host "      OK" -ForegroundColor Green

# 5. Disparar deploy na VPS (EasyPanel)
Write-Host ""
Write-Host "[5/6] Disparando deploy na VPS..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri "http://31.97.16.52:3000/api/deploy/447d005aee8f0ee2a0f5c313e3eb84a8afaaed356c389272" -UseBasicParsing -TimeoutSec 15 | Out-Null
    Write-Host "      OK" -ForegroundColor Green
} catch {
    Write-Host "      Aviso: nao foi possivel disparar o deploy na VPS ($($_.Exception.Message))" -ForegroundColor DarkYellow
}

# 6. Concluido
Write-Host ""
Write-Host "[6/6] Deploy enviado para Vercel + VPS..." -ForegroundColor Yellow
Write-Host "      A Vercel atualiza automaticamente em ~1 minuto." -ForegroundColor Gray
Write-Host "      A VPS leva ~1-2 minutos para rebuildar o container." -ForegroundColor Gray

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Concluido! Acesse em ~1 minuto:" -ForegroundColor Cyan
Write-Host "  https://dashboardgmbc.com.br" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
pause
