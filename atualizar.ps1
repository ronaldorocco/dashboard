# ──────────────────────────────────────────────────
#  Atualizar Dashboard — Guarda Municipal BC
#  Uso: .\atualizar.ps1
#  Ou com mensagem: .\atualizar.ps1 "sua mensagem"
# ──────────────────────────────────────────────────

param([string]$mensagem = "")

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Dashboard Seguranca BC — Atualizando..." -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# 1. Gerar HTML
Write-Host ""
Write-Host "[1/4] Gerando dashboard HTML..." -ForegroundColor Yellow
python build_dashboard.py
if ($LASTEXITCODE -ne 0) { Write-Host "ERRO ao gerar o dashboard." -ForegroundColor Red; pause; exit 1 }
Write-Host "      OK" -ForegroundColor Green

# 2. Copiar para index.html
Write-Host ""
Write-Host "[2/4] Copiando para index.html..." -ForegroundColor Yellow
Copy-Item dashboard_interativo.html index.html -Force
Write-Host "      OK" -ForegroundColor Green

# 3. Commit
Write-Host ""
Write-Host "[3/4] Fazendo commit..." -ForegroundColor Yellow
git add dashboard_interativo.html index.html build_dashboard.py

if ($mensagem -eq "") {
    $data = Get-Date -Format "dd/MM/yyyy HH:mm"
    $mensagem = "Atualizar dashboard — $data"
}
git commit -m $mensagem
if ($LASTEXITCODE -ne 0) { Write-Host "      Nada novo para commitar." -ForegroundColor DarkYellow }
else { Write-Host "      OK — $mensagem" -ForegroundColor Green }

# 4. Push para GitHub (Vercel atualiza automaticamente)
Write-Host ""
Write-Host "[4/4] Enviando para GitHub..." -ForegroundColor Yellow
git push origin master:main
if ($LASTEXITCODE -ne 0) { Write-Host "ERRO ao fazer push." -ForegroundColor Red; pause; exit 1 }
Write-Host "      OK" -ForegroundColor Green

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Concluido! O Vercel vai atualizar em breve." -ForegroundColor Cyan
Write-Host "  https://gmbcdashboard.vercel.app" -ForegroundColor White
Write-Host "  https://gmbcdashboard.com.br" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
pause
