$ErrorActionPreference = "Stop"

# ============================================================================
# Projeto Euclides - Interface web (FastAPI + Uvicorn)
# Sobe o disparador HTTP em 0.0.0.0:8080 para o Gilson acessar pela rede
# interna via http://<ip-da-vm>:8080
# ============================================================================

Set-Location -Path $PSScriptRoot

# 1. Promove env vars de usuario para o processo (mesma logica dos runners PROD).
foreach ($name in @("ETCM_USERNAME","ETCM_PASSWORD","ETCM_USER","ETCM_LOGIN","ETCM_PASS","ETCM_SENHA")) {
    $val = [Environment]::GetEnvironmentVariable($name, "User")
    if (-not [string]::IsNullOrEmpty($val)) { Set-Item -Path "Env:$name" -Value $val }
}
if (
    [string]::IsNullOrWhiteSpace($env:ETCM_USERNAME) -and
    [string]::IsNullOrWhiteSpace($env:ETCM_USER) -and
    [string]::IsNullOrWhiteSpace($env:ETCM_LOGIN)
) {
    Write-Host "[ERRO] Usuario do e-TCM nao configurado em variavel de ambiente." -ForegroundColor Red
    Write-Host "       Configure ETCM_USERNAME/ETCM_PASSWORD (variaveis de usuario) e reabra o PowerShell." -ForegroundColor Yellow
    exit 1
}

$env:PYTHONUNBUFFERED = "1"
$env:PYTHONIOENCODING = "utf-8"

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado. Rode:" -ForegroundColor Red
    Write-Host "       python -m venv .venv" -ForegroundColor Yellow
    Write-Host "       .\.venv\Scripts\pip install -r requirements.txt" -ForegroundColor Yellow
    exit 1
}

# 2. Garante deps web instaladas (idempotente).
$needs = & $pythonExe -c "import fastapi, uvicorn, multipart, jinja2" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[INFO] Instalando dependencias web..." -ForegroundColor Cyan
    & (Join-Path ".venv\Scripts" "pip.exe") install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { Write-Host "[ERRO] Falha no pip install" -ForegroundColor Red; exit 1 }
}

# 3. Descobre IP da VM para imprimir o link util para o Gilson.
$ip = (Get-NetIPAddress -AddressFamily IPv4 -PrefixOrigin Dhcp,Manual `
        | Where-Object { $_.IPAddress -notmatch '^127\.' -and $_.IPAddress -notmatch '^169\.254' } `
        | Select-Object -First 1 -ExpandProperty IPAddress)

$port = if ($env:EUCLIDES_WEB_PORT) { $env:EUCLIDES_WEB_PORT } else { "8080" }

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host " Projeto Euclides - interface web                         " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host (" Local : http://localhost:{0}" -f $port) -ForegroundColor White
if ($ip) {
    Write-Host (" Rede  : http://{0}:{1}    <-- enviar para o Gilson" -f $ip, $port) -ForegroundColor Yellow
}
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""

# 4. Sobe o uvicorn.
& $pythonExe -m uvicorn web.app:app --host 0.0.0.0 --port $port --log-level info
