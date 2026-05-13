$ErrorActionPreference = "Stop"

foreach ($name in @("ETCM_USERNAME", "ETCM_PASSWORD", "ETCM_USER", "ETCM_LOGIN", "ETCM_PASS", "ETCM_SENHA")) {
    $val = [Environment]::GetEnvironmentVariable($name, "User")
    if (-not [string]::IsNullOrEmpty($val)) {
        Set-Item -Path "Env:$name" -Value $val
    }
}

if (
    [string]::IsNullOrWhiteSpace($env:ETCM_USERNAME) -and
    [string]::IsNullOrWhiteSpace($env:ETCM_USER) -and
    [string]::IsNullOrWhiteSpace($env:ETCM_LOGIN)
) {
    Write-Host "[ERRO] Usuario do e-TCM nao configurado em variavel de ambiente." -ForegroundColor Red
    exit 1
}

$env:ETCM_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
$env:BASE_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"

$env:PROCESSOS_LIST = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

$env:USE_CAIXA_CORREIO = "true"
$env:REQUEST_SIGNATURE = "true"
$env:ASSINANTE_NOME = "Roseli Chaves"
$env:TRAMITAR_DESTINO = "Em assinatura"

$env:OFICIO_TEMPLATE_MODE = "auto"
$env:OFICIO_PRESERVE_AT_TOKENS = "true"
$env:OFICIO_ADD_EUCLIDES_MARKER = "true"
$env:OFICIO_EUCLIDES_MARKER = "\euclides"
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "true"
$env:REUSE_EXISTING_OFICIO = "false"

$env:ENVIRONMENT = "producao"
$env:SAFE_DELETE_OWN_DRAFTS = "true"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "true"
$env:FORCE_RECREATE_COMUNICACAO = "true"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "true"
$env:FORCE_REVOKE_PENDING_ROSELI_SIGNATURE = "true"
$env:NEVER_DELETE_FINAL_SIGNED = "true"

$env:MAX_PROCESSOS = "0"

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "[1/3] Removendo/ignorando artefatos temporarios de analise..." -ForegroundColor Cyan
if (Test-Path "_zip_analysis") { Remove-Item "_zip_analysis" -Recurse -Force }

Write-Host "[2/3] Rodando pytest oficial..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] pytest falhou. Abortando producao." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "[3/3] Rodando recriacao em producao nos 5 processos..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\run_5_processos_PROD_RECREATE_$ts.log"
$errFile = "logs\run_5_processos_PROD_RECREATE_$ts.err.log"

$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile

Write-Host "--- stdout ultimas 80 linhas ---"
if (Test-Path $outFile) { Get-Content $outFile -Tail 80 }

Write-Host "--- stderr ultimas 40 linhas ---"
if (Test-Path $errFile) { Get-Content $errFile -Tail 40 }

exit $proc.ExitCode
