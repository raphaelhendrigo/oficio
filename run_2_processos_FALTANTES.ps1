$ErrorActionPreference = "Stop"

# Script TEMPORARIO 2026-05-14 - apenas TC/007902/2022 e TC/008636/2022
# (que falharam na rodada das 18:07). Reusa toda a configuracao do script
# completo, mas restringe PROCESSOS_LIST aos 2 pendentes.
# ONLY_PROCESSOS_AUTHORIZED continua com os 5 para que o cleanup destrutivo
# seja permitido sem fora-de-escopo.

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

$env:PYTHONUNBUFFERED = "1"
$env:PYTHONIOENCODING = "utf-8"

# Retry de TC/007902/2022 (form submit timing intermitente nesse processo)
$env:PROCESSOS_LIST = "TC/007902/2022"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

$env:USE_CAIXA_CORREIO = "true"
$env:FORCE_RECREATE_COMUNICACAO = "true"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "true"
$env:SAFE_DELETE_OWN_DRAFTS = "true"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "true"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"

$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

$env:REQUEST_SIGNATURE = "false"
$env:ASSINANTE_NOME = ""
$env:TRAMITAR_DESTINO = ""
$env:SKIP_SIGNATURE = "true"
$env:SKIP_TRAMITACAO = "true"
$env:STOP_AFTER_OFICIO_CONCLUIDO = "true"

$env:OFICIO_TEMPLATE_MODE = "auto"
$env:OFICIO_PRESERVE_AT_TOKENS = "true"
$env:OFICIO_ADD_EUCLIDES_MARKER = "true"
$env:OFICIO_EUCLIDES_MARKER = "/euclides"
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "true"
$env:OFICIO_ENCAMINHA_TEXT_BOLD = "false"

$env:MAX_PROCESSOS = "0"

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "Rodando TC/007902/2022 e TC/008636/2022..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\run_2_processos_FALTANTES_$ts.log"
$errFile = "logs\run_2_processos_FALTANTES_$ts.err.log"

$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile

Write-Host "--- stdout ultimas 80 linhas ---"
if (Test-Path $outFile) { Get-Content $outFile -Tail 80 }

Write-Host "--- stderr ultimas 40 linhas ---"
if (Test-Path $errFile) { Get-Content $errFile -Tail 40 }

exit $proc.ExitCode
