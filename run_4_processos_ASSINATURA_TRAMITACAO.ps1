$ErrorActionPreference = "Stop"

# Script 2026-05-15 (ASSINATURA + TRAMITACAO):
# Para os 4 processos com Of'icio SSG ja concluido (TC/007902, TC/008084,
# TC/013838, TC/018149), solicita assinatura para a Dra. Roseli Chaves
# e tramita 'Em assinatura'.
#
# Estrategia: reaproveita o pipeline existente desligando o que nao precisa:
#   - SAFE_DELETE_OWN_DRAFTS=false      -> nao apaga of'icio/comm existentes
#   - USE_CAIXA_CORREIO=false           -> nao mexe em comm (ja criadas)
#   - REUSE_EXISTING_OFICIO=true        -> reusa o Ofi'cio SSG ja concluido
#   - STOP_AFTER_OFICIO_CONCLUIDO=false -> avanca para assinar + tramitar
#   - SKIP_SIGNATURE=false              -> pede assinatura
#   - SKIP_TRAMITACAO=false             -> tramita
#   - ASSINANTE_NOME="Roseli Chaves"    -> assinante (matcher tolera variacoes)
#   - TRAMITAR_DESTINO="Em assinatura"  -> destino

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

# Os 4 processos com Of'icio SSG concluido vinculado a comm:
$env:PROCESSOS_LIST = "TC/007902/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# DESLIGADO: nao mexer em ofi'cio/comm existentes
$env:USE_CAIXA_CORREIO = "false"
$env:SAFE_DELETE_OWN_DRAFTS = "false"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "false"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "false"
$env:FORCE_RECREATE_COMUNICACAO = "false"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"
$env:REUSE_EXISTING_OFICIO = "true"

# LIGADO: pedir assinatura + tramitar
$env:STOP_AFTER_OFICIO_CONCLUIDO = "false"
$env:SKIP_SIGNATURE = "false"
$env:SKIP_TRAMITACAO = "false"
$env:REQUEST_SIGNATURE = "true"
$env:ASSINANTE_NOME = "Roseli Chaves"
$env:SIGNER_NAME = "Roseli Chaves"
$env:DISTRIBUIR_PARA = "Roseli Chaves"
$env:TRAMITAR_DESTINO = "Em assinatura"

# Esses params nao serao usados (USE_CAIXA_CORREIO=false), mas mantemos coerencia.
$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

$env:OFICIO_TEMPLATE_MODE = "auto"
$env:OFICIO_PRESERVE_AT_TOKENS = "true"
$env:OFICIO_ADD_EUCLIDES_MARKER = "true"
$env:OFICIO_EUCLIDES_MARKER = "/euclides"
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "true"
$env:OFICIO_ENCAMINHA_TEXT_BOLD = "false"
$env:OFICIO_ENCAMINHA_FONT_NAME = "Times New Roman"
$env:OFICIO_ENCAMINHA_FONT_SIZE_PT = "12"

$env:MAX_PROCESSOS = "0"

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "[1/2] Rodando pytest oficial..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] pytest falhou. Abortando." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "[2/2] Solicitando assinatura (Roseli Chaves) + tramitando (Em assinatura) para os 4 processos..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\run_4_processos_ASSINATURA_TRAMITACAO_$ts.log"
$errFile = "logs\run_4_processos_ASSINATURA_TRAMITACAO_$ts.err.log"

$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile

Write-Host "--- stdout ultimas 80 linhas ---"
if (Test-Path $outFile) { Get-Content $outFile -Tail 80 }

Write-Host "--- stderr ultimas 40 linhas ---"
if (Test-Path $errFile) { Get-Content $errFile -Tail 40 }

exit $proc.ExitCode
