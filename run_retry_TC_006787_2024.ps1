$ErrorActionPreference = "Stop"

# Retry isolado 2026-05-22 (TC/006787/2024 - Urbanismo, template Geral):
# Falhou nas 3 tentativas do lote run_24_processos_UTAP_2026_05_22 no mesmo ponto
# (comunicacao processual nao confirmada na grid - bug intermitente do form DevExpress).
# O cleanup destrutivo JA rodou no lote (comm/oficio antigos apagados), entao o
# processo esta incompleto; este retry recria tudo do zero.
# Mesmo fluxo: cleanup destrutivo + comm + Oficio SSG + concluido + assinatura
# para Roseli Chaves, SEM tramitar. Data 22/05/2026. Secretaria: Geral (auto).

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
    exit 1
}

$env:ETCM_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
$env:BASE_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"

$env:PYTHONUNBUFFERED = "1"
$env:PYTHONIOENCODING = "utf-8"

$env:PROCESSOS_LIST = "TC/006787/2024"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/006787/2024"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

$env:USE_CAIXA_CORREIO = "true"
$env:SAFE_DELETE_OWN_DRAFTS = "true"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "true"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "true"
$env:FORCE_RECREATE_COMUNICACAO = "true"
$env:FORCE_REVOKE_PENDING_ROSELI_SIGNATURE = "true"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"
$env:SKIP_COMUNICACAO_CLEANUP = "false"
$env:REUSE_EXISTING_OFICIO = "false"

$env:STOP_AFTER_OFICIO_CONCLUIDO = "false"
$env:SKIP_SIGNATURE = "false"
$env:REQUEST_SIGNATURE = "true"
$env:ASSINANTE_NOME = "Roseli Chaves"
$env:SIGNER_NAME = "Roseli Chaves"

$env:SKIP_TRAMITACAO = "true"
$env:TRAMITAR_DESTINO = ""
$env:DISTRIBUIR_PARA = ""

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
$env:DATA_OFICIO = "22/05/2026"
$env:OFICIO_DATA = "22/05/2026"

$env:MAX_PROCESSOS = "0"

Set-Location -Path $PSScriptRoot

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "[1/1] Retry isolado TC/006787/2024 (cleanup + comm + oficio + assinatura SEM tramitar, data 22/05/2026)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "logs\run_retry_TC_006787_2024_$ts.log"

$sw = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.UTF8Encoding]::new($false))
try {
    & $pythonExe ".\src\main.py" 2>&1 | ForEach-Object {
        $line = "{0}`t{1}" -f (Get-Date).ToString("o"), $_
        Write-Host $line
        $sw.WriteLine($line)
        $sw.Flush()
    }
} finally {
    $sw.Close()
}

$success = Select-String -Path $logPath -Pattern "Assinatura solicitada para .* no processo TC/006787/2024" -Quiet
if ($success) {
    Write-Host "RESULTADO: TC/006787/2024 -> OK (assinatura solicitada)." -ForegroundColor Green
    exit 0
} else {
    Write-Host "RESULTADO: TC/006787/2024 -> FALHA (assinatura nao solicitada). Ver $logPath" -ForegroundColor Red
    exit 1
}
