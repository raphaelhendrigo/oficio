$ErrorActionPreference = "Stop"

# ============================================================================
# DEMO 2026-06-12 - TC/008177/2023 (DILACAO) com Chrome VISIVEL
# Espelha run_4_DILACAO_GILSON_2026_05_25.ps1 para um unico processo,
# trocando HEADLESS por modo visivel para demonstracao do fluxo completo.
# ============================================================================

$processo = "TC/008177/2023"

foreach ($name in @("ETCM_USERNAME","ETCM_PASSWORD","ETCM_USER","ETCM_LOGIN","ETCM_PASS","ETCM_SENHA")) {
    $val = [Environment]::GetEnvironmentVariable($name, "User")
    if (-not [string]::IsNullOrEmpty($val)) { Set-Item -Path "Env:$name" -Value $val }
}
if ([string]::IsNullOrWhiteSpace($env:ETCM_USERNAME) -and [string]::IsNullOrWhiteSpace($env:ETCM_USER)) {
    Write-Host "[ERRO] credenciais nao configuradas." -ForegroundColor Red; exit 1
}

$env:ETCM_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
$env:BASE_URL = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
$env:PYTHONUNBUFFERED = "1"
$env:PYTHONIOENCODING = "utf-8"

$env:ENVIRONMENT = "producao"
# Demo: Chrome VISIVEL com slowmo para acompanhar.
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:DEVTOOLS = "false"
$env:SLOWMO_MS = "300"
$env:LOGIN_MANUAL_WAIT_MS = "30000"
$env:PAUSE_AFTER_LOGIN_MS = "3000"

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
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "false"
$env:OFICIO_ENCAMINHA_TEXT_BOLD = "false"
$dataHoje = (Get-Date).ToString('dd/MM/yyyy')
$env:DATA_OFICIO = $dataHoje
$env:OFICIO_DATA = $dataHoje

$env:MAX_PROCESSOS = "0"
$env:PROCESSOS_LIST = $processo
$env:ONLY_PROCESSOS_AUTHORIZED = $processo
$env:FORCE_TIPO = "DILACAO"
$env:EUCLIDES_JOB_ID = "demo_2026-06-12_TC008177_2023"

Set-Location -Path $PSScriptRoot
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
$logPath = "logs\DEMO_TC_008177_2023_$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " DEMO - TC/008177/2023 - DILACAO - Chrome VISIVEL         " -ForegroundColor Cyan
Write-Host " Data Oficio: $dataHoje | Assinante: Roseli Chaves        " -ForegroundColor Cyan
Write-Host " Log: $logPath                                            " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

& ".venv\Scripts\python.exe" "-u" ".\src\main.py" 2>&1 | ForEach-Object {
    $line = "{0}`t{1}" -f (Get-Date).ToString('HH:mm:ss'), $_
    Write-Host $line
    Add-Content -Path $logPath -Value $line -Encoding utf8
}

Write-Host ""
Write-Host "Exit code: $LASTEXITCODE" -ForegroundColor Yellow
Write-Host "Log salvo em: $logPath" -ForegroundColor Yellow
exit $LASTEXITCODE
