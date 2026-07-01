$ErrorActionPreference = "Stop"

# ============================================================================
# DEMO 2026-07-01 - TC/008940/2023 (REITERACAO) com Chrome VISIVEL
# Objetivo: validar em PROD o NOVO fluxo do ENCAMINHAMENTO adicionado hoje.
# Encaminhamento e gerado + anexado apos o oficio SSG (mesma comunicacao).
# Assinante ativo: Daniela Shimizu. Diretor: automatico pela secretaria.
# ============================================================================

$processo = "TC/008940/2023"

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
# Demo: Chrome VISIVEL com slowmo pra acompanhar.
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

# Assinante: Daniela Shimizu (ativa hoje na home).
$env:ASSINANTE_NOME = "Daniela Shimizu"
$env:SIGNER_NAME = "Daniela Shimizu"
$env:SIGNER_MATCH_TOKENS = "daniela,shimizu"
$env:SKIP_TRAMITACAO = "true"
$env:TRAMITAR_DESTINO = ""
$env:DISTRIBUIR_PARA = ""

# Templates de Daniela (defaults do runner web).
$env:OFICIO_TEMPLATES_DIR_UTAP = "modelos daniela/manutap"
$env:OFICIO_TEMPLATES_DIR_DILACAO = "modelos daniela/dilação"
$env:OFICIO_TEMPLATES_DIR_REITERACAO = "modelos daniela/reiteração"
$env:OFICIO_SECRETARIA_FALLBACK = "educacao"

$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

$env:OFICIO_TEMPLATE_MODE = "auto"
$env:OFICIO_PRESERVE_AT_TOKENS = "true"
$env:OFICIO_ADD_EUCLIDES_MARKER = "true"
$env:OFICIO_EUCLIDES_MARKER = "/euclides"
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "true"
$env:OFICIO_ENCAMINHA_TEXT_BOLD = "false"
$env:REITERACAO_REQUIRE_AUTO_EXTRACT = "true"
$dataHoje = (Get-Date).ToString('dd/MM/yyyy')
$env:DATA_OFICIO = $dataHoje
$env:OFICIO_DATA = $dataHoje

$env:MAX_PROCESSOS = "0"
$env:PROCESSOS_LIST = $processo
$env:ONLY_PROCESSOS_AUTHORIZED = $processo
$env:FORCE_TIPO = "REITERACAO"
$env:EUCLIDES_JOB_ID = "demo_2026-07-01_TC008940_2023_encaminhamento"

# Habilita o novo fluxo do encaminhamento (default = habilitado).
$env:SKIP_ENCAMINHAMENTO = "false"

Set-Location -Path $PSScriptRoot
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
$logPath = "logs\DEMO_TC_008940_2023_ENCAMINHAMENTO_$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " DEMO ENCAMINHAMENTO - TC/008940/2023 - REITERACAO        " -ForegroundColor Cyan
Write-Host " Assinante: Daniela Shimizu | Data: $dataHoje             " -ForegroundColor Cyan
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
