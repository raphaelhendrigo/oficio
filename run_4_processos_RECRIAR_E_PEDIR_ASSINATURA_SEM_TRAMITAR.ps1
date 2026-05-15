$ErrorActionPreference = "Stop"

# Script 2026-05-15 (RECRIAR + ASSINATURA, SEM TRAMITAR):
# Refaz do zero (cleanup completo + nova comm + novo ofi'cio) os 4 processos
# abaixo e ao final SOLICITA assinatura para a Dra. Roseli Chaves.
# IMPORTANTE: nao tramita o processo apos pedir assinatura.
#
# Estado atual desses processos (rodada 11:17): estao em fila 'Em assinatura'
# com ato concluido + solicitacao de assinatura pendente. O cleanup vai:
#   1. Estornar a solicitacao de assinatura (ato volta a 'Concluido')
#   2. Estornar a conclusao (ato volta a 'Rascunho')
#   3. Excluir o ato rascunho
#   4. Excluir a comm robot-marker
# Apos isso, o pipeline recria comm + ofi'cio com TNR 12pt no Encaminha,
# anexa via ?ntf=<cod_notif> (vinculado a comm) e pede assinatura.

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

# Retry dos 2 que falharam no form submit (bug intermitente do DevExpress
# em TC/007902 e TC/013838).
$env:PROCESSOS_LIST = "TC/013838/2023"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# LIGADOS: cleanup completo + recriar comm + reanexar oficio
$env:USE_CAIXA_CORREIO = "true"
$env:SAFE_DELETE_OWN_DRAFTS = "true"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "true"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "true"
$env:FORCE_RECREATE_COMUNICACAO = "true"
$env:FORCE_REVOKE_PENDING_ROSELI_SIGNATURE = "true"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"
$env:SKIP_COMUNICACAO_CLEANUP = "false"
$env:REUSE_EXISTING_OFICIO = "false"

# LIGADO: pedir assinatura
$env:STOP_AFTER_OFICIO_CONCLUIDO = "false"
$env:SKIP_SIGNATURE = "false"
$env:REQUEST_SIGNATURE = "true"
$env:ASSINANTE_NOME = "Roseli Chaves"
$env:SIGNER_NAME = "Roseli Chaves"

# DESLIGADO: NAO tramitar
$env:SKIP_TRAMITACAO = "true"
$env:TRAMITAR_DESTINO = ""
$env:DISTRIBUIR_PARA = ""

# Comunicacao Processual
$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

# Ofi'cio (TNR 12pt no Encaminha, marker /euclides)
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

Write-Host "[2/2] Recriando do zero (cleanup + nova comm + novo of'icio + pedir assinatura)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\run_4_processos_RECRIAR_E_PEDIR_ASSINATURA_SEM_TRAMITAR_$ts.log"
$errFile = "logs\run_4_processos_RECRIAR_E_PEDIR_ASSINATURA_SEM_TRAMITAR_$ts.err.log"

$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile

Write-Host "--- stdout ultimas 80 linhas ---"
if (Test-Path $outFile) { Get-Content $outFile -Tail 80 }

Write-Host "--- stderr ultimas 40 linhas ---"
if (Test-Path $errFile) { Get-Content $errFile -Tail 40 }

exit $proc.ExitCode
