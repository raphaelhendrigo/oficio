$ErrorActionPreference = "Stop"

# RECOVERY MODE 2026-05-25 (TC/009208/2023 - Educacao, REITERACAO):
# A comm 6692/2026 (robot-created, marker "gerado automaticamente") foi salva
# server-side na ultima tentativa, mas o script declarou falha por nao detectar
# a confirmacao na grid em 120s. A evidencia HTML confirma a existencia.
#
# Estrategia: rodar com RUN_PROD_DESTRUCTIVE_CLEANUP=false. Sem cleanup, o
# pre-create audit do main.py detecta a comm robo existente (1 robot-created)
# e segue o caminho "REUSANDO" — anexa o DOCX a comm 6692/2026, conclui o ato
# e pede assinatura, sem tentar criar uma nova comm.
#
# REITERACAO Educacao, data 25/05/2026. Mesma configuracao de auto-extracao da
# corrida anterior — o robo confirma SSG/Referencia/pecas Encaminha do nome
# das pecas.

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

$env:PROCESSOS_LIST = "TC/009208/2023"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/009208/2023"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# DESLIGADO: cleanup destrutivo. Queremos PRESERVAR a comm 6692/2026 ja salva
# server-side para o pre-create audit detectar e seguir caminho REUSANDO.
$env:USE_CAIXA_CORREIO = "true"
$env:SAFE_DELETE_OWN_DRAFTS = "false"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "false"
$env:FORCE_DELETE_OLD_OFICIO_SSG = "false"
$env:FORCE_RECREATE_COMUNICACAO = "false"
$env:FORCE_REVOKE_PENDING_ROSELI_SIGNATURE = "false"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"
$env:SKIP_COMUNICACAO_CLEANUP = "true"
$env:REUSE_EXISTING_OFICIO = "false"

# LIGADO: pedir assinatura
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
$env:DATA_OFICIO = "25/05/2026"
$env:OFICIO_DATA = "25/05/2026"

$env:MAX_PROCESSOS = "0"

Set-Location -Path $PSScriptRoot

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) { Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red; exit 1 }

Write-Host "[1/2] Rodando pytest oficial..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) { Write-Host "[ERRO] pytest falhou. Abortando." -ForegroundColor Red; exit $LASTEXITCODE }

# REITERACAO + exige auto-extracao (Referencia / SSG / pecas Encaminha).
$env:FORCE_TIPO = "REITERACAO"
$env:REITERACAO_REQUIRE_AUTO_EXTRACT = "true"

Write-Host "[2/2] REUSE mode TC/009208/2023 (REITERACAO Educacao, reaproveitando comm 6692/2026)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "logs\run_retry_TC_009208_2023_REUSE_$ts.log"
$sw = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.UTF8Encoding]::new($false))
try {
    & $pythonExe ".\src\main.py" 2>&1 | ForEach-Object {
        $line = "{0}`t{1}" -f (Get-Date).ToString("o"), $_
        Write-Host $line
        $sw.WriteLine($line); $sw.Flush()
    }
} finally { $sw.Close() }

$ok = Select-String -Path $logPath -Pattern "Assinatura solicitada para .* no processo TC/009208/2023" -Quiet
if ($ok) { Write-Host "RESULTADO: TC/009208/2023 -> OK (reuse)" -ForegroundColor Green; exit 0 }
else { Write-Host "RESULTADO: TC/009208/2023 -> FALHA (ver $logPath)" -ForegroundColor Red; exit 1 }
