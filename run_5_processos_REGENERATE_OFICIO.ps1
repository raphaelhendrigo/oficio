$ErrorActionPreference = "Stop"

# Script 2026-05-14 (Regen-only): regenera apenas o DOCX + Ofício SSG ato
# nos 5 processos, REUSANDO as comunicacoes processuais ja criadas (5953,
# 5954, 5955, 5956). Objetivo: corrigir a fonte do Encaminha para
# Times New Roman 12pt sem negrito.
#
# Diferencas em relacao ao script principal:
#   - USE_CAIXA_CORREIO=false  -> nao cria nova comm
#   - SKIP_COMUNICACAO_CLEANUP=true -> cleanup nao tenta apagar comm
#   - Cleanup do Ofício SSG continua ligado (estorna conclusao + delete)
#   - Nova geracao do DOCX usa o set_encaminha_text_without_bold atualizado
#     (forca Times New Roman 12pt). Override por env disponivel se precisar.

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

$env:PROCESSOS_LIST = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# Regen-only: nao mexer em comm (reusar existente)
$env:USE_CAIXA_CORREIO = "false"
$env:SKIP_COMUNICACAO_CLEANUP = "true"

# Cleanup do Ofício SSG fica ligado (estornar conclusao + delete + recriar)
$env:FORCE_DELETE_OLD_OFICIO_SSG = "true"
$env:SAFE_DELETE_OWN_DRAFTS = "true"
$env:RUN_PROD_DESTRUCTIVE_CLEANUP = "true"
$env:FORCE_DELETE_ALL_COMUNICACOES_AUTHORIZED = "false"

# Esses params sao ignorados quando USE_CAIXA_CORREIO=false, mas mantidos
# para coerencia.
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

# Brief 2026-05-14: forca Times New Roman 12pt sem negrito no conteudo de
# Encaminha (default ja eh esse no codigo; aqui apenas explicito para log).
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
    Write-Host "[ERRO] pytest falhou. Abortando execucao." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "[2/2] Regenerando Oficio SSG nos 5 processos (font Encaminha = Times New Roman 12pt)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\run_5_processos_REGENERATE_OFICIO_$ts.log"
$errFile = "logs\run_5_processos_REGENERATE_OFICIO_$ts.err.log"

$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile

Write-Host "--- stdout ultimas 80 linhas ---"
if (Test-Path $outFile) { Get-Content $outFile -Tail 80 }

Write-Host "--- stderr ultimas 40 linhas ---"
if (Test-Path $errFile) { Get-Content $errFile -Tail 40 }

exit $proc.ExitCode
