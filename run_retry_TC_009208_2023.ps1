$ErrorActionPreference = "Stop"

# Teste 2026-05-25 (TC/009208/2023 - Educacao, REITERACAO):
# Primeiro caso usando o fluxo de REITERACAO permanente (briefing Gilson Nobrega).
#
# Regras de REITERACAO (auto-extraidas do NOME das pecas pelo main.py, sem
# precisar setar nada manualmente):
#   - Referencia (cabecalho)  : "Ofício SSG <num/ano>, encaminhado eletronicamente em <DD/MM/AAAA>."
#                               -> peca SSG logo apos o ULTIMO MANUTAP-OF.
#   - Encaminha               : "Cópia das peças <MANUTAP> e <DES> dos autos."
#                               -> ultima peca MANUTAP-OF + ultima peca DES.
#   - Corpo do oficio (SSG)   : mesmo nº/ano da Referencia.
# Quem faz tudo isso: src/main.py -> _extract_reiteracao_data_from_pieces().
# Setamos REITERACAO_REQUIRE_AUTO_EXTRACT=true para falhar cedo se nao
# conseguir extrair (em vez de gerar oficio com campos vazios).
#
# Mesmo fluxo destrutivo + assinatura: cleanup + comm + Oficio SSG (template
# REITERACAO Educacao) + concluido + assinatura Roseli Chaves, SEM tramitar.
# Data 25/05/2026.

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

# Comunicacao Processual: 60 dias, descricao "reiteracao" (auto pelo normalize)
$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

# Oficio: REITERACAO, template auto por secretaria (Educacao). O template tem
# "Cópia da(s) peça(s) XX dos autos." e exigimos que a substituicao funcione
# (true). O XX e substituido pelos numeros vindos da auto-extracao via
# OFICIO_ENCAMINHA_PIECE_NUMBERS (setado pelo main.py).
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

Write-Host "[1/2] Rodando pytest oficial (com FORCE_TIPO vazio)..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) { Write-Host "[ERRO] pytest falhou. Abortando." -ForegroundColor Red; exit $LASTEXITCODE }

# Override REITERACAO + exige auto-extracao (Referencia / SSG / pecas Encaminha).
$env:FORCE_TIPO = "REITERACAO"
$env:REITERACAO_REQUIRE_AUTO_EXTRACT = "true"
# IMPORTANTE: NAO setar OFICIO_REFERENCIA_TEXT / OFICIO_SSG_REF /
# OFICIO_ENCAMINHA_PIECE_NUMBERS aqui — a auto-extracao no main.py preenche
# esses tres a partir do nome das pecas do processo. Setar manualmente
# atrapalha a verificacao do fluxo permanente.

Write-Host "[2/2] Retry isolado TC/009208/2023 (REITERACAO Educacao, data 25/05/2026)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "logs\run_retry_TC_009208_2023_$ts.log"
$sw = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.UTF8Encoding]::new($false))
try {
    & $pythonExe ".\src\main.py" 2>&1 | ForEach-Object {
        $line = "{0}`t{1}" -f (Get-Date).ToString("o"), $_
        Write-Host $line
        $sw.WriteLine($line); $sw.Flush()
    }
} finally { $sw.Close() }

$ok = Select-String -Path $logPath -Pattern "Assinatura solicitada para .* no processo TC/009208/2023" -Quiet
if ($ok) { Write-Host "RESULTADO: TC/009208/2023 -> OK" -ForegroundColor Green; exit 0 }
else { Write-Host "RESULTADO: TC/009208/2023 -> FALHA (ver $logPath)" -ForegroundColor Red; exit 1 }
