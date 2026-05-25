$ErrorActionPreference = "Stop"

# Retry isolado 2026-05-25 (TC/000524/2022 - Saude, REITERACAO):
# Originalmente caiu no batch run_32_processos_DILACAO_2026_05_25, mas a sonda
# read-only (probe_pieces_TC_000524_2022.py) revelou que ESTE PROCESSO NAO E
# DILACAO. Conforme:
#   - Peca 23 (DES 897/2026, 21/05/2026, Conselheiro Eduardo Tuma):
#     "Objeto: Aposentadoria. Prazo transcorrido in albis. Alerta de 2a reiteracao.
#      DESPACHO. ... RENOVE-SE o oficio de pecas 15 e 20, servindo este como alerta
#      de 2a reiteracao."
#   - Peca 22 (INF 3403/2026, 11/05/2026): "Transcorreu o prazo sem manifestacao."
#   - Nao existe peca REQUERIMENTO da SMS no autos (SMS nunca respondeu).
# Portanto: FORCE_TIPO=REITERACAO, template modelos_reiteracao - Saude.
# Referencia (cabecalho) = ofcio SSG anterior (peca 19 = OF SSG 13845/2026),
# seguindo o mesmo padrao que a peca 19 usou (1a reiteracao).
# Mesmo fluxo: cleanup destrutivo + comm + Oficio SSG + concluido + assinatura
# Roseli Chaves, SEM tramitar. Data 25/05/2026.

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

$env:PROCESSOS_LIST = "TC/000524/2022"
$env:ONLY_PROCESSOS_AUTHORIZED = "TC/000524/2022"

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

# Oficio: REITERACAO, template auto por secretaria (Saude). O template de
# reiteracao tem "Cópia da(s) peça(s) XX dos autos" no Encaminha — XX e o
# numero da peca preferida (MANUTAP-OF), preenchido pelo proprio main.py.
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

# Override de REITERACAO apos pytest. Referencia/SSG manuais (nao tem
# auto-extracao para reiteracao; nao existe REQUERIMENTO no autos).
$env:FORCE_TIPO = "REITERACAO"
$env:OFICIO_REFERENCIA_TEXT = "Ofício SSG 13845/2026, encaminhado eletronicamente em 03/03/2026."
$env:OFICIO_SSG_REF = "13845/2026"

Write-Host "[2/2] Retry isolado TC/000524/2022 (REITERACAO Saude, data 25/05/2026)..." -ForegroundColor Cyan
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "logs\run_retry_TC_000524_2022_$ts.log"
$sw = [System.IO.StreamWriter]::new($logPath, $false, [System.Text.UTF8Encoding]::new($false))
try {
    & $pythonExe ".\src\main.py" 2>&1 | ForEach-Object {
        $line = "{0}`t{1}" -f (Get-Date).ToString("o"), $_
        Write-Host $line
        $sw.WriteLine($line); $sw.Flush()
    }
} finally { $sw.Close() }

$ok = Select-String -Path $logPath -Pattern "Assinatura solicitada para .* no processo TC/000524/2022" -Quiet
if ($ok) { Write-Host "RESULTADO: TC/000524/2022 -> OK" -ForegroundColor Green; exit 0 }
else { Write-Host "RESULTADO: TC/000524/2022 -> FALHA (ver $logPath)" -ForegroundColor Red; exit 1 }
