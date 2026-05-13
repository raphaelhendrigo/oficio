# run_5_processos_corrigidos.ps1
#
# Executa o fluxo APO-PEN nos 5 processos pendentes em homologacao,
# usando os modelos novos (utap/dilacao/reiteracao/juizo x educacao/saude/geral)
# e a regra de preservacao de tokens @@... (preenchimento pelo e-TCM).
#
# IMPORTANTE - Seguranca:
#   - Este script NUNCA contem senha literal.
#   - As credenciais sao lidas das variaveis de ambiente de USUARIO do
#     Windows (ETCM_USERNAME, ETCM_PASSWORD, com aliases ETCM_USER,
#     ETCM_LOGIN, ETCM_PASS, ETCM_SENHA todas suportadas pelo src/config.py).
#   - Para configurar a senha de forma segura (sem ela passar pelo chat ou
#     pelo repo), execute uma unica vez:
#       Copy-Item scripts\set_local_user_env.template.ps1 scripts\set_local_user_env.ps1
#       powershell -ExecutionPolicy Bypass -File scripts\set_local_user_env.ps1
#     Depois feche e reabra o PowerShell antes de rodar este script.
#
# Pre-condicoes:
#   - ETCM_USERNAME e ETCM_PASSWORD configurados em "User" scope.
#   - .venv ja criado e dependencias instaladas (requirements.txt + playwright install).
#   - Repo na branch fix/preserve-at-tokens-modelos-apopen (ou main apos merge).

$ErrorActionPreference = "Stop"

# --------------------------- Sanidade de credenciais ------------------------

# Quando o PowerShell e iniciado a partir de outro processo (sandbox, etc),
# pode nao herdar as vars de User scope automaticamente. Le explicitamente
# do registro e injeta no processo atual (e nos filhos).
foreach ($name in @("ETCM_USERNAME", "ETCM_PASSWORD", "ETCM_USER", "ETCM_LOGIN", "ETCM_PASS", "ETCM_SENHA")) {
    $val = [Environment]::GetEnvironmentVariable($name, "User")
    if (-not [string]::IsNullOrEmpty($val)) {
        Set-Item -Path "Env:$name" -Value $val
    }
}

$user = $env:ETCM_USERNAME
$pass = $env:ETCM_PASSWORD
if ([string]::IsNullOrWhiteSpace($user)) {
    Write-Host "[ERRO] ETCM_USERNAME nao esta configurado em User scope." -ForegroundColor Red
    Write-Host "      Rode scripts\set_local_user_env.ps1 antes de executar." -ForegroundColor Yellow
    exit 1
}
if ([string]::IsNullOrWhiteSpace($pass)) {
    Write-Host "[AVISO] ETCM_PASSWORD vazio." -ForegroundColor Yellow
    Write-Host "        O fluxo de src/main.py permite login MANUAL na janela aberta" -ForegroundColor Yellow
    Write-Host "        (timeout via LOGIN_MANUAL_WAIT_MS); voce vai precisar digitar a senha" -ForegroundColor Yellow
    Write-Host "        no navegador quando ele abrir." -ForegroundColor Yellow
}

# --------------------------- Ambiente alvo (PROD) ---------------------------

# Os 5 processos alvo estao em PRODUCAO. URL correta confirmada pelo usuario
# em 2026-05-13: https://etcm.tcm.sp.gov.br/ (sem o prefixo "homologacao-").
# config.py aceita ETCM_URL OR BASE_URL — definimos ambos por seguranca.
$env:ETCM_URL  = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"
$env:BASE_URL  = "https://etcm.tcm.sp.gov.br/paginas/login.aspx"

# --------------------------- Lista de processos -----------------------------

$env:PROCESSOS_LIST = "TC/007902/2022,TC/008636/2022,TC/008084/2023,TC/013838/2023,TC/018149/2024"

# --------------------------- Modo navegador ---------------------------------

$env:HEADLESS = "false"          # janela visivel
$env:SHOW_BROWSER = "true"       # forca janela mesmo se algo virar headless
$env:WATCH_MODE = "true"         # acompanhamento humano
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# --------------------------- Fluxo de comunicacao ---------------------------

$env:USE_CAIXA_CORREIO = "true"
$env:REQUEST_SIGNATURE = "true"
$env:ASSINANTE_NOME = "Roseli Chaves"
$env:TRAMITAR_DESTINO = "Em assinatura"

# --------------------------- Selecao de modelo + tokens @@ ------------------

# OFICIO_TEMPLATE_MODE=auto: ignora OFICIO_TEMPLATE fixo, usa classificacao
# por tipo (utap/dilacao/reiteracao/juizo) x secretaria (educacao/saude/geral).
$env:OFICIO_TEMPLATE_MODE = "auto"

# OFICIO_PRESERVE_AT_TOKENS=true: bloqueia upload se algum @@ do modelo sumir
# do DOCX gerado (validacao via docx_utils.assert_at_tokens_preserved).
$env:OFICIO_PRESERVE_AT_TOKENS = "true"

# Nao reaproveitar oficio SSG existente — vamos gerar nova minuta.
$env:REUSE_EXISTING_OFICIO = "false"

# --------------------------- Limpeza de minutas antigas ---------------------

# AUTORIZADO PELO OPERADOR em 2026-05-13: os oficios anteriores foram
# considerados incorretos e devem ser derrubados antes de recriar com a
# logica nova (preservacao @@). Em PROD a flag SAFE_DELETE_OWN_DRAFTS=true
# atua como autorizacao explicita (conforme regra do brief original).
#
# O cleanup eh feito por processo via:
#   - delete_ato_oficio_ssg_qualquer_estado (estorna assinatura -> estorna
#     conclusao -> excluir; NUNCA toca em status 'Assinado').
#   - delete_comunicacao_processual (exclui na Caixa de Correio).
$env:ENVIRONMENT = "producao"
$env:SAFE_DELETE_OWN_DRAFTS = "true"

# --------------------------- Limite de lote ---------------------------------

$env:MAX_PROCESSOS = "0"   # sem limite (vai processar todos os 5)

# --------------------------- Diretorios -------------------------------------

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs"))   { New-Item -ItemType Directory -Path "logs"   | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

# --------------------------- Pre-validacao DOCX (offline) -------------------

Write-Host ""
Write-Host "[1/2] Rodando pytest para garantir preservacao @@ antes de tocar no e-TCM..." -ForegroundColor Cyan
$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado em .venv\Scripts\python.exe. Rode python -m venv .venv e instale dependencias." -ForegroundColor Red
    exit 1
}
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] pytest falhou. Abortando execucao no e-TCM para evitar regressao." -ForegroundColor Red
    exit $LASTEXITCODE
}

# --------------------------- Execucao do fluxo ------------------------------

Write-Host ""
Write-Host "[2/2] Executando fluxo APO-PEN nos 5 processos..." -ForegroundColor Cyan
Write-Host ("Processos: " + $env:PROCESSOS_LIST) -ForegroundColor Cyan
Write-Host ""

# Usa Start-Process com stdout/stderr separados em arquivos. Evita o bug do
# PowerShell 5.1 onde linhas em stderr de native exes viram NativeCommandError
# (que aborta a pipeline mesmo quando o python terminou com sucesso).
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile = "logs\main_stdout_$ts.log"
$errFile = "logs\main_stderr_$ts.log"
Write-Host "stdout: $outFile"
Write-Host "stderr: $errFile"
$proc = Start-Process -FilePath $pythonExe -ArgumentList ".\src\main.py" `
    -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput $outFile `
    -RedirectStandardError $errFile
$rc = $proc.ExitCode

Write-Host ""
Write-Host "--- stdout (ultimas 40 linhas) ---" -ForegroundColor Cyan
if (Test-Path $outFile) { Get-Content $outFile -Tail 40 }
Write-Host ""
Write-Host "--- stderr (ultimas 20 linhas) ---" -ForegroundColor Cyan
if (Test-Path $errFile) { Get-Content $errFile -Tail 20 }

Write-Host ""
if ($rc -eq 0) {
    Write-Host "[OK] Fluxo executado." -ForegroundColor Green
} else {
    Write-Host "[FALHA] Fluxo retornou codigo $rc." -ForegroundColor Red
}
exit $rc
