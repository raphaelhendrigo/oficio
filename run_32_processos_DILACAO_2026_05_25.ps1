$ErrorActionPreference = "Stop"

# ============================================================================
# Script 2026-05-25 (32 processos DILACAO - 1 SAUDE + 31 EDUCACAO)
# Origem: planilha Processos_24_05_2026_DILACAO_SME_60.pdf (coluna "N° Processo").
#
# Fluxo completo PROD destrutivo (cleanup recria comm + Oficio SSG dilacao) +
# comm + Oficio SSG (template DILACAO por secretaria, auto) + concluido +
# assinatura Roseli Chaves, SEM tramitar. Data fixa no Oficio: 25/05/2026.
# Prazo da comunicacao: 60 dias.
#
# Especificidades de DILACAO (igual ao run_dilacao_TC_006237_2023.ps1):
#   - FORCE_TIPO=DILACAO  -> modelo de DILACAO + descricao "dilação".
#     Setado APOS o pytest para nao vazar override para os testes de classificacao.
#   - Referencia (cabecalho) e nº do Oficio SSG (corpo) sao AUTO-EXTRAIDOS por
#     processo (1a linha "Ofício nº..." da peca REQUERIMENTO + nº da peca apos
#     o ultimo MANUTAP-OF). NAO setamos OFICIO_REFERENCIA_TEXT/OFICIO_SSG_REF
#     globalmente, senao todos os 32 receberiam o MESMO valor.
#   - OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "false" (Encaminha de dilacao
#     fica "s/n").
#
# Orquestrador com timestamp por linha + RETRY automatico (ate 3 tentativas)
# + relatorio por processo.
#
# AGENDADO para 00:10 de 25/05/2026 (segunda) via Windows Task Scheduler (headful):
#   exige PC ligado, logado em TCM\20386 e SEM suspensao.
#   Login no e-TCM e AUTOMATICO via credenciais (ETCM_USERNAME/ETCM_PASSWORD).
# ============================================================================

# --- Lotes (a secretaria esperada serve so para conferencia no relatorio) ---
$saude = @(
    "TC/000524/2022"
)
$educacao = @(
    "TC/020122/2024","TC/008626/2023",
    "TC/000209/2022","TC/000381/2022","TC/000557/2022","TC/001037/2022","TC/001242/2022","TC/010634/2021",
    "TC/008009/2023","TC/008221/2023","TC/008331/2023","TC/008414/2023","TC/008895/2023","TC/006268/2023",
    "TC/006204/2024","TC/008918/2022","TC/008293/2023","TC/006298/2023","TC/006330/2023","TC/008278/2023",
    "TC/005642/2023","TC/005458/2023","TC/005390/2023","TC/005305/2023","TC/011771/2022","TC/009075/2023",
    "TC/009014/2023","TC/008199/2023","TC/008342/2023","TC/008516/2023","TC/001090/2024"
)
$all = @($saude + $educacao)

$expectedSec = @{}
foreach ($p in $saude)    { $expectedSec[$p] = "Saúde" }
foreach ($p in $educacao) { $expectedSec[$p] = "Educação" }

$maxAttempts = 3

# --- Credenciais via variaveis de ambiente de USUARIO ---
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

$env:ENVIRONMENT = "producao"
$env:HEADLESS = "false"
$env:SHOW_BROWSER = "true"
$env:WATCH_MODE = "true"
$env:SLOWMO_MS = "200"
$env:LOGIN_MANUAL_WAIT_MS = "60000"
$env:PAUSE_AFTER_LOGIN_MS = "5000"

# LIGADOS: cleanup completo + criar comm + anexar oficio
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

# Comunicacao Processual: prazo 60 dias (confirmado pelo operador)
$env:COMUNICACAO_PRAZO_DIAS = "60"
$env:COMUNICACAO_REFERENCIA = "gerado automaticamente"
$env:STATUS_ENTREGA = "Normal"

# Oficio - dilacao: NAO setar FORCE_TIPO aqui (afetaria os testes de classificacao
# do pytest). FORCE_TIPO=DILACAO e setado APOS o pytest, abaixo.
$env:OFICIO_TEMPLATE_MODE = "auto"
$env:OFICIO_PRESERVE_AT_TOKENS = "true"
$env:OFICIO_ADD_EUCLIDES_MARKER = "true"
$env:OFICIO_EUCLIDES_MARKER = "/euclides"
$env:OFICIO_REQUIRE_PIECE_NUMBER_IN_ENCAMINHA = "false"
$env:OFICIO_ENCAMINHA_TEXT_BOLD = "false"
$env:DATA_OFICIO = "25/05/2026"
$env:OFICIO_DATA = "25/05/2026"

# IMPORTANTE: NAO setar OFICIO_REFERENCIA_TEXT nem OFICIO_SSG_REF aqui.
# Em batch, esses valores devem ser AUTO-EXTRAIDOS por processo (peca
# REQUERIMENTO + peca apos MANUTAP-OF). Setar globalmente faria os 32
# processos receberem a mesma Referencia/Oficio SSG -> bug grave.

$env:MAX_PROCESSOS = "0"

# Autoriza cleanup destrutivo para TODOS os 32
$env:ONLY_PROCESSOS_AUTHORIZED = ($all -join ",")

# Garante diretorio de trabalho correto mesmo quando lancado pelo Task Scheduler
Set-Location -Path $PSScriptRoot

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "[1/2] Rodando pytest oficial (com FORCE_TIPO ainda VAZIO)..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] pytest falhou. Abortando." -ForegroundColor Red
    exit $LASTEXITCODE
}

# Override de DILACAO APOS o pytest (mesmo padrao do run_dilacao_TC_006237_2023.ps1)
$env:FORCE_TIPO = "DILACAO"

# ---------------------------------------------------------------------------
# Funcao: roda main.py para a lista informada, carimbando timestamp por linha.
# ---------------------------------------------------------------------------
function Invoke-Lote {
    param([string[]]$Lista, [string]$LogPath)

    $env:PROCESSOS_LIST = ($Lista -join ",")
    $sw = [System.IO.StreamWriter]::new($LogPath, $false, [System.Text.UTF8Encoding]::new($false))
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
}

# ---------------------------------------------------------------------------
# Funcao: parseia o log de uma tentativa -> info por processo.
# ---------------------------------------------------------------------------
function Parse-LoteLog {
    param([string]$LogPath, [string[]]$Lista)

    $result = @{}
    foreach ($p in $Lista) {
        $result[$p] = @{ Success=$false; Secretaria=""; Template=""; Start=$null; End=$null }
    }
    if (-not (Test-Path $LogPath)) { return $result }

    $current = $null
    foreach ($raw in Get-Content -Path $LogPath -Encoding utf8) {
        $tab = $raw.IndexOf("`t")
        if ($tab -lt 0) { continue }
        $tsStr = $raw.Substring(0, $tab)
        $text  = $raw.Substring($tab + 1)
        $ts = $null
        try { $ts = [datetime]::Parse($tsStr, $null, [System.Globalization.DateTimeStyles]::RoundtripKind) } catch { $ts = $null }

        if ($text -match "Iniciando pipeline do processo (TC/\d+/\d+)") {
            $current = $Matches[1]
            if ($result.ContainsKey($current) -and $ts) {
                if (-not $result[$current].Start) { $result[$current].Start = $ts }
                $result[$current].End = $ts
            }
            continue
        }
        if ($null -ne $current -and $result.ContainsKey($current)) {
            if ($ts) { $result[$current].End = $ts }
            if ($text -match "Modelo selecionado:\s*tipo=([^,]+),\s*secretaria=([^,]+),\s*arquivo=(.+)$") {
                $result[$current].Secretaria = $Matches[2].Trim()
                $result[$current].Template   = $Matches[3].Trim()
            }
        }
        if ($text -match "Assinatura solicitada para .* no processo (TC/\d+/\d+)") {
            $sp = $Matches[1]
            if ($result.ContainsKey($sp)) {
                $result[$sp].Success = $true
                if ($ts) { $result[$sp].End = $ts }
            }
        }
    }
    return $result
}

# ---------------------------------------------------------------------------
# Loop de tentativas com retry automatico
# ---------------------------------------------------------------------------
$runTs = Get-Date -Format 'yyyyMMdd_HHmmss'
$pending = [System.Collections.Generic.List[string]]::new()
$all | ForEach-Object { $pending.Add($_) }

$final = @{}
$all | ForEach-Object { $final[$_] = @{ Success=$false; Secretaria=""; Template=""; DurationSec=0; Attempt=0 } }

$attemptWall = @()

for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    if ($pending.Count -eq 0) { break }
    $lista = @($pending.ToArray())
    Write-Host ""
    Write-Host "==================================================================" -ForegroundColor Yellow
    Write-Host ("[TENTATIVA $attempt/$maxAttempts] {0} processo(s): {1}" -f $lista.Count, ($lista -join ", ")) -ForegroundColor Yellow
    Write-Host "==================================================================" -ForegroundColor Yellow

    $logPath = "logs\run_32_DILACAO_2026_05_25_${runTs}_attempt${attempt}.log"
    $swA = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-Lote -Lista $lista -LogPath $logPath
    $swA.Stop()
    $attemptWall += $swA.Elapsed.TotalSeconds

    $parsed = Parse-LoteLog -LogPath $logPath -Lista $lista
    $okThisAttempt = @()
    foreach ($p in $lista) {
        $info = $parsed[$p]
        if ($info.Start -and $info.End) { $dur = ($info.End - $info.Start).TotalSeconds } else { $dur = 0 }
        $final[$p].Secretaria  = $info.Secretaria
        $final[$p].Template    = $info.Template
        $final[$p].Attempt     = $attempt
        if ($dur -gt 0) { $final[$p].DurationSec = [math]::Round($dur, 1) }
        if ($info.Success) {
            $final[$p].Success = $true
            $okThisAttempt += $p
        }
    }
    foreach ($p in $okThisAttempt) { [void]$pending.Remove($p) }

    Write-Host ""
    Write-Host ("[TENTATIVA $attempt] sucesso: {0} | ainda pendentes: {1}" -f $okThisAttempt.Count, $pending.Count) -ForegroundColor Cyan
}

$totalSeconds = ($attemptWall | Measure-Object -Sum).Sum

function Fmt-Dur { param([double]$s)
    $t = [TimeSpan]::FromSeconds($s)
    return ('{0:00}:{1:00}:{2:00}' -f [int]$t.TotalHours, $t.Minutes, $t.Seconds)
}

$okCount = (($all | Where-Object { $final[$_].Success }) | Measure-Object).Count
$failCount = $all.Count - $okCount
$avgPerProc = if ($all.Count -gt 0) { $totalSeconds / $all.Count } else { 0 }

$lines = @()
$lines += "RELATORIO - 32 PROCESSOS DILACAO (SAUDE + EDUCACAO) - $runTs"
$lines += "Data do Oficio: 25/05/2026 | Prazo: 60 dias | Assinante: Roseli Chaves | SEM tramitar | PROD destrutivo"
$lines += "Tipo forcado: DILACAO (FORCE_TIPO=DILACAO apos pytest)"
$lines += ""
$lines += ("Tempo total (soma do relogio das {0} tentativa(s)): {1} ({2:N0}s)" -f $attemptWall.Count, (Fmt-Dur $totalSeconds), $totalSeconds)
$lines += ("Tempo medio por processo (total / {0}): {1} ({2:N1}s)" -f $all.Count, (Fmt-Dur $avgPerProc), $avgPerProc)
$lines += ("Sucesso: $okCount/$($all.Count) | Falha: $failCount")
$lines += ""
$lines += ("{0,-18} {1,-8} {2,-10} {3,-10} {4,-8} {5}" -f "PROCESSO","STATUS","ESPERADA","DETECTADA","TEMPO","MODELO")
$lines += ("-" * 110)
foreach ($p in $all) {
    $i = $final[$p]
    $status = if ($i.Success) { "OK" } else { "FALHA" }
    $exp = $expectedSec[$p]
    $det = if ($i.Secretaria) { $i.Secretaria } else { "?" }
    $flag = ""
    if ($i.Secretaria -and ($i.Secretaria -notlike "*$exp*") -and ($exp -notlike "*$($i.Secretaria)*")) { $flag = " <-- DIVERGENCIA" }
    $lines += ("{0,-18} {1,-8} {2,-10} {3,-10} {4,-8} {5}{6}" -f $p, $status, $exp, $det, (Fmt-Dur $i.DurationSec), $i.Template, $flag)
}
$lines += ""
if ($failCount -gt 0) {
    $lines += "PROCESSOS COM FALHA (apos $maxAttempts tentativas):"
    foreach ($p in ($all | Where-Object { -not $final[$_].Success })) { $lines += "  - $p" }
}

$reportPath = "logs\RELATORIO_32_DILACAO_2026_05_25_${runTs}.txt"
$lines | Set-Content -Path $reportPath -Encoding utf8

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Green
$lines | ForEach-Object { Write-Host $_ }
Write-Host "==================================================================" -ForegroundColor Green
Write-Host "Relatorio salvo em: $reportPath" -ForegroundColor Green

if ($failCount -gt 0) { exit 1 } else { exit 0 }
