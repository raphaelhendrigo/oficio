$ErrorActionPreference = "Stop"

# ============================================================================
# 2026-05-25 - 4 processos de DILACAO (lote do Gilson, todos Educacao)
# Origem: planilha Processos_25_05_2026.pdf, distribuidos para GILSON DE NOBREGA.
#
# Fluxo completo PROD destrutivo + comm + Oficio SSG (DILACAO Educacao, auto
# por secretaria) + concluido + assinatura Roseli Chaves, SEM tramitar.
# Data do Oficio: 25/05/2026. Prazo: 60 dias.
#
# FORCE_TIPO=DILACAO (apos pytest). Auto-extracao por processo via
# _extract_referencia_from_requerimento + _extract_ssg_ref_after_manutap.
# ============================================================================

$saude = @()
$educacao = @(
    "TC/001098/2024",  # Joao Antonio
    "TC/000383/2022",  # Roberto Braguim
    "TC/008629/2023",  # Roberto Braguim
    "TC/004898/2023"   # Roberto Braguim
)
$all = @($saude + $educacao)

$expectedSec = @{}
foreach ($p in $saude)    { $expectedSec[$p] = "Saúde" }
foreach ($p in $educacao) { $expectedSec[$p] = "Educação" }

$maxAttempts = 3

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
$env:DATA_OFICIO = "26/05/2026"
$env:OFICIO_DATA = "26/05/2026"

$env:MAX_PROCESSOS = "0"
$env:ONLY_PROCESSOS_AUTHORIZED = ($all -join ",")

Set-Location -Path $PSScriptRoot

if (-not (Test-Path "output")) { New-Item -ItemType Directory -Path "output" | Out-Null }
if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }
if (-not (Test-Path "artifacts\evidence")) { New-Item -ItemType Directory -Path "artifacts\evidence" -Force | Out-Null }

$pythonExe = Join-Path ".venv\Scripts" "python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERRO] .venv nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "[1/2] Rodando pytest oficial (FORCE_TIPO vazio)..." -ForegroundColor Cyan
& $pythonExe -m pytest tests -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERRO] pytest falhou. Abortando." -ForegroundColor Red
    exit $LASTEXITCODE
}

# Override DILACAO APOS o pytest
$env:FORCE_TIPO = "DILACAO"

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
    } finally { $sw.Close() }
}

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
    $logPath = "logs\run_4_DILACAO_GILSON_2026_05_25_${runTs}_attempt${attempt}.log"
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
$lines += "RELATORIO - 4 PROCESSOS DILACAO LOTE GILSON (EDUCACAO) - $runTs"
$lines += "Data Oficio: 26/05/2026 | Prazo: 60 dias | Assinante: Roseli Chaves | SEM tramitar | PROD destrutivo"
$lines += "Tipo forcado: DILACAO"
$lines += ""
$lines += ("Tempo total ({0} tent): {1} ({2:N0}s)" -f $attemptWall.Count, (Fmt-Dur $totalSeconds), $totalSeconds)
$lines += ("Tempo medio: {0} ({1:N1}s)" -f (Fmt-Dur $avgPerProc), $avgPerProc)
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
    $lines += "FALHAS apos $maxAttempts tentativas:"
    foreach ($p in ($all | Where-Object { -not $final[$_].Success })) { $lines += "  - $p" }
}

$reportPath = "logs\RELATORIO_4_DILACAO_GILSON_2026_05_25_${runTs}.txt"
$lines | Set-Content -Path $reportPath -Encoding utf8

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Green
$lines | ForEach-Object { Write-Host $_ }
Write-Host "==================================================================" -ForegroundColor Green
Write-Host "Relatorio: $reportPath" -ForegroundColor Green

if ($failCount -gt 0) { exit 1 } else { exit 0 }
