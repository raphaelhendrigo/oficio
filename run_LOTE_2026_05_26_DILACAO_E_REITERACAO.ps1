$ErrorActionPreference = "Continue"

# Wrapper agendado 2026-05-26 00:10:
# Encadeia os dois batches em sequencia, cada um em um processo PowerShell
# isolado (powershell.exe -NoProfile -File). Necessario porque dilacao e
# reiteracao precisam de FORCE_TIPO diferente -- rodar como subprocessos
# garante que env vars (especialmente FORCE_TIPO) nao vazem entre os dois.
#
#  1) BATCH 1: run_4_DILACAO_GILSON_2026_05_25.ps1   (4 dilacao Educacao)
#  2) BATCH 2: run_9_REITERACAO_2026_05_25.ps1       (1 Saude + 8 Educacao)
#
# Total ~55 min no caminho feliz. Cada batch roda pytest no inicio.
#
# IMPORTANTE: este arquivo eh 100% ASCII. PowerShell 5.1 ler arquivos sem
# BOM usa o codepage do sistema (Windows-1252 em pt-BR). Caracteres
# multi-byte UTF-8 (acentos e em-dash) viram bytes que podem incluir aspas
# direitas (U+201D), o que quebra o lexer de strings PowerShell. Na primeira
# versao desse wrapper havia em-dash dentro de uma chamada Log-Both, e a
# tarefa agendada de 2026-05-26 00:10 saiu com 0x80070001 em 0.7s sem
# escrever nenhum log. Diagnostico via [Parser]::ParseFile mostrou 7 erros
# de parse. Daqui em diante, NADA de acentos / em-dash neste arquivo.

Set-Location -Path $PSScriptRoot

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$wrapperLog = "logs\run_LOTE_2026_05_26_WRAPPER_$ts.log"

if (-not (Test-Path "logs")) { New-Item -ItemType Directory -Path "logs" | Out-Null }

function Log-Both {
    param([string]$Msg)
    $line = "{0}`t{1}" -f (Get-Date).ToString("o"), $Msg
    Write-Host $line
    Add-Content -Path $wrapperLog -Value $line -Encoding utf8
}

Log-Both "=========================================================="
Log-Both "WRAPPER 2026-05-26 - INICIO (Dilacao 4 + Reiteracao 9 = 13 procs)"
Log-Both "=========================================================="

# --- BATCH 1: 4 DILACAO ---
Log-Both "[1/2] Lancando batch DILACAO (4 procs Educacao)..."
$swA = [System.Diagnostics.Stopwatch]::StartNew()
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\run_4_DILACAO_GILSON_2026_05_25.ps1"
$rc1 = $LASTEXITCODE
$swA.Stop()
Log-Both ("[1/2] DILACAO terminou: exit={0} | duracao={1}" -f $rc1, $swA.Elapsed)

# --- BATCH 2: 9 REITERACAO ---
Log-Both "[2/2] Lancando batch REITERACAO (1 Saude + 8 Educacao)..."
$swB = [System.Diagnostics.Stopwatch]::StartNew()
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\run_9_REITERACAO_2026_05_25.ps1"
$rc2 = $LASTEXITCODE
$swB.Stop()
Log-Both ("[2/2] REITERACAO terminou: exit={0} | duracao={1}" -f $rc2, $swB.Elapsed)

Log-Both "=========================================================="
Log-Both ("RESUMO FINAL: DILACAO exit={0} | REITERACAO exit={1}" -f $rc1, $rc2)
Log-Both ("Wrapper log: {0}" -f $wrapperLog)
Log-Both "=========================================================="

if ($rc1 -ne 0 -or $rc2 -ne 0) { exit 1 } else { exit 0 }
