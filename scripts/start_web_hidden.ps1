$ErrorActionPreference = "SilentlyContinue"

# Heartbeat-friendly launcher do uvicorn:
# - se ja tem algo escutando na 8080, sai (Task Scheduler chama de 5 em 5 min)
# - senao, sobe o servidor em janela escondida e fica vigiando

$root = Split-Path -Parent $PSScriptRoot
Set-Location -Path $root

$logDir = Join-Path $root 'logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd_HHmm')
$logFile = Join-Path $logDir ("web_$stamp.log")

# Se ja esta escutando, nada a fazer.
$listening = Get-NetTCPConnection -State Listen -LocalPort 8080 -ErrorAction SilentlyContinue
if ($listening) { exit 0 }

# Promove credenciais do escopo User para o processo.
foreach ($name in @('ETCM_USERNAME','ETCM_PASSWORD','ETCM_USER','ETCM_LOGIN','ETCM_PASS','ETCM_SENHA')) {
    $val = [Environment]::GetEnvironmentVariable($name, 'User')
    if (-not [string]::IsNullOrEmpty($val)) { Set-Item -Path "Env:$name" -Value $val }
}

$env:PYTHONUNBUFFERED = '1'
$env:PYTHONIOENCODING = 'utf-8'

$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path $python)) { exit 1 }

# Cria janela CMD escondida com a saida indo para o log.
$cmdLine = "/c `"`"$python`" -m uvicorn web.app:app --host 0.0.0.0 --port 8080 --log-level info >> `"$logFile`" 2>&1`""
Start-Process -FilePath 'cmd.exe' -ArgumentList $cmdLine -WindowStyle Hidden -WorkingDirectory $root
