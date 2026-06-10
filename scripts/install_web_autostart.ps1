$ErrorActionPreference = "Stop"

# ============================================================================
# Projeto Euclides - registra a interface web como Task Scheduler 24/7.
# Roda no boot + a cada 5 min checa se esta viva (auto-restart).
# Precisa ser executado como ADMINISTRADOR uma unica vez.
# ============================================================================

$root      = Split-Path -Parent $PSScriptRoot
$taskName  = 'ProjetoEuclidesWeb'
$psExe     = (Get-Command powershell.exe).Source
$startCmd  = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$root\scripts\start_web_hidden.ps1`""

# 1. Action principal: roda o script de start em janela escondida.
$action = New-ScheduledTaskAction -Execute $psExe -Argument $startCmd -WorkingDirectory $root

# 2. Triggers: ao ligar a VM + repeticao a cada 5 min para auto-heal.
$trigBoot = New-ScheduledTaskTrigger -AtStartup
$trigRepeat = New-ScheduledTaskTrigger -Once -At ([DateTime]::Now.AddMinutes(1)) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration ([TimeSpan]::FromDays(365 * 10))

# 3. Settings: nao bloqueia se ja rodando, restart se falhar.
$settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartInterval (New-TimeSpan -Minutes 2) `
    -RestartCount 5

# 4. Principal: roda como o usuario atual, com privilegios maximos, mesmo
#    sem login interativo (necessario para iniciar no boot antes do logon).
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest -LogonType S4U

# 5. Cria / sobrescreve a task.
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "[INFO] Removendo task existente '$taskName'..."
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

Register-ScheduledTask -TaskName $taskName `
    -Action $action `
    -Trigger @($trigBoot, $trigRepeat) `
    -Settings $settings `
    -Principal $principal `
    -Description 'Projeto Euclides - interface web FastAPI (boot + heartbeat 5min)' | Out-Null

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host " Task '$taskName' registrada." -ForegroundColor Green
Write-Host " - Inicia no boot da VM" -ForegroundColor Green
Write-Host " - Auto-restart a cada 5min se cair (start_web_hidden.ps1)" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Iniciar agora? Run: Start-ScheduledTask -TaskName $taskName" -ForegroundColor Yellow
Write-Host "Para remover:    Unregister-ScheduledTask -TaskName $taskName -Confirm:`$false" -ForegroundColor Yellow
