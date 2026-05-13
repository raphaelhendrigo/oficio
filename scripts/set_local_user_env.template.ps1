# scripts/set_local_user_env.template.ps1
#
# Template seguro para configurar credenciais do e-TCM como variaveis de
# ambiente do USUARIO no Windows (escopo "User"). Persiste entre sessoes
# e reboots, ficando fora do repositorio.
#
# Como usar:
#   1) Copie este arquivo para scripts/set_local_user_env.ps1
#      (o .gitignore ja bloqueia o arquivo sem o sufixo .template)
#   2) Execute no PowerShell:
#        powershell -ExecutionPolicy Bypass -File scripts\set_local_user_env.ps1
#   3) Reabra o PowerShell/VS Code antes de rodar o robo.
#
# REGRAS DE SEGURANCA (NAO QUEBRE):
#   - Nunca coloque a senha literal neste arquivo nem em nenhum outro do repo.
#   - Nunca commite scripts/set_local_user_env.ps1 (sem .template).
#   - Nunca imprima a senha no console, log ou screenshot.
#   - Se desconfiar que a senha vazou, troque-a imediatamente no e-TCM.

# --- Usuario -----------------------------------------------------------------
$Username = "20386"
[Environment]::SetEnvironmentVariable("ETCM_USERNAME", $Username, "User")

# --- Senha (lida mascarada) --------------------------------------------------
$secure = Read-Host "Senha do e-TCM" -AsSecureString
if ($null -eq $secure) {
    Write-Host "Senha nao pode ser vazia. Abortando." -ForegroundColor Red
    exit 1
}

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringUni($bstr)
    if ([string]::IsNullOrWhiteSpace($plain)) {
        Write-Host "Senha nao pode ser vazia. Abortando." -ForegroundColor Red
        exit 1
    }

    # Nomes oficiais lidos pelo robo
    [Environment]::SetEnvironmentVariable("ETCM_USERNAME", $Username, "User")
    [Environment]::SetEnvironmentVariable("ETCM_PASSWORD", $plain,    "User")

    # Aliases legados (config.py ja resolve qualquer um deles)
    [Environment]::SetEnvironmentVariable("ETCM_USER",  $Username, "User")
    [Environment]::SetEnvironmentVariable("ETCM_LOGIN", $Username, "User")
    [Environment]::SetEnvironmentVariable("ETCM_PASS",  $plain,    "User")
    [Environment]::SetEnvironmentVariable("ETCM_SENHA", $plain,    "User")

    Write-Host ""
    Write-Host "[OK] ETCM_USERNAME configurado: $Username" -ForegroundColor Green
    Write-Host "[OK] ETCM_PASSWORD configurado." -ForegroundColor Green
    Write-Host "[OK] Aliases ETCM_USER, ETCM_LOGIN, ETCM_PASS, ETCM_SENHA configurados." -ForegroundColor Green
    Write-Host ""
    Write-Host "IMPORTANTE: reabra o PowerShell / VS Code antes de rodar o robo." -ForegroundColor Yellow
}
finally {
    if ($null -ne $bstr -and $bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    # Limpa a variavel local da memoria do processo.
    if (Get-Variable -Name plain -Scope Local -ErrorAction SilentlyContinue) {
        Remove-Variable -Name plain -Scope Local -ErrorAction SilentlyContinue
    }
}
