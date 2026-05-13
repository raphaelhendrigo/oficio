$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$folders = @("modelos_utap", "modelos_dilacao", "modelos_reiteracao", "modelos_juizo")

function Decode-ZipUnicodeEscapeName {
    param([Parameter(Mandatory = $true)][string]$Name)
    return [regex]::Replace($Name, '#U([0-9A-Fa-f]{4})', {
        param($m)
        [char]([Convert]::ToInt32($m.Groups[1].Value, 16))
    })
}

foreach ($folder in $folders) {
    $path = Join-Path $root $folder
    if (-not (Test-Path -LiteralPath $path)) { continue }
    Get-ChildItem -LiteralPath $path -File | ForEach-Object {
        $decoded = Decode-ZipUnicodeEscapeName -Name $_.Name
        if ($decoded -ne $_.Name) {
            $target = Join-Path $_.DirectoryName $decoded
            if (Test-Path -LiteralPath $target) {
                Write-Host "[AVISO] Destino ja existe, pulando: $target" -ForegroundColor Yellow
            }
            else {
                Rename-Item -LiteralPath $_.FullName -NewName $decoded
                Write-Host "[OK] $($_.Name) -> $decoded"
            }
        }
    }
}
