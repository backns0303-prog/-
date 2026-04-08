param(
    [Parameter(Mandatory = $true)]
    [string]$CloudflaredPath,

    [Parameter(Mandatory = $true)]
    [string]$LogPath
)

$ErrorActionPreference = "Stop"

if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

& $CloudflaredPath tunnel --url http://localhost:8501 *>&1 |
    Tee-Object -FilePath $LogPath -Append
