param()

$ErrorActionPreference = 'Stop'
$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dashboardVbs = Join-Path $baseDir 'run_dashboard_hidden.vbs'
$logPath = Join-Path $baseDir 'public_tunnel.log'
$errPath = Join-Path $baseDir 'public_tunnel.err.log'
$linkPath = Join-Path $baseDir 'public_link.txt'
$downloadUrl = 'https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/'

Add-Type -AssemblyName PresentationFramework

function Find-Cloudflared {
    $cmd = Get-Command cloudflared -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $wingetRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    if (Test-Path $wingetRoot) {
        $found = Get-ChildItem $wingetRoot -Recurse -Filter cloudflared.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }

    $localCopy = Join-Path $baseDir 'cloudflared.exe'
    if (Test-Path $localCopy) { return $localCopy }
    return $null
}

$cloudflaredPath = Find-Cloudflared
if (-not $cloudflaredPath) {
    Start-Process $downloadUrl
    [System.Windows.MessageBox]::Show("cloudflared가 설치되지 않아 다운로드 페이지를 열었습니다.`n설치 후 run_public_link_hidden.vbs를 다시 실행해 주세요.", '임시 공개 링크') | Out-Null
    exit 1
}

if (Test-Path $linkPath) { Remove-Item $linkPath -Force }
if (Test-Path $logPath) { Remove-Item $logPath -Force }
if (Test-Path $errPath) { Remove-Item $errPath -Force }

Start-Process wscript.exe -ArgumentList ('"' + $dashboardVbs + '"') -WindowStyle Hidden
Start-Sleep -Seconds 4

Start-Process -FilePath $cloudflaredPath -ArgumentList @('tunnel','--url','http://localhost:8501') -WindowStyle Hidden -RedirectStandardOutput $logPath -RedirectStandardError $errPath | Out-Null

$publicUrl = $null
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 1
    if (Test-Path $logPath) {
        $content = Get-Content -LiteralPath $logPath -Raw -ErrorAction SilentlyContinue
        if (Test-Path $errPath) {
            $content += "`n" + (Get-Content -LiteralPath $errPath -Raw -ErrorAction SilentlyContinue)
        }
        $match = [regex]::Match($content, 'https://[-a-zA-Z0-9]+\.trycloudflare\.com')
        if ($match.Success) {
            $publicUrl = $match.Value
            break
        }
    }
}

if (-not $publicUrl) {
    [System.Windows.MessageBox]::Show("임시 공개 링크를 찾지 못했습니다.`npublic_tunnel.log를 확인해 주세요.", '임시 공개 링크') | Out-Null
    if (Test-Path $logPath) { Start-Process notepad.exe $logPath }
    if (Test-Path $errPath) { Start-Process notepad.exe $errPath }
    exit 1
}

Set-Content -Path $linkPath -Value $publicUrl -Encoding utf8
Set-Clipboard -Value $publicUrl
[System.Windows.MessageBox]::Show("임시 공개 링크를 클립보드에 복사했습니다.`n`n$publicUrl", '임시 공개 링크') | Out-Null
Start-Process $publicUrl
