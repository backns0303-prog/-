$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$htmlPath = Join-Path $root "dashboard_presentation.html"
$pdfPath = Join-Path $root "dashboard_presentation.pdf"
$edgeCandidates = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)

$edgePath = $edgeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edgePath) {
    throw "Microsoft Edge executable not found."
}

if (-not (Test-Path $htmlPath)) {
    throw "HTML presentation not found: $htmlPath"
}

$htmlUri = [System.Uri]::new($htmlPath).AbsoluteUri

& $edgePath `
    --headless `
    --disable-gpu `
    --print-to-pdf="$pdfPath" `
    --print-to-pdf-no-header `
    $htmlUri | Out-Null

if (-not (Test-Path $pdfPath)) {
    throw "PDF export failed: $pdfPath"
}

Write-Output "PDF export complete: $pdfPath"
