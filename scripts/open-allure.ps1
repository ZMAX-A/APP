param(
    [string]$RunId
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Command allure -ErrorAction SilentlyContinue)) {
    throw 'Allure command line was not found.'
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$reportRoot = Join-Path $projectRoot 'reports\allure-report'
if ($RunId) {
    $report = Join-Path $reportRoot $RunId
}
else {
    $latest = Get-ChildItem -LiteralPath $reportRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latest) {
        throw 'No generated Allure reports were found.'
    }
    $report = $latest.FullName
}

if (-not (Test-Path -LiteralPath (Join-Path $report 'index.html'))) {
    throw "Allure report is incomplete or does not exist: $report"
}

Write-Host "Opening Allure report: $report"
Write-Host 'Press Ctrl+C to stop the report server.'
& allure open $report
