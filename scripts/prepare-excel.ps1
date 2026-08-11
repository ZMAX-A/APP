param(
    [string]$Workbook = "test_case.xlsx",
    [switch]$ReplaceCurrentSteps
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
    throw "Project virtual environment was not found: $python"
}

$arguments = @(
    (Join-Path $PSScriptRoot 'prepare_excel.py'),
    '--workbook',
    (Join-Path $projectRoot $Workbook)
)
if ($ReplaceCurrentSteps) {
    $arguments += '--replace-current-steps'
}

& $python @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Excel preparation failed with exit code $LASTEXITCODE."
}
