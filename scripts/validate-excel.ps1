param(
    [string]$Workbook = "test_case.xlsx"
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$projectRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
    throw "Project virtual environment was not found: $python"
}

$workbookPath = if ([System.IO.Path]::IsPathRooted($Workbook)) {
    [System.IO.Path]::GetFullPath($Workbook)
}
else {
    [System.IO.Path]::GetFullPath((Join-Path $projectRoot $Workbook))
}

& $python (Join-Path $PSScriptRoot 'validate_excel.py') $workbookPath
if ($LASTEXITCODE -ne 0) {
    throw "Excel validation failed with exit code $LASTEXITCODE."
}
