param(
    [switch]$RunSeeded,
    [switch]$AllowMutation,
    [switch]$AllowDestructive,
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$runner = Join-Path $PSScriptRoot 'run-excel.ps1'
$arguments = @('-Tags', 'smoke')

if ($RunSeeded) {
    $arguments += '-RunSeeded'
}
if ($AllowMutation) {
    $arguments += '-AllowMutation'
}
if ($AllowDestructive) {
    $arguments += '-AllowDestructive'
}
if ($OpenReport) {
    $arguments += '-OpenReport'
}

& $runner @arguments
exit $LASTEXITCODE
