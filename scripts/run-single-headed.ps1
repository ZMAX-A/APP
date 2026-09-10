param(
    [string]$CaseId,
    [switch]$NoWriteBack,
    [switch]$OpenReport,
    [switch]$DryRun
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
$runner = Join-Path $PSScriptRoot 'run-excel.ps1'
$profileScript = Join-Path $PSScriptRoot 'get-case-profile.py'
$workbook = Join-Path $projectRoot 'test_case.xlsx'

if (-not $CaseId) {
    $CaseId = Read-Host '请输入单条用例 ID，例如 TC-HOME-001'
}
$CaseId = $CaseId.Trim()
if (-not $CaseId) {
    throw '用例 ID 不能为空。'
}
if ($CaseId.Contains(',') -or $CaseId.Contains('*') -or $CaseId.Contains('?')) {
    throw '该入口只允许运行一个精确用例 ID，不接受逗号或通配符。'
}
if (-not (Test-Path -LiteralPath $python)) {
    throw "Project virtual environment was not found: $python"
}

$profileOutput = @(
    & $python $profileScript --workbook $workbook --case-id $CaseId 2>&1
)
if ($LASTEXITCODE -ne 0) {
    throw ($profileOutput -join [Environment]::NewLine)
}
$profile = ($profileOutput -join [Environment]::NewLine) | ConvertFrom-Json

Write-Host ''
Write-Host ("用例：{0}" -f $profile.title) -ForegroundColor Cyan
Write-Host ("标签：{0}" -f ($profile.tags -join ', '))
Write-Host ("Excel 回写：{0}" -f (-not $NoWriteBack))

$runnerParameters = @{
    CaseId = [string]$profile.caseId
    RunSeeded = $true
    SkipExcelValidation = $true
}
$runnerArguments = @('-CaseId', $profile.caseId, '-RunSeeded', '-SkipExcelValidation')
$needsPreflight = $profile.dedicatedCustomer -and (
    $profile.mutating -or $profile.destructive
)

if ($profile.mutating -or $profile.destructive) {
    $runnerParameters.AllowMutation = $true
    $runnerArguments += '-AllowMutation'
}
if ($profile.destructive) {
    $runnerParameters.AllowDestructive = $true
    $runnerArguments += '-AllowDestructive'
}
if ($needsPreflight) {
    $runnerParameters.CustomerPreflightVerified = $true
    $runnerArguments += '-CustomerPreflightVerified'
}
if ($NoWriteBack) {
    $runnerParameters.NoWriteBack = $true
    $runnerArguments += '-NoWriteBack'
}
if ($OpenReport) {
    $runnerParameters.OpenReport = $true
    $runnerArguments += '-OpenReport'
}

if ($DryRun) {
    Write-Output ("preflight_required={0}" -f $needsPreflight)
    Write-Output ("planned=run-excel.ps1 " + ($runnerArguments -join ' '))
    exit 0
}

if (-not $NoWriteBack) {
    try {
        $workbookProbe = [System.IO.File]::Open(
            $workbook,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $workbookProbe.Close()
    }
    catch {
        throw 'test_case.xlsx 正在被 Excel/WPS 占用。请关闭表格后重新运行。'
    }
}

if ($profile.persistent) {
    Write-Host '该用例会持久化新增业务数据，测试后不会自动删除。' -ForegroundColor Yellow
}
elseif ($profile.destructive) {
    Write-Host '该用例包含删除动作，只允许删除本轮创建的可恢复数据。' -ForegroundColor Yellow
}
if ($needsPreflight) {
    Write-Host ''
    Write-Host '正在执行专用客户只读唯一性预检...' -ForegroundColor Cyan
    $preflightParameters = @{
        CaseId = 'TC-HOME-007'
        RunSeeded = $true
        NoWriteBack = $true
        SkipExcelValidation = $true
    }
    & $runner @preflightParameters
    if ($LASTEXITCODE -ne 0) {
        throw "唯一性预检失败（exit $LASTEXITCODE），目标用例未执行。"
    }
}

Write-Host ''
Write-Host '正在运行目标用例...' -ForegroundColor Cyan
& $runner @runnerParameters
$exitCode = $LASTEXITCODE

Write-Host ''
if ($exitCode -eq 0) {
    Write-Host '单条用例运行完成。' -ForegroundColor Green
}
else {
    Write-Host ("单条用例运行失败，exit code: {0}" -f $exitCode) -ForegroundColor Red
}
exit $exitCode
