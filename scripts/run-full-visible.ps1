param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUNBUFFERED = '1'

$projectRoot = Split-Path -Parent $PSScriptRoot
$runner = Join-Path $PSScriptRoot 'run-excel.ps1'
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$retrySelector = Join-Path $PSScriptRoot 'select_retry_cases.py'
$workbook = Join-Path $projectRoot 'test_case.xlsx'
$outputRoot = Join-Path $projectRoot 'outputs'
$logRoot = Join-Path $outputRoot 'launcher-logs'
$launcherLockPath = Join-Path $outputRoot '.full-visible-run.lock'
$latestRunPath = Join-Path $projectRoot 'reports\allure-report\latest-run.txt'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$appiumProcessIdPath = Join-Path $outputRoot "launcher-appium-$stamp.pid"
$logPath = Join-Path $logRoot "full-run-$stamp.log"
$statusPath = Join-Path $logRoot "full-run-$stamp.status.json"
$latestStatusPath = Join-Path $logRoot 'full-run-latest.status.json'
$launcherLock = $null
$transcriptStarted = $false
$preflightRunId = ''
$fullRunId = ''
$retryRunId = ''
$retryCaseIds = @()
$blockedRetryCaseIds = @()
$remainingFailedCaseIds = @()
$restorationFailed = $false
$firstRoundExitCode = -1
$retryExitCode = -1
$script:visibleRunnerExitCode = 2

function Write-Status(
    [string]$Phase,
    [int]$ExitCode,
    [string]$Message,
    [bool]$UpdateLatest = $true
) {
    $payload = [ordered]@{
        phase = $Phase
        updatedAt = (Get-Date).ToString('o')
        exitCode = $ExitCode
        message = $Message
        preflightRunId = $preflightRunId
        fullRunId = $fullRunId
        retryRunId = $retryRunId
        retryCaseIds = @($retryCaseIds)
        blockedRetryCaseIds = @($blockedRetryCaseIds)
        remainingFailedCaseIds = @($remainingFailedCaseIds)
        restorationFailed = $restorationFailed
        firstRoundExitCode = $firstRoundExitCode
        retryExitCode = $retryExitCode
        logPath = $logPath
        workbook = $workbook
    }
    $json = $payload | ConvertTo-Json
    [System.IO.File]::WriteAllText($statusPath, $json, $utf8)
    if ($UpdateLatest) {
        [System.IO.File]::WriteAllText($latestStatusPath, $json, $utf8)
    }
}

function Invoke-VisibleRunner(
    [string]$CaseId,
    [switch]$RunSeeded,
    [switch]$AllowMutation,
    [switch]$AllowDestructive,
    [switch]$CustomerPreflightVerified,
    [switch]$NoWriteBack,
    [switch]$ReadOnlyRetry,
    [switch]$SkipExcelValidation,
    [switch]$VerboseProgress,
    [switch]$KeepAppium,
    [string]$AppiumProcessIdFile
) {
    $runnerParameters = @{}
    $displayArguments = @()
    if ($CaseId) {
        $runnerParameters.CaseId = $CaseId
        $displayArguments += @('-CaseId', $CaseId)
    }
    foreach ($switchName in @(
        'RunSeeded',
        'AllowMutation',
        'AllowDestructive',
        'CustomerPreflightVerified',
        'NoWriteBack',
        'ReadOnlyRetry',
        'SkipExcelValidation',
        'VerboseProgress',
        'KeepAppium'
    )) {
        if ((Get-Variable -Name $switchName -ValueOnly).IsPresent) {
            $runnerParameters[$switchName] = $true
            $displayArguments += "-$switchName"
        }
    }
    if ($AppiumProcessIdFile) {
        $runnerParameters.AppiumProcessIdFile = $AppiumProcessIdFile
    }

    Write-Host ''
    Write-Host ("run-excel.ps1 " + ($displayArguments -join ' ')) -ForegroundColor DarkGray
    # Do not pipe or capture this call: Python detects a real console and flushes
    # each pytest progress line immediately. Store the exit code out-of-band so
    # pytest output cannot pollute the function return value.
    & $runner @runnerParameters
    $script:visibleRunnerExitCode = [int]$LASTEXITCODE
    Write-Host (
        "run-excel.ps1 exit code: {0}" -f $script:visibleRunnerExitCode
    ) -ForegroundColor DarkGray
}

function Read-LatestRunId {
    if (-not (Test-Path -LiteralPath $latestRunPath)) {
        return ''
    }
    return (Get-Content -Raw -LiteralPath $latestRunPath).Trim()
}

function Get-RetryPlan([string]$RunId, [string]$RetryRunId = '') {
    if (-not $RunId) {
        throw '首轮 Run ID 缺失，无法选择复跑用例。'
    }
    $resultsPath = Join-Path $projectRoot "reports\allure-results\$RunId"
    $selectorArguments = @('--results-dir', $resultsPath)
    if ($RetryRunId) {
        $selectorArguments += @(
            '--retry-results-dir',
            (Join-Path $projectRoot "reports\allure-results\$RetryRunId")
        )
    }
    $planOutput = @(& $python $retrySelector @selectorArguments)
    if ($LASTEXITCODE -ne 0) {
        throw 'Allure 证据校验失败，无法确认复跑范围或最终结果。'
    }
    return ($planOutput -join [Environment]::NewLine) | ConvertFrom-Json
}

function Stop-OwnedAppium {
    if (-not (Test-Path -LiteralPath $appiumProcessIdPath)) {
        return
    }
    try {
        $processId = 0
        $rawProcessId = (Get-Content -LiteralPath $appiumProcessIdPath -Raw).Trim()
        if ([int]::TryParse($rawProcessId, [ref]$processId)) {
            $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($process) {
                Write-Host "正在关闭全量启动器创建的 Appium（PID $processId）..."
                Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
                Wait-Process -Id $processId -Timeout 5 -ErrorAction SilentlyContinue
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $appiumProcessIdPath -Force -ErrorAction SilentlyContinue
    }
}

if ($DryRun) {
    Write-Output 'mode=visible;window=normal;writeback=True'
    Write-Output 'preflight=run-excel.ps1 -CaseId TC-HOME-007 -RunSeeded -NoWriteBack -SkipExcelValidation -VerboseProgress -KeepAppium'
    Write-Output 'first_round=run-excel.ps1 -RunSeeded -AllowMutation -AllowDestructive -CustomerPreflightVerified -SkipExcelValidation -VerboseProgress -KeepAppium'
    Write-Output 'retry=wait 30 seconds; rerun first-round failures once with the same authorizations; overwrite Excel with second-round results'
    exit 0
}

New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
try {
    $launcherLock = [System.IO.File]::Open(
        $launcherLockPath,
        [System.IO.FileMode]::OpenOrCreate,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
}
catch {
    $message = '已有全量启动器正在运行，请等待其结束。'
    Write-Host $message -ForegroundColor Red
    Write-Status 'launcher_busy' 2 $message $false
    exit 2
}

try {
    Start-Transcript -LiteralPath $logPath -Append | Out-Null
    $transcriptStarted = $true
    Write-Status 'starting' 0 'Visible full run is starting.'

    Write-Host 'YanJia Android 全量测试' -ForegroundColor Cyan
    Write-Host ("工作簿：{0}" -f $workbook)
    Write-Host ("日志：{0}" -f $logPath)

    if (-not (Test-Path -LiteralPath $runner)) {
        throw "Runner not found: $runner"
    }
    if (-not (Test-Path -LiteralPath $workbook)) {
        throw "Workbook not found: $workbook"
    }

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

    Write-Host ''
    Write-Host '阶段 1/3：专用客户只读唯一性预检' -ForegroundColor Cyan
    Write-Status 'preflight' 0 'Running dedicated-customer read-only preflight.'
    Invoke-VisibleRunner `
        -CaseId 'TC-HOME-007' `
        -RunSeeded `
        -NoWriteBack `
        -SkipExcelValidation `
        -VerboseProgress `
        -KeepAppium `
        -AppiumProcessIdFile $appiumProcessIdPath
    $preflightExit = $script:visibleRunnerExitCode
    $preflightRunId = Read-LatestRunId
    if ($preflightExit -ne 0) {
        $message = '唯一客户预检失败，全量测试未启动。'
        Write-Host $message -ForegroundColor Red
        Write-Status 'preflight_failed' $preflightExit $message
        exit $preflightExit
    }

    Write-Host ''
    Write-Host '阶段 2/3：84 条授权全量测试并回写 Excel' -ForegroundColor Cyan
    Write-Status 'full_run' 0 'Preflight passed; running the authorized full suite with Excel writeback.'
    $latestBeforeFull = Read-LatestRunId
    Invoke-VisibleRunner `
        -RunSeeded `
        -AllowMutation `
        -AllowDestructive `
        -CustomerPreflightVerified `
        -SkipExcelValidation `
        -VerboseProgress `
        -KeepAppium `
        -AppiumProcessIdFile $appiumProcessIdPath
    $fullExit = $script:visibleRunnerExitCode
    $firstRoundExitCode = $fullExit
    $fullRunCandidate = Read-LatestRunId
    if ($fullRunCandidate -and $fullRunCandidate -ne $latestBeforeFull) {
        $fullRunId = $fullRunCandidate
    }

    if ($fullExit -eq 1) {
        $retryPlan = Get-RetryPlan $fullRunId
        $retryCaseIds = @($retryPlan.retryCaseIds)
        $blockedRetryCaseIds = @($retryPlan.blockedRetryCaseIds)
        $remainingFailedCaseIds = @($retryPlan.remainingFailedCaseIds)
        $restorationFailed = [bool]$retryPlan.restorationFailed
        if ($blockedRetryCaseIds.Count -gt 0) {
            Write-Warning (
                "以下失败用例因写入、删除、持久化或 no_retry 标签未自动复跑：{0}" -f
                ($blockedRetryCaseIds -join ',')
            )
        }
        if ($retryCaseIds.Count -gt 0) {
            $retryList = $retryCaseIds -join ','
            Write-Host ''
            Write-Host (
                "首轮有 {0} 条可安全复跑的只读失败用例，30 秒后开始第二轮：{1}" -f
                $retryCaseIds.Count,
                $retryList
            ) -ForegroundColor Yellow
            Write-Status 'retry_wait' 0 'Waiting 30 seconds before retrying first-round failures.'
            for ($remaining = 30; $remaining -gt 0; $remaining--) {
                if ($remaining -eq 30 -or $remaining % 5 -eq 0 -or $remaining -le 3) {
                    Write-Host ("第二轮复跑倒计时：{0} 秒" -f $remaining) -ForegroundColor DarkGray
                }
                Start-Sleep -Seconds 1
            }

            Write-Host ''
            Write-Host '阶段 3/3：复跑首轮可重复只读失败用例，结果覆盖主表' -ForegroundColor Cyan
            Write-Status 'retry_run' 0 'Retrying repeatable read-only failures once; results overwrite Excel.'
            $latestBeforeRetry = Read-LatestRunId
            Invoke-VisibleRunner `
                -CaseId $retryList `
                -RunSeeded `
                -ReadOnlyRetry `
                -SkipExcelValidation `
                -VerboseProgress `
                -KeepAppium `
                -AppiumProcessIdFile $appiumProcessIdPath
            $retryExitCode = $script:visibleRunnerExitCode
            $retryRunCandidate = Read-LatestRunId
            if ($retryRunCandidate -and $retryRunCandidate -ne $latestBeforeRetry) {
                $retryRunId = $retryRunCandidate
            }
            if ($retryRunId) {
                $verifiedPlan = Get-RetryPlan $fullRunId $retryRunId
                $remainingFailedCaseIds = @($verifiedPlan.remainingFailedCaseIds)
            }
        }
        else {
            Write-Warning '没有符合自动复跑安全条件的首轮失败用例。'
        }
    }

    Write-Host ''
    $finalExit = if ($retryExitCode -ge 0) { $retryExitCode } else { $fullExit }
    if ($fullExit -ne 0 -and ($remainingFailedCaseIds.Count -gt 0 -or -not $retryRunId)) {
        if ($finalExit -eq 0) { $finalExit = 1 }
    }
    if ($fullExit -eq 0) {
        $message = '全量测试全部通过，结果已回写 Excel。'
        Write-Host $message -ForegroundColor Green
        Write-Status 'completed' 0 $message
    }
    elseif ($retryRunId -and $finalExit -eq 0) {
        $message = '第二轮复跑全部通过，已用第二轮结果覆盖 Excel 对应首轮结果。'
        Write-Host $message -ForegroundColor Green
        Write-Status 'completed_after_retry' 0 $message
    }
    else {
        $message = if ($retryExitCode -ge 0) {
            '第二轮复跑已结束，仍有未通过用例；Excel 对应用例以第二轮实际结果为准。'
        }
        else {
            '全量测试已结束，但未能启动有效复跑；现有结果已回写 Excel。'
        }
        Write-Host $message -ForegroundColor Yellow
        Write-Status 'completed_with_failures' $finalExit $message
    }
    if ($fullRunId) {
        Write-Host ("首轮 Run ID：{0}" -f $fullRunId)
        Write-Host ("首轮报告：reports\allure-report\{0}" -f $fullRunId)
    }
    if ($retryRunId) {
        Write-Host ("复跑 Run ID：{0}" -f $retryRunId)
        Write-Host ("复跑报告：reports\allure-report\{0}" -f $retryRunId)
    }
    exit $finalExit
}
catch {
    $message = $_.Exception.Message
    Write-Host ("启动器失败：{0}" -f $message) -ForegroundColor Red
    Write-Status 'launcher_failed' 2 $message
    exit 2
}
finally {
    Stop-OwnedAppium
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    if ($launcherLock) {
        $launcherLock.Dispose()
        Remove-Item -LiteralPath $launcherLockPath -Force -ErrorAction SilentlyContinue
    }
}
