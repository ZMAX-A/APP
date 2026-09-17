param(
    [string]$ExcelFile = "test_case.xlsx",
    [string]$ExcelOutput,
    [string]$ReportsRoot,
    [string]$RunId,
    [string]$CaseId,
    [string]$Tags,
    [switch]$RunSeeded,
    [switch]$AllowMutation,
    [switch]$AllowDestructive,
    [switch]$CustomerPreflightVerified,
    [switch]$NoWriteBack,
    [switch]$ReadOnlyRetry,
    [switch]$OpenReport,
    [switch]$SkipExcelValidation,
    [switch]$VerboseProgress,
    [switch]$SkipPreflight,
    [switch]$NoAutoStartAppium,
    [switch]$KeepAppium,
    [switch]$IgnoreDotEnv,
    [string]$AppiumProcessIdFile,
    [ValidateRange(5, 300)]
    [int]$AppiumStartupTimeoutSeconds = 60
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUNBUFFERED = '1'
$env:ALLURE_NO_ANALYTICS = '1'
if ($IgnoreDotEnv) {
    $env:YANJIA_IGNORE_DOTENV = 'true'
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$effectiveReportsRoot = if ($ReportsRoot) {
    if ([System.IO.Path]::IsPathRooted($ReportsRoot)) {
        [System.IO.Path]::GetFullPath($ReportsRoot)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path $projectRoot $ReportsRoot))
    }
}
else {
    Join-Path $projectRoot 'reports'
}
if ($RunId -and $RunId -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$') {
    throw 'RunId 只能包含字母、数字、点、下划线和连字符，且最长 128 个字符。'
}
$effectiveRunId = if ($RunId) { $RunId } else { Get-Date -Format 'yyyyMMdd_HHmmss' }
$runLock = $null
$appiumProcess = $null
$pushedLocation = $false
$pytestExitCode = 2
$appiumProcessIdPath = $null

function Resolve-ProjectPath([string]$Value) {
    if ([System.IO.Path]::IsPathRooted($Value)) {
        return [System.IO.Path]::GetFullPath($Value)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $projectRoot $Value))
}

function Get-ProjectSetting([string]$Name, [string]$DefaultValue) {
    $environmentValue = [Environment]::GetEnvironmentVariable($Name)
    if ($environmentValue) {
        return $environmentValue.Trim()
    }
    $dotenv = Join-Path $projectRoot '.env'
    if (-not $IgnoreDotEnv -and (Test-Path -LiteralPath $dotenv)) {
        $line = Get-Content -LiteralPath $dotenv -Encoding UTF8 |
            Where-Object { $_ -match ("^\s*" + [regex]::Escape($Name) + "\s*=") } |
            Select-Object -First 1
        if ($line) {
            $value = ($line -split '=', 2)[1].Trim().Trim('"').Trim("'")
            if ($value) {
                return $value
            }
        }
    }
    return $DefaultValue
}

function Test-AppiumReady([string]$ServerUrl) {
    try {
        $status = Invoke-RestMethod -Uri ($ServerUrl.TrimEnd('/') + '/status') -TimeoutSec 2
        return ($null -ne $status.value -and $status.value.ready -eq $true)
    }
    catch {
        return $false
    }
}

function Stop-ExactProcess([int]$ProcessId) {
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    Wait-Process -Id $ProcessId -Timeout 5 -ErrorAction SilentlyContinue
}

function Mask-DeviceId([string]$DeviceId) {
    if ($DeviceId.Length -le 4) {
        return $DeviceId
    }
    return ('***' + $DeviceId.Substring($DeviceId.Length - 4))
}

if (-not (Test-Path -LiteralPath $python)) {
    throw "Project virtual environment was not found: $python"
}
if ($IgnoreDotEnv) {
    # TOP runs must match the dependencies covered by the signed Release SBOM.
    # This check deliberately precedes workbook writes, ADB and Appium access.
    & $python -I (Join-Path $PSScriptRoot 'runtime_dependencies.py')
    if ($LASTEXITCODE -ne 0) {
        throw 'TOP runtime dependency verification failed. Install the signed release runtime before running.'
    }
}
$allureCommand = Get-Command allure -ErrorAction SilentlyContinue
if ($OpenReport -and -not $allureCommand) {
    throw 'OpenReport requires the Allure command line. Install allure-commandline before running.'
}

$excelPath = Resolve-ProjectPath $ExcelFile
$appiumProcessIdPath = if ($AppiumProcessIdFile) {
    Resolve-ProjectPath $AppiumProcessIdFile
}
else {
    $null
}
if (-not (Test-Path -LiteralPath $excelPath)) {
    throw "Excel test workbook was not found: $excelPath"
}

if (-not $NoWriteBack -and -not $ExcelOutput) {
    try {
        $lockProbe = [System.IO.File]::Open(
            $excelPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $lockProbe.Close()
    }
    catch {
        throw "Excel用例文件正在被WPS/Excel占用，请关闭文件后再运行：$excelPath"
    }
}

New-Item -ItemType Directory -Path $effectiveReportsRoot -Force | Out-Null
$lockPath = Join-Path $effectiveReportsRoot '.excel-run.lock'
try {
    $runLock = [System.IO.File]::Open(
        $lockPath,
        [System.IO.FileMode]::OpenOrCreate,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
}
catch {
    throw '已有 Excel 自动化任务正在运行。请等待其结束后再试。'
}

try {
    if (-not $SkipExcelValidation) {
        Write-Host '正在离线校验 Excel 用例...'
        $validationArguments = @($excelPath)
        if ($CaseId) {
            $validationArguments += @('--case-id', $CaseId)
        }
        & $python (Join-Path $PSScriptRoot 'validate_excel.py') @validationArguments
        if ($LASTEXITCODE -ne 0) {
            throw "Excel validation failed with exit code $LASTEXITCODE."
        }
    }

    $appiumServerUrl = Get-ProjectSetting 'APPIUM_SERVER_URL' 'http://127.0.0.1:4723'
    $configuredAppiumStartupTimeout = Get-ProjectSetting `
        'APPIUM_STARTUP_TIMEOUT_SECONDS' `
        ([string]$AppiumStartupTimeoutSeconds)
    $effectiveAppiumStartupTimeout = 0
    if (
        -not [int]::TryParse(
            $configuredAppiumStartupTimeout,
            [ref]$effectiveAppiumStartupTimeout
        ) -or
        $effectiveAppiumStartupTimeout -lt 5 -or
        $effectiveAppiumStartupTimeout -gt 300
    ) {
        throw 'APPIUM_STARTUP_TIMEOUT_SECONDS 必须是 5~300 之间的整数。'
    }
    $appPackage = Get-ProjectSetting 'YANJIA_APP_PACKAGE' 'com.xiaofutech.yanjia_ai'
    $configuredUdid = Get-ProjectSetting 'ANDROID_UDID' ''
    $deviceId = $configuredUdid
    $deviceModel = 'unknown'
    $androidVersion = 'unknown'
    $androidSdk = 'unknown'
    $appVersion = 'unknown'
    $appVersionCode = 'unknown'

    if (-not $SkipPreflight) {
        $adb = $null
        if ($env:ANDROID_HOME) {
            $candidate = Join-Path $env:ANDROID_HOME 'platform-tools\adb.exe'
            if (Test-Path -LiteralPath $candidate) {
                $adb = $candidate
            }
        }
        if (-not $adb) {
            $adbCommand = Get-Command adb -ErrorAction SilentlyContinue
            if ($adbCommand) {
                $adb = $adbCommand.Source
            }
        }
        if (-not $adb) {
            throw '找不到 adb。请设置 ANDROID_HOME，或将 platform-tools 加入 PATH。'
        }

        $deviceLines = @(& $adb devices -l)
        if ($LASTEXITCODE -ne 0) {
            throw 'adb devices 执行失败。'
        }
        $readyDevices = @(
            $deviceLines |
                Where-Object { $_ -match '^\S+\s+device(?:\s|$)' } |
                ForEach-Object { ($_ -split '\s+')[0] }
        )
        if ($configuredUdid) {
            $readyDevices = @($readyDevices | Where-Object { $_ -eq $configuredUdid })
        }
        if ($readyDevices.Count -ne 1) {
            throw "必须且只能有一台已授权目标设备，当前找到 $($readyDevices.Count) 台。请检查 adb devices -l。"
        }
        $deviceId = $readyDevices[0]
        $env:ANDROID_UDID = $deviceId

        $packagePath = & $adb -s $deviceId shell pm path $appPackage
        if ($LASTEXITCODE -ne 0 -or -not ($packagePath -match '^package:')) {
            throw "目标应用未安装：$appPackage"
        }
        $deviceModel = ((& $adb -s $deviceId shell getprop ro.product.model) -join '').Trim()
        $androidVersion = ((& $adb -s $deviceId shell getprop ro.build.version.release) -join '').Trim()
        $androidSdk = ((& $adb -s $deviceId shell getprop ro.build.version.sdk) -join '').Trim()
        $packageDump = (& $adb -s $deviceId shell dumpsys package $appPackage) -join [Environment]::NewLine
        $versionNameMatch = [regex]::Match($packageDump, '(?m)^\s*versionName=([^\r\n]+)')
        $versionCodeMatch = [regex]::Match($packageDump, '(?m)^\s*versionCode=(\d+)')
        if ($versionNameMatch.Success) {
            $appVersion = $versionNameMatch.Groups[1].Value.Trim()
        }
        if ($versionCodeMatch.Success) {
            $appVersionCode = $versionCodeMatch.Groups[1].Value
        }

        Write-Host "设备预检通过：$deviceModel / Android $androidVersion (SDK $androidSdk)"
        Write-Host "应用版本：$appVersion ($appVersionCode)"
    }

    if (-not (Test-AppiumReady $appiumServerUrl)) {
        if ($NoAutoStartAppium) {
            throw "Appium 未就绪：$appiumServerUrl"
        }
        $serverUri = [Uri]$appiumServerUrl
        if ($serverUri.Host -notin @('127.0.0.1', 'localhost', '::1')) {
            throw "远程 Appium 未就绪，无法自动启动：$appiumServerUrl"
        }
        $appiumLogDir = Join-Path $effectiveReportsRoot 'appium'
        New-Item -ItemType Directory -Path $appiumLogDir -Force | Out-Null
        $stdoutLog = Join-Path $appiumLogDir "$effectiveRunId.stdout.log"
        $stderrLog = Join-Path $appiumLogDir "$effectiveRunId.stderr.log"
        $appiumCommand = Get-Command appium -ErrorAction Stop
        $appiumInstallRoot = Split-Path -Parent $appiumCommand.Source
        $appiumEntry = Join-Path $appiumInstallRoot 'node_modules\appium\index.js'
        if (-not (Test-Path -LiteralPath $appiumEntry)) {
            throw "找不到 Appium Node 入口：$appiumEntry"
        }
        $nodeExecutable = Join-Path $appiumInstallRoot 'node.exe'
        if (-not (Test-Path -LiteralPath $nodeExecutable)) {
            $nodeExecutable = (Get-Command node -ErrorAction Stop).Source
        }
        $bindAddress = if ($serverUri.Host -eq 'localhost') {
            '127.0.0.1'
        }
        else {
            $serverUri.Host
        }
        $startAppium = @{
            FilePath = $nodeExecutable
            ArgumentList = @(
                ('"{0}"' -f $appiumEntry),
                '--address', $bindAddress,
                '--port', $serverUri.Port,
                '--base-path', '/',
                '--log-level', 'warn'
            )
            RedirectStandardOutput = $stdoutLog
            RedirectStandardError = $stderrLog
            WindowStyle = 'Hidden'
            PassThru = $true
        }
        $appiumProcess = Start-Process @startAppium
        if ($appiumProcessIdPath) {
            $appiumProcessIdDirectory = Split-Path -Parent $appiumProcessIdPath
            New-Item -ItemType Directory -Path $appiumProcessIdDirectory -Force | Out-Null
            [System.IO.File]::WriteAllText(
                $appiumProcessIdPath,
                [string]$appiumProcess.Id,
                $utf8
            )
        }
        Write-Host "正在自动启动 Appium（PID $($appiumProcess.Id)）..."
        $deadline = (Get-Date).AddSeconds($effectiveAppiumStartupTimeout)
        while ((Get-Date) -lt $deadline -and -not (Test-AppiumReady $appiumServerUrl)) {
            if ($appiumProcess.HasExited) {
                throw "Appium 启动进程意外退出，请查看：$stderrLog"
            }
            Start-Sleep -Milliseconds 500
        }
        if (-not (Test-AppiumReady $appiumServerUrl)) {
            throw (
                "Appium 在${effectiveAppiumStartupTimeout}秒内未就绪，" +
                "请查看：$stdoutLog 和 $stderrLog"
            )
        }
        Write-Host 'Appium 已就绪。'
    }
    else {
        Write-Host '复用已运行的 Appium 服务。'
    }

    $allureResults = Join-Path $effectiveReportsRoot "allure-results\$effectiveRunId"
    $allureReportRoot = Join-Path $effectiveReportsRoot 'allure-report'
    $allureReport = Join-Path $allureReportRoot $effectiveRunId
    New-Item -ItemType Directory -Path $allureResults -Force | Out-Null
    New-Item -ItemType Directory -Path $allureReportRoot -Force | Out-Null

    $previousReport = Get-ChildItem -LiteralPath $allureReportRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne $effectiveRunId -and (Test-Path -LiteralPath (Join-Path $_.FullName 'history')) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($previousReport) {
        Copy-Item -LiteralPath (Join-Path $previousReport.FullName 'history') -Destination (Join-Path $allureResults 'history') -Recurse -Force
    }

    $categoriesPath = Join-Path $projectRoot 'config\allure-categories.json'
    if (Test-Path -LiteralPath $categoriesPath) {
        Copy-Item -LiteralPath $categoriesPath -Destination (Join-Path $allureResults 'categories.json')
    }

    $environment = @(
        'Project=YanJia AI Android Automation'
        "RunId=$effectiveRunId"
        "ExcelWorkbook=$excelPath"
        'Platform=Android'
        'Framework=Appium + pytest + Excel'
        "Device=$deviceModel"
        "DeviceId=$(Mask-DeviceId $deviceId)"
        "Android=$androidVersion"
        "AndroidSDK=$androidSdk"
        "AppPackage=$appPackage"
        "AppVersion=$appVersion"
        "AppVersionCode=$appVersionCode"
    )
    [System.IO.File]::WriteAllLines(
        (Join-Path $allureResults 'environment.properties'),
        $environment,
        $utf8
    )

    $pytestArguments = @(
        '-m', 'excel_driven',
        '--excel-file', $excelPath,
        '--excel-run-id', $effectiveRunId,
        '--allure-report-dir', $allureReport,
        '--alluredir', $allureResults
    )
    if ($VerboseProgress) {
        $pytestArguments += '-v'
    }
    if ($ExcelOutput) {
        $pytestArguments += @('--excel-output', (Resolve-ProjectPath $ExcelOutput))
    }
    if ($CaseId) {
        $pytestArguments += @('--excel-case-id', $CaseId)
    }
    if ($Tags) {
        $pytestArguments += @('--excel-tags', $Tags)
    }
    if ($RunSeeded) {
        $pytestArguments += '--run-seeded'
    }
    if ($AllowMutation) {
        $pytestArguments += '--allow-mutation'
    }
    if ($AllowDestructive) {
        $pytestArguments += '--allow-destructive'
    }
    if ($CustomerPreflightVerified) {
        $pytestArguments += '--customer-preflight-verified'
    }
    if ($NoWriteBack) {
        $pytestArguments += '--no-excel-writeback'
    }
    if ($ReadOnlyRetry) {
        $pytestArguments += '--readonly-retry'
    }

    Write-Host "Run ID: $effectiveRunId"
    Write-Host "Excel:  $excelPath"
    Write-Host "Allure results: $allureResults"
    if ($VerboseProgress) {
        Write-Host '正在启动 pytest；下面将逐条显示用例名称、结果和百分比...' -ForegroundColor Cyan
    }

    Push-Location $projectRoot
    $pushedLocation = $true
    & $python -m pytest @pytestArguments
    $pytestExitCode = $LASTEXITCODE

    $resultFiles = @(
        Get-ChildItem -LiteralPath $allureResults -Filter '*-result.json' -ErrorAction SilentlyContinue
    )
    if ($resultFiles.Count -gt 0 -and $allureCommand) {
        & $allureCommand.Source generate $allureResults --output $allureReport --clean
        if ($LASTEXITCODE -ne 0) {
            throw "Allure report generation failed with exit code $LASTEXITCODE."
        }
        Write-Host "Allure report generated: $allureReport"
        Write-Host "View later: .\scripts\open-allure.ps1 -RunId $effectiveRunId"
        if ($OpenReport) {
            Write-Host 'Allure report server is starting. Press Ctrl+C to stop it.'
            & $allureCommand.Source open $allureReport
        }
    }
    elseif ($resultFiles.Count -gt 0) {
        Write-Warning "Allure CLI 未安装；测试已完成，原始结果保留在：$allureResults"
    }
    else {
        Write-Warning 'No Allure result JSON files were generated.'
    }
    if ($resultFiles.Count -gt 0) {
        [System.IO.File]::WriteAllText(
            (Join-Path $allureReportRoot 'latest-run.txt'),
            $effectiveRunId,
            $utf8
        )
    }
}
finally {
    if ($pushedLocation) {
        Pop-Location
    }
    if ($appiumProcess -and -not $KeepAppium) {
        Write-Host "正在关闭本次自动启动的 Appium（PID $($appiumProcess.Id)）..."
        Stop-ExactProcess -ProcessId $appiumProcess.Id
        if ($appiumProcessIdPath) {
            Remove-Item -LiteralPath $appiumProcessIdPath -Force -ErrorAction SilentlyContinue
        }
    }
    if ($runLock) {
        $runLock.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}

exit $pytestExitCode
