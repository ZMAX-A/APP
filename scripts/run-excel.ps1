param(
    [string]$ExcelFile = "test_case.xlsx",
    [string]$ExcelOutput,
    [string]$CaseId,
    [string]$Tags,
    [switch]$RunSeeded,
    [switch]$AllowMutation,
    [switch]$AllowDestructive,
    [switch]$NoWriteBack,
    [switch]$OpenReport,
    [switch]$SkipPreflight,
    [switch]$NoAutoStartAppium,
    [switch]$KeepAppium
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:ALLURE_NO_ANALYTICS = '1'

$projectRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$reportsRoot = Join-Path $projectRoot 'reports'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$runLock = $null
$appiumProcess = $null
$pushedLocation = $false
$pytestExitCode = 2

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
    if (Test-Path -LiteralPath $dotenv) {
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
if (-not (Get-Command allure -ErrorAction SilentlyContinue)) {
    throw 'Allure command line was not found. Install allure-commandline before running.'
}

$excelPath = Resolve-ProjectPath $ExcelFile
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

New-Item -ItemType Directory -Path $reportsRoot -Force | Out-Null
$lockPath = Join-Path $reportsRoot '.excel-run.lock'
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
    Write-Host '正在离线校验 Excel 用例...'
    $validationArguments = @($excelPath)
    if ($CaseId) {
        $validationArguments += @('--case-id', $CaseId)
    }
    & $python (Join-Path $PSScriptRoot 'validate_excel.py') @validationArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Excel validation failed with exit code $LASTEXITCODE."
    }

    $appiumServerUrl = Get-ProjectSetting 'APPIUM_SERVER_URL' 'http://127.0.0.1:4723'
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
        $appiumLogDir = Join-Path $reportsRoot 'appium'
        New-Item -ItemType Directory -Path $appiumLogDir -Force | Out-Null
        $stdoutLog = Join-Path $appiumLogDir "$runId.stdout.log"
        $stderrLog = Join-Path $appiumLogDir "$runId.stderr.log"
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
        Write-Host "正在自动启动 Appium（PID $($appiumProcess.Id)）..."
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and -not (Test-AppiumReady $appiumServerUrl)) {
            if ($appiumProcess.HasExited) {
                throw "Appium 启动进程意外退出，请查看：$stderrLog"
            }
            Start-Sleep -Milliseconds 500
        }
        if (-not (Test-AppiumReady $appiumServerUrl)) {
            throw "Appium 在30秒内未就绪，请查看：$stderrLog"
        }
        Write-Host 'Appium 已就绪。'
    }
    else {
        Write-Host '复用已运行的 Appium 服务。'
    }

    $allureResults = Join-Path $reportsRoot "allure-results\$runId"
    $allureReportRoot = Join-Path $reportsRoot 'allure-report'
    $allureReport = Join-Path $allureReportRoot $runId
    New-Item -ItemType Directory -Path $allureResults -Force | Out-Null

    $previousReport = Get-ChildItem -LiteralPath $allureReportRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne $runId -and (Test-Path -LiteralPath (Join-Path $_.FullName 'history')) } |
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
        "RunId=$runId"
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
        '--excel-run-id', $runId,
        '--allure-report-dir', $allureReport,
        '--alluredir', $allureResults
    )
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
    if ($NoWriteBack) {
        $pytestArguments += '--no-excel-writeback'
    }

    Write-Host "Run ID: $runId"
    Write-Host "Excel:  $excelPath"
    Write-Host "Allure results: $allureResults"

    Push-Location $projectRoot
    $pushedLocation = $true
    & $python -m pytest @pytestArguments
    $pytestExitCode = $LASTEXITCODE

    $resultFiles = @(
        Get-ChildItem -LiteralPath $allureResults -Filter '*-result.json' -ErrorAction SilentlyContinue
    )
    if ($resultFiles.Count -gt 0) {
        & allure generate $allureResults --output $allureReport --clean
        if ($LASTEXITCODE -ne 0) {
            throw "Allure report generation failed with exit code $LASTEXITCODE."
        }
        [System.IO.File]::WriteAllText(
            (Join-Path $allureReportRoot 'latest-run.txt'),
            $runId,
            $utf8
        )
        Write-Host "Allure report generated: $allureReport"
        Write-Host "View later: .\scripts\open-allure.ps1 -RunId $runId"
        if ($OpenReport) {
            Write-Host 'Allure report server is starting. Press Ctrl+C to stop it.'
            & allure open $allureReport
        }
    }
    else {
        Write-Warning 'No Allure result JSON files were generated.'
    }
}
finally {
    if ($pushedLocation) {
        Pop-Location
    }
    if ($appiumProcess -and -not $KeepAppium) {
        Write-Host "正在关闭本次自动启动的 Appium（PID $($appiumProcess.Id)）..."
        Stop-ExactProcess -ProcessId $appiumProcess.Id
    }
    if ($runLock) {
        $runLock.Dispose()
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
    }
}

exit $pytestExitCode
