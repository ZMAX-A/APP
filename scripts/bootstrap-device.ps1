$ErrorActionPreference = 'Stop'

if (-not $env:ANDROID_HOME) {
    throw 'ANDROID_HOME is not set.'
}

$adbExecutable = Join-Path $env:ANDROID_HOME 'platform-tools\adb.exe'
if (-not (Test-Path -LiteralPath $adbExecutable)) {
    throw "adb.exe was not found under ANDROID_HOME: $adbExecutable"
}

$deviceLines = & $adbExecutable devices -l
$readyDevices = @($deviceLines | Where-Object { $_ -match '\sdevice(\s|$)' })
if ($readyDevices.Count -ne 1) {
    $deviceLines | ForEach-Object { Write-Host $_ }
    throw "Expected exactly one authorized Android device, found $($readyDevices.Count)."
}

Write-Host 'Checking io.appium.settings...'
$settingsPackagePath = & $adbExecutable shell pm path io.appium.settings
if (-not ($settingsPackagePath -match '^package:')) {
    $defaultAppiumHome = Join-Path $env:USERPROFILE '.appium'
    $settingsApk = Get-ChildItem -LiteralPath $defaultAppiumHome -Recurse -Filter 'settings_apk-debug.apk' |
        Select-Object -First 1 -ExpandProperty FullName
    if (-not $settingsApk) {
        throw 'settings_apk-debug.apk was not found. Install the UiAutomator2 driver first.'
    }

    Write-Host 'Installing the official Appium Settings helper...'
    Write-Host 'Approve the installation prompt on the tablet if it appears.'
    & $adbExecutable install -r -g $settingsApk
    if ($LASTEXITCODE -ne 0) {
        throw 'Appium Settings installation failed. On Xiaomi/HyperOS enable USB installation and USB debugging security settings.'
    }
}

# Android 16 requires location runtime permission before the helper may start
# its location-type foreground service. Xiaomi does not always grant these via
# `adb install -g`, so grant them explicitly and idempotently.
Write-Host 'Granting Appium Settings foreground-location prerequisites...'
& $adbExecutable shell pm grant io.appium.settings android.permission.ACCESS_COARSE_LOCATION
& $adbExecutable shell pm grant io.appium.settings android.permission.ACCESS_FINE_LOCATION
& $adbExecutable shell appops set io.appium.settings COARSE_LOCATION allow
& $adbExecutable shell appops set io.appium.settings FINE_LOCATION allow

& $adbExecutable shell am force-stop io.appium.settings
& $adbExecutable shell am start `
    -n io.appium.settings/.Settings `
    -a android.intent.action.MAIN `
    -c android.intent.category.LAUNCHER | Out-Null
Start-Sleep -Seconds 2

$serviceState = & $adbExecutable shell dumpsys activity services io.appium.settings
if (-not ($serviceState -match 'isForeground=true')) {
    throw 'Appium Settings did not start as a foreground service. Inspect AndroidRuntime logcat before running tests.'
}

Write-Host 'Appium Settings foreground service is ready.'
Write-Host 'The first Appium session may now install the UiAutomator2 server helper packages.'
