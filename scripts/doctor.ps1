$ErrorActionPreference = 'Stop'

Write-Host 'Checking Appium and UiAutomator2...'
appium --version
appium driver list --installed
appium driver doctor uiautomator2

if (-not $env:ANDROID_HOME) {
    throw 'ANDROID_HOME is not set.'
}

$adbExecutable = Join-Path $env:ANDROID_HOME 'platform-tools\adb.exe'
if (-not (Test-Path -LiteralPath $adbExecutable)) {
    throw "adb.exe was not found under ANDROID_HOME: $adbExecutable"
}

Write-Host 'Checking attached Android devices...'
$deviceLines = & $adbExecutable devices -l
$deviceLines | ForEach-Object { Write-Host $_ }
$readyDevices = @($deviceLines | Where-Object { $_ -match '\sdevice(\s|$)' })
if ($readyDevices.Count -ne 1) {
    throw "Expected exactly one authorized device, found $($readyDevices.Count)."
}

Write-Host 'Checking the YanJia AI package...'
$packagePath = & $adbExecutable shell pm path com.xiaofutech.yanjia_ai
if (-not ($packagePath -match '^package:')) {
    throw 'com.xiaofutech.yanjia_ai is not installed on the connected device.'
}

& $adbExecutable shell dumpsys package com.xiaofutech.yanjia_ai |
    Select-String -Pattern 'versionName=|versionCode=|minSdk=|targetSdk='

Write-Host 'Environment check passed.'
