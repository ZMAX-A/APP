$ErrorActionPreference = 'Stop'

Write-Host 'Starting Appium on http://127.0.0.1:4723 with warning-only logs.'
Write-Host 'Warning-only logging prevents credential values from appearing in HTTP request logs.'

appium `
    --address 127.0.0.1 `
    --port 4723 `
    --base-path / `
    --log-level warn
