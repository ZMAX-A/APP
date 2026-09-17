param(
    [string]$Python = 'python',
    [string]$EnvironmentPath = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = if ($EnvironmentPath) {
    [System.IO.Path]::GetFullPath($EnvironmentPath)
} else {
    Join-Path $projectRoot '.venv'
}
if (Test-Path -LiteralPath $runtimeRoot) {
    throw 'Runtime destination already exists. Choose a new release directory; existing environments are preserved.'
}
& $Python -I -c "import platform,sys; assert sys.platform == 'win32' and sys.version_info[:2] == (3,12) and sys.implementation.name == 'cpython' and platform.machine().lower() in ('amd64','x86_64'), 'Windows x64 CPython 3.12 is required'"
if ($LASTEXITCODE -ne 0) { throw 'Unsupported Python runtime.' }
& $Python -I -m venv $runtimeRoot
if ($LASTEXITCODE -ne 0) { throw 'Could not create an isolated runtime.' }
$runtimePython = Join-Path $runtimeRoot 'Scripts\python.exe'
& $runtimePython -I -m pip --isolated install --disable-pip-version-check --no-input `
    --no-cache-dir --only-binary=:all: --require-hashes --index-url https://pypi.org/simple `
    -r (Join-Path $projectRoot 'config\runtime-requirements.lock')
if ($LASTEXITCODE -ne 0) { throw 'Hash-locked runtime installation failed; do not start the Worker.' }
# Link only this release's signed src directory, without an editable install or
# build-time dependencies. ASCII escaping also supports non-ASCII Windows paths.
& $runtimePython -I -c 'import pathlib,sys; source=pathlib.Path(sys.argv[1],"src").resolve(); assert source.is_dir(); target=pathlib.Path(sys.argv[2],"Lib","site-packages","yanjia_runtime_source.pth"); target.write_text("import sys; sys.path.insert(0, " + ascii(str(source)) + ")\n",encoding="ascii")' $projectRoot $runtimeRoot
if ($LASTEXITCODE -ne 0) { throw 'Could not link the signed release source directory.' }
& $runtimePython -I (Join-Path $PSScriptRoot 'runtime_dependencies.py')
if ($LASTEXITCODE -ne 0) { throw 'Installed runtime did not match the signed dependency lock.' }
Write-Host 'Hash-locked Windows runtime installed and verified.'
