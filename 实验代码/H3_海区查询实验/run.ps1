param(
    [string]$PythonPath,
    [switch]$Install,
    [switch]$TestOnly
)
$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
if (-not $PythonPath) {
    $bundled = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundled) {
        $PythonPath = $bundled
    } else {
        $command = Get-Command python -ErrorAction SilentlyContinue
        if ($command -and $command.Source -notmatch 'WindowsApps') {
            $PythonPath = $command.Source
        } else {
            throw 'Python 3.11+ is required. Use -PythonPath to specify it.'
        }
    }
}
if ($Install) {
    $runtime = Join-Path $PSScriptRoot '.runtime'
    $env:TEMP = Join-Path $runtime 'tmp'
    $env:TMP = $env:TEMP
    $env:PIP_CACHE_DIR = Join-Path $runtime 'cache'
    New-Item -ItemType Directory -Force -Path $env:TEMP,$env:PIP_CACHE_DIR | Out-Null
    & $PythonPath -m pip install --upgrade --target (Join-Path $runtime 'packages') --only-binary=:all: --no-deps -r (Join-Path $PSScriptRoot 'requirements.txt') --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
if ($TestOnly) {
    & $PythonPath (Join-Path $PSScriptRoot 'test_experiment.py')
} else {
    & $PythonPath (Join-Path $PSScriptRoot 'run_experiment.py') --prepare-repeats 7 --query-repeats 31
}
exit $LASTEXITCODE
