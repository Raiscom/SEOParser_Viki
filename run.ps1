param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AppArgs
)

$pythonPath = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$mainPath = Join-Path $PSScriptRoot "main.py"

if (-not (Test-Path $pythonPath)) {
    Write-Error "Python not found: $pythonPath"
    exit 1
}

if (-not (Test-Path $mainPath)) {
    Write-Error "Main script not found: $mainPath"
    exit 1
}

& $pythonPath $mainPath @AppArgs
exit $LASTEXITCODE
