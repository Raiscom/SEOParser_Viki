param(
    [string]$ProjectName = "SEOParser_Viki"
)

$pythonPath = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$mainPath = Join-Path $PSScriptRoot "main.py"
$hookPath = Join-Path $PSScriptRoot "pyinstaller-hooks"

if (-not (Test-Path $pythonPath)) {
    Write-Error "Python not found: $pythonPath. Create .venv and install dependencies first."
    exit 1
}

$distPath = Join-Path $PSScriptRoot "dist"
$dataPath = Join-Path $PSScriptRoot "data"
$envExamplePath = Join-Path $PSScriptRoot ".env.example"

if (-not (Test-Path $dataPath)) {
    Write-Error "Data directory not found: $dataPath"
    exit 1
}

if (-not (Test-Path $mainPath)) {
    Write-Error "Main script not found: $mainPath"
    exit 1
}

if (-not (Test-Path $hookPath)) {
    Write-Error "PyInstaller hook directory not found: $hookPath"
    exit 1
}

$tkInfoLines = & $pythonPath -c @'
import sys
from pathlib import Path

import _tkinter

print(Path(sys.base_prefix).resolve())
print(Path(_tkinter.__file__).resolve())
print(_tkinter.TCL_VERSION)
print(_tkinter.TK_VERSION)
'@

if ($LASTEXITCODE -ne 0) {
    Write-Error "Unable to inspect Tcl/Tk from: $pythonPath"
    exit $LASTEXITCODE
}

$tkInfoLines = @($tkInfoLines)
if ($tkInfoLines.Count -lt 4) {
    Write-Error "Unexpected Tcl/Tk inspection output from: $pythonPath"
    exit 1
}

$runtimePythonPath = [string]$tkInfoLines[0]
$tkinterBinary = [string]$tkInfoLines[1]
$tclVersion = [string]$tkInfoLines[2]
$tkVersion = [string]$tkInfoLines[3]

$tclDataPath = Join-Path $runtimePythonPath ("tcl\tcl{0}" -f $tclVersion)
$tkDataPath = Join-Path $runtimePythonPath ("tcl\tk{0}" -f $tkVersion)
$tclModulePath = Join-Path $runtimePythonPath ("tcl\tcl{0}" -f ($tclVersion -replace '\..*$', ''))
$tclBinary = Join-Path $runtimePythonPath ("DLLs\tcl{0}t.dll" -f ($tclVersion -replace '\.', ''))
$tkBinary = Join-Path $runtimePythonPath ("DLLs\tk{0}t.dll" -f ($tkVersion -replace '\.', ''))

foreach ($requiredPath in @($tclDataPath, $tkDataPath, $tclModulePath, $tkinterBinary, $tclBinary, $tkBinary)) {
    if (-not (Test-Path $requiredPath)) {
        Write-Error "Required Tk resource not found: $requiredPath"
        exit 1
    }
}

& $pythonPath -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name $ProjectName `
    --additional-hooks-dir $hookPath `
    --hidden-import tkinter `
    --hidden-import tkinter.ttk `
    --hidden-import _tkinter `
    --add-data "$tclDataPath;_tcl_data" `
    --add-data "$tkDataPath;_tk_data" `
    --add-data "$tclModulePath;tcl8" `
    --add-binary "$tkinterBinary;." `
    --add-binary "$tclBinary;." `
    --add-binary "$tkBinary;." `
    $mainPath

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$bundlePath = Join-Path $distPath $ProjectName
$bundleDataPath = Join-Path $bundlePath "data"

if (Test-Path $bundleDataPath) {
    Remove-Item -LiteralPath $bundleDataPath -Recurse -Force
}

Copy-Item $dataPath $bundleDataPath -Recurse -Force

if (Test-Path $envExamplePath) {
    Copy-Item $envExamplePath (Join-Path $bundlePath ".env.example") -Force
}

Write-Host "Portable build created in: $bundlePath"
