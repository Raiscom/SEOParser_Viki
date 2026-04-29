param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,
    [switch]$UpdateReferences
)

$ErrorActionPreference = "Stop"

$sourcePath = $PSScriptRoot
$targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
$sourceFullPath = [System.IO.Path]::GetFullPath($sourcePath)

if ($targetFullPath.TrimEnd('\') -eq $sourceFullPath.TrimEnd('\')) {
    Write-Error "TargetPath must point to the old portable folder, not to this new folder."
    exit 1
}

if (-not (Test-Path (Join-Path $sourceFullPath "SEOParser_Viki.exe"))) {
    Write-Error "This script must be run from a new SEOParser_Viki portable folder."
    exit 1
}

if (-not (Test-Path $targetFullPath)) {
    New-Item -ItemType Directory -Path $targetFullPath | Out-Null
}

$preserveNames = @(".env", "parser.log")
foreach ($item in Get-ChildItem -LiteralPath $sourceFullPath -Force) {
    if ($preserveNames -contains $item.Name) {
        continue
    }
    if ($item.Name -eq "data" -and -not $UpdateReferences -and (Test-Path (Join-Path $targetFullPath "data"))) {
        continue
    }
    Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $targetFullPath $item.Name) -Recurse -Force
}

if (-not (Test-Path (Join-Path $targetFullPath ".env")) -and (Test-Path (Join-Path $sourceFullPath ".env.example"))) {
    Copy-Item (Join-Path $sourceFullPath ".env.example") (Join-Path $targetFullPath ".env") -Force
}

Write-Host "Updated portable folder: $targetFullPath"
Write-Host "Preserved user files: .env, parser.log"
if (-not $UpdateReferences) {
    Write-Host "Preserved existing data folder. Use -UpdateReferences to overwrite bundled references."
}
