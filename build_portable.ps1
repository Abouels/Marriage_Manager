param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$env:PYTHONUTF8 = "1"
$env:PIP_PROGRESS_BAR = "off"
py -m pip install --upgrade pip --progress-bar off
py -m pip install -r requirements-build.txt --progress-bar off

$SupportEndpointFile = Join-Path $ProjectRoot "support_endpoint.txt"
$SupportEndpoint = "$env:MARRIAGE_MANAGER_SUPPORT_URL".Trim()
if ($SupportEndpoint) {
    Set-Content -Path $SupportEndpointFile -Value $SupportEndpoint -Encoding UTF8
} elseif (Test-Path $SupportEndpointFile) {
    Remove-Item -LiteralPath $SupportEndpointFile -Force
}

$VersionFile = Join-Path $ProjectRoot "version.txt"
$ResolvedVersion = "$Version".Trim()
if (-not $ResolvedVersion) {
    $ResolvedVersion = "$env:MARRIAGE_MANAGER_VERSION".Trim()
}
if ($ResolvedVersion) {
    Set-Content -Path $VersionFile -Value $ResolvedVersion -Encoding UTF8
} elseif (Test-Path $VersionFile) {
    Remove-Item -LiteralPath $VersionFile -Force
}

$BuildDir = Join-Path $ProjectRoot "build"
$DistDir = Join-Path $ProjectRoot "dist"
foreach ($PathToClean in @($BuildDir, $DistDir)) {
    if (Test-Path $PathToClean) {
        attrib -R "$PathToClean\*" /S /D
        Remove-Item -LiteralPath $PathToClean -Recurse -Force
    }
}

py -m PyInstaller .\packaging\MarriageExpensesManager.spec --noconfirm --clean

$PortableDir = Join-Path $ProjectRoot "dist\MarriageExpensesManager"
New-Item -ItemType File -Path (Join-Path $PortableDir "portable.flag") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $PortableDir "app_data") -Force | Out-Null

Write-Host ""
Write-Host "Portable build is ready:"
Write-Host $PortableDir
