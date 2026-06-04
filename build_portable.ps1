$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m pip install pyinstaller

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
