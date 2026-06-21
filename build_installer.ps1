param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"

Set-Location $ProjectRoot

if (!(Test-Path $Iscc)) {
    throw "Inno Setup 6 is not installed. Install it first, then run this script again: https://jrsoftware.org/isinfo.php"
}

if (!(Test-Path (Join-Path $ProjectRoot "dist\MarriageExpensesManager\MarriageExpensesManager.exe"))) {
    & (Join-Path $ProjectRoot "build_portable.ps1") -Version $Version
}

$InstallerVersion = $Version.TrimStart([char[]]"vV")
if ($InstallerVersion -notmatch '^\d+(\.\d+){0,3}$') {
    $InstallerVersion = "1.0.0"
}

& $Iscc "/DMyAppVersion=$InstallerVersion" (Join-Path $ProjectRoot "packaging\MarriageExpensesManager.iss")

Write-Host ""
Write-Host "Installer build is ready:"
Write-Host (Join-Path $ProjectRoot "dist\installer")
