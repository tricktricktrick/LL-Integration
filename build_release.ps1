param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Command,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed with exit code $LASTEXITCODE"
    }
}

function Remove-IfExists {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }
}

Write-Host "Building native messaging helper..."
Remove-IfExists "dist\ll_integration_native.exe"
Invoke-Checked {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --console `
        --name "ll_integration_native" `
        "native-app\main.py"
} "Native helper build"

Copy-Item -LiteralPath "dist\ll_integration_native.exe" -Destination "native-app\ll_integration_native.exe" -Force

$commonInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--icon", "installer_icon.png",
    "--add-data", "installer_icon.png;.",
    "--add-data", "native-app;native-app",
    "--add-data", "mo2-plugin;mo2-plugin"
)

Write-Host "Building stable installer..."
Remove-IfExists "dist\LLIntegrationInstaller.exe"
Invoke-Checked {
    & $Python -m PyInstaller @commonInstallerArgs --name "LLIntegrationInstaller" "installer.py"
} "Stable installer build"

Write-Host "Building installer with experimental toolbar..."
Remove-IfExists "dist\LLIntegrationInstaller-WithToolbar.exe"
Invoke-Checked {
    & $Python -m PyInstaller @commonInstallerArgs --name "LLIntegrationInstaller-WithToolbar" "installer.py"
} "WithToolbar installer build"

Remove-Item -LiteralPath "native-app\ll_integration_native.exe" -Force
Remove-IfExists "dist\ll_integration_native.exe"

Write-Host "Packaging Opera/Chromium extension..."
Remove-IfExists "dist\LLIntegration-Opera-MV3.zip"
Compress-Archive -Path "opera-extension\*" -DestinationPath "dist\LLIntegration-Opera-MV3.zip" -Force

Write-Host ""
Write-Host "Release artifacts:"
Write-Host "  dist\LLIntegrationInstaller.exe"
Write-Host "  dist\LLIntegrationInstaller-WithToolbar.exe"
Write-Host "  dist\LLIntegration-Opera-MV3.zip"
