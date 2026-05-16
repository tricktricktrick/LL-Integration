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

function New-ExtensionZip {
    param(
        [Parameter(Mandatory = $true)][string]$SourceDir,
        [Parameter(Mandatory = $true)][string]$DestinationZip,
        [string[]]$Exclude = @()
    )

    Add-Type -AssemblyName System.IO.Compression
    Remove-IfExists $DestinationZip

    $sourceRoot = (Resolve-Path -LiteralPath $SourceDir).Path.TrimEnd("\", "/")
    $excludeSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $Exclude) {
        [void]$excludeSet.Add(($item -replace "\\", "/"))
    }

    $zipStream = [System.IO.File]::Open($DestinationZip, [System.IO.FileMode]::CreateNew)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | ForEach-Object {
                $relativePath = $_.FullName.Substring($sourceRoot.Length).TrimStart("\", "/") -replace "\\", "/"
                if ($excludeSet.Contains($relativePath)) {
                    return
                }

                $entry = $zip.CreateEntry($relativePath, [System.IO.Compression.CompressionLevel]::Optimal)
                $inputStream = [System.IO.File]::OpenRead($_.FullName)
                try {
                    $entryStream = $entry.Open()
                    try {
                        $inputStream.CopyTo($entryStream)
                    } finally {
                        $entryStream.Dispose()
                    }
                } finally {
                    $inputStream.Dispose()
                }
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $zipStream.Dispose()
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
        --icon "native-app\icons\ll_integration.ico" `
        --name "ll_integration_native" `
        "native-app\main.py"
} "Native helper build"

Copy-Item -LiteralPath "dist\ll_integration_native.exe" -Destination "native-app\ll_integration_native.exe" -Force

Write-Host "Building floating capture controls..."
Remove-IfExists "dist\ll_integration_overlay.exe"
Invoke-Checked {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --icon "native-app\icons\ll_integration.ico" `
        --add-data "native-app\icons;icons" `
        --name "ll_integration_overlay" `
        "native-app\overlay.py"
} "Floating controls build"

Copy-Item -LiteralPath "dist\ll_integration_overlay.exe" -Destination "native-app\ll_integration_overlay.exe" -Force

Write-Host "Building Vortex manager..."
Remove-IfExists "dist\ll_integration_vortex_manager.exe"
Invoke-Checked {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --icon "native-app\icons\ll_integration.ico" `
        --add-data "native-app\icons;icons" `
        --name "ll_integration_vortex_manager" `
        "native-app\manager_vortex.py"
} "Vortex manager build"

Copy-Item -LiteralPath "dist\ll_integration_vortex_manager.exe" -Destination "native-app\ll_integration_vortex_manager.exe" -Force

$commonInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--icon", "native-app\icons\ll_integration.ico",
    "--add-data", "native-app\icons;icons",
    "--add-data", "native-app;native-app",
    "--add-data", "mo2-plugin;mo2-plugin",
    "--add-data", "vortex-extension;vortex-extension"
)

Write-Host "Building stable installer..."
Remove-IfExists "dist\LLIntegrationInstaller.exe"
Invoke-Checked {
    & $Python -m PyInstaller @commonInstallerArgs --name "LLIntegrationInstaller" "installer.py"
} "Stable installer build"

Remove-Item -LiteralPath "native-app\ll_integration_native.exe" -Force
Remove-Item -LiteralPath "native-app\ll_integration_overlay.exe" -Force
Remove-Item -LiteralPath "native-app\ll_integration_vortex_manager.exe" -Force
Remove-IfExists "dist\ll_integration_native.exe"
Remove-IfExists "dist\ll_integration_overlay.exe"
Remove-IfExists "dist\ll_integration_vortex_manager.exe"

Write-Host "Packaging Opera/Chromium extension..."
New-ExtensionZip `
    -SourceDir "opera-extension" `
    -DestinationZip "dist\LLIntegration-Opera-MV3.zip" `
    -Exclude @("icons/ll_integration.svg")

Write-Host ""
Write-Host "Release artifacts:"
Write-Host "  dist\LLIntegrationInstaller.exe"
Write-Host "  dist\LLIntegration-Opera-MV3.zip"
