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
New-ExtensionZip `
    -SourceDir "opera-extension" `
    -DestinationZip "dist\LLIntegration-Opera-MV3.zip" `
    -Exclude @("icons/ll_integration.svg")

Write-Host ""
Write-Host "Release artifacts:"
Write-Host "  dist\LLIntegrationInstaller.exe"
Write-Host "  dist\LLIntegrationInstaller-WithToolbar.exe"
Write-Host "  dist\LLIntegration-Opera-MV3.zip"
