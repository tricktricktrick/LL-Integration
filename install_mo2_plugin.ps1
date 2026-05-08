param(
    [Parameter(Mandatory = $true)]
    [string]$Mo2PluginsPath,

    [switch]$WithToolbar
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $root "mo2-plugin"
$target = Join-Path $Mo2PluginsPath "ll_integration"

if (-not (Test-Path -LiteralPath $Mo2PluginsPath)) {
    throw "MO2 plugins path does not exist: $Mo2PluginsPath"
}

New-Item -ItemType Directory -Force -Path $target | Out-Null

$files = @(
    "__init__.py",
    "plugin.py",
    "utils.py",
    "check_update.py",
    "LL.sample.ini"
)

foreach ($file in $files) {
    Copy-Item -LiteralPath (Join-Path $source $file) -Destination (Join-Path $target $file) -Force
}

$iconsSource = Join-Path $source "icons"
$iconsTarget = Join-Path $target "icons"
if (Test-Path -LiteralPath $iconsSource) {
    New-Item -ItemType Directory -Force -Path $iconsTarget | Out-Null
    Copy-Item -Path (Join-Path $iconsSource "*") -Destination $iconsTarget -Force
}

$experimentalSource = Join-Path $source "experimental"
$experimentalTarget = Join-Path $target "experimental"
if ($WithToolbar -and (Test-Path -LiteralPath $experimentalSource)) {
    New-Item -ItemType Directory -Force -Path $experimentalTarget | Out-Null
    Copy-Item -Path (Join-Path $experimentalSource "*") -Destination $experimentalTarget -Recurse -Force
} elseif (Test-Path -LiteralPath $experimentalTarget) {
    Remove-Item -LiteralPath $experimentalTarget -Recurse -Force
}

$nativeRoot = Join-Path $env:LOCALAPPDATA "LLIntegration\native-app"
if (-not (Test-Path -LiteralPath $nativeRoot)) {
    $nativeRoot = Join-Path $root "native-app"
}

$paths = @{
    ll_ini_path = Join-Path $nativeRoot "downloads_storage\latest_ll_download.ini"
    cookies_path = Join-Path $nativeRoot "cookies_storage\cookies_ll.json"
    experimental_toolbar = [bool]$WithToolbar
} | ConvertTo-Json -Depth 2

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText((Join-Path $target "plugin_paths.json"), $paths, $utf8NoBom)

Write-Host "Installed LL Integration MO2 plugin to: $target"
if ($WithToolbar) {
    Write-Host "Experimental toolbar button enabled. Restart MO2."
} else {
    Write-Host "Stable mode installed. Restart MO2, then open Tools > LL Integration."
}
