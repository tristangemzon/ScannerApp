#Requires -Version 5.1
<#
.SYNOPSIS
    Builds ScannerApp in Release configuration and copies minimum files to Deploy folder.

.DESCRIPTION
    This script:
    1. Locates MSBuild from Visual Studio installation
    2. Builds the project in Release configuration for the specified platform
    3. Copies only essential runtime files (no PDB, XML, or documentation) to Deploy folder

.PARAMETER Platform
    Target platform: x86 (32-bit) or x64 (64-bit). Default is x86 for maximum scanner compatibility.

.PARAMETER Clean
    If specified, cleans the Deploy folder before copying new files.

.EXAMPLE
    .\Build-Deploy.ps1
    Builds x86 (32-bit) and deploys to the Deploy folder.

.EXAMPLE
    .\Build-Deploy.ps1 -Platform x64
    Builds x64 (64-bit) and deploys to the Deploy folder.

.EXAMPLE
    .\Build-Deploy.ps1 -Platform x86 -Clean
    Cleans Deploy folder, then builds x86 and deploys.
#>

param(
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x86",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

# Paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = $ScriptDir
$ProjectFile = Join-Path $ProjectDir "ScannerApp.csproj"
$OutputDir = Join-Path $ProjectDir "bin\$Platform\Release"
$DeployDir = Join-Path $ProjectDir "Deploy"

# Find MSBuild
function Find-MSBuild {
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vsWhere) {
        $msbuildPath = & $vsWhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if ($msbuildPath) { return $msbuildPath }
    }
    
    # Fallback paths
    $fallbackPaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2019\*\MSBuild\Current\Bin\MSBuild.exe"
    )
    
    foreach ($path in $fallbackPaths) {
        $found = Get-ChildItem $path -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    
    throw "MSBuild not found. Please install Visual Studio with .NET desktop development workload."
}

$bitness = if ($Platform -eq "x86") { "32-bit" } else { "64-bit" }
Write-Host "=== ScannerApp Build & Deploy ($bitness) ===" -ForegroundColor Cyan
Write-Host ""

# Step 1: Find MSBuild
Write-Host "Finding MSBuild..." -ForegroundColor Yellow
$msbuild = Find-MSBuild
Write-Host "  Found: $msbuild" -ForegroundColor Green

# Step 2: Build
Write-Host ""
Write-Host "Building Release|$Platform..." -ForegroundColor Yellow
& $msbuild $ProjectFile /p:Configuration=Release /p:Platform=$Platform /t:Rebuild /v:minimal /nologo

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "  Build succeeded" -ForegroundColor Green

# Step 3: Prepare Deploy folder
Write-Host ""
Write-Host "Preparing Deploy folder..." -ForegroundColor Yellow

if ($Clean -and (Test-Path $DeployDir)) {
    Remove-Item $DeployDir -Recurse -Force
    Write-Host "  Cleaned existing Deploy folder" -ForegroundColor Gray
}

if (-not (Test-Path $DeployDir)) {
    New-Item -ItemType Directory -Path $DeployDir -Force | Out-Null
}

# Step 4: Copy only essential files (no PDB, no XML documentation)
Write-Host ""
Write-Host "Copying deployment files..." -ForegroundColor Yellow

$extensions = @("*.exe", "*.dll", "*.config")
$copiedFiles = @()

foreach ($ext in $extensions) {
    $files = Get-ChildItem (Join-Path $OutputDir $ext) -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        Copy-Item $file.FullName -Destination $DeployDir -Force
        $copiedFiles += $file
    }
}

# Step 5: Summary
Write-Host ""
Write-Host "=== Deployment Summary ===" -ForegroundColor Cyan
Write-Host "  Platform: $Platform ($bitness)" -ForegroundColor White
Write-Host "  Output: $DeployDir" -ForegroundColor White
Write-Host "  Files:  $($copiedFiles.Count)" -ForegroundColor White

$totalSize = ($copiedFiles | Measure-Object -Property Length -Sum).Sum
Write-Host "  Size:   $([math]::Round($totalSize / 1MB, 2)) MB" -ForegroundColor White

Write-Host ""
Write-Host "Deployed files:" -ForegroundColor Gray
$copiedFiles | Sort-Object Name | ForEach-Object {
    $sizeKB = [math]::Round($_.Length / 1KB, 1)
    Write-Host "  $($_.Name) ($sizeKB KB)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Deploy completed successfully!" -ForegroundColor Green
