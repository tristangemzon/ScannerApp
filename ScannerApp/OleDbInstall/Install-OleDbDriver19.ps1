<#
.SYNOPSIS
    Checks and installs Microsoft OLE DB Driver 19 for SQL Server (x64) and its prerequisites.

.DESCRIPTION
    This script verifies the installation of:
    1. Visual C++ Redistributable 2015-2022 (x86) - Required prerequisite
    2. Visual C++ Redistributable 2015-2022 (x64) - Required prerequisite
    3. Microsoft OLE DB Driver 19 for SQL Server (x64)
    
    Both x86 and x64 VC++ Redistributables are required for OLE DB Driver 19.
    If components are missing, it downloads and installs them in the correct order.

.PARAMETER DownloadPath
    Path where installers will be downloaded. Defaults to user's TEMP folder.

.PARAMETER Force
    Forces reinstallation even if components are already installed.

.EXAMPLE
    .\Install-OleDbDriver19.ps1
    
.EXAMPLE
    .\Install-OleDbDriver19.ps1 -Force
    
.NOTES
    Requires: PowerShell 5.1, Administrator privileges for installation
    Author: Auto-generated
    Date: 2026-04-10
#>

[CmdletBinding()]
param(
    [string]$DownloadPath = $env:TEMP,
    [switch]$Force
)

#Requires -Version 5.1

# URLs for downloads
$VCRedistX64Url = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$VCRedistX86Url = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
$OleDbUrl = "https://go.microsoft.com/fwlink/?linkid=2318101"  # MSOLEDBSQL19 v19.4.1 (x64/Arm64)

# Minimum required versions
$MinVCRedistVersion = [Version]"14.34.0.0"  # VS 2022 minimum required for MSOLEDBSQL19 (per 19.3.0 release notes)
$MinOleDbVersion = [Version]"19.0.0.0"

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )
    
    $colors = @{
        "Info"    = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error"   = "Red"
    }
    
    $prefix = @{
        "Info"    = "[*]"
        "Success" = "[+]"
        "Warning" = "[!]"
        "Error"   = "[-]"
    }
    
    Write-Host "$($prefix[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-InstalledVCRedist {
    <#
    .SYNOPSIS
        Checks if Visual C++ Redistributable 2015-2022 (x86 and x64) are installed
    .PARAMETER Architecture
        Specify 'x86', 'x64', or 'Both' to check specific architecture(s)
    #>
    param(
        [ValidateSet("x86", "x64", "Both")]
        [string]$Architecture = "Both"
    )
    
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $allEntries = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue
    
    # Check x64
    $vcRedistX64 = $allEntries | 
        Where-Object { 
            $_.DisplayName -like "*Visual C++ 2015-2022*x64*" -or
            $_.DisplayName -like "*Visual C++ 2022*x64*Redistributable*"
        } |
        Select-Object -First 1
    
    # Check x86
    $vcRedistX86 = $allEntries | 
        Where-Object { 
            $_.DisplayName -like "*Visual C++ 2015-2022*x86*" -or
            $_.DisplayName -like "*Visual C++ 2022*x86*Redistributable*"
        } |
        Select-Object -First 1
    
    $result = @{
        x64 = @{
            Installed = $false
            Version = $null
            DisplayName = $null
        }
        x86 = @{
            Installed = $false
            Version = $null
            DisplayName = $null
        }
        BothInstalled = $false
    }
    
    if ($vcRedistX64) {
        $result.x64 = @{
            Installed = $true
            Version = [Version]$vcRedistX64.DisplayVersion
            DisplayName = $vcRedistX64.DisplayName
        }
    }
    
    if ($vcRedistX86) {
        $result.x86 = @{
            Installed = $true
            Version = [Version]$vcRedistX86.DisplayVersion
            DisplayName = $vcRedistX86.DisplayName
        }
    }
    
    $result.BothInstalled = $result.x64.Installed -and $result.x86.Installed
    
    return $result
}

function Get-InstalledOleDbDriver {
    <#
    .SYNOPSIS
        Checks if Microsoft OLE DB Driver 19 for SQL Server (x64) is installed
    #>
    
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $oleDb = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue | 
        Where-Object { 
            $_.DisplayName -like "*OLE DB Driver 19*SQL Server*" -or
            $_.DisplayName -like "*MSOLEDBSQL19*"
        } |
        Select-Object -First 1
    
    if ($oleDb) {
        return @{
            Installed = $true
            Version = [Version]$oleDb.DisplayVersion
            DisplayName = $oleDb.DisplayName
        }
    }
    
    return @{ Installed = $false; Version = $null; DisplayName = $null }
}

function Install-VCRedist {
    param(
        [string]$DownloadPath,
        [ValidateSet("x86", "x64")]
        [string]$Architecture = "x64"
    )
    
    $url = if ($Architecture -eq "x64") { $VCRedistX64Url } else { $VCRedistX86Url }
    $installerPath = Join-Path $DownloadPath "vc_redist.$Architecture.exe"
    
    Write-Status "Downloading Visual C++ Redistributable 2015-2022 ($Architecture)..." -Type Info
    
    try {
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Status "Download complete. Installing ($Architecture)..." -Type Info
        
        $process = Start-Process -FilePath $installerPath -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Status "Visual C++ Redistributable ($Architecture) installed successfully." -Type Success
            if ($process.ExitCode -eq 3010) {
                Write-Status "A system restart may be required." -Type Warning
            }
            return $true
        } else {
            Write-Status "Installation ($Architecture) failed with exit code: $($process.ExitCode)" -Type Error
            return $false
        }
    }
    catch {
        Write-Status "Error: $($_.Exception.Message)" -Type Error
        return $false
    }
    finally {
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-OleDbDriver {
    param(
        [string]$DownloadPath
    )
    
    $installerPath = Join-Path $DownloadPath "msoledbsql19.msi"
    
    Write-Status "Downloading Microsoft OLE DB Driver 19 for SQL Server..." -Type Info
    
    try {
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($OleDbUrl, $installerPath)
        
        Write-Status "Download complete. Installing..." -Type Info
        
        $arguments = "/i `"$installerPath`" /quiet /norestart IACCEPTMSOLEDBSQLLICENSETERMS=YES"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Status "OLE DB Driver 19 installed successfully." -Type Success
            if ($process.ExitCode -eq 3010) {
                Write-Status "A system restart may be required." -Type Warning
            }
            return $true
        } else {
            Write-Status "Installation failed with exit code: $($process.ExitCode)" -Type Error
            return $false
        }
    }
    catch {
        Write-Status "Error: $($_.Exception.Message)" -Type Error
        return $false
    }
    finally {
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# MAIN SCRIPT
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  OLE DB Driver 19 for SQL Server - Installation Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check for administrator privileges
$isAdmin = Test-Administrator
if (-not $isAdmin) {
    Write-Status "This script requires Administrator privileges for installation." -Type Warning
    Write-Status "Currently running in check-only mode." -Type Info
    Write-Host ""
}

$needsInstall = $false
$restartRequired = $false

# ============================================================================
# STEP 1: Check Visual C++ Redistributable x86 (PREREQUISITE)
# ============================================================================
Write-Host "--- Step 1a: Checking Visual C++ Redistributable 2015-2022 (x86) ---" -ForegroundColor White
$vcStatus = Get-InstalledVCRedist

if ($vcStatus.x86.Installed -and -not $Force) {
    Write-Status "INSTALLED: $($vcStatus.x86.DisplayName)" -Type Success
    Write-Status "Version: $($vcStatus.x86.Version)" -Type Info
    
    if ($vcStatus.x86.Version -lt $MinVCRedistVersion) {
        Write-Status "Version is below minimum required ($MinVCRedistVersion). Update recommended." -Type Warning
    }
} else {
    if ($Force -and $vcStatus.x86.Installed) {
        Write-Status "Force flag set. Will reinstall VC++ Redistributable (x86)." -Type Warning
    } else {
        Write-Status "NOT INSTALLED: Visual C++ Redistributable 2015-2022 (x86)" -Type Warning
    }
    
    if ($isAdmin) {
        $vcX86Installed = Install-VCRedist -DownloadPath $DownloadPath -Architecture "x86"
        if (-not $vcX86Installed) {
            Write-Status "Failed to install Visual C++ Redistributable (x86). Cannot proceed with OLE DB installation." -Type Error
            exit 1
        }
    } else {
        Write-Status "Run script as Administrator to install." -Type Warning
    }
}

Write-Host ""

# ============================================================================
# STEP 1b: Check Visual C++ Redistributable x64 (PREREQUISITE)
# ============================================================================
Write-Host "--- Step 1b: Checking Visual C++ Redistributable 2015-2022 (x64) ---" -ForegroundColor White
# Refresh status after potential x86 install
$vcStatus = Get-InstalledVCRedist

if ($vcStatus.x64.Installed -and -not $Force) {
    Write-Status "INSTALLED: $($vcStatus.x64.DisplayName)" -Type Success
    Write-Status "Version: $($vcStatus.x64.Version)" -Type Info
    
    if ($vcStatus.x64.Version -lt $MinVCRedistVersion) {
        Write-Status "Version is below minimum required ($MinVCRedistVersion). Update recommended." -Type Warning
    }
} else {
    if ($Force -and $vcStatus.x64.Installed) {
        Write-Status "Force flag set. Will reinstall VC++ Redistributable (x64)." -Type Warning
    } else {
        Write-Status "NOT INSTALLED: Visual C++ Redistributable 2015-2022 (x64)" -Type Warning
    }
    
    if ($isAdmin) {
        $vcX64Installed = Install-VCRedist -DownloadPath $DownloadPath -Architecture "x64"
        if (-not $vcX64Installed) {
            Write-Status "Failed to install Visual C++ Redistributable (x64). Cannot proceed with OLE DB installation." -Type Error
            exit 1
        }
    } else {
        Write-Status "Run script as Administrator to install." -Type Warning
    }
}

Write-Host ""

# ============================================================================
# STEP 2: Check OLE DB Driver 19 (MAIN COMPONENT)
# ============================================================================
Write-Host "--- Step 2: Checking Microsoft OLE DB Driver 19 for SQL Server (x64) ---" -ForegroundColor White
Write-Status "Note: OLE DB Driver 19 requires BOTH x86 and x64 VC++ Redistributables" -Type Info
$oleDbStatus = Get-InstalledOleDbDriver

if ($oleDbStatus.Installed -and -not $Force) {
    Write-Status "INSTALLED: $($oleDbStatus.DisplayName)" -Type Success
    Write-Status "Version: $($oleDbStatus.Version)" -Type Info
} else {
    if ($Force -and $oleDbStatus.Installed) {
        Write-Status "Force flag set. Will reinstall OLE DB Driver." -Type Warning
    } else {
        Write-Status "NOT INSTALLED: Microsoft OLE DB Driver 19 for SQL Server" -Type Warning
    }
    
    if ($isAdmin) {
        # Verify BOTH VC++ redistributables are now installed before proceeding
        $vcRecheck = Get-InstalledVCRedist
        if (-not $vcRecheck.BothInstalled) {
            Write-Status "Both x86 and x64 Visual C++ Redistributables are required but not fully installed. Cannot proceed." -Type Error
            if (-not $vcRecheck.x86.Installed) { Write-Status "Missing: VC++ Redistributable (x86)" -Type Error }
            if (-not $vcRecheck.x64.Installed) { Write-Status "Missing: VC++ Redistributable (x64)" -Type Error }
            exit 1
        }
        
        $oleDbInstalled = Install-OleDbDriver -DownloadPath $DownloadPath
        if (-not $oleDbInstalled) {
            Write-Status "Failed to install OLE DB Driver 19." -Type Error
            exit 1
        }
    } else {
        Write-Status "Run script as Administrator to install." -Type Warning
    }
}

Write-Host ""

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Recheck all components
$finalVcStatus = Get-InstalledVCRedist
$finalOleDbStatus = Get-InstalledOleDbDriver

Write-Host ""
Write-Host "Component                                  Status" -ForegroundColor White
Write-Host "---------                                  ------" -ForegroundColor White

if ($finalVcStatus.x86.Installed) {
    Write-Host "Visual C++ Redistributable 2015-2022 (x86) " -NoNewline
    Write-Host "INSTALLED ($($finalVcStatus.x86.Version))" -ForegroundColor Green
} else {
    Write-Host "Visual C++ Redistributable 2015-2022 (x86) " -NoNewline
    Write-Host "NOT INSTALLED" -ForegroundColor Red
}

if ($finalVcStatus.x64.Installed) {
    Write-Host "Visual C++ Redistributable 2015-2022 (x64) " -NoNewline
    Write-Host "INSTALLED ($($finalVcStatus.x64.Version))" -ForegroundColor Green
} else {
    Write-Host "Visual C++ Redistributable 2015-2022 (x64) " -NoNewline
    Write-Host "NOT INSTALLED" -ForegroundColor Red
}

if ($finalOleDbStatus.Installed) {
    Write-Host "Microsoft OLE DB Driver 19 for SQL Server  " -NoNewline
    Write-Host "INSTALLED ($($finalOleDbStatus.Version))" -ForegroundColor Green
} else {
    Write-Host "Microsoft OLE DB Driver 19 for SQL Server  " -NoNewline
    Write-Host "NOT INSTALLED" -ForegroundColor Red
}

Write-Host ""

if ($finalVcStatus.BothInstalled -and $finalOleDbStatus.Installed) {
    Write-Status "All components are installed and ready." -Type Success
    exit 0
} else {
    if (-not $isAdmin) {
        Write-Status "Run this script as Administrator to install missing components." -Type Warning
    }
    exit 1
}
