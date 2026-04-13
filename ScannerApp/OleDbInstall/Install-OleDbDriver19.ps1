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
    [switch]$Force,
    [switch]$DiagnoseVCRedist
)

#Requires -Version 5.1

# URLs for downloads (permalinks for latest supported versions)
$VCRedistX64Url = "https://aka.ms/vc14/vc_redist.x64.exe"  # Latest VC++ v14 (14.50.35719.0+)
$VCRedistX86Url = "https://aka.ms/vc14/vc_redist.x86.exe"  # Latest VC++ v14 (14.50.35719.0+)
$OleDbUrl = "https://go.microsoft.com/fwlink/?linkid=2318101"  # MSOLEDBSQL19 v19.4.1 (x64/Arm64)

# Minimum required versions
$MinVCRedistVersion = [Version]"14.34.0.0"  # VS 2022 minimum required for MSOLEDBSQL19 (per 19.3.0 release notes)
$MinOleDbVersion = [Version]"19.0.0.0"

function Show-VCRedistDiagnostics {
    <#
    .SYNOPSIS
        Shows all Visual C++ Redistributables found in the registry for troubleshooting
    #>
    Write-Host ""
    Write-Host "=== VC++ Redistributable Diagnostics ===" -ForegroundColor Cyan
    
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $allEntries = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Visual C++*" } |
        Select-Object DisplayName, DisplayVersion, PSPath |
        Sort-Object DisplayName
    
    if ($allEntries) {
        Write-Host "Found the following Visual C++ entries in registry:" -ForegroundColor Yellow
        Write-Host ""
        foreach ($entry in $allEntries) {
            Write-Host "  Name: $($entry.DisplayName)" -ForegroundColor White
            Write-Host "  Version: $($entry.DisplayVersion)" -ForegroundColor Gray
            Write-Host "  Path: $($entry.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::', '')" -ForegroundColor DarkGray
            Write-Host ""
        }
    } else {
        Write-Host "No Visual C++ Redistributables found in registry!" -ForegroundColor Red
    }
    
    Write-Host "=== End Diagnostics ===" -ForegroundColor Cyan
    Write-Host ""
}

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

function Test-InstallerBusy {
    <#
    .SYNOPSIS
        Checks if Windows Installer (msiexec) is currently running another installation
    #>
    $msiProcesses = Get-Process -Name msiexec -ErrorAction SilentlyContinue | 
        Where-Object { $_.Id -ne $PID }
    
    # Also check for any vc_redist installers running
    $vcRedistProcesses = Get-Process -Name "vc_redist*" -ErrorAction SilentlyContinue
    
    return ($msiProcesses -and $msiProcesses.Count -gt 0) -or 
           ($vcRedistProcesses -and $vcRedistProcesses.Count -gt 0)
}

function Wait-InstallerFree {
    <#
    .SYNOPSIS
        Waits for Windows Installer to become available
    .PARAMETER MaxWaitSeconds
        Maximum time to wait in seconds (default: 120)
    .PARAMETER CheckIntervalSeconds
        How often to check in seconds (default: 5)
    #>
    param(
        [int]$MaxWaitSeconds = 120,
        [int]$CheckIntervalSeconds = 5
    )
    
    $elapsed = 0
    while ((Test-InstallerBusy) -and ($elapsed -lt $MaxWaitSeconds)) {
        if ($elapsed -eq 0) {
            Write-Status "Another installer is running. Waiting for it to complete..." -Type Warning
        }
        Start-Sleep -Seconds $CheckIntervalSeconds
        $elapsed += $CheckIntervalSeconds
        Write-Host "." -NoNewline
    }
    
    if ($elapsed -gt 0) {
        Write-Host ""  # New line after dots
    }
    
    if (Test-InstallerBusy) {
        Write-Status "Installer is still busy after waiting $MaxWaitSeconds seconds." -Type Warning
        return $false
    }
    
    return $true
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
    
    # Check x64 - expanded patterns to catch various naming conventions
    $vcRedistX64 = $allEntries | 
        Where-Object { 
            ($_.DisplayName -like "*Visual C++ 2015-2022*x64*") -or
            ($_.DisplayName -like "*Visual C++ 2022*x64*Redistributable*") -or
            ($_.DisplayName -like "*Visual C++ v14*Redistributable*(x64)*") -or
            ($_.DisplayName -like "*Microsoft Visual C++*Redistributable*(x64)*" -and $_.DisplayName -match "201[5-9]|202[0-9]|v14") -or
            ($_.DisplayName -like "*VC++ 2015-2022*x64*")
        } |
        Sort-Object { [Version]$_.DisplayVersion } -Descending |
        Select-Object -First 1
    
    # Check x86 - expanded patterns to catch various naming conventions
    $vcRedistX86 = $allEntries | 
        Where-Object { 
            ($_.DisplayName -like "*Visual C++ 2015-2022*x86*") -or
            ($_.DisplayName -like "*Visual C++ 2022*x86*Redistributable*") -or
            ($_.DisplayName -like "*Visual C++ v14*Redistributable*(x86)*") -or
            ($_.DisplayName -like "*Microsoft Visual C++*Redistributable*(x86)*" -and $_.DisplayName -match "201[5-9]|202[0-9]|v14") -or
            ($_.DisplayName -like "*VC++ 2015-2022*x86*")
        } |
        Sort-Object { [Version]$_.DisplayVersion } -Descending |
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
        [string]$Architecture = "x64",
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 30
    )
    
    $url = if ($Architecture -eq "x64") { $VCRedistX64Url } else { $VCRedistX86Url }
    $installerPath = Join-Path $DownloadPath "vc_redist.$Architecture.exe"
    $logPath = Join-Path $DownloadPath "vc_redist_$Architecture.log"
    
    # Wait for any existing installer to complete before starting
    if (-not (Wait-InstallerFree -MaxWaitSeconds 120)) {
        Write-Status "Cannot proceed while another installer is running. Please close other installers and try again." -Type Error
        return $false
    }
    
    Write-Status "Downloading Visual C++ Redistributable 2015-2022 ($Architecture)..." -Type Info
    
    try {
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Status "Download complete. Installing ($Architecture)..." -Type Info
        
        $attempt = 0
        $installSuccess = $false
        
        while (-not $installSuccess -and $attempt -lt $MaxRetries) {
            $attempt++
            
            if ($attempt -gt 1) {
                Write-Status "Retry attempt $attempt of $MaxRetries..." -Type Info
                # Wait for installer to be free before retrying
                if (-not (Wait-InstallerFree -MaxWaitSeconds 60)) {
                    Write-Status "Installer still busy, waiting $RetryDelaySeconds seconds before retry..." -Type Warning
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
            
            # Use /passive for progress bar without user interaction (more reliable than /quiet)
            $arguments = "/install /passive /norestart /log `"$logPath`""
            $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
            
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010 -or $process.ExitCode -eq 1638) {
                # 0 = Success, 3010 = Reboot required, 1638 = Another version already installed
                if ($process.ExitCode -eq 1638) {
                    Write-Status "A newer version of Visual C++ Redistributable ($Architecture) is already installed." -Type Success
                } else {
                    Write-Status "Visual C++ Redistributable ($Architecture) installer completed with exit code: $($process.ExitCode)" -Type Info
                }
                if ($process.ExitCode -eq 3010) {
                    Write-Status "A system restart may be required." -Type Warning
                }
                
                # Verify installation actually succeeded by checking registry
                Start-Sleep -Seconds 2  # Brief delay for registry to be updated
                $verifyStatus = Get-InstalledVCRedist -Architecture $Architecture
                $archKey = $Architecture.ToLower()
                if ($verifyStatus.$archKey.Installed) {
                    Write-Status "Visual C++ Redistributable ($Architecture) verified installed: $($verifyStatus.$archKey.Version)" -Type Success
                    $installSuccess = $true
                } else {
                    Write-Status "WARNING: Installer reported success but verification failed!" -Type Error
                    Write-Status "The ($Architecture) redistributable may require a system reboot to complete installation." -Type Warning
                    Write-Status "Check log file: $logPath" -Type Info
                    $installSuccess = $false
                }
            } elseif ($process.ExitCode -eq 1618 -or $process.ExitCode -eq 1602) {
                # 1618 = Another installation in progress, 1602 = User cancelled (may indicate installer conflict)
                Write-Status "Another installation is in progress (exit code: $($process.ExitCode))." -Type Warning
                if ($attempt -lt $MaxRetries) {
                    Write-Status "Will retry after waiting..." -Type Info
                }
                $installSuccess = $false
            } else {
                Write-Status "Installation ($Architecture) failed with exit code: $($process.ExitCode)" -Type Error
                Write-Status "Check log file: $logPath" -Type Info
                # Don't retry for other failures
                break
            }
        }  # End retry loop
        
        return $installSuccess
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
        [string]$DownloadPath,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 30
    )
    
    $installerPath = Join-Path $DownloadPath "msoledbsql19.msi"
    $logPath = Join-Path $DownloadPath "msoledbsql19.log"
    
    # Wait for any existing installer to complete before starting
    if (-not (Wait-InstallerFree -MaxWaitSeconds 120)) {
        Write-Status "Cannot proceed while another installer is running. Please close other installers and try again." -Type Error
        return $false
    }
    
    Write-Status "Downloading Microsoft OLE DB Driver 19 for SQL Server..." -Type Info
    
    try {
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($OleDbUrl, $installerPath)
        
        Write-Status "Download complete. Installing..." -Type Info
        
        $attempt = 0
        $installSuccess = $false
        
        while (-not $installSuccess -and $attempt -lt $MaxRetries) {
            $attempt++
            
            if ($attempt -gt 1) {
                Write-Status "Retry attempt $attempt of $MaxRetries..." -Type Info
                # Wait for installer to be free before retrying
                if (-not (Wait-InstallerFree -MaxWaitSeconds 60)) {
                    Write-Status "Installer still busy, waiting $RetryDelaySeconds seconds before retry..." -Type Warning
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
            
            $arguments = "/i `"$installerPath`" /quiet /norestart /log `"$logPath`" IACCEPTMSOLEDBSQLLICENSETERMS=YES"
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru
            
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                Write-Status "OLE DB Driver 19 installed successfully." -Type Success
                if ($process.ExitCode -eq 3010) {
                    Write-Status "A system restart may be required." -Type Warning
                }
                $installSuccess = $true
            } elseif ($process.ExitCode -eq 1618) {
                # 1618 = Another installation in progress
                Write-Status "Another installation is in progress (exit code: 1618)." -Type Warning
                if ($attempt -lt $MaxRetries) {
                    Write-Status "Will retry after waiting..." -Type Info
                }
                $installSuccess = $false
            } else {
                Write-Status "Installation failed with exit code: $($process.ExitCode)" -Type Error
                Write-Status "Check log file: $logPath" -Type Info
                # Don't retry for other failures
                break
            }
        }  # End retry loop
        
        return $installSuccess
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

# Show diagnostics if requested
if ($DiagnoseVCRedist) {
    Show-VCRedistDiagnostics
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
