# ScannerApp Developer Guide

## Project Overview

ScannerApp is a Windows desktop application for TWAIN-based document scanning with PDF output.

- **Framework**: .NET Framework 4.8
- **Architecture**: x86 (32-bit) default, x64 (64-bit) optional
- **Key Dependencies**: NTwain, PdfSharp 6.x

---

## Build Configurations

| Configuration | Platform | Output Path | Debug Symbols | Use For |
|---------------|----------|-------------|---------------|---------|
| Debug | AnyCPU | bin\Debug\ | Full PDB | Development |
| Release | AnyCPU | bin\Release\ | PDB only | Testing |
| **Release** | **x86** | **bin\x86\Release\** | **None** | **Deployment (32-bit scanners)** |
| Release | x64 | bin\x64\Release\ | None | Deployment (64-bit scanners) |

> **Note**: Use **x86 (32-bit)** for maximum compatibility with older scanners that only have 32-bit TWAIN drivers.

### Platform Selection Guide

| Platform | Runtime Behavior | TWAIN Compatibility | Recommended For |
|----------|-----------------|---------------------|-----------------|
| **x86** | Always 32-bit | ✓ All 32-bit drivers | Default - most scanners |
| x64 | Always 64-bit | 64-bit drivers only | Scanners with 64-bit drivers |
| AnyCPU | 64-bit on 64-bit OS | 64-bit drivers only | Advanced scenarios only |

> **Performance Note**: There is virtually no performance difference between x86, x64, and AnyCPU. The IL code is JIT-compiled to native code at runtime. Choose based on TWAIN driver compatibility, not performance.

---

## Building for Deployment

### Option 1: PowerShell Script (Recommended)

Run the build script from the `ScannerApp` project folder:

```powershell
cd ScannerApp

# Build x86 (32-bit) - DEFAULT for older scanner compatibility
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1

# Build x64 (64-bit) - only if scanner has 64-bit drivers
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1 -Platform x64

# Build AnyCPU - WARNING: runs as 64-bit on 64-bit Windows
# Use only if you have 64-bit TWAIN drivers or want flexible deployment
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1 -Platform AnyCPU

# Clean deploy folder first, then build
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1 -Clean
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1 -Platform x64 -Clean
powershell -ExecutionPolicy Bypass -File .\Build-Deploy.ps1 -Platform AnyCPU -Clean
```

> **Note**: The `-ExecutionPolicy Bypass` flag is required because the script is not digitally signed.

Output is placed in platform-specific folders:
- `ScannerApp\Deploy\x86\` - 32-bit deployment
- `ScannerApp\Deploy\x64\` - 64-bit deployment
- `ScannerApp\Deploy\AnyCPU\` - AnyCPU deployment

> **WARNING - AnyCPU and TWAIN Drivers**: AnyCPU builds run as 64-bit on 64-bit Windows by default. Most scanners only have 32-bit TWAIN drivers. A 64-bit process **cannot** load 32-bit TWAIN drivers. Use **x86** unless you are certain your scanner has 64-bit drivers.

### Option 2: Visual Studio

1. Open `ScannerApp.slnx` in Visual Studio
2. Set configuration to **Release** and platform to **x86** (or x64)
3. Build → Rebuild Solution
4. Output is in `ScannerApp\bin\x86\Release\` (or `bin\x64\Release\`)

### Option 3: Command Line (MSBuild)

```powershell
# Find MSBuild (Visual Studio 2022)
$msbuild = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"

# Build x86 (32-bit)
& $msbuild ScannerApp\ScannerApp.csproj /p:Configuration=Release /p:Platform=x86 /t:Rebuild

# Build x64 (64-bit)
& $msbuild ScannerApp\ScannerApp.csproj /p:Configuration=Release /p:Platform=x64 /t:Rebuild
```

---

## Deployment Files

Each `Deploy\<Platform>\` folder contains only essential runtime files:

| File Type | Included | Notes |
|-----------|----------|-------|
| *.exe | ✓ | Main application |
| *.dll | ✓ | Dependencies |
| *.config | ✓ | App configuration |
| *.pdb | ✗ | Debug symbols (excluded) |
| *.xml | ✗ | XML documentation (excluded) |

### Required Files (~2.2 MB)

- `ScannerApp_x86.exe` or `ScannerApp_x64.exe` - Main executable (platform-specific name)
- `ScannerApp_x86.exe.config` or `ScannerApp_x64.exe.config` - Configuration
- `NTwain.dll` - TWAIN scanning library
- `PdfSharp.dll` + related DLLs - PDF generation
- `Microsoft.*.dll` - .NET extensions
- `System.*.dll` - .NET support libraries

---

## System Requirements

### Runtime Requirements
- Windows 10/11
- .NET Framework 4.8
- TWAIN-compatible scanner (32-bit or 64-bit drivers)

### Development Requirements
- Visual Studio 2022 or later
- .NET Framework 4.8 SDK
- Windows SDK (for TWAIN support)

---

## Project Structure

```
ScannerApp/
├── ScannerApp.slnx        # Solution file
├── packages/              # NuGet packages (restored)
└── ScannerApp/
    ├── Build-Deploy.ps1  # Build & deploy script
    ├── Deploy/            # Deployment output (generated)
    │   ├── x86/          # 32-bit deployment files
    │   ├── x64/          # 64-bit deployment files
    │   └── AnyCPU/       # AnyCPU deployment files
    ├── DEVELOPER-GUIDE.md # This guide
    ├── ScannerApp.csproj  # Project file
    ├── Program.cs         # Entry point
    ├── TwainScanner.cs    # TWAIN scanning implementation
    ├── TwainScannerExt.cs # Scanner capability extensions
    ├── PdfCreator.cs      # PDF document generation
    └── Logger.cs          # Logging utility
```

---

## Adding New Build Configurations

To add a new configuration (e.g., Debug|x64), edit `ScannerApp.csproj`:

```xml
<PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|x64' ">
  <PlatformTarget>x64</PlatformTarget>
  <DebugSymbols>true</DebugSymbols>
  <DebugType>full</DebugType>
  <Optimize>false</Optimize>
  <OutputPath>bin\x64\Debug\</OutputPath>
  <DefineConstants>DEBUG;TRACE</DefineConstants>
  <ErrorReport>prompt</ErrorReport>
  <WarningLevel>4</WarningLevel>
</PropertyGroup>
```

Then use Configuration Manager in Visual Studio to create the matching solution platform.

---

## Troubleshooting

### "x64 platform not showing in Visual Studio"
1. Build → Configuration Manager
2. Active solution platform → \<New...\>
3. Select x64, copy settings from AnyCPU
4. Reload solution

### "MSBuild not found"
Install Visual Studio with ".NET desktop development" workload, or use the Visual Studio Installer to add MSBuild tools.

### "TWAIN source not found"
This usually means the app bitness doesn't match the scanner driver:
- **Older scanners**: Use the **x86 (32-bit)** build - most older scanners only have 32-bit TWAIN drivers
- **Newer scanners**: May work with either x86 or x64, but x86 is safest

To check if your scanner has 64-bit drivers:
1. Look in `C:\Windows\twain_64` for 64-bit drivers
2. Look in `C:\Windows\twain_32` for 32-bit drivers
3. If only `twain_32` exists, use the x86 build

### Execute With Parameters Example
```
ScannerApp_x64.exe --sourceindex 0 --output "" --feeder true --duplex true --colormode color --resolution high --pagewidth 8500 --pageheight 11000
```