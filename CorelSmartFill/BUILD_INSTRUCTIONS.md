# Build Instructions - CorelDRAW Smart Fill Tool

## Quick Build (Windows)

### Prerequisites Check

```powershell
# Check if .NET 6 SDK is installed
dotnet --version
# Should show: 6.0.x or higher

# If not installed, download from:
# https://dotnet.microsoft.com/download/dotnet/6.0
```

### Build Steps

```bash
# 1. Navigate to project folder
cd CorelSmartFill

# 2. Restore dependencies
dotnet restore

# 3. Build (Debug)
dotnet build

# 4. Build (Release - recommended)
dotnet build --configuration Release

# 5. Run
dotnet run
# OR directly:
.\bin\Release\net6.0-windows\CorelSmartFill.exe
```

---

## Publish for Distribution

### Create Single-File Executable

```bash
# This creates ONE .exe file that includes everything
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true

# Output location:
# bin\Release\net6.0-windows\win-x64\publish\CorelSmartFill.exe

# This file can be copied to any Windows PC and run without installing .NET!
```

### Create Framework-Dependent Build

```bash
# Smaller file size, but requires .NET 6 Runtime on target PC
dotnet publish -c Release -r win-x64 --self-contained false

# Output location:
# bin\Release\net6.0-windows\win-x64\publish\
```

---

## Using Visual Studio 2022

### Method 1: Build from IDE

1. **Open Project**
   - Launch Visual Studio 2022
   - File → Open → Project/Solution
   - Select `CorelSmartFill.csproj`

2. **Set Build Configuration**
   - Top toolbar: Change "Debug" to "Release"

3. **Build**
   - Build → Build Solution (Ctrl+Shift+B)
   - Wait for "Build succeeded" message

4. **Find Output**
   - Right-click project in Solution Explorer
   - Open Folder in File Explorer
   - Navigate to: `bin\Release\net6.0-windows\`
   - Run: `CorelSmartFill.exe`

### Method 2: Publish from IDE

1. **Right-click project** in Solution Explorer
2. **Select "Publish..."**
3. **Choose target:**
   - Folder
   - Click "Next"
4. **Choose location:**
   - Browse to desired output folder
   - Click "Finish"
5. **Click "Publish"**
6. **Find output** in chosen folder

---

## Troubleshooting Build Issues

### Error: SDK not found

**Problem:**
```
error NETSDK1045: The current .NET SDK does not support targeting .NET 6.0
```

**Solution:**
```bash
# Install .NET 6 SDK
# Download from: https://dotnet.microsoft.com/download/dotnet/6.0
# Choose "SDK" not "Runtime"
```

### Error: COM Reference issues

**Problem:**
```
error: CorelDRAW COM reference cannot be resolved
```

**Solution:**
This is OK! COM reference will be resolved at runtime when CorelDRAW is installed. The app will still build.

If you want to avoid the warning:
1. Remove the COM reference from .csproj
2. Use late binding (dynamic) only

### Error: Windows Forms not supported

**Problem:**
```
error NETSDK1135: Windows Forms is only supported on Windows
```

**Solution:**
You must build on Windows. This is a Windows-only application.

If on Linux/Mac:
```bash
# Use a Windows VM or dual-boot
# OR use Remote Desktop to Windows machine
# OR use GitHub Actions (see CI/CD section)
```

### Warning: Platform target mismatch

**Problem:**
```
warning MSB3270: There was a mismatch between processor architecture
```

**Solution:**
```bash
# Force x64 build
dotnet build -r win-x64

# OR edit .csproj, add:
<PropertyGroup>
  <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>
```

---

## Continuous Integration (CI/CD)

### GitHub Actions Example

Create `.github/workflows/build.yml`:

```yaml
name: Build CorelSmartFill

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build:
    runs-on: windows-latest

    steps:
    - uses: actions/checkout@v3

    - name: Setup .NET
      uses: actions/setup-dotnet@v3
      with:
        dotnet-version: 6.0.x

    - name: Restore dependencies
      run: dotnet restore CorelSmartFill/CorelSmartFill.csproj

    - name: Build
      run: dotnet build CorelSmartFill/CorelSmartFill.csproj --configuration Release --no-restore

    - name: Publish
      run: dotnet publish CorelSmartFill/CorelSmartFill.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true

    - name: Upload artifact
      uses: actions/upload-artifact@v3
      with:
        name: CorelSmartFill-exe
        path: CorelSmartFill/bin/Release/net6.0-windows/win-x64/publish/CorelSmartFill.exe
```

---

## Build Configurations

### Debug Build
- **Purpose:** Development, testing
- **Features:**
  - Debug symbols included
  - No optimizations
  - Larger file size
- **Speed:** Slower
- **Command:** `dotnet build`

### Release Build
- **Purpose:** Distribution to users
- **Features:**
  - Optimized code
  - No debug symbols
  - Smaller file size
- **Speed:** Faster
- **Command:** `dotnet build -c Release`

---

## Output Files Explained

After successful build in `bin\Release\net6.0-windows\`:

```
CorelSmartFill.exe          - Main executable (framework-dependent)
CorelSmartFill.dll          - Application library
CorelSmartFill.deps.json    - Dependency information
CorelSmartFill.runtimeconfig.json - Runtime configuration
System.Drawing.Common.dll   - Required library
```

**To distribute:**
- **Framework-dependent:** Copy entire folder
- **Self-contained:** Just the single .exe from publish folder

---

## File Size Comparison

| Build Type | Size | Requires .NET? |
|------------|------|----------------|
| Debug | ~50 KB | Yes |
| Release | ~50 KB | Yes |
| Self-contained | ~140 MB | No |
| Self-contained + Trimmed | ~80 MB | No |

**Reduce self-contained size:**

```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:PublishTrimmed=true /p:TrimMode=link
```

**Warning:** Trimming may cause issues with COM interop. Test thoroughly!

---

## Testing the Build

### Quick Test

```bash
# 1. Build
dotnet build -c Release

# 2. Run
dotnet run --configuration Release

# 3. Verify:
# - Window opens
# - Click "Connect to CorelDRAW"
# - Should show connection message
```

### Full Test Checklist

- [ ] Application starts without errors
- [ ] Window displays correctly
- [ ] "Connect to CorelDRAW" button works
- [ ] Connection succeeds when CorelDRAW is running
- [ ] Can select shapes in CorelDRAW
- [ ] Fill operation works
- [ ] Sliders update values
- [ ] Status display shows correct info
- [ ] Application closes cleanly

---

## Distribution Checklist

Before distributing to users:

- [ ] Build in Release mode
- [ ] Test on clean Windows 10/11 VM
- [ ] Verify CorelDRAW integration works
- [ ] Check all features work
- [ ] Include README.md
- [ ] Include license file
- [ ] Create installer (optional)
- [ ] Sign executable (optional, for production)

---

## Creating Installer (Optional)

### Using WiX Toolset

1. **Install WiX Toolset**
   - Download: https://wixtoolset.org/
   - Install WiX v3.11 or later

2. **Create installer.wxs:**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
  <Product Id="*" Name="CorelDRAW Smart Fill Tool"
           Language="1033" Version="1.0.0.0"
           Manufacturer="Your Name" UpgradeCode="PUT-GUID-HERE">

    <Package InstallerVersion="200" Compressed="yes"
             InstallScope="perMachine" />

    <Media Id="1" Cabinet="product.cab" EmbedCab="yes" />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLFOLDER" Name="CorelSmartFill" />
      </Directory>
    </Directory>

    <ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
      <Component Id="MainExecutable">
        <File Source="bin\Release\net6.0-windows\CorelSmartFill.exe" />
      </Component>
    </ComponentGroup>

    <Feature Id="ProductFeature" Title="CorelSmartFill" Level="1">
      <ComponentGroupRef Id="ProductComponents" />
    </Feature>
  </Product>
</Wix>
```

3. **Build installer:**

```bash
candle installer.wxs
light installer.wixobj -out CorelSmartFill-Setup.msi
```

---

## Code Signing (Production)

### Why sign?

- Prevents "Unknown Publisher" warning
- Users trust signed applications
- Required for some enterprise environments

### How to sign:

```bash
# 1. Obtain code signing certificate
# - From: DigiCert, Sectigo, GlobalSign, etc.
# - Export as .pfx file

# 2. Sign the executable
signtool sign /f YourCertificate.pfx /p CertPassword /t http://timestamp.digicert.com /fd SHA256 CorelSmartFill.exe

# 3. Verify signature
signtool verify /pa CorelSmartFill.exe
```

---

## Common Build Commands Reference

```bash
# Clean build folders
dotnet clean

# Restore packages
dotnet restore

# Build Debug
dotnet build

# Build Release
dotnet build -c Release

# Run without building
dotnet run --no-build

# Publish self-contained single file
dotnet publish -c Release -r win-x64 --self-contained /p:PublishSingleFile=true

# Publish framework-dependent
dotnet publish -c Release -r win-x64 --self-contained false

# Pack as NuGet (if making library)
dotnet pack

# Run tests (if test project exists)
dotnet test
```

---

## Performance Optimization for Build

### Faster builds:

```xml
<!-- Add to .csproj -->
<PropertyGroup>
  <!-- Disable package restore on build (restore separately) -->
  <RestorePackages>false</RestorePackages>

  <!-- Parallel build -->
  <BuildInParallel>true</BuildInParallel>

  <!-- Skip XML docs generation (faster) -->
  <GenerateDocumentationFile>false</GenerateDocumentationFile>
</PropertyGroup>
```

### Build once, run many:

```bash
# Build
dotnet build -c Release

# Run multiple times without rebuilding
dotnet run --no-build -c Release
```

---

## Need Help?

### Build fails?

1. Check Visual Studio version (2022 or later)
2. Verify .NET 6 SDK installed
3. Clean and rebuild: `dotnet clean && dotnet build`
4. Delete bin/ and obj/ folders, rebuild
5. Check for typos in .csproj file

### Runtime errors?

1. Make sure target PC has .NET 6 Runtime
2. OR use self-contained build
3. Run from command line to see error messages
4. Check Windows Event Viewer for crashes

---

**Ready to build!** 🚀

Follow the Quick Build section at the top for fastest results.
