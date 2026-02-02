<#
.SYNOPSIS
    Installs dependencies for Hermes EML Converter on Windows.
    
.DESCRIPTION
    This script automates the installation of:
    1. Python 3 (if missing)
    2. LibreOffice (via Winget)
    3. GTK3 Runtime (Required for WeasyPrint)
    4. Poppler (Required for pdf2image)
    5. Python packages defined in pyproject.toml

.NOTES
    Requires Administrative Privileges for some installers.
#>

$ErrorActionPreference = "Stop"

function Test-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Warning "This script requires Administrator privileges to install system components."
    Write-Warning "Please right-click the script and select 'Run as Administrator'."
    exit
}

Write-Host "Starting Hermes Installation..." -ForegroundColor Cyan

# 0. Configure Proxy (if set in Env Vars)
$proxyEnv = $env:HTTP_PROXY
if ([string]::IsNullOrWhiteSpace($proxyEnv)) {
    $proxyEnv = $env:HTTPS_PROXY
}

if (-not [string]::IsNullOrWhiteSpace($proxyEnv)) {
    Write-Host "`n[0/5] Configuring Proxy from Environment Variable..." -ForegroundColor Yellow
    Write-Host "Using Proxy: $proxyEnv"
    try {
        $proxy = New-Object System.Net.WebProxy($proxyEnv)
        $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
        [System.Net.WebRequest]::DefaultWebProxy = $proxy
        Write-Host "Proxy Configured." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to configure proxy: $_"
    }
}

# 1. System Dependencies
Write-Host "`n[1/5] Checking System Dependencies..." -ForegroundColor Yellow

# Function to check and install via winget
function Install-WingetApp ($Id, $Name) {
    Write-Host "Checking for $Name..." -NoNewline
    $check = winget list -e --id $Id 2>$null
    if ($check) {
        Write-Host " Installed." -ForegroundColor Green
    }
    else {
        Write-Host " Not found. Installing..." -ForegroundColor Yellow
        winget install -e --id $Id --accept-package-agreements --accept-source-agreements
    }
}

$hasWinget = (Get-Command winget -ErrorAction SilentlyContinue)

# Function to install LibreOffice directly if Winget is missing
function Install-LibreOffice-Direct {
    Write-Host "Winget not found. Downloading LibreOffice MSI..." -ForegroundColor Yellow
    $loUserVersion = "24.2.0"
    $loUrl = "https://download.documentfoundation.org/libreoffice/stable/7.6.4/win/x86_64/LibreOffice_7.6.4_Win_x86-64.msi"
    # Using a 7.6.4 stable link as example or find a permalink.
    # To ensure it works, let's use a known stable version.
    $loUrl = "https://downloadarchive.documentfoundation.org/libreoffice/old/7.6.4.1/win/x86_64/LibreOffice_7.6.4.1_Win_x86-64.msi"
    
    $loInstaller = "$env:TEMP\LibreOffice.msi"
    try {
        Invoke-WebRequest -Uri $loUrl -OutFile $loInstaller
        Write-Host "Installing LibreOffice (this may take a while)..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$loInstaller`" /qn /norestart" -Wait
        Write-Host "LibreOffice installed." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install LibreOffice directly: $_"
    }
}

# Function to install Python directly if Winget is missing
function Install-Python-Direct {
    Write-Host "Winget not found. Downloading Python Installer..." -ForegroundColor Yellow
    $pyUrl = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"
    $pyInstaller = "$env:TEMP\python_installer.exe"
    try {
        Invoke-WebRequest -Uri $pyUrl -OutFile $pyInstaller
        Write-Host "Installing Python..."
        Start-Process -FilePath $pyInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
        Write-Host "Python installed." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install Python directly: $_"
    }
}

# Install LibreOffice
if ($hasWinget) {
    Install-WingetApp "LibreOffice.LibreOffice" "LibreOffice"
}
else {
    # Check if installed checks
    if (-not (Test-Path "$env:ProgramFiles\LibreOffice\program\soffice.exe")) {
        Install-LibreOffice-Direct
    }
    else {
        Write-Host "LibreOffice already installed." -ForegroundColor Green
    }
}

# Install Python if needed
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    if ($hasWinget) {
        Install-WingetApp "Python.Python.3.11" "Python 3.11"
    }
    else {
        Install-Python-Direct
    }
    # Update Path for current session
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

# 2. GTK3 Runtime (Direct Download)
Write-Host "`n[2/5] Installing GTK3 Runtime..." -ForegroundColor Yellow
$gtkPath = "C:\Program Files\GTK3-Runtime Win64\bin\gtk-3.dll"
#if (-not (Test-Path $gtkPath)) {
if (-not (1 -eq 1)) {
    $gtkUrl = "https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases/download/2022-01-04/gtk3-runtime-3.24.31-2022-01-04-ts-win64.exe"
    $gtkInstaller = "$env:TEMP\gtk3-installer.exe"
    
    Write-Host "Downloading GTK3 installer..."
    Invoke-WebRequest -Uri $gtkUrl -OutFile $gtkInstaller
    
    Write-Host "Installing GTK3..."
    Start-Process -FilePath $gtkInstaller -ArgumentList "/S" -Wait
    
    # Add to PATH
    $gtkBin = "C:\Program Files\GTK3-Runtime Win64\bin"
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -notlike "*$gtkBin*") {
        [Environment]::SetEnvironmentVariable("Path", "$currentPath;$gtkBin", "Machine")
        Write-Host "Added GTK3 to PATH." -ForegroundColor Green
    }
}
else {
    Write-Host "GTK3 already installed." -ForegroundColor Green
}

# 3. Poppler (Direct Download)
Write-Host "`n[3/5] Installing Poppler..." -ForegroundColor Yellow
$popplerPfile = "C:\ProgramData\Poppler"
#if (-not (Test-Path "$popplerPfile\Library\bin\pdfinfo.exe")) {
if (-not (1 -eq 1)) {
    $popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    $popplerZip = "$env:TEMP\poppler.zip"
    
    Write-Host "Downloading Poppler..."
    Invoke-WebRequest -Uri $popplerUrl -OutFile $popplerZip
    
    Write-Host "Extracting Poppler to $popplerPfile..."
    Expand-Archive -Path $popplerZip -DestinationPath "$env:TEMP\poppler_extract" -Force
    
    # Move the inner folder
    $extracted = Get-ChildItem "$env:TEMP\poppler_extract" | Select-Object -First 1
    if (Test-Path $popplerPfile) { Remove-Item $popplerPfile -Recurse -Force }
    Move-Item $extracted.FullName $popplerPfile
    
    # Add to PATH
    $popplerBin = "$popplerPfile\Library\bin"
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -notlike "*$popplerBin*") {
        [Environment]::SetEnvironmentVariable("Path", "$currentPath;$popplerBin", "Machine")
        Write-Host "Added Poppler to PATH." -ForegroundColor Green
    }
}
else {
    Write-Host "Poppler already installed." -ForegroundColor Green
}

# 4. Python Requirements
Write-Host "`n[4/5] Installing Python Packages..." -ForegroundColor Yellow
try {
    # Upgrade pip
    python -m pip install --upgrade pip
    
    # Install from pyproject.toml
    # Use -e . to install in editable mode which installs deps from pyproject.toml
    python -m pip install -e .
    Write-Host "Python dependencies installed successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to install python dependencies."
}

Write-Host "`n[5/5] Installation Complete!" -ForegroundColor Cyan
Write-Host "Please restart your terminal/powershell to ensure PATH changes take effect."
Write-Host "You can verify installation by running: python -c 'import weasyprint; import pdf2image; print(""Success"")'"
