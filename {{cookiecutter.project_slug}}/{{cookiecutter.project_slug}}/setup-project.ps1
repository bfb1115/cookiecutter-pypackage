# PowerShell script for cookiecutter template
# This script sets up a Python project with venv and NSSM service

param(
    [Parameter(Mandatory=$false)]
    [string]$ProjectSlug = "{{cookiecutter.project_slug}}"
)

# Configuration
$BaseDir = "C:\automation"
$ProjectDir = Join-Path $BaseDir $ProjectSlug
$VenvDir = Join-Path $ProjectDir "venv"
$MainScript = Join-Path $ProjectDir "main.py"
$RequirementsFile = Join-Path $ProjectDir "requirements.txt"

Write-Host "Setting up project: $ProjectSlug" -ForegroundColor Green
Write-Host "Project directory: $ProjectDir" -ForegroundColor Yellow

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check for administrator privileges
if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator to install Windows services."
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

# Create base automation directory if it doesn't exist
if (-not (Test-Path $BaseDir)) {
    Write-Host "Creating base directory: $BaseDir" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null
}

# Create project directory if it doesn't exist
if (-not (Test-Path $ProjectDir)) {
    Write-Host "Creating project directory: $ProjectDir" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $ProjectDir -Force | Out-Null
}

# Check if Python is available
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Found Python: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Error "Python is not installed or not in PATH. Please install Python first."
    exit 1
}

# Create virtual environment
Write-Host "Creating virtual environment..." -ForegroundColor Cyan
if (Test-Path $VenvDir) {
    Write-Host "Virtual environment already exists. Removing old venv..." -ForegroundColor Yellow
    Remove-Item -Path $VenvDir -Recurse -Force
}

python -m venv $VenvDir
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to create virtual environment"
    exit 1
}

# Activate virtual environment and install requirements
Write-Host "Activating virtual environment and installing requirements..." -ForegroundColor Cyan
$ActivateScript = Join-Path $VenvDir "Scripts\Activate.ps1"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$PythonwExe = Join-Path $VenvDir "Scripts\pythonw.exe"

if (Test-Path $RequirementsFile) {
    & $ActivateScript
    & $PythonExe -m pip install --upgrade pip
    & $PythonExe -m pip install -r $RequirementsFile
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install requirements"
        exit 1
    }
    Write-Host "Requirements installed successfully" -ForegroundColor Green
} else {
    Write-Host "No requirements.txt found, skipping package installation" -ForegroundColor Yellow
}

# Check if main.py exists
if (-not (Test-Path $MainScript)) {
    Write-Warning "main.py not found at $MainScript"
    Write-Host "You'll need to create main.py before the service can run properly" -ForegroundColor Yellow
}

# Check if NSSM is available
try {
    $nssmVersion = nssm version 2>&1
    Write-Host "Found NSSM: $nssmVersion" -ForegroundColor Green
} catch {
    Write-Error "NSSM is not installed or not in PATH."
    Write-Host "Please install NSSM from https://nssm.cc/download or use chocolatey: choco install nssm" -ForegroundColor Yellow
    exit 1
}

# Install service with NSSM
$ServiceName = $ProjectSlug
Write-Host "Installing Windows service: $ServiceName" -ForegroundColor Cyan

# Remove existing service if it exists
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Removing existing service: $ServiceName" -ForegroundColor Yellow
    nssm stop $ServiceName
    nssm remove $ServiceName confirm
}

# Install new service
nssm install $ServiceName $PythonwExe
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install service with NSSM"
    exit 1
}

# Configure service parameters
nssm set $ServiceName AppDirectory $ProjectDir
nssm set $ServiceName AppParameters $MainScript
# Convert project_slug to PascalCase for display name
$DisplayName = (Get-Culture).TextInfo.ToTitleCase($ProjectSlug.Replace('_', ' ')).Replace(' ', '')
nssm set $ServiceName DisplayName $DisplayName
nssm set $ServiceName Description "{{cookiecutter.project_short_description}}"
nssm set $ServiceName Start SERVICE_AUTO_START

# Set up logging
$LogDir = Join-Path $ProjectDir "logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

nssm set $ServiceName AppStdout (Join-Path $LogDir "stdout.log")
nssm set $ServiceName AppStderr (Join-Path $LogDir "stderr.log")
nssm set $ServiceName AppRotateFiles 1
nssm set $ServiceName AppRotateOnline 1
nssm set $ServiceName AppRotateSeconds 86400
nssm set $ServiceName AppRotateBytes 1048576

Write-Host "Service '$ServiceName' installed successfully!" -ForegroundColor Green
Write-Host "Service configuration:" -ForegroundColor Cyan
Write-Host "  - Executable: $PythonwExe" -ForegroundColor White
Write-Host "  - Script: $MainScript" -ForegroundColor White
Write-Host "  - Working Directory: $ProjectDir" -ForegroundColor White
Write-Host "  - Logs Directory: $LogDir" -ForegroundColor White

Write-Host "`nTo manage the service:" -ForegroundColor Cyan
Write-Host "  Start:   nssm start $ServiceName" -ForegroundColor White
Write-Host "  Stop:    nssm stop $ServiceName" -ForegroundColor White
Write-Host "  Status:  nssm status $ServiceName" -ForegroundColor White
Write-Host "  Remove:  nssm remove $ServiceName confirm" -ForegroundColor White

Write-Host "`nSetup completed successfully!" -ForegroundColor Green

# Pause so user can see the results
Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")