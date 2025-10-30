# SiteAgentWatcher Installation Script for Windows
# This script builds, deploys, and installs SiteAgentWatcher as a Windows service using NSSM

param(
    [string]$ServiceName = "SiteAgentWatcher",
    [string]$InstallPath = "C:\Program Files (x86)\SiteAgentWatcher",
    [string]$DisplayName = "SiteAgent Watcher Service",
    [string]$Description = "Monitors and manages hotel door lock services",
    [int]$Port = 7331,
    [switch]$SkipBuild = $false,
    [string]$NssmVersion = "2.24"
)

$ErrorActionPreference = "Stop"

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    throw "This script must be run as Administrator!"
}

Write-Host "===========================================================" -ForegroundColor Cyan
Write-Host "  SiteAgentWatcher Installation" -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan
Write-Host ""

# Get repository root
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not (Test-Path "$repoRoot\package.json")) {
    throw "Not in SiteAgentWatcher repository!"
}

# Step 1: Download latest release from S3
Write-Host "Step 1: Downloading latest release from S3..." -ForegroundColor Cyan
Write-Host ""

# S3 configuration (can be overridden via parameters later if needed)
$S3Bucket = "siteagent-watcher-prod"
$S3Region = "eu-north-1"

$latestUrl = "https://$S3Bucket.s3.$S3Region.amazonaws.com/latest.json"
Write-Host "Fetching latest version info: $latestUrl" -ForegroundColor Yellow

try {
    # Ensure TLS 1.2 for older PowerShell environments
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $latestJson = Invoke-WebRequest -Uri $latestUrl -UseBasicParsing -ErrorAction Stop
    $latest = $latestJson.Content | ConvertFrom-Json
    $latestVersion = $latest.version
    if (-not $latestVersion) { throw "latest.json does not contain 'version'" }
    Write-Host "Latest version on S3: $latestVersion" -ForegroundColor Green
}
catch {
    throw "Failed to download latest.json from S3: $_"
}

# Ensure installation directory exists early for downloads
if (-not (Test-Path $InstallPath)) {
    Write-Host "Creating: $InstallPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
}

# Download main executable (SiteAgentWatcher.exe) from S3
$siteAgentExeUrl = "https://$S3Bucket.s3.$S3Region.amazonaws.com/releases/$latestVersion/SiteAgentWatcher.exe"

Write-Host "Downloading SiteAgentWatcher.exe..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $siteAgentExeUrl -OutFile "$InstallPath\SiteAgentWatcher.exe" -UseBasicParsing -ErrorAction Stop
    Write-Host "Success: SiteAgentWatcher.exe downloaded" -ForegroundColor Green
}
catch {
    throw "Failed to download SiteAgentWatcher.exe from $siteAgentExeUrl: $_"
}

# Step 2: Build updater locally
Write-Host "Step 2: Building updater locally..." -ForegroundColor Cyan
Write-Host ""

try {
    $nodeVersion = node --version 2>&1
    Write-Host "Detected Node.js: $nodeVersion" -ForegroundColor Green
}
catch {
    Write-Error "Node.js not found! Install Node.js 18+ from https://nodejs.org"
    exit 1
}

Write-Host "Installing updater dependencies..." -ForegroundColor Yellow
Push-Location "$repoRoot\updater"
npm ci
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "Failed to install updater dependencies"
}

Write-Host "Compiling updater TypeScript..." -ForegroundColor Yellow
npm run build
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "Failed to build updater"
}

Write-Host "Packaging updater.exe..." -ForegroundColor Yellow
npm install -g pkg
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "Failed to install pkg for updater packaging"
}

pkg . --targets node18-win-x64 --output "$InstallPath\updater.exe"
if ($LASTEXITCODE -ne 0) {
    Pop-Location
    throw "Failed to package updater.exe"
}
Pop-Location
Write-Host "Success: updater.exe built locally" -ForegroundColor Green

# Step 3: Download and setup NSSM
Write-Host "Step 3: Setting up NSSM (Non-Sucking Service Manager)..." -ForegroundColor Cyan
$nssmPath = "$InstallPath\nssm.exe"

if (-not (Test-Path $nssmPath)) {
    Write-Host "Downloading NSSM $NssmVersion..." -ForegroundColor Yellow
    $nssmZip = "$env:TEMP\nssm-$NssmVersion.zip"
    $nssmUrl = "https://nssm.cc/release/nssm-$NssmVersion.zip"
    
    try {
        Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip -UseBasicParsing
        
        Write-Host "Extracting NSSM..." -ForegroundColor Yellow
        $nssmTemp = "$env:TEMP\nssm-$NssmVersion"
        Expand-Archive -Path $nssmZip -DestinationPath $nssmTemp -Force
        
        $nssmExe = Get-ChildItem -Path $nssmTemp -Recurse -Filter "nssm.exe" | Where-Object { $_.Directory.Name -eq "win64" } | Select-Object -First 1
        
        if (-not $nssmExe) {
    throw "Could not find nssm.exe in downloaded archive!"
        }
        
        if (-not (Test-Path $InstallPath)) {
            New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        }
        
        Copy-Item $nssmExe.FullName -Destination $nssmPath -Force
        Remove-Item $nssmZip -Force -ErrorAction SilentlyContinue
        Remove-Item $nssmTemp -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Host "Success: NSSM downloaded and installed" -ForegroundColor Green
    }
    catch {
    throw "Failed to download NSSM: $_"
    }
} else {
    Write-Host "NSSM already exists at $nssmPath" -ForegroundColor Gray
}

# Step 4: Prepare installation directory (already created if downloads ran)
Write-Host "Step 4: Preparing installation directory..." -ForegroundColor Cyan
if (-not (Test-Path $InstallPath)) {
    Write-Host "Creating: $InstallPath" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
} else {
    Write-Host "Install path exists: $InstallPath" -ForegroundColor Gray
}

Write-Host "Creating version.json..." -ForegroundColor Yellow
# Use version from latest.json
$version = $latestVersion
Write-Host "Using version from S3 latest.json: $version" -ForegroundColor Gray

$versionJson = @{
    version = $version
    updatedAt = (Get-Date).ToString("o")
} | ConvertTo-Json

# Write without BOM to avoid JSON parsing issues
[System.IO.File]::WriteAllText("$InstallPath\version.json", $versionJson, [System.Text.UTF8Encoding]::new($false))

Write-Host "Success: Executables and version file copied" -ForegroundColor Green

Write-Host "Creating required directories..." -ForegroundColor Yellow
$DataPath = "$InstallPath\data"
$LogsPath = "$InstallPath\logs"
$UpdatesPath = "$InstallPath\updates"
$BackupsPath = "$InstallPath\backups"

New-Item -ItemType Directory -Path $DataPath -Force | Out-Null
    New-Item -ItemType Directory -Path $LogsPath -Force | Out-Null
New-Item -ItemType Directory -Path $UpdatesPath -Force | Out-Null
New-Item -ItemType Directory -Path $BackupsPath -Force | Out-Null
Write-Host "Success: Directories created" -ForegroundColor Green

Write-Host "Copying Prisma files..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path "$InstallPath\prisma" -Force | Out-Null
Copy-Item "$repoRoot\prisma\schema.prisma" -Destination "$InstallPath\prisma\schema.prisma" -Force
Copy-Item "$repoRoot\prisma\migrations" -Destination "$InstallPath\prisma\migrations" -Recurse -Force
Write-Host "Success: Prisma files copied" -ForegroundColor Green

# Step 5: Configure environment
Write-Host ""
Write-Host "Step 5: Configuring environment..." -ForegroundColor Cyan
$EnvFile = "$InstallPath\.env"

# Always prepare the database path with forward slashes for DATABASE_URL
$dbPath = "$InstallPath/data/siteagent-watcher.db" -replace '\\', '/'

$envExists = Test-Path $EnvFile

if (-not $envExists) {
    Write-Host "Creating .env file with default configuration..." -ForegroundColor Yellow
    Write-Host ""
    
    # Prompt for API Key
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host "   API Key Configuration" -ForegroundColor Yellow
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please enter a secure API key for authentication." -ForegroundColor Cyan
    Write-Host "This key will be required for all API requests." -ForegroundColor Gray
    Write-Host "Leave empty to generate a random key automatically." -ForegroundColor Gray
    Write-Host ""
    
    $apiKey = Read-Host -Prompt "API Key (press Enter for auto-generate)"
    
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        # Generate random API key
        $apiKey = -join ((65..90) + (97..122) + (48..57) | Get-Random -Count 32 | ForEach-Object {[char]$_})
        Write-Host "Generated API Key: $apiKey" -ForegroundColor Green
        Write-Host "IMPORTANT: Save this key securely!" -ForegroundColor Yellow
    }
    else {
        Write-Host "Using provided API Key" -ForegroundColor Green
    }
    
    Write-Host ""
    
    # Create default .env content with all required variables and absolute paths
    $envContent = @"
# Server Configuration
PORT=$Port
WEBSOCKET_PORT=8081
NODE_ENV=production
API_VERSION=v1
API_PREFIX=/api

# Database
DATABASE_URL=file:$dbPath

# Service Monitoring
SERVICE_CHECK_INTERVAL=30000
SERVICE_RESTART_MAX_ATTEMPTS=3
SERVICE_RESTART_DELAY=5000

# SiteAgent Configuration
SITEAGENT_LOGS_PATH=C:\Program Files (x86)\Dialock\HMS\1.1\SiteAgent\logs

# Logging
LOG_LEVEL=info
LOG_DIR=./logs
LOG_MAX_SIZE=10m
LOG_MAX_FILES=7d

# Security
API_KEY=$apiKey
ALLOWED_IPS=*

# Windows Services to Monitor (comma-separated)
MONITORED_SERVICES=Dialock HMS Site Agent

# Auto-Update Configuration
AUTO_UPDATE_ENABLED=true
AUTO_UPDATE_CHECK_INTERVAL=3600000
AUTO_UPDATE_AUTO_INSTALL=false
AUTO_UPDATE_S3_BUCKET=siteagent-watcher-prod
AUTO_UPDATE_S3_REGION=eu-north-1
"@
    
    Set-Content -Path $EnvFile -Value $envContent -Encoding UTF8
    Write-Host "Success: .env created with absolute paths" -ForegroundColor Green
}
else {
    Write-Host "Success: .env file already exists" -ForegroundColor Green
}

Write-Host ""
Write-Host "Configuration file location: $EnvFile" -ForegroundColor Cyan
Write-Host ""

# Initialize database
Write-Host "Initializing database..." -ForegroundColor Yellow
Push-Location $InstallPath
$env:DATABASE_URL = "file:$dbPath"

try {
    npx prisma migrate deploy --schema=prisma/schema.prisma 2>&1 | Out-Host
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Success: Database initialized" -ForegroundColor Green
    }
    else {
        Write-Host "Note: Database will be initialized on first run by the application" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Note: Database will be initialized on first run by the application" -ForegroundColor Yellow
}
Pop-Location

# Step 6: Install Windows Service using NSSM
Write-Host ""
Write-Host "Step 6: Installing Windows Service with NSSM..." -ForegroundColor Cyan

# Check if service already exists
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  ERROR: Service Already Exists" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "A service named '$ServiceName' already exists." -ForegroundColor Yellow
    Write-Host "Please run the cleanup script first:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  .\scripts\cleanup.ps1" -ForegroundColor White
    Write-Host ""
    Write-Host "Then run this install script again." -ForegroundColor Yellow
    Write-Host ""
    throw "Service already exists"
}

Write-Host "Creating service with NSSM..." -ForegroundColor Yellow
$exePath = "$InstallPath\SiteAgentWatcher.exe"

# Install service with NSSM
$installResult = & $nssmPath install $ServiceName $exePath 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create service: $installResult"
}

# Configure service
& $nssmPath set $ServiceName DisplayName "$DisplayName"
& $nssmPath set $ServiceName Description "$Description"
& $nssmPath set $ServiceName Start SERVICE_AUTO_START

# Set working directory (this is the key feature of NSSM!)
& $nssmPath set $ServiceName AppDirectory $InstallPath

# Configure stdout and stderr logging
& $nssmPath set $ServiceName AppStdout "$LogsPath\service-stdout.log"
& $nssmPath set $ServiceName AppStderr "$LogsPath\service-stderr.log"

# Configure automatic restart on failure
# Exit code 0 = clean exit (e.g. for updates), don't restart
# Any other exit code = restart
& $nssmPath set $ServiceName AppExit 0 Exit
& $nssmPath set $ServiceName AppExit Default Restart
& $nssmPath set $ServiceName AppRestartDelay 5000

Write-Host "Success: Service created with NSSM" -ForegroundColor Green

Write-Host "Configuring firewall rules..." -ForegroundColor Yellow
try {
    $httpRule = Get-NetFirewallRule -DisplayName "$DisplayName HTTP" -ErrorAction SilentlyContinue
    if (-not $httpRule) {
        New-NetFirewallRule -DisplayName "$DisplayName HTTP" -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow -Profile Domain,Private -Description "Allow HTTP access" | Out-Null
    }
    
    $wsRule = Get-NetFirewallRule -DisplayName "$DisplayName WebSocket" -ErrorAction SilentlyContinue
    if (-not $wsRule) {
        New-NetFirewallRule -DisplayName "$DisplayName WebSocket" -Direction Inbound -Protocol TCP -LocalPort 8081 -Action Allow -Profile Domain,Private -Description "Allow WebSocket" | Out-Null
    }
    Write-Host "Success: Firewall rules configured" -ForegroundColor Green
}
catch {
    Write-Host "WARNING: Could not configure firewall rules" -ForegroundColor Yellow
}

# Step 7: Start service
Write-Host ""
Write-Host "Step 7: Starting service..." -ForegroundColor Cyan

Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

# Verify service status
$service = Get-Service -Name $ServiceName
Write-Host ""
Write-Host "===========================================================" -ForegroundColor Green

$isRunning = ($service.Status -eq "Running")
if ($isRunning) {
    Write-Host "SUCCESS: Installation successful!" -ForegroundColor Green
    Write-Host "===========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Service Status: Running" -ForegroundColor Green
    Write-Host "Installation Path: $InstallPath" -ForegroundColor White
    Write-Host "Configuration: $EnvFile" -ForegroundColor White
    Write-Host "Logs: $LogsPath" -ForegroundColor White
    Write-Host ""
    Write-Host "API Endpoints:" -ForegroundColor Cyan
    Write-Host "  Health: http://localhost:$Port/api/v1/health" -ForegroundColor White
    Write-Host "  Services: http://localhost:$Port/api/v1/services" -ForegroundColor White
    Write-Host "  Keys: http://localhost:$Port/api/v1/keys/generate" -ForegroundColor White
    Write-Host "  Logs: http://localhost:$Port/api/v1/logs/latest" -ForegroundColor White
    Write-Host ""
    Write-Host "Service is running and ready!" -ForegroundColor Green
    Write-Host "Test the API: curl http://localhost:$Port/api/v1/health" -ForegroundColor White
}

if (-not $isRunning) {
    Write-Host "WARNING: Installation completed but service not running" -ForegroundColor Yellow
    Write-Host "===========================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Service Status: $($service.Status)" -ForegroundColor Yellow
    Write-Host "Installation Path: $InstallPath" -ForegroundColor White
    Write-Host "Configuration File: $EnvFile" -ForegroundColor White
    Write-Host "Logs: $LogsPath" -ForegroundColor White
    Write-Host ""
    Write-Host "Next steps to start the service:" -ForegroundColor Cyan
    Write-Host "1. Edit configuration file and update required settings:" -ForegroundColor White
    Write-Host "   notepad $EnvFile" -ForegroundColor Gray
    Write-Host "   - Set API_KEY (for API authentication)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "2. After updating configuration, start the service:" -ForegroundColor White
    Write-Host "   Start-Service $ServiceName" -ForegroundColor Gray
    Write-Host ""
    Write-Host "3. Check service status:" -ForegroundColor White
    Write-Host "   Get-Service $ServiceName" -ForegroundColor Gray
    Write-Host ""
    Write-Host "4. View logs if service fails to start:" -ForegroundColor White
    Write-Host "   Get-Content $LogsPath\siteagent-watcher-*.log -Tail 50" -ForegroundColor Gray
    Write-Host ""
    Write-Host "5. Test the API once running:" -ForegroundColor White
    Write-Host "   curl http://localhost:$Port/api/v1/health" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Installation script completed." -ForegroundColor Cyan
