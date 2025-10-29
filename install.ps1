#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Bootstrap installer for SiteAgentWatcher service
.DESCRIPTION
    This script automates the complete installation process:
    - Installs/upgrades Node.js (if needed)
    - Downloads and extracts the pre-built SiteAgentWatcher repository ZIP from a public URL
    - Runs the service installation script
.EXAMPLE
    irm https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_PUBLIC_REPO/main/bootstrap-siteagent.ps1 | iex
#>

$ErrorActionPreference = "Stop"

# Configuration
$NodeJSVersion = "18.20.5"
$NodeJSUrl = "https://nodejs.org/dist/v$NodeJSVersion/node-v$NodeJSVersion-x64.msi"
$ArchiveUrl = "https://oleh-bucket.s3.eu-north-1.amazonaws.com/dh-doorlocks-siteagent-watcher-main.zip"  # Public S3 URL for the ZIP
$CloneDirectory = "C:\SiteAgentWatcher-Repo"
$TempDir = "$env:TEMP\SiteAgentWatcher-Bootstrap"
$ZipPath = Join-Path $TempDir "repo.zip"

# Colors for output
$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"
$ColorInfo = "Cyan"

# Ensure we're running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Please restart PowerShell as Administrator and try again."
    exit 1
}

# Set execution policy for this session
Set-ExecutionPolicy Bypass -Scope Process -Force

# Create temp directory
if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

Write-Host ""
Write-Host "===========================================================" -ForegroundColor $ColorInfo
Write-Host "   SiteAgentWatcher Bootstrap Installer" -ForegroundColor $ColorInfo
Write-Host "===========================================================" -ForegroundColor $ColorInfo
Write-Host ""

#region Functions

function Get-NodeJSVersion {
    try {
        $version = node --version 2>$null
        if ($version) {
            return $version -replace 'v', ''
        }
    }
    catch {
        return $null
    }
    return $null
}

function Get-NodeJSMajorVersion {
    $version = Get-NodeJSVersion
    if ($version) {
        return [int]($version.Split('.')[0])
    }
    return 0
}

function Install-NodeJS {
    param (
        [string]$Action = "Install"
    )
    
    Write-Host "[$Action Node.js $NodeJSVersion]" -ForegroundColor $ColorInfo
    
    $installerPath = Join-Path $TempDir "nodejs.msi"
    
    try {
        Write-Host "  Downloading Node.js installer..." -ForegroundColor Gray
        Invoke-WebRequest -Uri $NodeJSUrl -OutFile $installerPath -UseBasicParsing
        
        Write-Host "  Installing Node.js (this may take a few minutes)..." -ForegroundColor Gray
        $arguments = "/i `"$installerPath`" /quiet /norestart"
        Start-Process msiexec.exe -ArgumentList $arguments -Wait -NoNewWindow
        
        Write-Host "  Success: Node.js installed successfully" -ForegroundColor $ColorSuccess
        
        # Cleanup
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        Write-Host "  Failed: Failed to install Node.js: $_" -ForegroundColor $ColorError
        return $false
    }
}

function Refresh-EnvironmentPath {
    Write-Host "[Refreshing Environment Variables]" -ForegroundColor $ColorInfo
    
    try {
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + 
                    [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        Write-Host "  Success: Environment variables refreshed" -ForegroundColor $ColorSuccess
    }
    catch {
        Write-Host "  Warning: Could not refresh environment variables" -ForegroundColor $ColorWarning
    }
}

function Download-AndExtract-Repository {
    Write-Host ""
    Write-Host "[Downloading and Extracting SiteAgentWatcher Repository]" -ForegroundColor $ColorInfo
    
    try {
        # Remove existing directory if it exists
        if (Test-Path $CloneDirectory) {
            Write-Host "  Removing existing directory..." -ForegroundColor Gray
            Remove-Item -Path $CloneDirectory -Recurse -Force
        }
        
        Write-Host "  Downloading archive from public URL..." -ForegroundColor Gray
        Invoke-WebRequest -Uri $ArchiveUrl -OutFile $ZipPath -UseBasicParsing
        
        Write-Host "  Extracting archive..." -ForegroundColor Gray
        Expand-Archive -Path $ZipPath -DestinationPath $CloneDirectory -Force
        
        # Archives often extract to a subfolder (e.g., dh-doorlocks-siteagent-watcher-main); move contents up
        $extractedSubfolder = Get-ChildItem $CloneDirectory | Select-Object -First 1
        if ($extractedSubfolder) {
            Get-ChildItem (Join-Path $CloneDirectory $extractedSubfolder.Name) | Move-Item -Destination $CloneDirectory -Force
            Remove-Item (Join-Path $CloneDirectory $extractedSubfolder.Name) -Recurse -Force
        }
        
        Write-Host "  Success: Repository extracted to $CloneDirectory" -ForegroundColor $ColorSuccess
        
        # Cleanup
        Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        Write-Host "  Failed: Failed to download/extract repository: $_" -ForegroundColor $ColorError
        Write-Host "Troubleshooting:" -ForegroundColor $ColorWarning
        Write-Host "  1. Verify the public URL is accessible" -ForegroundColor $ColorWarning
        Write-Host "  2. Check your network connection" -ForegroundColor $ColorWarning
        Write-Host "  3. Download manually and place in $CloneDirectory" -ForegroundColor $ColorWarning
        return $false
    }
}

#endregion

#region Main Execution

try {
    # Step 1: Check and install/upgrade Node.js
    Write-Host ""
    Write-Host "Step 1: Checking Node.js installation..." -ForegroundColor $ColorInfo
    Write-Host ""
    
    $currentMajorVersion = Get-NodeJSMajorVersion
    
    if ($currentMajorVersion -eq 0) {
        Write-Host "  Node.js is not installed" -ForegroundColor $ColorWarning
        $installed = Install-NodeJS -Action "Install"
        if (-not $installed) {
            throw "Failed to install Node.js"
        }
        Refresh-EnvironmentPath
    }
    elseif ($currentMajorVersion -lt 18) {
        Write-Host "  Node.js v$currentMajorVersion detected (< 18)" -ForegroundColor $ColorWarning
        Write-Host "  Upgrading to Node.js v18..." -ForegroundColor $ColorWarning
        $installed = Install-NodeJS -Action "Upgrade"
        if (-not $installed) {
            throw "Failed to upgrade Node.js"
        }
        Refresh-EnvironmentPath
    }
    else {
        $currentVersion = Get-NodeJSVersion
        Write-Host "  Success: Node.js v$currentVersion is already installed (>= 18)" -ForegroundColor $ColorSuccess
    }
    
    # Step 2: Download and extract repository (no GitLab or auth needed)
    Write-Host ""
    Write-Host "Step 2: Downloading and extracting repository..." -ForegroundColor $ColorInfo
    $downloaded = Download-AndExtract-Repository
    
    if (-not $downloaded) {
        throw "Failed to download and extract repository"
    }
    
    # Step 3: Run installation script
    Write-Host ""
    Write-Host "Step 3: Running SiteAgentWatcher installation script..." -ForegroundColor $ColorInfo
    Write-Host ""
    
    $installScriptPath = Join-Path $CloneDirectory "scripts\install.ps1"
    
    if (-not (Test-Path $installScriptPath)) {
        throw "Installation script not found at: $installScriptPath"
    }
    
    Write-Host "  Executing install.ps1..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "===========================================================" -ForegroundColor $ColorInfo
    Write-Host "   Starting SiteAgentWatcher Installation" -ForegroundColor $ColorInfo
    Write-Host "===========================================================" -ForegroundColor $ColorInfo
    Write-Host ""
    
    # Change to scripts directory and run install.ps1
    Push-Location (Join-Path $CloneDirectory "scripts")
    try {
        & .\install.ps1
    }
    finally {
        Pop-Location
    }
    
    Write-Host ""
    Write-Host "===========================================================" -ForegroundColor $ColorSuccess
    Write-Host "   Bootstrap Installation Completed Successfully!" -ForegroundColor $ColorSuccess
    Write-Host "===========================================================" -ForegroundColor $ColorSuccess
    Write-Host ""
    Write-Host "Repository location: $CloneDirectory" -ForegroundColor $ColorInfo
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "===========================================================" -ForegroundColor $ColorError
    Write-Host "   Bootstrap Installation Failed" -ForegroundColor $ColorError
    Write-Host "===========================================================" -ForegroundColor $ColorError
    Write-Host ""
    Write-Host "Error: $_" -ForegroundColor $ColorError
    Write-Host ""
    Write-Host "Please check the error message above and try again." -ForegroundColor $ColorWarning
    Write-Host "If the problem persists, contact your administrator." -ForegroundColor $ColorWarning
    Write-Host ""
    exit 1
}
finally {
    # Cleanup temp directory
    if (Test-Path $TempDir) {
        Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#endregion
