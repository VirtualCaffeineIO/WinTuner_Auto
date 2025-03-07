# WinTuner Auto-Update Script
# Author: [Your Name]
# Website: [Your Website]
# Date: [MM/DD/YYYY]
# Version: 4.0
# Description: This script automates application updates and management using WinTuner in Microsoft Intune.
# Warniing - Not ready for production, still being developed
# Credits: Special thanks to the open-source community and contributors who made this possible.
# ------------------------------------------------------------

Function Write-Log {
    Param ([string]$Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp, $Message"
    Add-Content -Path $CsvLogFile -Value $LogEntry
    Write-Host $Message
}

# Prompt user to specify the packaging directory
$rootPackageFolder = Read-Host "Please enter the path to the packaging directory"
if (-not (Test-Path $rootPackageFolder)) {
    $createFolder = Read-Host "Directory does not exist. Would you like to create it? (Y/N)"
    if ($createFolder -match "^[Yy]$") {
        New-Item -ItemType Directory -Path $rootPackageFolder | Out-Null
        Write-Host "[SUCCESS] Created directory: $rootPackageFolder"
    } else {
        Write-Host "[EXIT] Directory not created. Exiting script."
        exit
    }
}

# Set logging directory
$CsvLogFile = Join-Path -Path $rootPackageFolder -ChildPath "WinTuner_UpdateLog.csv"

# Ensure PowerShell 7 is running
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Log "[WARNING] Not running in PowerShell 7. Relaunching in PowerShell 7..."
    $pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if ($pwshPath) {
        Start-Process -FilePath $pwshPath -ArgumentList "-File `"$PSCommandPath`"" -NoNewWindow -Wait
        exit
    } else {
        Write-Log "[ERROR] PowerShell 7 is not installed. Exiting script."
        exit
    }
}

# Verify if WinTuner is installed
if (-not (Get-Module -ListAvailable -Name WinTuner)) {
    Write-Log "[WARNING] WinTuner is not installed. Installing now..."
    try {
        Install-Module -Name WinTuner -Force -Scope CurrentUser
        Write-Log "[SUCCESS] WinTuner installed successfully."
    } catch {
        Write-Log "[ERROR] Failed to install WinTuner. Exiting script."
        exit
    }
}

# Ensure WinTuner module is up to date
Write-Log "[INFO] Checking for WinTuner updates..."
try {
    Import-Module WinTuner -Force
    Update-Module -Name WinTuner -Force -Scope CurrentUser -ErrorAction SilentlyContinue
    Write-Log "[SUCCESS] WinTuner is updated to the latest version."
} catch {
    Write-Log "[ERROR] Failed to update WinTuner. Proceeding with the current version."
}

# Connect to WinTuner (ensure authentication)
Write-Log "[CONNECT] Connecting to WinTuner..."
try {
    Connect-WtWinTuner | Out-Null
    Write-Log "[SUCCESS] Connected to WinTuner successfully."
} catch {
    Write-Log "[ERROR] Error: Failed to connect to WinTuner. Exiting script."
    exit
}

# Get installed applications in Intune
Write-Log "[INFO] Listing all installed applications..."
$allApps = Get-WtWin32Apps

if ($allApps.Count -gt 0) {
    Write-Host ""
    Write-Host "Installed Applications:"
    $i = 1
    foreach ($app in $allApps) {
        $appName = if ($app.DisplayName) { $app.DisplayName } elseif ($app.Name) { $app.Name } else { "Unknown App" }
        Write-Host "$i. $appName (Version: $($app.CurrentVersion))"
        $i++
    }
} else {
    Write-Log "[NOTICE] No installed applications found."
    exit
}

# Get applications that need updating
Write-Log "[INFO] Checking for applications that need updating..."
$updatedApps = Get-WtWin32Apps -Update $true -Superseded $false

if ($updatedApps.Count -gt 0) {
    Write-Host ""
    Write-Host "Applications that require updates:"
    $i = 1
    foreach ($app in $updatedApps) {
        $appName = if ($app.DisplayName) { $app.DisplayName } elseif ($app.Name) { $app.Name } else { "Unknown App" }
        Write-Host "$i. $appName (Current: $($app.CurrentVersion) → Latest: $($app.LatestVersion))"
        $i++
    }
    
    Write-Host ""
    Write-Host "Options:"
    Write-Host "0. Exit without making any changes"
    Write-Host "1. Update all applications"
    Write-Host "2. Select specific applications to update"
    Write-Host "3. Skip updates and exit"
    
    $choice = Read-Host "Enter your choice (0-3)"
    
    if ($choice -eq "0") {
        Write-Log "[EXIT] Exiting script without making any changes."
        exit
    } elseif ($choice -eq "3") {
        Write-Log "[NOTICE] Skipping updates and exiting script."
        exit
    } elseif ($choice -eq "1") {
        foreach ($app in $updatedApps) { 
            Write-Log "[UPDATE] Updating application: $($app.DisplayName)"
            New-WtWingetPackage -PackageId $($app.PackageId) -PackageFolder $rootPackageFolder -Version $($app.LatestVersion) | Deploy-WtWin32App -GraphId $($app.GraphId) 
            Write-Log "[SUCCESS] Updated: $($app.DisplayName) to version $($app.LatestVersion)"
        }
    } elseif ($choice -eq "2") {
        $selectedApps = Read-Host "Enter the numbers of the applications to update (comma-separated)"
        $selectedIndexes = $selectedApps -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
        
        foreach ($index in $selectedIndexes) {
            $app = $updatedApps[$index - 1]
            if ($app) {
                Write-Log "[UPDATE] Updating application: $($app.DisplayName)"
                New-WtWingetPackage -PackageId $($app.PackageId) -PackageFolder $rootPackageFolder -Version $($app.LatestVersion) | Deploy-WtWin32App -GraphId $($app.GraphId) 
                Write-Log "[SUCCESS] Updated: $($app.DisplayName) to version $($app.LatestVersion)"
            }
        }
    }
} else {
    Write-Log "[NOTICE] No applications needed updating."
}

Write-Host ""
Write-Host "Would you like to search for and add new applications to Intune?"
Write-Host "(This happens before updates are performed)"
Write-Host "1. Yes, search for new apps"
Write-Host "2. No, exit"

$searchChoice = Read-Host "Enter your choice (1-2)"

if ($searchChoice -eq "1") {
    $searchTerm = Read-Host "Enter the name or Package ID of the application to search for"
    Write-Log "[INFO] Searching for applications matching: $searchTerm"
    
    $searchResults = Search-WtWinGetPackage -PackageId $searchTerm
    
    if ($searchResults.Count -gt 0) {
        Write-Host "`nFound applications:"
        $i = 1
        foreach ($app in $searchResults) {
            Write-Host "$i. $($app.Name) (Package ID: $($app.PackageId))"
            $i++
        }
        
        $selectedAppIndex = Read-Host "Enter the number of the application to add (or press Enter to cancel)"
        
        if ($selectedAppIndex -match "^\d+$") {
            $selectedApp = $searchResults[$selectedAppIndex - 1]
            Write-Log "[INFO] Adding application: $($selectedApp.Name)"
            
            New-WtWingetPackage -PackageId $selectedApp.PackageId -PackageFolder $rootPackageFolder -Version $selectedApp.Version | Deploy-WtWin32App -GraphId $selectedApp.GraphId
            Write-Log "[SUCCESS] Added application: $($selectedApp.Name)"
        } else {
            Write-Log "[NOTICE] No application selected. Skipping new app addition."
        }
    } else {
        Write-Log "[NOTICE] No applications found matching search criteria."
    }
}

Write-Log "[SUCCESS] Script execution completed."
Write-Host "[LOG] Log saved to: $CsvLogFile"
