<#
.SYNOPSIS
  WinTuner Auto-Update Script (All-in-One)

.DESCRIPTION
  A single script that updates/installs WinTuner in a separate process, then relaunches itself.
  This avoids the "assembly already loaded" conflict by closing the session after updates.
  Once re-launched with -AlreadyUpdated, it proceeds directly to normal WinTuner logic.

.NOTES
  1) Must be run as Powershell 7
  2) Run as a standard user

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [string]$RootPackageFolder,

    [Parameter(Mandatory = $false)]
    [switch]$SilentMode,

    [Parameter(Mandatory = $false)]
    [switch]$UpdateAllApps,

    # This switch is used internally to detect if we've already performed the out-of-process update
    [Parameter(Mandatory = $false)]
    [switch]$AlreadyUpdated
)

# ------------------------------------------------------------
# 0. Determine This Script Path (Compatible with PS5)
# ------------------------------------------------------------
$scriptPath = $MyInvocation.MyCommand.Path

# ------------------------------------------------------------
# Early Check: Must be in PowerShell 7+
# ------------------------------------------------------------
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later."
    Write-Host "Would you like to attempt to launch this script with pwsh (PowerShell 7) if it's installed?"
    Write-Host "1. Yes"
    Write-Host "2. No (Exit)"
    $choice = Read-Host "Choice (1-2)"

    switch ($choice) {
        "1" {
            $pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
            if ($pwshPath) {
                Write-Host "Launching script with PowerShell 7..."
                Start-Process -FilePath $pwshPath -ArgumentList @("-File", $scriptPath) -NoNewWindow -Wait
                exit
            } else {
                Write-Host "[ERROR] PowerShell 7 not found on this system."
                Write-Host "Please install from:"
                Write-Host "  https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell"
                exit
            }
        }
        "2" {
            Write-Host "Exiting script. Please install PowerShell 7 from:";
            Write-Host "  https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell"
            exit
        }
        default {
            Write-Host "Invalid choice. Exiting..."
            exit
        }
    }
}

# ------------------------------------------------------------
# 1. Logging Function
# ------------------------------------------------------------
Function Write-Log {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$LogType = "INFO"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp, $LogType, $Message"

    if ([string]::IsNullOrWhiteSpace($CsvLogFile)) {
        Write-Host "[NO LOG FILE] $Message"
    } else {
        Add-Content -Path $CsvLogFile -Value $LogEntry
    }

    switch ($LogType) {
        "INFO"    { Write-Host "[INFO] $Message" -ForegroundColor Gray }
        "WARN"    { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "[ERROR] $Message" -ForegroundColor Red }
        "SUCCESS" { Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
        default   { Write-Host $Message }
    }
}

# ------------------------------------------------------------
# 2. Install/Update WinTuner (Out-of-Process) if not AlreadyUpdated
# ------------------------------------------------------------
if (-not $AlreadyUpdated) {
    if (-not $RootPackageFolder) {
        $RootPackageFolder = Read-Host "Please enter the path to the packaging directory"
    }

    Write-Host "[SETUP] Installing/updating WinTuner in a fresh session..."
    try {
        $updateArgs = @(
            "-NoProfile",
            "-Command",
            "if (-not (Get-InstalledModule -Name WinTuner -ErrorAction SilentlyContinue)) { Install-Module WinTuner -Scope CurrentUser -Force } else { Update-Module WinTuner -Scope CurrentUser -Force -ErrorAction SilentlyContinue }"
        )
        Start-Process pwsh -ArgumentList $updateArgs -Wait
    } catch {
        Write-Host "[SETUP ERROR] Failed to install/update WinTuner: $($_.Exception.Message)"
        exit
    }

    Write-Host "[SETUP] Relaunching this script in a fresh session..."

    $argList = @(
        "-NoProfile",
        "-File",
        $scriptPath
    )

    if ($RootPackageFolder) {
        $argList += "-RootPackageFolder"
        $argList += $RootPackageFolder
    }
    if ($SilentMode) {
        $argList += "-SilentMode"
    }
    if ($UpdateAllApps) {
        $argList += "-UpdateAllApps"
    }

    $argList += "-AlreadyUpdated"

    Start-Process pwsh -ArgumentList $argList -Wait
    Write-Host "[SETUP] Relaunched script finished. Exiting initial process..."
    exit
}

# ------------------------------------------------------------
# 3. Normal Script Logic (after -AlreadyUpdated)
# ------------------------------------------------------------

# 3A. Ensure Root Folder
if (-not $RootPackageFolder) {
    $RootPackageFolder = Read-Host "Please enter the path to the packaging directory"
}
if (-not (Test-Path $RootPackageFolder)) {
    $createFolder = if ($SilentMode) { "Y" } else { Read-Host "Directory does not exist. Create it? (Y/N)" }
    if ($createFolder -match "^[Yy]$") {
        New-Item -ItemType Directory -Path $RootPackageFolder -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[INFO] Created/confirmed directory: $RootPackageFolder"
    } else {
        Write-Host "[ERROR] Directory not created. Exiting."
        exit
    }
}

# 3B. CSV Log File
$CsvLogFile = Join-Path $RootPackageFolder "WinTuner_UpdateLog.csv"
Write-Log "Packaging folder set to: $RootPackageFolder" "INFO"

# 3D. Import WinTuner
try {
    Import-Module WinTuner -Force
    Write-Log "WinTuner module imported." "SUCCESS"
} catch {
    Write-Log "Failed to import WinTuner: $($_.Exception.Message)" "ERROR"
    exit
}

# 3E. Connect to WinTuner
Write-Log "Connecting to WinTuner..." "INFO"
try {
    Connect-WtWinTuner | Out-Null
    Write-Log "Connected to WinTuner." "SUCCESS"
} catch {
    Write-Log "Failed to connect: $($_.Exception.Message)" "ERROR"
    exit
}

# 3F. List Installed Apps
Write-Log "Listing installed applications..." "INFO"
$allApps = Get-WtWin32Apps -ErrorAction SilentlyContinue
if (!$allApps -or $allApps.Count -eq 0) {
    Write-Log "No installed apps found. Exiting." "WARN"
    exit
}
$i = 1
foreach ($app in $allApps) {
    $appName = $app.DisplayName ?? $app.Name ?? "Unknown App"
    Write-Host "$i. $appName (Version: $($app.CurrentVersion))"
    $i++
}

# 3G. Check for Updates
Write-Log "Checking for apps needing updates..." "INFO"
$updatedApps = Get-WtWin32Apps -Update $true -Superseded $false -ErrorAction SilentlyContinue
if ($updatedApps.Count -gt 0) {
    Write-Host "`nApps requiring updates:"
    $i = 1
    foreach ($app in $updatedApps) {
        $appName = $app.DisplayName ?? $app.Name ?? "Unknown App"
        Write-Host "$i. $appName (Current: $($app.CurrentVersion) -> Latest: $($app.LatestVersion))"
        $i++
    }
    if (-not $SilentMode) {
        Write-Host "`nOptions:";
        Write-Host "0. Exit";
        Write-Host "1. Update all";
        Write-Host "2. Select apps";
        Write-Host "3. Skip"

        $choice = Read-Host "Choice (0-3)"
        switch ($choice) {
            "0" { Write-Log "Exiting without changes."; exit }
            "1" {
                foreach ($app in $updatedApps) {
                    Write-Log "Updating: $($app.DisplayName)"

                    # Build intunewin
                    New-WtWingetPackage -PackageId $($app.PackageId) -PackageFolder $RootPackageFolder -Version $($app.LatestVersion)

                    # Deploy
                    Deploy-WtWin32App -PackageId $($app.PackageId) -Version $($app.LatestVersion) -RootPackageFolder $RootPackageFolder `
                                     -GraphId $($app.GraphId)

                    Write-Log "Updated: $($app.DisplayName) -> $($app.LatestVersion)"
                }
            }
            "2" {
                $selectedApps = Read-Host "Enter the numbers of the apps to update (comma-separated)"
                $selectedIndexes = $selectedApps -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
                foreach ($index in $selectedIndexes) {
                    if ($index -gt 0 -and $index -le $updatedApps.Count) {
                        $app = $updatedApps[$index - 1]
                        Write-Log "Updating: $($app.DisplayName)"

                        New-WtWingetPackage -PackageId $($app.PackageId) -PackageFolder $RootPackageFolder -Version $($app.LatestVersion)

                        Deploy-WtWin32App -PackageId $($app.PackageId) -Version $($app.LatestVersion) -RootPackageFolder $RootPackageFolder `
                                         -GraphId $($app.GraphId)

                        Write-Log "Updated: $($app.DisplayName) -> $($app.LatestVersion)"
                    } else {
                        Write-Log "Invalid selection: $index"
                    }
                }
            }
            "3" { Write-Log "Skipping updates."; break }
            default { Write-Log "Invalid choice. Exiting."; exit }
        }
    } else {
        if ($UpdateAllApps) {
            foreach ($app in $updatedApps) {
                Write-Log "Updating: $($app.DisplayName)"

                New-WtWingetPackage -PackageId $($app.PackageId) -PackageFolder $RootPackageFolder -Version $($app.LatestVersion)

                Deploy-WtWin32App -PackageId $($app.PackageId) -Version $($app.LatestVersion) -RootPackageFolder $RootPackageFolder `
                                 -GraphId $($app.GraphId)

                Write-Log "Updated: $($app.DisplayName) -> $($app.LatestVersion)"
            }
        } else {
            Write-Log "Silent mode, skipping updates."
        }
    }
} else {
    Write-Log "No apps need updating." "INFO"
}

# 3H. Search & Add Apps
if (-not $SilentMode) {
    # We'll allow multiple searches in a loop
    while ($true) {
        Write-Host "`nAdd new apps?"
        Write-Host "1. Yes"
        Write-Host "2. No (Finish)"
        $searchChoice = Read-Host "Choice (1-2)"
        if ($searchChoice -eq "2") {
            break
        } elseif ($searchChoice -ne "1") {
            Write-Host "Invalid choice. Please enter 1 or 2."
            continue
        }

        $searchTerm = Read-Host "Enter name or Package ID"
        Write-Log "Searching: $searchTerm" "INFO"

        $searchResults = Search-WtWinGetPackage -PackageId $searchTerm -ErrorAction SilentlyContinue
        if ($searchResults -and $searchResults.Count -gt 0) {
            Write-Host "`nFound:";
            $i=1
            foreach ($app in $searchResults) {
                Write-Host "$i. $($app.Name) (ID: $($app.PackageId))"
                $i++
            }

            $selectedAppIndex = Read-Host "Number to add (or Enter to cancel)"
            if ($selectedAppIndex -match "^\d+$") {
                $selectedApp = $searchResults[$selectedAppIndex - 1]
                Write-Log "Adding application: $($selectedApp.Name)" "INFO"

                # -----------------------------------------
                # Prompt for assignment (Assign vs. Require)
                # and target (AllUsers, AllDevices, group).
                # -----------------------------------------
                Write-Host ""
                Write-Host "Would you like to (1) Assign or (2) Require this new package?"
                $assignChoice = Read-Host "Select 1 or 2"

                Write-Host "Who should receive it?"
                Write-Host "1. All Users"
                Write-Host "2. All Devices"
                Write-Host "3. A specific group"
                $targetChoice = Read-Host "Select 1-3"

                [string[]]$targets = @()
                switch ($targetChoice) {
                  '1' { $targets = @("AllUsers") }
                  '2' { $targets = @("AllDevices") }
                  '3' {
                    $groupName = Read-Host "Enter the group name or ID"
                    $targets = @($groupName)
                  }
                  default {
                    Write-Host "Invalid choice. Defaulting to AllUsers..."
                    $targets = @("AllUsers")
                  }
                }

                # Build intunewin package with -PackageFolder
                New-WtWingetPackage -PackageId $selectedApp.PackageId -PackageFolder $RootPackageFolder -Version $selectedApp.Version

                # Deploy with -RootPackageFolder, using available vs. required
                Write-Log "Deploying package..." "INFO"
                try {
                    if ($assignChoice -eq '1') {
                        # Available
                        Deploy-WtWin32App -PackageId $selectedApp.PackageId -Version $selectedApp.Version -RootPackageFolder $RootPackageFolder `
                                          -GraphId $selectedApp.GraphId -AvailableFor $targets
                    } elseif ($assignChoice -eq '2') {
                        # Required
                        Deploy-WtWin32App -PackageId $selectedApp.PackageId -Version $selectedApp.Version -RootPackageFolder $RootPackageFolder `
                                          -GraphId $selectedApp.GraphId -RequiredFor $targets
                    } else {
                        Write-Host "Invalid choice. Defaulting to Assign AllUsers..."
                        Deploy-WtWin32App -PackageId $selectedApp.PackageId -Version $selectedApp.Version -RootPackageFolder $RootPackageFolder `
                                          -GraphId $selectedApp.GraphId -AvailableFor @("AllUsers")
                    }
                    Write-Log "Added: $($selectedApp.Name)" "SUCCESS"
                } catch {
                    Write-Log "Failed to deploy: $($_.Exception.Message)" "ERROR"
                }
            } else {
                Write-Log "No app selected." "WARN"
            }
        } else {
            Write-Log "No matches." "WARN"
        }
    }
}

Write-Log "Script execution complete." "SUCCESS"
Write-Host "[LOG] Log saved to: $CsvLogFile"
