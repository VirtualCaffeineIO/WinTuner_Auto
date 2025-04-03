# WinTuner Auto-Update Script
# Author: Virtual Caffeine IO
# Website: virtualcaffeine.io
# Version: 4.2

Function Write-Log {
    Param ([string]$Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp, $Message"
    if ($global:CsvLogFile) {
        Add-Content -Path $global:CsvLogFile -Value $LogEntry
    }
    Write-Host $Message
}

Function Cleanup-SupersededApps {
    $oldApps = Get-WtWin32Apps -Superseded $true
    if (-not $oldApps -or $oldApps.Count -eq 0) {
        Write-Log "[INFO] No superseded apps to clean up."
        return
    }

    Write-Host "`nSuperseded apps:"
    $i = 1
    foreach ($app in $oldApps) {
        $appName = if ($app.DisplayName) { $app.DisplayName } elseif ($app.Name) { $app.Name } else { "Unknown App" }
        Write-Host "$i. $appName ($($app.CurrentVersion))"
        $i++
    }

    Write-Host "`nOptions:"
    Write-Host "1. Remove all"
    Write-Host "2. Choose specific"
    Write-Host "3. Cancel"
    $choice = Read-Host "Enter choice (1-3)"

    if ($choice -eq "1") {
        foreach ($app in $oldApps) {
            if ($app.GraphId) {
                Remove-WtWin32App -AppId $app.GraphId
                Write-Log "[CLEANUP] Removed $($app.DisplayName)"
            } else {
                Write-Log "[SKIPPED] No GraphId found for $($app.DisplayName), skipping."
            }
        }
    } elseif ($choice -eq "2") {
        $selection = Read-Host "Enter numbers (comma-separated)"
        $indexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
        foreach ($index in $indexes) {
            $app = $oldApps[$index - 1]
            if ($app.GraphId) {
                Remove-WtWin32App -AppId $app.GraphId
                Write-Log "[CLEANUP] Removed $($app.DisplayName)"
            } else {
                Write-Log "[SKIPPED] No GraphId found for $($app.DisplayName), skipping."
            }
        }
    } else {
        Write-Log "[INFO] Cleanup cancelled."
    }
}

Function Show-MainMenu {
    do {
        Write-Host "`nWhat would you like to do next?"
        Write-Host "1. Update existing apps"
        Write-Host "2. Add new apps"
        Write-Host "3. Clean up superseded apps"
        Write-Host "4. Exit"
        $choice = Read-Host "Enter your choice (1-4)"
        switch ($choice) {
            '1' { Update-ExistingApps }
            '2' { Add-NewApps }
            '3' { Cleanup-SupersededApps }
            '4' {
                Write-Log "[EXIT] User exited script."
                return
            }
            default {
                Write-Host "Invalid selection, please choose 1-4." -ForegroundColor Yellow
            }
        }
    } while ($true)
}

Function Initialize-WinTuner {
    do {
        $global:rootPackageFolder = Read-Host "Please enter the path to the packaging directory"
        if ([string]::IsNullOrWhiteSpace($global:rootPackageFolder)) {
            Write-Host "[ERROR] Directory path cannot be empty." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($global:rootPackageFolder))

    if (-not (Test-Path $global:rootPackageFolder)) {
        $createFolder = Read-Host "Directory does not exist. Would you like to create it? (Y/N)"
        if ($createFolder -match "^[Yy]$") {
            try {
                New-Item -ItemType Directory -Path $global:rootPackageFolder | Out-Null
                Write-Host "[SUCCESS] Created directory: $global:rootPackageFolder"
            } catch {
                Write-Host "[ERROR] Failed to create directory. Exiting." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "[EXIT] Directory not created. Exiting script."
            exit
        }
    }

    $global:CsvLogFile = Join-Path -Path $global:rootPackageFolder -ChildPath "WinTuner_UpdateLog.csv"
    Write-Log "[LOGGING] Logging initialized."
# Ensure PowerShell 7
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Host "[ERROR] This script requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
        exit
    }

    # Ensure WinTuner module is installed and imported
    if (-not (Get-Module -ListAvailable -Name WinTuner)) {
        Write-Log "[INFO] WinTuner module not found. Installing..."
        try {
            Install-Module -Name WinTuner -Force -Scope CurrentUser -AllowClobber
            Write-Log "[SUCCESS] WinTuner module installed."
        } catch {
            Write-Log "[ERROR] Failed to install WinTuner module. Exiting."
            exit
        }
    }

    try {
        Import-Module WinTuner -Force -ErrorAction Stop
    } catch {
        Write-Log "[ERROR] Failed to import WinTuner module. Exiting."
        exit
    }

    try {
        Write-Log "[CONNECT] Connecting to WinTuner..."
        Connect-WtWinTuner
        Write-Log "[SUCCESS] Connected to WinTuner successfully."
    } catch {
        Write-Log "[ERROR] Failed to connect to WinTuner. Exiting."
        exit
    }
}

Function Update-ExistingApps {
    $updatedApps = Get-WtWin32Apps -Update $true -Superseded $false
    if (-not $updatedApps -or $updatedApps.Count -eq 0) {
        Write-Log "[INFO] No apps to update."
        return
    }

    Write-Host "`nApps to update:"
    $i = 1
    foreach ($app in $updatedApps) {
        $appName = if ($app.DisplayName) { $app.DisplayName } elseif ($app.Name) { $app.Name } else { "Unknown App" }
        Write-Host "$i. $appName ($($app.CurrentVersion) → $($app.LatestVersion))"
        $i++
    }
    Write-Host "`nOptions:"
    Write-Host "1. Update all"
    Write-Host "2. Choose specific"
    Write-Host "3. Cancel"
    $choice = Read-Host "Enter choice (1-3)"

    if ($choice -eq "1") {
        foreach ($app in $updatedApps) {
            New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $global:rootPackageFolder -Version $app.LatestVersion | Deploy-WtWin32App -GraphId $app.GraphId
            Write-Log "[SUCCESS] Updated $($app.DisplayName)"
        }
    } elseif ($choice -eq "2") {
        $selection = Read-Host "Enter numbers (comma-separated)"
        $indexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" }
        foreach ($index in $indexes) {
            $app = $updatedApps[$index - 1]
            if ($app) {
                New-WtWingetPackage -PackageId $app.PackageId -PackageFolder $global:rootPackageFolder -Version $app.LatestVersion | Deploy-WtWin32App -GraphId $app.GraphId
                Write-Log "[SUCCESS] Updated $($app.DisplayName)"
            }
        }
    } else {
        Write-Log "[INFO] Update cancelled."
    }
}

Function Add-NewApps {
    do {
        $term = Read-Host "Enter app name or ID to search (or 'done' to exit)"
        if ($term -eq "done") { break }

        $results = Search-WtWinGetPackage -PackageId $term
        if (-not $results) {
            Write-Log "[INFO] No matches found."
            continue
        }

        $i = 1
        foreach ($app in $results) {
            Write-Host "$i. $($app.Name) (ID: $($app.PackageId))"
            $i++
        }

        $select = Read-Host "Select app number to deploy (or Enter to skip)"
        if ($select -match "^\d+$") {
            $selected = $results[$select - 1]

            $assignType = Read-Host "Assignment type? (1 = Available, 2 = Required, 3 = Skip)"
            Write-Host "NOTE: A GUID (Azure AD Group Object ID) must be used — group *names* will not work." -ForegroundColor Yellow
            $targetGroup = Read-Host "Assign to who? (1 = All Users, 2 = All Devices, 3 = Enter Group ID)"

            $available = $null
            $required = $null
            $group = $null

            if ($targetGroup -eq "1") {
                $group = "AllUsers"
            } elseif ($targetGroup -eq "2") {
                $group = "AllDevices"
            } elseif ($targetGroup -eq "3") {
                $group = Read-Host "Enter Azure AD group Object ID"
            }

            if ($assignType -eq "1") {
                $available = $group
            } elseif ($assignType -eq "2") {
                $required = $group
            }

            if ($required) {
                New-WtWingetPackage -PackageId $selected.PackageId -PackageFolder $global:rootPackageFolder -Version $selected.Version |
                    Deploy-WtWin32App -GraphId $selected.GraphId -RequiredFor $required
            } elseif ($available) {
                New-WtWingetPackage -PackageId $selected.PackageId -PackageFolder $global:rootPackageFolder -Version $selected.Version |
                    Deploy-WtWin32App -GraphId $selected.GraphId -AvailableFor $available
            } else {
                New-WtWingetPackage -PackageId $selected.PackageId -PackageFolder $global:rootPackageFolder -Version $selected.Version |
                    Deploy-WtWin32App -GraphId $selected.GraphId
            }

            Write-Log "[SUCCESS] Deployed $($selected.Name)"
        }
    } while ($true)
}

        $i = 1

# Start script
Initialize-WinTuner
Show-MainMenu