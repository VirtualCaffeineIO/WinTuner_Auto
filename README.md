# WinTuner Intune Automation Script

This PowerShell script automates application lifecycle management in Microsoft Intune using the WinTuner module. It streamlines the process of:

- Checking for app updates and deploying newer versions
- Searching for and adding new apps from the Winget repository
- Cleaning up old or superseded Intune apps
- Assigning apps to users/devices (Available or Required)

## Features
- Interactive menu for update, add, and cleanup
- Supports app assignment to All Users, All Devices, or Group Object IDs
- Detects PowerShell 7
- Ensures WinTuner is installed and up to date
- Prompts for WinTuner authentication

## Requirements
- PowerShell 7 or later
- WinTuner module (auto-installed if missing)
- Admin rights to Microsoft Intune environment

## Usage
Run the script from an elevated PowerShell 7 terminal:

```powershell
./wintuner-auto.ps1
```

Follow the prompts to:
1. Choose a packaging directory
2. Authenticate with Intune via WinTuner
3. Select Update, Add, or Cleanup operations

## Author
- **Virtual Caffeine IO**
