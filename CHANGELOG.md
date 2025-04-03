# Changelog

## [1.0.1] - Latest
### Fixed
- Replaced `-JsonPath` with pipeline for Deploy-WtWin32App (fixes deploy errors)
- Added null check for GraphId during app cleanup
- Validates PowerShell 7 and auto-installs WinTuner module
- Restores proper app assignment prompts

### Added
- `Cleanup-SupersededApps` function with safety logic
- Full GitHub-ready packaging and README

## [1.0.0] - Initial release
- Functional update, add, and assignment via WinTuner