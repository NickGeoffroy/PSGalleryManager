# PSGalleryManager

A WPF-based graphical tool for managing PowerShell modules on Windows.

## Features

- **Search & Install** - Browse and install modules from PSGallery
- **Update Checking** - Scan installed modules for available updates
- **Per-Module Scope Detection** - Shows whether each module is installed as CurrentUser or AllUsers
- **Batch Operations** - Multi-select modules with Ctrl/Shift+click for bulk update or uninstall
- **Smart Admin Warnings** - Standard users get clear warnings when trying to modify AllUsers modules
- **JSON Cache** - Instant startup by caching module data locally
- **Detail Panel** - View description, author, downloads, tags, dependencies, license, and install location
- **Open Location** - Open the module's install folder directly in Explorer

## Installation

```powershell
Install-Module -Name PSGalleryManager -Scope CurrentUser
```

## Usage

```powershell
# Full command
Start-PSGalleryManager

# Short alias
psgm
```

Run as Administrator to manage modules installed in AllUsers (system-wide) scope.

## Keyboard Shortcuts

- **Enter** in search box triggers gallery search
- **Ctrl+Click** to select multiple modules
- **Shift+Click** to select a range of modules

## Cache

Module data is cached in `%APPDATA%\PSGalleryManager\` for fast startup:

- `installed.json` - Installed module list with metadata
- `updates.json` - Update check results

Use **Refresh (live scan)** to bypass cache, or **Clear Cache** to delete cached data.

## Requirements

- Windows PowerShell 5.1 or later
- Windows OS (WPF dependency)
- PowerShellGet module (included with PowerShell 5.1+)

## License

MIT License - See [LICENSE](LICENSE) for details.

## Author

Nick Geoffroy - Network-IT BV
