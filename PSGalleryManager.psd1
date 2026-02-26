@{
    # Module info
    RootModule        = 'PSGalleryManager.psm1'
    ModuleVersion     = '1.0.2'
    GUID              = 'a3f7b2c1-9d4e-4f8a-b6c5-1e2d3f4a5b6c'
    Author            = 'Nick Geoffroy'
    CompanyName       = 'Network-IT BV'
    Copyright         = '(c) 2025 Nick Geoffroy. MIT License.'
    Description       = 'A WPF-based graphical tool for managing PowerShell modules. Search and install from PSGallery, per-module scope detection (CurrentUser/AllUsers), update checking, batch update and uninstall, module details panel, and JSON cache for fast startup.'

    # Requirements
    PowerShellVersion = '5.1'
    CLRVersion        = '4.0'

    # Functions to export
    FunctionsToExport = @('Start-PSGalleryManager')
    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @('psgm')

    # Private data for PSGallery
    PrivateData = @{
        PSData = @{
            Tags         = @('GUI', 'Module', 'Manager', 'WPF', 'PSGallery', 'Install', 'Update', 'Uninstall', 'PowerShellGet', 'Windows')
            ProjectUri   = 'https://github.com/NickGeoffroy/PSGalleryManager'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            ReleaseNotes = @'
v1.0.2
- Initial public release
- WPF dark-themed GUI for managing PowerShell modules
- Search and install modules from PSGallery
- Per-module scope detection (CurrentUser vs AllUsers)
- Scope badges on module cards and detail panel
- Batch select, update, and uninstall with Ctrl/Shift+click
- Smart admin warnings for AllUsers scope modules
- JSON cache for instant startup (stored in %APPDATA%\PSGalleryManager)
- Cache age display, manual refresh, and clear cache
- Detail panel with description, downloads, tags, dependencies, license
- Open module install location in Explorer
- AcceptLicense parameter auto-detection for older PowerShellGet
'@
        }
    }
}
