@{
    # Module manifest for LAPS

    # Script module associated with this manifest
    RootModule = 'LAPS.psm1'

    # Version number
    ModuleVersion = '1.0.0'

    # ID used to uniquely identify this module
    GUID = 'b2c3d4e5-f6a7-8901-bcde-f12345678901'

    # Author
    Author = 'markorr321'

    # Company
    CompanyName = 'Orr365'

    # Copyright
    Copyright = '(c) 2026. All rights reserved.'

    # Description
    Description = 'PowerShell module for retrieving Microsoft LAPS (Local Administrator Password Solution) credentials from Entra ID devices via Microsoft Graph API. Features an interactive TUI for device search, password display, and clipboard copy. Browser-based authentication with optional custom app registration support. Cross-platform compatible with Windows, macOS, and Linux.'

    # Minimum PowerShell version
    PowerShellVersion = '7.0'

    # Functions to export
    FunctionsToExport = @('Start-LAPS', 'Configure-LAPS', 'Clear-LAPSConfig', 'Get-LAPSHelp')

    # Cmdlets to export
    CmdletsToExport = @()

    # Variables to export
    VariablesToExport = @()

    # Aliases to export
    AliasesToExport = @()

    # Private data
    PrivateData = @{
        PSData = @{
            Tags = @('LAPS', 'Entra', 'Azure', 'Identity', 'Security', 'MicrosoftGraph', 'LocalAdmin', 'Password', 'CrossPlatform', 'TUI')

            LicenseUri = 'https://github.com/markorr321/LAPS/blob/main/LICENSE'

            ProjectUri = 'https://github.com/markorr321/LAPS'

            ReleaseNotes = @'
## 1.0.0
- Initial release
- Interactive TUI for LAPS password retrieval
- Device search by name with Entra ID
- Password display with base64 decoding
- Copy password to clipboard
- Browser-based authentication via MSAL
- Custom app registration support (Configure-LAPS)
- Cross-platform support (Windows, macOS, Linux)
'@

            Prerelease = ''

            RequireLicenseAcceptance = $false
        }
    }
}
