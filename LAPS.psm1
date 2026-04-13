#Requires -Version 7.0

# LAPS PowerShell Module
# TUI for retrieving Local Administrator Password Solution credentials via Microsoft Graph

function Get-LAPSHelp {
    <#
    .SYNOPSIS
        Displays help information for LAPS commands.

    .DESCRIPTION
        Shows all available LAPS commands with examples and usage information.

    .EXAMPLE
        Get-LAPSHelp
    #>
    [CmdletBinding()]
    param()

    Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                    LAPS HELP & COMMANDS                      ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

    Write-Host "BASIC USAGE" -ForegroundColor Yellow
    Write-Host "═══════════" -ForegroundColor Yellow
    Write-Host "  Start-LAPS" -ForegroundColor White
    Write-Host "    Launch the LAPS password retrieval tool" -ForegroundColor Gray
    Write-Host "    Search for devices and retrieve local admin passwords`n" -ForegroundColor Gray

    Write-Host "CONFIGURATION COMMANDS" -ForegroundColor Yellow
    Write-Host "══════════════════════" -ForegroundColor Yellow

    Write-Host "`n  Configure-LAPS" -ForegroundColor White
    Write-Host "    Set up custom app registration for your organization" -ForegroundColor Gray
    Write-Host "    - Prompts for ClientId and TenantId" -ForegroundColor DarkGray
    Write-Host "    - Saves as environment variables (persists across sessions)" -ForegroundColor DarkGray
    Write-Host "    - After configuration, just run: Start-LAPS`n" -ForegroundColor DarkGray

    Write-Host "  Clear-LAPSConfig" -ForegroundColor White
    Write-Host "    Remove saved configuration and return to default auth" -ForegroundColor Gray
    Write-Host "    - Removes environment variables permanently`n" -ForegroundColor DarkGray

    Write-Host "ADVANCED USAGE" -ForegroundColor Yellow
    Write-Host "══════════════" -ForegroundColor Yellow
    Write-Host "  Start-LAPS -ClientId <id> -TenantId <id>" -ForegroundColor White
    Write-Host "    Use custom app registration for a single session" -ForegroundColor Gray
    Write-Host "    (Does not save configuration)`n" -ForegroundColor DarkGray

    Write-Host "APP REGISTRATION REQUIREMENTS" -ForegroundColor Yellow
    Write-Host "═════════════════════════════" -ForegroundColor Yellow
    Write-Host "  - Platform: Mobile and desktop applications" -ForegroundColor Gray
    Write-Host "  - Redirect URI: http://localhost" -ForegroundColor Gray
    Write-Host "  - Allow public client flows: Yes" -ForegroundColor Gray
    Write-Host "  - API Permissions (delegated):" -ForegroundColor Gray
    Write-Host "    - Device.Read.All" -ForegroundColor DarkGray
    Write-Host "    - DeviceLocalCredential.Read.All`n" -ForegroundColor DarkGray

    Write-Host "IN-APP CONTROLS" -ForegroundColor Yellow
    Write-Host "═══════════════" -ForegroundColor Yellow
    Write-Host "  Up/Down     Navigate device list" -ForegroundColor Gray
    Write-Host "  Enter       Select device / Confirm" -ForegroundColor Gray
    Write-Host "  C           Copy password to clipboard" -ForegroundColor Gray
    Write-Host "  S           New search" -ForegroundColor Gray
    Write-Host "  ESC         Go back" -ForegroundColor Gray
    Write-Host "  Ctrl+Q      Exit`n" -ForegroundColor Gray
}

function Configure-LAPS {
    <#
    .SYNOPSIS
        Configure LAPS with custom app registration credentials.

    .DESCRIPTION
        Interactively prompts for ClientId and TenantId and saves them as user-level
        environment variables. Once configured, Start-LAPS will automatically use
        these credentials without requiring parameters.

    .EXAMPLE
        Configure-LAPS
    #>
    [CmdletBinding()]
    param()

    Write-Host "`nLAPS Configuration" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    Write-Host "`nThis will configure your custom app registration for LAPS."
    Write-Host "These settings will be saved as user-level environment variables.`n"

    $clientId = Read-Host "Enter your App Registration Client ID"
    if ([string]::IsNullOrWhiteSpace($clientId)) {
        Write-Host "ClientId cannot be empty. Configuration cancelled." -ForegroundColor Yellow
        return
    }

    $tenantId = Read-Host "Enter your Tenant ID"
    if ([string]::IsNullOrWhiteSpace($tenantId)) {
        Write-Host "TenantId cannot be empty. Configuration cancelled." -ForegroundColor Yellow
        return
    }

    try {
        [System.Environment]::SetEnvironmentVariable('LAPS_CLIENTID', $clientId, 'User')
        [System.Environment]::SetEnvironmentVariable('LAPS_TENANTID', $tenantId, 'User')

        $env:LAPS_CLIENTID = $clientId
        $env:LAPS_TENANTID = $tenantId

        Write-Host "`nConfiguration saved successfully!" -ForegroundColor Green
        Write-Host "You can now run Start-LAPS without parameters.`n" -ForegroundColor Green

        # macOS-specific handling
        $isRunningOnMac = if ($null -ne $IsMacOS) { $IsMacOS } else { $PSVersionTable.OS -match 'Darwin' }
        if ($isRunningOnMac) {
            Write-Host "macOS Note:" -ForegroundColor Yellow
            Write-Host "Environment variables may not persist across terminal sessions on macOS." -ForegroundColor Gray
            Write-Host "To ensure persistence, add the following to your PowerShell profile:`n" -ForegroundColor Gray
            Write-Host "`$env:LAPS_CLIENTID = `"$clientId`"" -ForegroundColor Cyan
            Write-Host "`$env:LAPS_TENANTID = `"$tenantId`"`n" -ForegroundColor Cyan

            Write-Host "Would you like to:" -ForegroundColor Yellow
            Write-Host "  1) Add automatically to PowerShell profile" -ForegroundColor White
            Write-Host "  2) Do it manually later" -ForegroundColor White
            Write-Host ""
            $choice = Read-Host "Enter choice (1 or 2)"

            if ($choice -eq "1") {
                $profilePath = $PROFILE.CurrentUserAllHosts
                if (-not (Test-Path $profilePath)) {
                    New-Item -Path $profilePath -ItemType File -Force | Out-Null
                }

                $profileContent = @"

# LAPS Configuration
`$env:LAPS_CLIENTID = "$clientId"
`$env:LAPS_TENANTID = "$tenantId"
"@
                Add-Content -Path $profilePath -Value $profileContent
                Write-Host "`nAdded to PowerShell profile: $profilePath" -ForegroundColor Green
                Write-Host "Configuration will persist across sessions.`n" -ForegroundColor Green
            } else {
                Write-Host "`nYou can add it manually later to: $($PROFILE.CurrentUserAllHosts)`n" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "`nFailed to save configuration: $_" -ForegroundColor Red
    }
}

function Clear-LAPSConfig {
    <#
    .SYNOPSIS
        Clears the saved LAPS configuration.

    .DESCRIPTION
        Removes the user-level environment variables for ClientId and TenantId.
        After clearing, Start-LAPS will use the default authentication flow.

    .EXAMPLE
        Clear-LAPSConfig
    #>
    [CmdletBinding()]
    param()

    try {
        [System.Environment]::SetEnvironmentVariable('LAPS_CLIENTID', $null, 'User')
        [System.Environment]::SetEnvironmentVariable('LAPS_TENANTID', $null, 'User')

        $env:LAPS_CLIENTID = $null
        $env:LAPS_TENANTID = $null

        Write-Host "LAPS configuration cleared successfully." -ForegroundColor Green
        Write-Host "Start-LAPS will now use the default authentication flow.`n" -ForegroundColor Green

        # macOS-specific handling
        $isRunningOnMac = if ($null -ne $IsMacOS) { $IsMacOS } else { $PSVersionTable.OS -match 'Darwin' }
        if ($isRunningOnMac) {
            $profilePath = $PROFILE.CurrentUserAllHosts
            if (Test-Path $profilePath) {
                $profileContent = Get-Content -Path $profilePath -Raw
                if ($profileContent -match 'LAPS_CLIENTID' -or $profileContent -match 'LAPS_TENANTID') {
                    Write-Host "macOS Note:" -ForegroundColor Yellow
                    Write-Host "Configuration found in PowerShell profile." -ForegroundColor Gray
                    Write-Host "Would you like to remove it from your profile? (y/n)" -ForegroundColor Yellow
                    $choice = Read-Host

                    if ($choice -eq 'y' -or $choice -eq 'Y') {
                        $newContent = $profileContent -replace '(?ms)# LAPS Configuration.*?\$env:LAPS_TENANTID = ".*?"', ''
                        Set-Content -Path $profilePath -Value $newContent.Trim()
                        Write-Host "Removed from PowerShell profile: $profilePath`n" -ForegroundColor Green
                    } else {
                        Write-Host "Profile not modified. You can manually edit: $profilePath`n" -ForegroundColor Gray
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Failed to clear configuration: $_" -ForegroundColor Red
    }
}

function Start-LAPS {
    <#
    .SYNOPSIS
        Launch the LAPS password retrieval TUI.

    .DESCRIPTION
        Interactive console application for searching Entra ID devices and
        retrieving their LAPS (Local Administrator Password Solution) credentials
        via Microsoft Graph API.

        Supports browser-based authentication with optional custom app registration.

    .PARAMETER ClientId
        Optional. Client ID of a custom app registration for delegated auth.

    .PARAMETER TenantId
        Optional. Tenant ID to use with the specified app registration.

    .EXAMPLE
        Start-LAPS
        Launch with default authentication.

    .EXAMPLE
        Start-LAPS -ClientId "your-client-id" -TenantId "your-tenant-id"
        Launch with a custom app registration.
    #>
    [CmdletBinding()]
    param(
        [Parameter(HelpMessage = "Client ID of the app registration to use for delegated auth")]
        [string]$ClientId,

        [Parameter(HelpMessage = "Tenant ID to use with the specified app registration")]
        [string]$TenantId
    )

    # Check for saved configuration if no parameters provided
    if (-not $ClientId -and $env:LAPS_CLIENTID) {
        $ClientId = $env:LAPS_CLIENTID
    }
    if (-not $TenantId -and $env:LAPS_TENANTID) {
        $TenantId = $env:LAPS_TENANTID
    }

    # Build argument list for the script
    $scriptPath = Join-Path $PSScriptRoot "LAPS.ps1"
    $arguments = @{}
    if ($ClientId) { $arguments['ClientId'] = $ClientId }
    if ($TenantId) { $arguments['TenantId'] = $TenantId }

    & $scriptPath @arguments
}

Export-ModuleMember -Function @('Start-LAPS', 'Configure-LAPS', 'Clear-LAPSConfig', 'Get-LAPSHelp')
