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
    $scriptPath = Join-Path $PSScriptRoot "Intune-LAPS.ps1"
    $arguments = @{}
    if ($ClientId) { $arguments['ClientId'] = $ClientId }
    if ($TenantId) { $arguments['TenantId'] = $TenantId }

    & $scriptPath @arguments
}

Export-ModuleMember -Function @('Start-LAPS', 'Configure-LAPS', 'Clear-LAPSConfig', 'Get-LAPSHelp')
# SIG # Begin signature block
# MIIsCQYJKoZIhvcNAQcCoIIr+jCCK/YCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDV+r2VBizfa3A2
# L0/YmrVNEqAkPfCx9y0sVBH38YzAy6CCJRowggVvMIIEV6ADAgECAhBI/JO0YFWU
# jTanyYqJ1pQWMA0GCSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQI
# DBJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoM
# EUNvbW9kbyBDQSBMaW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2Vy
# dmljZXMwHhcNMjEwNTI1MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjBWMQswCQYDVQQG
# EwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYDVQQDEyRTZWN0aWdv
# IFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCN55QSIgQkdC7/FiMCkoq2rjaFrEfUI5ErPtx94jGgUW+s
# hJHjUoq14pbe0IdjJImK/+8Skzt9u7aKvb0Ffyeba2XTpQxpsbxJOZrxbW6q5KCD
# J9qaDStQ6Utbs7hkNqR+Sj2pcaths3OzPAsM79szV+W+NDfjlxtd/R8SPYIDdub7
# P2bSlDFp+m2zNKzBenjcklDyZMeqLQSrw2rq4C+np9xu1+j/2iGrQL+57g2extme
# me/G3h+pDHazJyCh1rr9gOcB0u/rgimVcI3/uxXP/tEPNqIuTzKQdEZrRzUTdwUz
# T2MuuC3hv2WnBGsY2HH6zAjybYmZELGt2z4s5KoYsMYHAXVn3m3pY2MeNn9pib6q
# RT5uWl+PoVvLnTCGMOgDs0DGDQ84zWeoU4j6uDBl+m/H5x2xg3RpPqzEaDux5mcz
# mrYI4IAFSEDu9oJkRqj1c7AGlfJsZZ+/VVscnFcax3hGfHCqlBuCF6yH6bbJDoEc
# QNYWFyn8XJwYK+pF9e+91WdPKF4F7pBMeufG9ND8+s0+MkYTIDaKBOq3qgdGnA2T
# OglmmVhcKaO5DKYwODzQRjY1fJy67sPV+Qp2+n4FG0DKkjXp1XrRtX8ArqmQqsV/
# AZwQsRb8zG4Y3G9i/qZQp7h7uJ0VP/4gDHXIIloTlRmQAOka1cKG8eOO7F/05QID
# AQABo4IBEjCCAQ4wHwYDVR0jBBgwFoAUoBEKIz6W8Qfs4q8p74Klf9AwpLQwHQYD
# VR0OBBYEFDLrkpr/NZZILyhAQnAgNpFcF4XmMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMDMBsGA1UdIAQUMBIwBgYE
# VR0gADAIBgZngQwBBAEwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybC5jb21v
# ZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEMBQADggEBABK/oe+LdJqYRLhpRrWrJAoMpIpnuDqBv0WKfVIHqI0fTiGF
# OaNrXi0ghr8QuK55O1PNtPvYRL4G2VxjZ9RAFodEhnIq1jIV9RKDwvnhXRFAZ/ZC
# J3LFI+ICOBpMIOLbAffNRk8monxmwFE2tokCVMf8WPtsAO7+mKYulaEMUykfb9gZ
# pk+e96wJ6l2CxouvgKe9gUhShDHaMuwV5KZMPWw5c9QLhTkg4IUaaOGnSDip0TYl
# d8GNGRbFiExmfS9jzpjoad+sPKhdnckcW67Y8y90z7h+9teDnRGWYpquRRPaf9xH
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggWNMIIEdaADAgECAhAOmxiO
# +dAt5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAi
# BgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAw
# MDBaFw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERp
# Z2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsb
# hA3EMB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iT
# cMKyunWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGb
# NOsFxl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclP
# XuU15zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCr
# VYJBMtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFP
# ObURWBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTv
# kpI6nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWM
# cCxBYKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls
# 5Q5SUUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBR
# a2+xq4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6
# MIIBNjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qY
# rhwPTzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8E
# BAMCAYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDig
# NoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCg
# v0NcVec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQT
# SnovLbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh
# 65ZyoUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSw
# uKFWjuyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAO
# QGPFmCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjD
# TZ9ztwGpn1eqXijiuZQwggYaMIIEAqADAgECAhBiHW0MUgGeO5B5FSCJIRwKMA0G
# CSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExp
# bWl0ZWQxLTArBgNVBAMTJFNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBSb290
# IFI0NjAeFw0yMTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5NTlaMFQxCzAJBgNVBAYT
# AkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28g
# UHVibGljIENvZGUgU2lnbmluZyBDQSBSMzYwggGiMA0GCSqGSIb3DQEBAQUAA4IB
# jwAwggGKAoIBgQCbK51T+jU/jmAGQ2rAz/V/9shTUxjIztNsfvxYB5UXeWUzCxEe
# AEZGbEN4QMgCsJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NVDgFigOMYzB2OKhdqfWGV
# oYW3haT29PSTahYkwmMv0b/83nbeECbiMXhSOtbam+/36F09fy1tsB8je/RV0mIk
# 8XL/tfCK6cPuYHE215wzrK0h1SWHTxPbPuYkRdkP05ZwmRmTnAO5/arnY83jeNzh
# P06ShdnRqtZlV59+8yv+KIhE5ILMqgOZYAENHNX9SJDm+qxp4VqpB3MV/h53yl41
# aHU5pledi9lCBbH9JeIkNFICiVHNkRmq4TpxtwfvjsUedyz8rNyfQJy/aOs5b4s+
# ac7IH60B+Ja7TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz44MPZ1f9+YEQIQty/NQd/
# 2yGgW+ufflcZ/ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBMdlyh2n5HirY4jKnFH/9g
# Rvd+QOfdRrJZb1sCAwEAAaOCAWQwggFgMB8GA1UdIwQYMBaAFDLrkpr/NZZILyhA
# QnAgNpFcF4XmMB0GA1UdDgQWBBQPKssghyi47G9IritUpimqF6TNDDAOBgNVHQ8B
# Af8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcD
# AzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsGA1UdHwREMEIwQKA+oDyG
# Omh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5n
# Um9vdFI0Ni5jcmwwewYIKwYBBQUHAQEEbzBtMEYGCCsGAQUFBzAChjpodHRwOi8v
# Y3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RSNDYu
# cDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG
# 9w0BAQwFAAOCAgEABv+C4XdjNm57oRUgmxP/BP6YdURhw1aVcdGRP4Wh60BAscjW
# 4HL9hcpkOTz5jUug2oeunbYAowbFC2AKK+cMcXIBD0ZdOaWTsyNyBBsMLHqafvIh
# rCymlaS98+QpoBCyKppP0OcxYEdU0hpsaqBBIZOtBajjcw5+w/KeFvPYfLF/ldYp
# mlG+vd0xqlqd099iChnyIMvY5HexjO2AmtsbpVn0OhNcWbWDRF/3sBp6fWXhz7Dc
# ML4iTAWS+MVXeNLj1lJziVKEoroGs9Mlizg0bUMbOalOhOfCipnx8CaLZeVme5yE
# Lg09Jlo8BMe80jO37PU8ejfkP9/uPak7VLwELKxAMcJszkyeiaerlphwoKx1uHRz
# NyE6bxuSKcutisqmKL5OTunAvtONEoteSiabkPVSZ2z76mKnzAfZxCl/3dq3dUNw
# 4rg3sTCggkHSRqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5JKdGvspbOrTfOXyXvmPL6
# E52z1NZJ6ctuMFBQZH3pwWvqURR8AgQdULUvrxjUYbHHj95Ejza63zdrEcxWLDX6
# xWls/GDnVNueKjWUH3fTv1Y8Wdho698YADR7TNx8X8z2Bev6SivBBOHY+uqiirZt
# g0y9ShQoPzmCcn63Syatatvx157YK9hlcPmVoa1oDE5/L9Uo2bC5a4CH2RwwggZL
# MIIEs6ADAgECAhEAh4S8tN9yByR3E9jATIZw9DANBgkqhkiG9w0BAQwFADBUMQsw
# CQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJT
# ZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTI2MDIyNDAwMDAw
# MFoXDTI3MDIyNDIzNTk1OVowRDELMAkGA1UEBhMCVVMxDzANBgNVBAgMBkthbnNh
# czERMA8GA1UECgwITWFyayBPcnIxETAPBgNVBAMMCE1hcmsgT3JyMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAx9tr2sjXvlV3KjWWeg0HYTDicFwZDZv2
# tI//RO1C9IL7uShmYN0eSeyWZW/GNy7fTOlIJ6poUe4R3/ApsNsw9hpOMXc92Bny
# Ds/UXHMYx2YdOO4XI35IxfhZnZhgIj2acQ0BZ542hmYAwtz8c1Xu9xH51eTArmFW
# HV8angRsuFMVyKQOraWQs37tqOVwXeH3FQIT0mFBTbmENhgyxAGLq8nZMFM+JqVV
# WeRgvTFO48UZf0BhgH84k2M44CcA9vVML7w4yueg6qD6D/k7Opy1OfCR1qxSXI0w
# ZeUXodJvgisDRScKZJfPID6PIxxvoeem4VKkV0y3eBF+UtdQ8+NZ7qmlRl2hE6H6
# efWSRNW2imxeVSg9FgQONnJYhkyJmaio/NnLyDB6PyoCDZQaYDiMRRiycHPbYvba
# s0THWB2NFsgr3h3QZxQfZnNB2F/ZVdNlfbGpxTK53Yhf5XT0iaEat9r82wwjlP9c
# /PEl1q8G53Pco/ykqBk/V2PfohhuwiXBHb5zL518lCPPZmOCdIqyvkgAUzWymHSi
# Twm/ZNTNEaHLaktfBJ52G03r7F1YHSxPDJpH84RrBQwNWA8olog3uvvWTWImDuQd
# 8PdvhOrluh11pvMWRn+ic6e2E7A4KQr0x4bZoL/gWBTE9tL8AuCJyjxsjiDAbJRx
# d3Di5Bi7pGsCAwEAAaOCAaYwggGiMB8GA1UdIwQYMBaAFA8qyyCHKLjsb0iuK1Sm
# KaoXpM0MMB0GA1UdDgQWBBRlBYoMei+jtIKM2eL9y3kX+l6hqzAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzBKBgNVHSAE
# QzBBMDUGDCsGAQQBsjEBAgEDAjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3Rp
# Z28uY29tL0NQUzAIBgZngQwBBAEwSQYDVR0fBEIwQDA+oDygOoY4aHR0cDovL2Ny
# bC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcmww
# eQYIKwYBBQUHAQEEbTBrMEQGCCsGAQUFBzAChjhodHRwOi8vY3J0LnNlY3RpZ28u
# Y29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNydDAjBggrBgEFBQcw
# AYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wGwYDVR0RBBQwEoEQbW9yckBvcnIz
# NjUudGVjaDANBgkqhkiG9w0BAQwFAAOCAYEAQYDywuGVM9hgCjKW/Til/gPycxB1
# XL4OH7/9jV72/HPbBKnwXwiFlgTO+Lo4UEbZNy+WQk60u0XtrBIKUbhlapRGQPrl
# 2OKpf9rYOyysg1puVTqnaxY9vevhgB4NVpHqYMi8+Kzpa2rXzXyrVdbVNIMn00ZA
# V6tBTr0fhMt3P4oxF0WYQ/GjfUa1/8O3uqeni36iMyCqP7ao9rJgCOgNvEBokRhh
# 7fFC5YVIjMKwvU/7CgbkgjIBHfX4UMxU2BNvCGTR2ZA5IznmLsRI/4MEP9LMLV8D
# Qm8wh2P1uCaGANSLQ0EQIZtMEm1i03zBwDOTBLVAo7p+2Pw2q7LEOQni6LeX5AzT
# nRvHwcisRM3Kpvx+H6wJnL6x7TXZ7YCHhJ4ZTuMWblXJjVKPueEQfIh04x7oVbIV
# 8LNqVyoP9gJZfkmn5IW8cwIFAzFMsNqW1URfArzJ5An9xIYCUJbzohgtE71NjqiZ
# PI1k4GxzsyeqTNaXEXnzZEfogAvEmHFMMNXGMIIGtDCCBJygAwIBAgIQDcesVwX/
# IZkuQEMiDDpJhjANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjUwNTA3MDAwMDAwWhcN
# MzgwMTE0MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5n
# IFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAtHgx0wqYQXK+PEbAHKx126NGaHS0URedTa2NDZS1mZaDLFTtQ2oR
# jzUXMmxCqvkbsDpz4aH+qbxeLho8I6jY3xL1IusLopuW2qftJYJaDNs1+JH7Z+Qd
# SKWM06qchUP+AbdJgMQB3h2DZ0Mal5kYp77jYMVQXSZH++0trj6Ao+xh/AS7sQRu
# QL37QXbDhAktVJMQbzIBHYJBYgzWIjk8eDrYhXDEpKk7RdoX0M980EpLtlrNyHw0
# Xm+nt5pnYJU3Gmq6bNMI1I7Gb5IBZK4ivbVCiZv7PNBYqHEpNVWC2ZQ8BbfnFRQV
# ESYOszFI2Wv82wnJRfN20VRS3hpLgIR4hjzL0hpoYGk81coWJ+KdPvMvaB0WkE/2
# qHxJ0ucS638ZxqU14lDnki7CcoKCz6eum5A19WZQHkqUJfdkDjHkccpL6uoG8pbF
# 0LJAQQZxst7VvwDDjAmSFTUms+wV/FbWBqi7fTJnjq3hj0XbQcd8hjj/q8d6ylgx
# CZSKi17yVp2NL+cnT6Toy+rN+nM8M7LnLqCrO2JP3oW//1sfuZDKiDEb1AQ8es9X
# r/u6bDTnYCTKIsDq1BtmXUqEG1NqzJKS4kOmxkYp2WyODi7vQTCBZtVFJfVZ3j7O
# gWmnhFr4yUozZtqgPrHRVHhGNKlYzyjlroPxul+bgIspzOwbtmsgY1MCAwEAAaOC
# AV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFO9vU0rp5AZ8esri
# kFb2L9RJ7MtOMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1Ud
# DwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJv
# b3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwB
# BAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQAXzvsWgBz+Bz0RdnEw
# vb4LyLU0pn/N0IfFiBowf0/Dm1wGc/Do7oVMY2mhXZXjDNJQa8j00DNqhCT3t+s8
# G0iP5kvN2n7Jd2E4/iEIUBO41P5F448rSYJ59Ib61eoalhnd6ywFLerycvZTAz40
# y8S4F3/a+Z1jEMK/DMm/axFSgoR8n6c3nuZB9BfBwAQYK9FHaoq2e26MHvVY9gCD
# A/JYsq7pGdogP8HRtrYfctSLANEBfHU16r3J05qX3kId+ZOczgj5kjatVB+NdADV
# ZKON/gnZruMvNYY2o1f4MXRJDMdTSlOLh0HCn2cQLwQCqjFbqrXuvTPSegOOzr4E
# Wj7PtspIHBldNE2K9i697cvaiIo2p61Ed2p8xMJb82Yosn0z4y25xUbI7GIN/TpV
# fHIqQ6Ku/qjTY6hc3hsXMrS+U0yy+GWqAXam4ToWd2UQ1KYT70kZjE4YtL8Pbzg0
# c1ugMZyZZd/BdHLiRu7hAWE6bTEm4XYRkA6Tl4KSFLFk43esaUeqGkH/wyW4N7Oi
# gizwJWeukcyIPbAvjSabnf7+Pu0VrFgoiovRDiyx3zEdmcif/sYQsfch28bZeUz2
# rtY/9TCA6TD8dC3JE3rYkrhLULy7Dc90G6e8BlqmyIjlgp2+VqsS9/wQD7yFylIz
# 0scmbKvFoW2jNrbM1pD2T7m3XDCCBu0wggTVoAMCAQICEAqA7xhLjfEFgtHEdqeV
# dGgwDQYJKoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTAeFw0yNTA2MDQwMDAwMDBaFw0z
# NjA5MDMyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgU0hBMjU2IFJTQTQwOTYgVGltZXN0YW1w
# IFJlc3BvbmRlciAyMDI1IDEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDQRqwtEsae0OquYFazK1e6b1H/hnAKAd/KN8wZQjBjMqiZ3xTWcfsLwOvRxUwX
# cGx8AUjni6bz52fGTfr6PHRNv6T7zsf1Y/E3IU8kgNkeECqVQ+3bzWYesFtkepEr
# vUSbf+EIYLkrLKd6qJnuzK8Vcn0DvbDMemQFoxQ2Dsw4vEjoT1FpS54dNApZfKY6
# 1HAldytxNM89PZXUP/5wWWURK+IfxiOg8W9lKMqzdIo7VA1R0V3Zp3DjjANwqAf4
# lEkTlCDQ0/fKJLKLkzGBTpx6EYevvOi7XOc4zyh1uSqgr6UnbksIcFJqLbkIXIPb
# cNmA98Oskkkrvt6lPAw/p4oDSRZreiwB7x9ykrjS6GS3NR39iTTFS+ENTqW8m6TH
# uOmHHjQNC3zbJ6nJ6SXiLSvw4Smz8U07hqF+8CTXaETkVWz0dVVZw7knh1WZXOLH
# gDvundrAtuvz0D3T+dYaNcwafsVCGZKUhQPL1naFKBy1p6llN3QgshRta6Eq4B40
# h5avMcpi54wm0i2ePZD5pPIssoszQyF4//3DoK2O65Uck5Wggn8O2klETsJ7u8xE
# ehGifgJYi+6I03UuT1j7FnrqVrOzaQoVJOeeStPeldYRNMmSF3voIgMFtNGh86w3
# ISHNm0IaadCKCkUe2LnwJKa8TIlwCUNVwppwn4D3/Pt5pwIDAQABo4IBlTCCAZEw
# DAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQU5Dv88jHt/f3X85FxYxlQQ89hjOgwHwYD
# VR0jBBgwFoAU729TSunkBnx6yuKQVvYv1Ensy04wDgYDVR0PAQH/BAQDAgeAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMIGVBggrBgEFBQcBAQSBiDCBhTAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMF0GCCsGAQUFBzAChlFodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3Rh
# bXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcnQwXwYDVR0fBFgwVjBUoFKgUIZO
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAZSqt8RwnBLmuYEHs
# 0QhEnmNAciH45PYiT9s1i6UKtW+FERp8FgXRGQ/YAavXzWjZhY+hIfP2JkQ38U+w
# tJPBVBajYfrbIYG+Dui4I4PCvHpQuPqFgqp1PzC/ZRX4pvP/ciZmUnthfAEP1HSh
# TrY+2DE5qjzvZs7JIIgt0GCFD9ktx0LxxtRQ7vllKluHWiKk6FxRPyUPxAAYH2Vy
# 1lNM4kzekd8oEARzFAWgeW3az2xejEWLNN4eKGxDJ8WDl/FQUSntbjZ80FU3i54t
# px5F/0Kr15zW/mJAxZMVBrTE2oi0fcI8VMbtoRAmaaslNXdCG1+lqvP4FbrQ6IwS
# BXkZagHLhFU9HCrG/syTRLLhAezu/3Lr00GrJzPQFnCEH1Y58678IgmfORBPC1JK
# kYaEt2OdDh4GmO0/5cHelAK2/gTlQJINqDr6JfwyYHXSd+V08X1JUPvB4ILfJdmL
# +66Gp3CSBXG6IwXMZUXBhtCyIaehr0XkBoDIGMUG1dUtwq1qmcwbdUfcSYCn+Own
# cVUXf53VJUNOaMWMts0VlRYxe5nK+At+DI96HAlXHAL5SlfYxJ7La54i71McVWRP
# 66bW+yERNpbJCjyCYG2j+bdpxo/1Cy4uPcU3AWVPGrbn5PhDBf3Froguzzhk++am
# i+r3Qrx5bIbY3TVzgiFI7Gq3zWcxggZFMIIGQQIBATBpMFQxCzAJBgNVBAYTAkdC
# MRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVi
# bGljIENvZGUgU2lnbmluZyBDQSBSMzYCEQCHhLy033IHJHcT2MBMhnD0MA0GCWCG
# SAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# LwYJKoZIhvcNAQkEMSIEIEoDnYw7EhLaPKzHjIWzEBZhxhMEfyLOtJK0pYCs8wj9
# MA0GCSqGSIb3DQEBAQUABIICAJCnwfUYp32lRct/QFECHz8rDUgfgX5gP9yvR3Ue
# 1/B67dUjXBFnpuTuE2lObln6o8g+U/kAAtoVWgi0TELVD5010cRmLNliMyxwJsX3
# tqtL1eZHR9fqdM1BRp1KKZtMFNx8hp1+G7RczUiHzK+UERhVqaZnPqbct9cT1V7g
# sF782u7E6/FPp/jkCaNj1JK1F3coZUmyHx6sxRMe8aCog0wcAUBWLKGQDDWZU+dV
# 88lYVZhLMQa+6Uu9i3/m5saupZbK2gfAdkX8eOb4SxWz0lSWbsbpXMNc2RyD058G
# 56TmeUmppwwQUK1z+xoOsw27lLpq3OWO1w6g/IZ8DUo3ktBaoacruXBA1bgKxbs0
# WO3kcMN8XM/w6zv9Cv4NqHDFAWvUNegMgBiQiZdClS+aGNvFTRH+ZKYd8b9GyJZe
# kaoAAQ1mUBIOsMoHbo4kMKD6aU7cibsSsAbDVjbhAVopnBnYzZNR/k+obAjAKLJa
# HpJOt/npdqE0pxCc44eMf8p4xYcev9osydHlaGwC/1aimSQF1yjDTkYScNALatPu
# ZBaiWL/0qfVizQQgEPhhXkknvOxze9M6xC4ZoAxR2zKPYO0v6mZqpOyceDkzEP46
# fjMvg6OQy50uxjcqzAP6Ewx3HQF55hbuuKKfguR/e+vj8teDHNRgG/3Bs1UD9PxP
# gvw1oYIDJjCCAyIGCSqGSIb3DQEJBjGCAxMwggMPAgEBMH0waTELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBU
# cnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMQIQ
# CoDvGEuN8QWC0cR2p5V0aDANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI2MDQxNDEwNTA0NlowLwYJKoZI
# hvcNAQkEMSIEII/xHqMaoBHjWYxbMq0ND758fE2w1VSGYn2YTCmT2VX6MA0GCSqG
# SIb3DQEBAQUABIICAI6XoZYUdfENqEfLqN/1tx6Zq2qoX0tZM2esQqFa7J3wY8y0
# 5l1hME0lfOcWVX0vSXmeRgPDSS3WzXQtFQYWEZiExyo5oFVFuDr0fibldYJYk9A9
# J/n1MHLB6/7h59SorLnVup/vlekcHqheUpBHZPzJv0nuBR7dlxvcFRjtocFqqV/A
# LWLwe14v2TUflEyKSl6IMgc7u0iiPBSuz6cehnFvbw3acs9CaDvp7trunTi396Sj
# pMFOeJzoUSPGgHrsQ/dsSpGzMSREJUJetFps8XzhB2o4Df14o/pRpIDdHtMA8Z9n
# 49noUi4HzGUEN4vDgClDy5HVcVu/ZQyIbR/2ozpstLr5df1RbEVLvFaL/MKvGAKu
# JjRl3LXrppYcH56HFWmZaIbeFse1LjzrQUL7/HY7arRFrMa+uWDS5desq2OVWEhT
# ge2AjWBxYkPOfwKuNMhIrmKMH9Q08VdIQxhpsSNWsXxRa2T0HFIxQEEVGTCds6nI
# IMxA/klpmST3rd5xerTb36be7apKziaF9U8QfuaFwmSc8CkfUZt54yS2HxZGcqmI
# 878j6bsItKDsryIjzsphwdrNzQ0ks9SmU1fd2Mn11IgqCuaDjuUQ6nWF1XKAymSc
# Qk4C8AkX3Bkoq2UR99AruFgjo5qVWtTKE8z7Gh2yVmekGZ2/f5MUeIyWkFj9
# SIG # End signature block
