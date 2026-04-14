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

# SIG # Begin signature block
# MIIr0AYJKoZIhvcNAQcCoIIrwTCCK70CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtlRIn+FNA4iPz
# IkhpYP7qqAxFxXtMFjAJhBJW9lKK0qCCJOQwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# +9/DUp/mBlXpnYzyOmJRvOwkDynUWICE5EV7WtgwggYUMIID/KADAgECAhB6I67a
# U2mWD5HIPlz0x+M/MA0GCSqGSIb3DQEBDAUAMFcxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxLjAsBgNVBAMTJVNlY3RpZ28gUHVibGljIFRp
# bWUgU3RhbXBpbmcgUm9vdCBSNDYwHhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1
# OTU5WjBVMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSww
# KgYDVQQDEyNTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIENBIFIzNjCCAaIw
# DQYJKoZIhvcNAQEBBQADggGPADCCAYoCggGBAM2Y2ENBq26CK+z2M34mNOSJjNPv
# IhKAVD7vJq+MDoGD46IiM+b83+3ecLvBhStSVjeYXIjfa3ajoW3cS3ElcJzkyZlB
# nwDEJuHlzpbN4kMH2qRBVrjrGJgSlzzUqcGQBaCxpectRGhhnOSwcjPMI3G0hedv
# 2eNmGiUbD12OeORN0ADzdpsQ4dDi6M4YhoGE9cbY11XxM2AVZn0GiOUC9+XE0wI7
# CQKfOUfigLDn7i/WeyxZ43XLj5GVo7LDBExSLnh+va8WxTlA+uBvq1KO8RSHUQLg
# zb1gbL9Ihgzxmkdp2ZWNuLc+XyEmJNbD2OIIq/fWlwBp6KNL19zpHsODLIsgZ+WZ
# 1AzCs1HEK6VWrxmnKyJJg2Lv23DlEdZlQSGdF+z+Gyn9/CRezKe7WNyxRf4e4bwU
# trYE2F5Q+05yDD68clwnweckKtxRaF0VzN/w76kOLIaFVhf5sMM/caEZLtOYqYad
# tn034ykSFaZuIBU9uCSrKRKTPJhWvXk4CllgrwIDAQABo4IBXDCCAVgwHwYDVR0j
# BBgwFoAU9ndq3T/9ARP/FqFsggIv0Ao9FCUwHQYDVR0OBBYEFF9Y7UwxeqJhQo1S
# gLqzYZcZojKbMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMG
# A1UdJQQMMAoGCCsGAQUFBwMIMBEGA1UdIAQKMAgwBgYEVR0gADBMBgNVHR8ERTBD
# MEGgP6A9hjtodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LmNybDB8BggrBgEFBQcBAQRwMG4wRwYIKwYBBQUHMAKG
# O2h0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY1RpbWVTdGFtcGlu
# Z1Jvb3RSNDYucDdjMCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNv
# bTANBgkqhkiG9w0BAQwFAAOCAgEAEtd7IK0ONVgMnoEdJVj9TC1ndK/HYiYh9lVU
# acahRoZ2W2hfiEOyQExnHk1jkvpIJzAMxmEc6ZvIyHI5UkPCbXKspioYMdbOnBWQ
# Un733qMooBfIghpR/klUqNxx6/fDXqY0hSU1OSkkSivt51UlmJElUICZYBodzD3M
# /SFjeCP59anwxs6hwj1mfvzG+b1coYGnqsSz2wSKr+nDO+Db8qNcTbJZRAiSazr7
# KyUJGo1c+MScGfG5QHV+bps8BX5Oyv9Ct36Y4Il6ajTqV2ifikkVtB3RNBUgwu/m
# SiSUice/Jp/q8BMk/gN8+0rNIE+QqU63JoVMCMPY2752LmESsRVVoypJVt8/N3qQ
# 1c6FibbcRabo3azZkcIdWGVSAdoLgAIxEKBeNh9AQO1gQrnh1TA8ldXuJzPSuALO
# z1Ujb0PCyNVkWk7hkhVHfcvBfI8NtgWQupiaAeNHe0pWSGH2opXZYKYG4Lbukg7H
# pNi/KqJhue2Keak6qH9A8CeEOB7Eob0Zf+fU+CCQaL0cJqlmnx9HCDxF+3BLbUuf
# rV64EbTI40zqegPZdA+sXCmbcZy6okx/SjwsusWRItFA3DE8MORZeFb6BmzBtqKJ
# 7l939bbKBy2jvxcJI98Va95Q5JnlKor3m0E7xpMeYRriWklUPsetMSf2NvUQa/E5
# vVyefQIwggYaMIIEAqADAgECAhBiHW0MUgGeO5B5FSCJIRwKMA0GCSqGSIb3DQEB
# DAUAMFYxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLTAr
# BgNVBAMTJFNlY3RpZ28gUHVibGljIENvZGUgU2lnbmluZyBSb290IFI0NjAeFw0y
# MTAzMjIwMDAwMDBaFw0zNjAzMjEyMzU5NTlaMFQxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBDQSBSMzYwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAwggGKAoIB
# gQCbK51T+jU/jmAGQ2rAz/V/9shTUxjIztNsfvxYB5UXeWUzCxEeAEZGbEN4QMgC
# sJLZUKhWThj/yPqy0iSZhXkZ6Pg2A2NVDgFigOMYzB2OKhdqfWGVoYW3haT29PST
# ahYkwmMv0b/83nbeECbiMXhSOtbam+/36F09fy1tsB8je/RV0mIk8XL/tfCK6cPu
# YHE215wzrK0h1SWHTxPbPuYkRdkP05ZwmRmTnAO5/arnY83jeNzhP06ShdnRqtZl
# V59+8yv+KIhE5ILMqgOZYAENHNX9SJDm+qxp4VqpB3MV/h53yl41aHU5pledi9lC
# BbH9JeIkNFICiVHNkRmq4TpxtwfvjsUedyz8rNyfQJy/aOs5b4s+ac7IH60B+Ja7
# TVM+EKv1WuTGwcLmoU3FpOFMbmPj8pz44MPZ1f9+YEQIQty/NQd/2yGgW+ufflcZ
# /ZE9o1M7a5Jnqf2i2/uMSWymR8r2oQBMdlyh2n5HirY4jKnFH/9gRvd+QOfdRrJZ
# b1sCAwEAAaOCAWQwggFgMB8GA1UdIwQYMBaAFDLrkpr/NZZILyhAQnAgNpFcF4Xm
# MB0GA1UdDgQWBBQPKssghyi47G9IritUpimqF6TNDDAOBgNVHQ8BAf8EBAMCAYYw
# EgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcDAzAbBgNVHSAE
# FDASMAYGBFUdIAAwCAYGZ4EMAQQBMEsGA1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9j
# cmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVTaWduaW5nUm9vdFI0Ni5j
# cmwwewYIKwYBBQUHAQEEbzBtMEYGCCsGAQUFBzAChjpodHRwOi8vY3J0LnNlY3Rp
# Z28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RSNDYucDdjMCMGCCsG
# AQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOC
# AgEABv+C4XdjNm57oRUgmxP/BP6YdURhw1aVcdGRP4Wh60BAscjW4HL9hcpkOTz5
# jUug2oeunbYAowbFC2AKK+cMcXIBD0ZdOaWTsyNyBBsMLHqafvIhrCymlaS98+Qp
# oBCyKppP0OcxYEdU0hpsaqBBIZOtBajjcw5+w/KeFvPYfLF/ldYpmlG+vd0xqlqd
# 099iChnyIMvY5HexjO2AmtsbpVn0OhNcWbWDRF/3sBp6fWXhz7DcML4iTAWS+MVX
# eNLj1lJziVKEoroGs9Mlizg0bUMbOalOhOfCipnx8CaLZeVme5yELg09Jlo8BMe8
# 0jO37PU8ejfkP9/uPak7VLwELKxAMcJszkyeiaerlphwoKx1uHRzNyE6bxuSKcut
# isqmKL5OTunAvtONEoteSiabkPVSZ2z76mKnzAfZxCl/3dq3dUNw4rg3sTCggkHS
# RqTqlLMS7gjrhTqBmzu1L90Y1KWN/Y5JKdGvspbOrTfOXyXvmPL6E52z1NZJ6ctu
# MFBQZH3pwWvqURR8AgQdULUvrxjUYbHHj95Ejza63zdrEcxWLDX6xWls/GDnVNue
# KjWUH3fTv1Y8Wdho698YADR7TNx8X8z2Bev6SivBBOHY+uqiirZtg0y9ShQoPzmC
# cn63Syatatvx157YK9hlcPmVoa1oDE5/L9Uo2bC5a4CH2RwwggZLMIIEs6ADAgEC
# AhEAh4S8tN9yByR3E9jATIZw9DANBgkqhkiG9w0BAQwFADBUMQswCQYDVQQGEwJH
# QjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1
# YmxpYyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTI2MDIyNDAwMDAwMFoXDTI3MDIy
# NDIzNTk1OVowRDELMAkGA1UEBhMCVVMxDzANBgNVBAgMBkthbnNhczERMA8GA1UE
# CgwITWFyayBPcnIxETAPBgNVBAMMCE1hcmsgT3JyMIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEAx9tr2sjXvlV3KjWWeg0HYTDicFwZDZv2tI//RO1C9IL7
# uShmYN0eSeyWZW/GNy7fTOlIJ6poUe4R3/ApsNsw9hpOMXc92BnyDs/UXHMYx2Yd
# OO4XI35IxfhZnZhgIj2acQ0BZ542hmYAwtz8c1Xu9xH51eTArmFWHV8angRsuFMV
# yKQOraWQs37tqOVwXeH3FQIT0mFBTbmENhgyxAGLq8nZMFM+JqVVWeRgvTFO48UZ
# f0BhgH84k2M44CcA9vVML7w4yueg6qD6D/k7Opy1OfCR1qxSXI0wZeUXodJvgisD
# RScKZJfPID6PIxxvoeem4VKkV0y3eBF+UtdQ8+NZ7qmlRl2hE6H6efWSRNW2imxe
# VSg9FgQONnJYhkyJmaio/NnLyDB6PyoCDZQaYDiMRRiycHPbYvbas0THWB2NFsgr
# 3h3QZxQfZnNB2F/ZVdNlfbGpxTK53Yhf5XT0iaEat9r82wwjlP9c/PEl1q8G53Pc
# o/ykqBk/V2PfohhuwiXBHb5zL518lCPPZmOCdIqyvkgAUzWymHSiTwm/ZNTNEaHL
# aktfBJ52G03r7F1YHSxPDJpH84RrBQwNWA8olog3uvvWTWImDuQd8PdvhOrluh11
# pvMWRn+ic6e2E7A4KQr0x4bZoL/gWBTE9tL8AuCJyjxsjiDAbJRxd3Di5Bi7pGsC
# AwEAAaOCAaYwggGiMB8GA1UdIwQYMBaAFA8qyyCHKLjsb0iuK1SmKaoXpM0MMB0G
# A1UdDgQWBBRlBYoMei+jtIKM2eL9y3kX+l6hqzAOBgNVHQ8BAf8EBAMCB4AwDAYD
# VR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzBKBgNVHSAEQzBBMDUGDCsG
# AQQBsjEBAgEDAjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQ
# UzAIBgZngQwBBAEwSQYDVR0fBEIwQDA+oDygOoY4aHR0cDovL2NybC5zZWN0aWdv
# LmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIzNi5jcmwweQYIKwYBBQUH
# AQEEbTBrMEQGCCsGAQUFBzAChjhodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3Rp
# Z29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNydDAjBggrBgEFBQcwAYYXaHR0cDov
# L29jc3Auc2VjdGlnby5jb20wGwYDVR0RBBQwEoEQbW9yckBvcnIzNjUudGVjaDAN
# BgkqhkiG9w0BAQwFAAOCAYEAQYDywuGVM9hgCjKW/Til/gPycxB1XL4OH7/9jV72
# /HPbBKnwXwiFlgTO+Lo4UEbZNy+WQk60u0XtrBIKUbhlapRGQPrl2OKpf9rYOyys
# g1puVTqnaxY9vevhgB4NVpHqYMi8+Kzpa2rXzXyrVdbVNIMn00ZAV6tBTr0fhMt3
# P4oxF0WYQ/GjfUa1/8O3uqeni36iMyCqP7ao9rJgCOgNvEBokRhh7fFC5YVIjMKw
# vU/7CgbkgjIBHfX4UMxU2BNvCGTR2ZA5IznmLsRI/4MEP9LMLV8DQm8wh2P1uCaG
# ANSLQ0EQIZtMEm1i03zBwDOTBLVAo7p+2Pw2q7LEOQni6LeX5AzTnRvHwcisRM3K
# pvx+H6wJnL6x7TXZ7YCHhJ4ZTuMWblXJjVKPueEQfIh04x7oVbIV8LNqVyoP9gJZ
# fkmn5IW8cwIFAzFMsNqW1URfArzJ5An9xIYCUJbzohgtE71NjqiZPI1k4Gxzsyeq
# TNaXEXnzZEfogAvEmHFMMNXGMIIGYjCCBMqgAwIBAgIRAKQpO24e3denNAiHrXpO
# tyQwDQYJKoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGlu
# ZyBDQSBSMzYwHhcNMjUwMzI3MDAwMDAwWhcNMzYwMzIxMjM1OTU5WjByMQswCQYD
# VQQGEwJHQjEXMBUGA1UECBMOV2VzdCBZb3Jrc2hpcmUxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGlu
# ZyBTaWduZXIgUjM2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA04SV
# 9G6kU3jyPRBLeBIHPNyUgVNnYayfsGOyYEXrn3+SkDYTLs1crcw/ol2swE1TzB2a
# R/5JIjKNf75QBha2Ddj+4NEPKDxHEd4dEn7RTWMcTIfm492TW22I8LfH+A7Ehz0/
# safc6BbsNBzjHTt7FngNfhfJoYOrkugSaT8F0IzUh6VUwoHdYDpiln9dh0n0m545
# d5A5tJD92iFAIbKHQWGbCQNYplqpAFasHBn77OqW37P9BhOASdmjp3IijYiFdcA0
# WQIe60vzvrk0HG+iVcwVZjz+t5OcXGTcxqOAzk1frDNZ1aw8nFhGEvG0ktJQknnJ
# ZE3D40GofV7O8WzgaAnZmoUn4PCpvH36vD4XaAF2CjiPsJWiY/j2xLsJuqx3JtuI
# 4akH0MmGzlBUylhXvdNVXcjAuIEcEQKtOBR9lU4wXQpISrbOT8ux+96GzBq8Tdbh
# oFcmYaOBZKlwPP7pOp5Mzx/UMhyBA93PQhiCdPfIVOCINsUY4U23p4KJ3F1HqP3H
# 6Slw3lHACnLilGETXRg5X/Fp8G8qlG5Y+M49ZEGUp2bneRLZoyHTyynHvFISpefh
# BCV0KdRZHPcuSL5OAGWnBjAlRtHvsMBrI3AAA0Tu1oGvPa/4yeeiAyu+9y3SLC98
# gDVbySnXnkujjhIh+oaatsk/oyf5R2vcxHahajMCAwEAAaOCAY4wggGKMB8GA1Ud
# IwQYMBaAFF9Y7UwxeqJhQo1SgLqzYZcZojKbMB0GA1UdDgQWBBSIYYyhKjdkgShg
# oZsx0Iz9LALOTzAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUB
# Af8EDDAKBggrBgEFBQcDCDBKBgNVHSAEQzBBMDUGDCsGAQQBsjEBAgEDCDAlMCMG
# CCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQUzAIBgZngQwBBAIwSgYD
# VR0fBEMwQTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nQ0FSMzYuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEF
# BQcwAoY5aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0
# YW1waW5nQ0FSMzYuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdv
# LmNvbTANBgkqhkiG9w0BAQwFAAOCAYEAAoE+pIZyUSH5ZakuPVKK4eWbzEsTRJOE
# jbIu6r7vmzXXLpJx4FyGmcqnFZoa1dzx3JrUCrdG5b//LfAxOGy9Ph9JtrYChJaV
# HrusDh9NgYwiGDOhyyJ2zRy3+kdqhwtUlLCdNjFjakTSE+hkC9F5ty1uxOoQ2Zkf
# I5WM4WXA3ZHcNHB4V42zi7Jk3ktEnkSdViVxM6rduXW0jmmiu71ZpBFZDh7Kdens
# +PQXPgMqvzodgQJEkxaION5XRCoBxAwWwiMm2thPDuZTzWp/gUFzi7izCmEt4pE3
# Kf0MOt3ccgwn4Kl2FIcQaV55nkjv1gODcHcD9+ZVjYZoyKTVWb4VqMQy/j8Q3aaY
# d/jOQ66Fhk3NWbg2tYl5jhQCuIsE55Vg4N0DUbEWvXJxtxQQaVR5xzhEI+BjJKzh
# 3TQ026JxHhr2fuJ0mV68AluFr9qshgwS5SpN5FFtaSEnAwqZv3IS+mlG50rK7W3q
# XbWwi4hmpylUfygtYLEdLQukNEX1jiOKMIIGgjCCBGqgAwIBAgIQNsKwvXwbOuej
# s902y8l1aDANBgkqhkiG9w0BAQwFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Ck5ldyBKZXJzZXkxFDASBgNVBAcTC0plcnNleSBDaXR5MR4wHAYDVQQKExVUaGUg
# VVNFUlRSVVNUIE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBSU0EgQ2VydGlm
# aWNhdGlvbiBBdXRob3JpdHkwHhcNMjEwMzIyMDAwMDAwWhcNMzgwMTE4MjM1OTU5
# WjBXMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS4wLAYD
# VQQDEyVTZWN0aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIFJvb3QgUjQ2MIICIjAN
# BgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAiJ3YuUVnnR3d6LkmgZpUVMB8SQWb
# zFoVD9mUEES0QUCBdxSZqdTkdizICFNeINCSJS+lV1ipnW5ihkQyC0cRLWXUJzod
# qpnMRs46npiJPHrfLBOifjfhpdXJ2aHHsPHggGsCi7uE0awqKggE/LkYw3sqaBia
# 67h/3awoqNvGqiFRJ+OTWYmUCO2GAXsePHi+/JUNAax3kpqstbl3vcTdOGhtKShv
# ZIvjwulRH87rbukNyHGWX5tNK/WABKf+Gnoi4cmisS7oSimgHUI0Wn/4elNd40BF
# dSZ1EwpuddZ+Wr7+Dfo0lcHflm/FDDrOJ3rWqauUP8hsokDoI7D/yUVI9DAE/WK3
# Jl3C4LKwIpn1mNzMyptRwsXKrop06m7NUNHdlTDEMovXAIDGAvYynPt5lutv8lZe
# I5w3MOlCybAZDpK3Dy1MKo+6aEtE9vtiTMzz/o2dYfdP0KWZwZIXbYsTIlg1YIet
# Cpi5s14qiXOpRsKqFKqav9R1R5vj3NgevsAsvxsAnI8Oa5s2oy25qhsoBIGo/zi6
# GpxFj+mOdh35Xn91y72J4RGOJEoqzEIbW3q0b2iPuWLA911cRxgY5SJYubvjay3n
# SMbBPPFsyl6mY4/WYucmyS9lo3l7jk27MAe145GWxK4O3m3gEFEIkv7kRmefDR7O
# e2T1HxAnICQvr9sCAwEAAaOCARYwggESMB8GA1UdIwQYMBaAFFN5v1qqK0rPVIDh
# 2JvAnfKyA2bLMB0GA1UdDgQWBBT2d2rdP/0BE/8WoWyCAi/QCj0UJTAOBgNVHQ8B
# Af8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUEDDAKBggrBgEFBQcDCDAR
# BgNVHSAECjAIMAYGBFUdIAAwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC51
# c2VydHJ1c3QuY29tL1VTRVJUcnVzdFJTQUNlcnRpZmljYXRpb25BdXRob3JpdHku
# Y3JsMDUGCCsGAQUFBwEBBCkwJzAlBggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNl
# cnRydXN0LmNvbTANBgkqhkiG9w0BAQwFAAOCAgEADr5lQe1oRLjlocXUEYfktzsl
# jOt+2sgXke3Y8UPEooU5y39rAARaAdAxUeiX1ktLJ3+lgxtoLQhn5cFb3GF2SSZR
# X8ptQ6IvuD3wz/LNHKpQ5nX8hjsDLRhsyeIiJsms9yAWnvdYOdEMq1W61KE9JlBk
# B20XBee6JaXx4UBErc+YuoSb1SxVf7nkNtUjPfcxuFtrQdRMRi/fInV/AobE8Gw/
# 8yBMQKKaHt5eia8ybT8Y/Ffa6HAJyz9gvEOcF1VWXG8OMeM7Vy7Bs6mSIkYeYtdd
# U1ux1dQLbEGur18ut97wgGwDiGinCwKPyFO7ApcmVJOtlw9FVJxw/mL1TbyBns4z
# OgkaXFnnfzg4qbSvnrwyj1NiurMp4pmAWjR+Pb/SIduPnmFzbSN/G8reZCL4fvGl
# vPFk4Uab/JVCSmj59+/mB2Gn6G/UYOy8k60mKcmaAZsEVkhOFuoj4we8CYyaR9vd
# 9PGZKSinaZIkvVjbH/3nlLb0a7SBIkiRzfPfS9T+JesylbHa1LtRV9U/7m0q7Ma2
# CQ/t392ioOssXW7oKLdOmMBl14suVFBmbzrt5V5cQPnwtd3UOTpS9oCG+ZZheiIv
# PgkDmA8FzPsnfXW5qHELB43ET7HHFHeRPRYrMBKjkb8/IN7Po0d0hQoF4TeMM+zY
# AJzoKQnVKOLg8pZVPT8xggZCMIIGPgIBATBpMFQxCzAJBgNVBAYTAkdCMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENv
# ZGUgU2lnbmluZyBDQSBSMzYCEQCHhLy033IHJHcT2MBMhnD0MA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIHpcvQuE+8kf5CAvWKkiGOTPPGB9Oucxm/aU31jgMD6WMA0GCSqG
# SIb3DQEBAQUABIICADvQYO5dayn+1dSr6NBmgHb91WzpUoQo6pnA1mobTShRklBL
# SfaDGUQdWLHE8HNBJz7D+2UpszuZJbXrq2AtcCWxliHcEFKbDsxldk5q1oaXAxy+
# wQ31R5P13MBLWP9bamgSbbmssmtP6qnTYxMzuvDANeyQfqg65Ldk4T9qwPvqLyT0
# aRxcYKGFDZAAoHE3e/tkx3169/xvMTd8smTE+QSkiXd3FzCLFynyIRkcG+4mj7PY
# toVxySS8/sr3SEe91F2lafPbzUO5QdshhTExAQTBl7yDxpYiKZcbv+I5NoC2SVGv
# qq2VRK1lRyv7wXlmjmujNqImk4QFh+BGFrbLyYLAfJZLl46+JvYuqyavudN5Mibo
# w/xXzwFDniL72qhKioLdfPOPt0hoaptj5sSBbnBZDJpv+mZRHdmfTUcUf2tJYM3p
# lKn1gLzWbEg4aDE3D6Dlrk81vdfHpQNtg8/p9kwmr7m+1KLpZDfDC/BpOa4PySuH
# BtRET1UU+tiTLxQ0MgFKw34C1YUzK4/38APEOuIXJ5gvLOVgaIatu95f0K4VWUUM
# bJhih6/g9lZNPpwGiuiXteqM3FTYyw2YvIlDzOlFyQCZtGKur4gNKklxhS9QeZni
# h556/d6UVw+bECaqvdvso2qf0S1FUeQgiBKfzDVr7CbY8pKXLl9oh441u0k4oYID
# IzCCAx8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAW
# BgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMg
# VGltZSBTdGFtcGluZyBDQSBSMzYCEQCkKTtuHt3XpzQIh616TrckMA0GCWCGSAFl
# AwQCAgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjYwNDE0MDcyODMzWjA/BgkqhkiG9w0BCQQxMgQwzhX1Wzj0Hz5DxWPQN2RO
# 1xGrk+e+tY3SzfpmzhW9wnlCvq1GRz5yZ3Ltxo/StH6uMA0GCSqGSIb3DQEBAQUA
# BIICAB2zqjEShxbm0ABD2UseIp2bbr7NNIH36rwf5HGGtskM0i739Nq3dij/LijB
# AiCaskncrkQG1MONR0w7cTRZJeHLHjqJO7bnyJf8T2HytwTJWUJ0WMDJd3yWCtmN
# 2F+zN0tGGu7id9i5qSXkaOcWf1HCKf8C5HuF/YumbvdhFb6ClUiW62hO7J16QIJ1
# s6FddRMytnvhGI4HmXmaWpTAPaNHiG3sNBJPIRMFpkh5Xjt6PA3ZwUj4h9AKu0EV
# HASaL6mongns0t+W5tRst8kb0Og3bQ7OKO7le3GUOOYRAgw4bQWS5Ej5TvLtw+bt
# 4xVSb5deBW0Z9tXiIYCSuesgZKGqQAFpY4E+u8YepMj2V8iK6hFsNAmjvi254raZ
# ZyYnIWT6kBbHRHdt1BH5eIji3TY2JCKp1PmUfN+Z22RPxfUm1RVwxEQz8wcOzPru
# o9wcT38/uF+0bjCAA8MDqv5XKPOTJwhzMTuGTU7AC46Oo6FbJXVj2w9s+8O9BBZS
# 9edasdSBhGkPYAOqV9ZD9Gww1Xn7ctz0q81p1600VxDfRzL9XmN3h+kt+xxsIH+V
# Nf9T92zUQk5JpJ8g2tc+qkcPVUXpGEryxdMVjMRoZFO6tw9Gs1qDZDJukvajTEh4
# WRKv3/IzaVeXYvp7/Cm4HdZn2HcllOlCQkKWC4m5yhVhXtRo
# SIG # End signature block
