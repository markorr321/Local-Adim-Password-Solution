# ========================= Script Parameters =========================
param(
    [Parameter(HelpMessage = "Client ID of the app registration to use for delegated auth")]
    [string]$ClientId,

    [Parameter(HelpMessage = "Tenant ID to use with the specified app registration")]
    [string]$TenantId
)

# Store parameters at script scope for use in authentication functions
$script:CustomClientId = $ClientId
$script:CustomTenantId = $TenantId

# ========================= Version =========================
$script:Version = "1.0.2"

# ========================= Cross-Platform Keyboard Shortcuts =========================
$script:IsRunningOnMac = if ($null -ne $IsMacOS) { $IsMacOS } else { $PSVersionTable.OS -match 'Darwin' }

# Enable Ctrl+C as input on all platforms (prevents SIGINT, allows capture via ReadKey)
[Console]::TreatControlCAsInput = $true
Start-Sleep -Milliseconds 100
while ([Console]::KeyAvailable) {
    $null = [Console]::ReadKey($true)
}

function Test-QuitShortcut {
    param([System.ConsoleKeyInfo]$Key)
    return (($Key.Key -eq 'Q' -and ($Key.Modifiers -band [ConsoleModifiers]::Control)) -or
            ($Key.Key -eq 'C' -and ($Key.Modifiers -band [ConsoleModifiers]::Control)))
}

function Get-QuitShortcutText {
    return "Ctrl+Q Exit"
}

# ========================= Authentication =========================
$script:MSALAssemblyPaths = @{}

function Initialize-MSALAssemblies {
    <#
    .SYNOPSIS
        Loads MSAL assemblies for browser-based authentication.
    #>

    $userHome = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }

    # Try nuget cache first
    $nugetPath = Join-Path $userHome ".nuget/packages/microsoft.identity.client"
    $msalDll = $null
    $abstractionsDll = $null

    if (Test-Path $nugetPath) {
        $latestVersion = Get-ChildItem $nugetPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
        if ($latestVersion) {
            $msalDll = Join-Path $latestVersion.FullName "lib/net6.0/Microsoft.Identity.Client.dll"
            if (-not (Test-Path $msalDll)) {
                $msalDll = Join-Path $latestVersion.FullName "lib/netstandard2.0/Microsoft.Identity.Client.dll"
            }
        }

        $abstractionsPath = Join-Path $userHome ".nuget/packages/microsoft.identitymodel.abstractions"
        if (Test-Path $abstractionsPath) {
            $latestAbstractions = Get-ChildItem $abstractionsPath -Directory | Sort-Object Name -Descending | Select-Object -First 1
            if ($latestAbstractions) {
                $abstractionsDll = Join-Path $latestAbstractions.FullName "lib/net6.0/Microsoft.IdentityModel.Abstractions.dll"
                if (-not (Test-Path $abstractionsDll)) {
                    $abstractionsDll = Join-Path $latestAbstractions.FullName "lib/netstandard2.0/Microsoft.IdentityModel.Abstractions.dll"
                }
            }
        }
    }

    # Fallback to Az.Accounts
    if (-not $msalDll -or -not (Test-Path $msalDll)) {
        $LoadedAzAccountsModule = Get-Module -Name Az.Accounts
        if ($null -eq $LoadedAzAccountsModule) {
            $AzAccountsModule = Get-Module -Name Az.Accounts -ListAvailable | Select-Object -First 1
            if ($null -eq $AzAccountsModule) {
                Write-Verbose "Neither nuget cache nor Az.Accounts module found for MSAL"
                return $false
            }
            Import-Module Az.Accounts -ErrorAction SilentlyContinue -Verbose:$false
        }

        $LoadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Select-Object -ExpandProperty Location -ErrorAction SilentlyContinue
        $AzureCommon = $LoadedAssemblies | Where-Object { $_ -match "[/\\]Modules[/\\]Az.Accounts[/\\]" -and $_ -match "Microsoft.Azure.Common" }

        if ($AzureCommon) {
            $AzureCommonLocation = Split-Path -Parent $AzureCommon
            $foundMsal = Get-ChildItem -Path $AzureCommonLocation -Filter "Microsoft.Identity.Client.dll" -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
            $foundAbstractions = Get-ChildItem -Path $AzureCommonLocation -Filter "Microsoft.IdentityModel.Abstractions.dll" -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($foundMsal) { $msalDll = $foundMsal.FullName }
            if ($foundAbstractions) { $abstractionsDll = $foundAbstractions.FullName }
        }
    }

    if (-not $msalDll -or -not (Test-Path $msalDll)) {
        Write-Verbose "Could not find Microsoft.Identity.Client.dll"
        return $false
    }

    $loadedAssembliesCheck = [System.AppDomain]::CurrentDomain.GetAssemblies()

    if ($abstractionsDll -and (Test-Path $abstractionsDll)) {
        $alreadyLoaded = $loadedAssembliesCheck | Where-Object { $_.GetName().Name -eq 'Microsoft.IdentityModel.Abstractions' } | Select-Object -First 1
        if (-not $alreadyLoaded) {
            try {
                [void][System.Reflection.Assembly]::LoadFrom($abstractionsDll)
                $script:MSALAssemblyPaths['Microsoft.IdentityModel.Abstractions'] = $abstractionsDll
            } catch { }
        } else {
            $script:MSALAssemblyPaths['Microsoft.IdentityModel.Abstractions'] = $alreadyLoaded.Location
        }
    }

    $alreadyLoaded = $loadedAssembliesCheck | Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } | Select-Object -First 1
    if (-not $alreadyLoaded) {
        try {
            [void][System.Reflection.Assembly]::LoadFrom($msalDll)
            $script:MSALAssemblyPaths['Microsoft.Identity.Client'] = $msalDll
        } catch {
            Write-Verbose "Failed to load MSAL: $_"
            return $false
        }
    } else {
        $script:MSALAssemblyPaths['Microsoft.Identity.Client'] = $alreadyLoaded.Location
    }

    return $true
}

$script:MSALHelperCompiled = $false

function Initialize-MSALHelper {
    <#
    .SYNOPSIS
        Pre-compiles the MSAL helper C# code for browser-based authentication.
    #>

    if ($script:MSALHelperCompiled) { return $true }

    $referencedAssemblies = @(
        $script:MSALAssemblyPaths['Microsoft.IdentityModel.Abstractions'],
        $script:MSALAssemblyPaths['Microsoft.Identity.Client']
    ) | Where-Object { $_ }

    if ($referencedAssemblies.Count -lt 1) {
        throw "Missing required MSAL assemblies"
    }

    $referencedAssemblies += @("netstandard", "System.Linq", "System.Threading.Tasks", "System.Collections")

    $code = @"
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

public class LAPSBrowserAuth
{
    public static string GetAccessToken(string clientId, string[] scopes, string tenantId = null)
    {
        try
        {
            var task = Task.Run(async () => await GetAccessTokenAsync(clientId, scopes, tenantId));
            if (task.Wait(TimeSpan.FromSeconds(180)))
            {
                return task.Result;
            }
            throw new TimeoutException("Authentication timed out");
        }
        catch (AggregateException ae)
        {
            if (ae.InnerException != null) throw ae.InnerException;
            throw;
        }
    }

    private static async Task<string> GetAccessTokenAsync(string clientId, string[] scopes, string tenantId)
    {
        var builder = PublicClientApplicationBuilder.Create(clientId)
            .WithRedirectUri("http://localhost");

        if (!string.IsNullOrEmpty(tenantId))
        {
            builder = builder.WithAuthority(string.Format("https://login.microsoftonline.com/{0}", tenantId));
        }

        IPublicClientApplication publicClientApp = builder.Build();

        using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(180)))
        {
            var webViewOptions = new SystemWebViewOptions
            {
                HtmlMessageSuccess = @"
<html>
<head>
    <meta charset='UTF-8'>
    <title>Authentication Successful - LAPS</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #1a5276 0%, #2e86c1 100%); }
        .container { text-align: center; color: white; }
        .brand { font-size: 14px; letter-spacing: 4px; margin-bottom: 30px; opacity: 0.9; }
        .checkmark { font-size: 64px; margin-bottom: 20px; }
        h1 { margin: 0 0 10px 0; font-weight: 300; font-size: 28px; }
        p { margin: 0; opacity: 0.9; font-size: 16px; }
    </style>
</head>
<body>
    <div class='container'>
        <div class='brand'>[ L A P S ]</div>
        <div class='checkmark'>&#10003;</div>
        <h1>Authentication Successful</h1>
        <p>You can close this window and return to PowerShell.</p>
    </div>
</body>
</html>",
                HtmlMessageError = @"
<html>
<head>
    <meta charset='UTF-8'>
    <title>Authentication Failed - LAPS</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%); }
        .container { text-align: center; color: white; }
        .brand { font-size: 14px; letter-spacing: 4px; margin-bottom: 30px; opacity: 0.9; }
        .icon { font-size: 64px; margin-bottom: 20px; }
        h1 { margin: 0 0 10px 0; font-weight: 300; font-size: 28px; }
        p { margin: 0; opacity: 0.9; font-size: 16px; }
    </style>
</head>
<body>
    <div class='container'>
        <div class='brand'>[ L A P S ]</div>
        <div class='icon'>&#10005;</div>
        <h1>Authentication Failed</h1>
        <p>Please close this window and try again.</p>
    </div>
</body>
</html>"
            };

            var tokenBuilder = publicClientApp.AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount)
                .WithUseEmbeddedWebView(false)
                .WithSystemWebViewOptions(webViewOptions);

            if (!string.IsNullOrEmpty(tenantId))
            {
                tokenBuilder = tokenBuilder.WithExtraQueryParameters(string.Format("domain_hint={0}", tenantId));
            }

            var result = await tokenBuilder
                .ExecuteAsync(cts.Token)
                .ConfigureAwait(false);

            return result.AccessToken;
        }
    }
}
"@

    try {
        $null = [LAPSBrowserAuth]
        $script:MSALHelperCompiled = $true
        return $true
    } catch { }

    Add-Type -ReferencedAssemblies $referencedAssemblies -TypeDefinition $code -Language CSharp -ErrorAction Stop -IgnoreWarnings 3>$null

    $script:MSALHelperCompiled = $true
    return $true
}

function Get-BrowserAccessToken {
    param(
        [string[]]$Scopes
    )

    if (-not $script:MSALHelperCompiled) {
        $null = Initialize-MSALHelper
    }

    # Use custom ClientId if provided, otherwise use Microsoft's well-known PowerShell public client ID
    $clientId = if ($script:CustomClientId) { $script:CustomClientId } else { "14d82eec-204b-4c2f-b7e8-296a70dab67e" }
    $tenantId = $script:CustomTenantId

    $scopeArray = $Scopes | ForEach-Object {
        if ($_ -notlike "https://*") { "https://graph.microsoft.com/$_" } else { $_ }
    }

    $accessToken = [LAPSBrowserAuth]::GetAccessToken($clientId, $scopeArray, $tenantId)
    return $accessToken
}

# ========================= Graph API Helpers =========================
$script:AccessToken = $null

function Connect-LAPSGraph {
    <#
    .SYNOPSIS
        Authenticates to Microsoft Graph with LAPS-required scopes.
    #>

    try {
        if ($script:CustomClientId) {
            Write-Host "Using custom app registration..." -ForegroundColor Cyan
            Write-Host "  Client ID: $($script:CustomClientId)" -ForegroundColor Gray
            if ($script:CustomTenantId) {
                Write-Host "  Tenant ID: $($script:CustomTenantId)" -ForegroundColor Gray
            }
        } else {
            Write-Host "Using default Microsoft Graph authentication..." -ForegroundColor Cyan
        }

        Write-Host "Opening browser for authentication..." -ForegroundColor Cyan

        if ($script:MSALHelperCompiled) {
            Write-Host "Waiting for authentication response..." -ForegroundColor Yellow
            $script:AccessToken = Get-BrowserAccessToken -Scopes @(
                "Device.Read.All",
                "DeviceLocalCredential.Read.All",
                "DeviceManagementManagedDevices.PrivilegedOperations.All"
            )
            if ($script:AccessToken) {
                # Decode JWT to verify granted scopes
                $tokenParts = $script:AccessToken.Split('.')
                $payload = $tokenParts[1]
                # Fix base64 padding
                switch ($payload.Length % 4) {
                    2 { $payload += '==' }
                    3 { $payload += '=' }
                }
                $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload)) | ConvertFrom-Json
                $grantedScopes = $decoded.scp -split ' '

                if ($grantedScopes -notcontains 'DeviceLocalCredential.Read.All') {
                    Write-Host ""
                    Write-Host "  WARNING: DeviceLocalCredential.Read.All scope was NOT granted." -ForegroundColor Red
                    Write-Host "  Your app registration needs this delegated permission with admin consent." -ForegroundColor Yellow
                    Write-Host ""
                }

                Write-Host "  Connected" -ForegroundColor Green
                return $true
            } else {
                throw "Failed to get access token"
            }
        } else {
            throw "MSAL helper not initialized"
        }
    } catch {
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Invoke-LAPSGraphRequest {
    <#
    .SYNOPSIS
        Makes an authenticated request to Microsoft Graph API.
    #>
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [hashtable]$ExtraHeaders = @{}
    )

    $headers = @{
        'Authorization' = "Bearer $($script:AccessToken)"
        'Content-Type'  = 'application/json'
        'ocp-client-name' = 'LAPS-PowerShell'
        'ocp-client-version' = $script:Version
    }

    foreach ($key in $ExtraHeaders.Keys) {
        $headers[$key] = $ExtraHeaders[$key]
    }

    try {
        $response = Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -ErrorAction Stop
        return $response
    } catch {
        # PowerShell 7 stores the response body in ErrorDetails.Message
        $errorDetail = $null
        $errorCode = $null
        $statusCode = $null
        try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { }
        try {
            $errorBody = $_.ErrorDetails.Message | ConvertFrom-Json
            $errorDetail = $errorBody.error.message
            $errorCode = $errorBody.error.code
        } catch { }

        if ($statusCode -eq 401 -or $statusCode -eq 403) {
            $msg = "Access denied ($statusCode)"
            if ($errorDetail) { $msg += ": $errorDetail" }
            throw $msg
        }
        if ($errorDetail) {
            throw "$errorCode ($statusCode): $errorDetail"
        }
        throw $_.Exception.Message
    }
}

# ========================= LAPS Functions =========================

function Search-Device {
    <#
    .SYNOPSIS
        Searches for devices in Entra ID by display name.
    #>
    param(
        [string]$SearchTerm
    )

    # Advanced queries on devices require ConsistencyLevel: eventual and $count=true
    $filter = [System.Uri]::EscapeDataString("startsWith(displayName,'$SearchTerm')")
    $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=$filter&`$select=id,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled&`$count=true&`$top=50&`$orderby=displayName"

    try {
        $result = Invoke-LAPSGraphRequest -Uri $uri -ExtraHeaders @{ 'ConsistencyLevel' = 'eventual' }
        return $result.value
    } catch {
        Write-Host "  Search failed: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function Get-DeviceLAPSCredential {
    <#
    .SYNOPSIS
        Retrieves LAPS credentials for a device using the /directory/ endpoint.
        Looks up by device name first to get the correct credential info ID.
    #>
    param(
        [string]$DeviceName
    )

    # Step 1: Find the deviceLocalCredentialInfo by device name
    $filter = [System.Uri]::EscapeDataString("deviceName eq '$DeviceName'")
    $listUri = "https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials?`$filter=$filter"

    try {
        $listResult = Invoke-LAPSGraphRequest -Uri $listUri
        if (-not $listResult.value -or $listResult.value.Count -eq 0) {
            throw "No LAPS credentials found for device '$DeviceName'"
        }

        $credentialInfoId = $listResult.value[0].id

        # Step 2: Get the full credential details including passwords
        $uri = "https://graph.microsoft.com/v1.0/directory/deviceLocalCredentials/${credentialInfoId}?`$select=credentials"
        $result = Invoke-LAPSGraphRequest -Uri $uri
        return $result
    } catch {
        throw $_.Exception.Message
    }
}

function Get-IntuneManagedDeviceId {
    <#
    .SYNOPSIS
        Looks up the Intune managed device ID by device name.
    #>
    param(
        [string]$DeviceName
    )

    $filter = [System.Uri]::EscapeDataString("deviceName eq '$DeviceName'")
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter&`$select=id,deviceName"

    $result = Invoke-LAPSGraphRequest -Uri $uri
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0].id
    }
    return $null
}

function Invoke-LAPSPasswordRotation {
    <#
    .SYNOPSIS
        Triggers an on-demand LAPS password rotation for an Intune-managed device.
    #>
    param(
        [string]$ManagedDeviceId
    )

    # rotateLocalAdminPassword is only available on the beta endpoint
    $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/${ManagedDeviceId}/rotateLocalAdminPassword"
    Invoke-LAPSGraphRequest -Uri $uri -Method "POST"
}

# ========================= TUI Functions =========================

# Dynamic Control Bar System
$script:LastControlBarLine = -1

function Show-DynamicControlBar {
    param(
        [string]$ControlsText,
        [switch]$Force
    )

    $currentTop = [Console]::CursorTop

    # Clear previous control bar if it exists and is at a different line
    if ($script:LastControlBarLine -ge 0 -and $script:LastControlBarLine -ne ($currentTop + 1)) {
        try {
            [Console]::SetCursorPosition(0, $script:LastControlBarLine)
            Write-Host (" " * [Console]::WindowWidth) -NoNewline
        } catch { }
    }

    $targetTop = $currentTop + 1

    # Ensure buffer is tall enough
    if ($targetTop -ge [Console]::BufferHeight) {
        [Console]::BufferHeight = $targetTop + 2
    }

    # Draw the control bar below current content
    try {
        [Console]::SetCursorPosition(0, $targetTop)
        Write-Host "  $ControlsText" -ForegroundColor DarkGray
        $script:LastControlBarLine = $targetTop

        # Always return cursor above the control bar
        [Console]::SetCursorPosition(0, $currentTop)
    } catch { }
}

function Write-LAPSHost {
    <#
    .SYNOPSIS
        Write-Host wrapper that keeps the dynamic control bar below content.
    #>
    param(
        [string]$Object = "",
        [ConsoleColor]$ForegroundColor,
        [switch]$NoNewline,
        [string]$ControlsText = $null
    )

    $params = @{}
    if ($ForegroundColor) { $params['ForegroundColor'] = $ForegroundColor }
    if ($NoNewline) { $params['NoNewline'] = $true }

    Write-Host $Object @params

    if ($ControlsText) {
        Show-DynamicControlBar -ControlsText $ControlsText
    }
}

function Show-LAPSHeader {
    Write-Host "[ L A P S ]" -ForegroundColor Cyan -NoNewline
    Write-Host "  v$script:Version" -ForegroundColor DarkGray
    Write-Host "    Local Administrator Password Solution" -ForegroundColor DarkGray
    Write-Host "    with " -ForegroundColor DarkGray -NoNewline
    Write-Host "PowerShell" -ForegroundColor Blue
}

function Show-LAPSHeaderMinimal {
    Write-Host "[ L A P S ]" -ForegroundColor Cyan -NoNewline
    Write-Host "  v$script:Version" -ForegroundColor DarkGray
}

function Invoke-LAPSExit {
    param(
        [string]$Message = "Exiting..."
    )

    [Console]::CursorVisible = $true
    [Console]::TreatControlCAsInput = $false
    Clear-Host

    Show-LAPSHeaderMinimal
    Write-Host ""

    # Clear the access token
    $script:AccessToken = $null

    Write-Host "  Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "  Disconnected from Microsoft Graph." -ForegroundColor Green
    } catch {
        Write-Host "  Already disconnected from Microsoft Graph." -ForegroundColor DarkGray
    }

    Write-Host ""
    exit 0
}

function Test-GlobalShortcut {
    param(
        [System.ConsoleKeyInfo]$Key
    )

    if (Test-QuitShortcut -Key $Key) {
        Invoke-LAPSExit
        return $true
    }

    return $false
}

function Read-LAPSInput {
    <#
    .SYNOPSIS
        Enhanced input reader with global shortcut handling.
    #>
    param(
        [string]$Prompt,
        [switch]$Required,
        [string]$ControlsText
    )

    [Console]::CursorVisible = $true

    Write-Host "${Prompt}: " -ForegroundColor Cyan -NoNewline

    # Show control bar below the prompt
    if ($ControlsText) {
        $inputLeft = [Console]::CursorLeft
        $inputTop = [Console]::CursorTop
        Write-Host ""
        Write-Host "  $ControlsText" -ForegroundColor DarkGray
        $script:LastControlBarLine = [Console]::CursorTop - 1
        [Console]::SetCursorPosition($inputLeft, $inputTop)
    }

    $userInput = ""
    do {
        $key = [Console]::ReadKey($true)

        if (Test-GlobalShortcut -Key $key) {
            return $null
        }

        if ($key.Key -eq 'Escape') {
            Write-Host ""
            return $null
        }

        if ($key.Key -eq 'Enter') {
            Write-Host ""
            break
        }

        if ($key.Key -eq 'Backspace' -and $userInput.Length -gt 0) {
            $userInput = $userInput.Substring(0, $userInput.Length - 1)
            Write-Host "`b `b" -NoNewline
        }
        elseif ($key.KeyChar -ne "`0" -and [char]::IsControl($key.KeyChar) -eq $false) {
            $userInput += $key.KeyChar
            Write-Host $key.KeyChar -NoNewline -ForegroundColor White
        }
    } while ($true)

    if ($Required -and [string]::IsNullOrWhiteSpace($userInput)) {
        Write-Host "Input is required." -ForegroundColor DarkRed
        return Read-LAPSInput -Prompt $Prompt -Required:$Required
    }

    return $userInput
}

function Show-SimpleMenu {
    <#
    .SYNOPSIS
        Arrow-key driven single-selection menu.
    #>
    param(
        [array]$Items,
        [string]$Title = "Select an option",
        [int]$DefaultSelection = 0
    )

    if ($Items.Count -eq 0) {
        Write-Host "No items to select from." -ForegroundColor Red
        return -1
    }

    $currentIndex = $DefaultSelection
    [Console]::CursorVisible = $false

    try {
        do {
            Clear-Host
            Show-LAPSHeader
            Write-Host ""
            Write-Host $Title -ForegroundColor Cyan
            Write-Host ""

            for ($i = 0; $i -lt $Items.Count; $i++) {
                $arrow = if ($i -eq $currentIndex) { ">" } else { " " }
                $color = if ($i -eq $currentIndex) { "Yellow" } else { "White" }
                Write-Host " $arrow $($Items[$i])" -ForegroundColor $color
            }

            Write-Host ""
            Write-Host "  Up/Down Navigate | ENTER Select | $(Get-QuitShortcutText)" -ForegroundColor DarkGray

            $key = [Console]::ReadKey($true)

            if (Test-GlobalShortcut -Key $key) { return -1 }

            switch ($key.Key) {
                "UpArrow" {
                    $currentIndex = if ($currentIndex -gt 0) { $currentIndex - 1 } else { $Items.Count - 1 }
                }
                "DownArrow" {
                    $currentIndex = if ($currentIndex -lt ($Items.Count - 1)) { $currentIndex + 1 } else { 0 }
                }
                "Enter" {
                    return $currentIndex
                }
                "Escape" {
                    return -1
                }
            }
        } while ($true)
    } finally {
        [Console]::CursorVisible = $true
    }
}

function Show-DeviceSelectionMenu {
    <#
    .SYNOPSIS
        Displays device search results in a selectable menu.
    #>
    param(
        [array]$Devices
    )

    $currentIndex = 0
    [Console]::CursorVisible = $false

    try {
        do {
            Clear-Host
            Show-LAPSHeaderMinimal
            Write-Host ""
            Write-Host "  Select a device ($($Devices.Count) found)" -ForegroundColor Cyan
            Write-Host ""

            for ($i = 0; $i -lt $Devices.Count; $i++) {
                $device = $Devices[$i]
                $arrow = if ($i -eq $currentIndex) { ">" } else { " " }

                $enabled = if ($device.accountEnabled) { "Enabled" } else { "Disabled" }
                $enabledColor = if ($device.accountEnabled) { "Green" } else { "Red" }
                $trust = if ($device.trustType) { $device.trustType } else { "Unknown" }
                $os = if ($device.operatingSystem) { "$($device.operatingSystem) $($device.operatingSystemVersion)" } else { "Unknown OS" }

                if ($i -eq $currentIndex) {
                    Write-Host " $arrow " -NoNewline -ForegroundColor Yellow
                    Write-Host "$($device.displayName)" -NoNewline -ForegroundColor Yellow
                    Write-Host "  $os" -NoNewline -ForegroundColor DarkGray
                    Write-Host "  $trust" -NoNewline -ForegroundColor DarkGray
                    Write-Host "  $enabled" -ForegroundColor $enabledColor
                } else {
                    Write-Host " $arrow " -NoNewline
                    Write-Host "$($device.displayName)" -NoNewline -ForegroundColor White
                    Write-Host "  $os" -NoNewline -ForegroundColor DarkGray
                    Write-Host "  $trust" -NoNewline -ForegroundColor DarkGray
                    Write-Host "  $enabled" -ForegroundColor $enabledColor
                }
            }

            Write-Host ""
            Write-Host "  Up/Down Navigate | ENTER Select | ESC Back | $(Get-QuitShortcutText)" -ForegroundColor DarkGray

            $key = [Console]::ReadKey($true)

            if (Test-GlobalShortcut -Key $key) { return $null }

            switch ($key.Key) {
                "UpArrow" {
                    $currentIndex = if ($currentIndex -gt 0) { $currentIndex - 1 } else { $Devices.Count - 1 }
                }
                "DownArrow" {
                    $currentIndex = if ($currentIndex -lt ($Devices.Count - 1)) { $currentIndex + 1 } else { 0 }
                }
                "Enter" {
                    return $Devices[$currentIndex]
                }
                "Escape" {
                    return $null
                }
            }
        } while ($true)
    } finally {
        [Console]::CursorVisible = $true
    }
}

function Show-PasswordResult {
    <#
    .SYNOPSIS
        Displays the LAPS password result with copy-to-clipboard option.
    #>
    param(
        [object]$Device,
        [object]$Credential
    )

    $controlsText = "Ctrl+C Copy | R Rotate | S New Search | Ctrl+Q Exit"
    $noCredsControlsText = "S New Search | Ctrl+Q Exit"

    # Decode credential data once outside the loop
    $password = $null
    $accountName = $null
    $backupTime = $null
    $hasCreds = $false

    if ($Credential.credentials -and $Credential.credentials.Count -gt 0) {
        $hasCreds = $true
        $latestCred = $Credential.credentials | Sort-Object { [datetime]$_.backupDateTime } -Descending | Select-Object -First 1

        $accountName = $latestCred.accountName
        $passwordRaw = $latestCred.passwordBase64
        $password = if ($passwordRaw) {
            try {
                $bytes = [System.Convert]::FromBase64String($passwordRaw)
                $utf8 = [System.Text.Encoding]::UTF8.GetString($bytes)
                if ($utf8 -match '[^\x20-\x7E]') {
                    [System.Text.Encoding]::ASCII.GetString($bytes)
                } else {
                    $utf8
                }
            } catch {
                $passwordRaw
            }
        } else { "N/A" }

        $backupTime = if ($latestCred.backupDateTime) {
            try { ([datetime]$latestCred.backupDateTime).ToString("yyyy-MM-dd h:mm:ss tt") } catch { $latestCred.backupDateTime }
        } else { "N/A" }
    }

    do {
        Clear-Host
        $script:LastControlBarLine = -1
        Show-LAPSHeaderMinimal
        Write-Host ""
        Write-Host "  LAPS Credential Retrieved" -ForegroundColor Green
        Write-Host "  $("=" * 40)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Device:       " -NoNewline -ForegroundColor Gray
        Write-Host "$($Device.displayName)" -ForegroundColor White

        if ($hasCreds) {
            Write-Host ""
            Write-Host "  Account:      " -NoNewline -ForegroundColor Gray
            Write-Host "$accountName" -ForegroundColor Yellow
            Write-Host "  Password:     " -NoNewline -ForegroundColor Gray
            Write-Host "$password" -ForegroundColor Green
            Write-Host ""
            Write-Host "  Last Rotated: " -NoNewline -ForegroundColor Gray
            Write-Host "$backupTime" -ForegroundColor White
            Write-Host ""
            Write-Host "  $("-" * 40)" -ForegroundColor DarkGray

            # Remember where content ends so we can write below it
            $script:ContentLine = [Console]::CursorTop

            # Show dynamic control bar
            Show-DynamicControlBar -ControlsText $controlsText

            # Position cursor back at content line for any action output
            [Console]::SetCursorPosition(0, $script:ContentLine)

            $key = [Console]::ReadKey($true)

            # On password screen, Ctrl+C copies instead of quitting
            if ($key.Key -eq 'C' -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
                try {
                    Set-Clipboard -Value $password
                    Write-Host "  Password copied to clipboard!" -ForegroundColor Green
                    Show-DynamicControlBar -ControlsText $controlsText
                    Start-Sleep -Seconds 1
                } catch {
                    Write-Host "  Failed to copy: $_" -ForegroundColor Red
                    Show-DynamicControlBar -ControlsText $controlsText
                    Start-Sleep -Seconds 2
                }
            }
            elseif ($key.Key -eq 'R') {
                Write-Host "  Rotate password for $($Device.displayName)? (Y/N): " -ForegroundColor Yellow -NoNewline
                # Save cursor position at end of Y/N prompt
                $ynLeft = [Console]::CursorLeft
                $ynTop = [Console]::CursorTop
                Write-Host ""  # blank line between prompt and control bar
                Show-DynamicControlBar -ControlsText $controlsText
                # Restore cursor to right after "(Y/N): " so input appears inline
                [Console]::SetCursorPosition($ynLeft, $ynTop)
                $confirmKey = [Console]::ReadKey($true)
                if ($confirmKey.Key -eq 'Y') {
                    Write-Host "Y"
                    Write-Host "  Requesting password rotation..." -ForegroundColor Gray
                    try {
                        $managedDeviceId = Get-IntuneManagedDeviceId -DeviceName $Device.displayName
                        if (-not $managedDeviceId) {
                            Write-Host "  Device not found in Intune. Rotation requires an Intune-managed device." -ForegroundColor Red
                            Show-DynamicControlBar -ControlsText $controlsText
                            Start-Sleep -Seconds 3
                        } else {
                            Invoke-LAPSPasswordRotation -ManagedDeviceId $managedDeviceId
                            Write-Host "  Password rotation initiated!" -ForegroundColor Green
                            Write-Host "  The new password will appear after the device checks in." -ForegroundColor DarkGray
                            Show-DynamicControlBar -ControlsText $controlsText
                            Start-Sleep -Seconds 3
                        }
                    } catch {
                        Write-Host "  Rotation failed: $($_.Exception.Message)" -ForegroundColor Red
                        Show-DynamicControlBar -ControlsText $controlsText
                        Start-Sleep -Seconds 3
                    }
                } else {
                    Write-Host "N"
                }
            }
            elseif ($key.Key -eq 'Q' -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
                return 'Quit'
            }
            elseif ($key.Key -eq 'S') {
                return 'Search'
            }
            elseif ($key.Key -eq 'Escape') {
                return 'Search'
            }
        } else {
            Write-Host ""
            Write-Host "  No LAPS credentials found for this device." -ForegroundColor Yellow
            Write-Host "  The device may not have LAPS configured or credentials have not been backed up." -ForegroundColor DarkGray

            Show-DynamicControlBar -ControlsText $noCredsControlsText

            $key = [Console]::ReadKey($true)

            if (Test-GlobalShortcut -Key $key) { return 'Quit' }

            if ($key.Key -eq 'S' -or $key.Key -eq 'Escape' -or $key.Key -eq 'Enter') {
                return 'Search'
            }
        }
    } while ($true)
}

# ========================= Prerequisite Checks =========================

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Checks and installs required dependencies.
    #>

    $requiresInstall = $false

    # Check for Microsoft.Graph.Authentication (needed for device queries via SDK fallback)
    # We use direct REST calls, so we mainly need MSAL
    Write-Host "Checking prerequisites..." -ForegroundColor Gray

    # Initialize MSAL assemblies
    $msalReady = Initialize-MSALAssemblies
    if (-not $msalReady) {
        Write-Host "  MSAL assemblies not found. Attempting to install Az.Accounts..." -ForegroundColor Yellow
        try {
            if (Get-Command Install-PSResource -ErrorAction SilentlyContinue) {
                Install-PSResource -Name Az.Accounts -Scope CurrentUser -TrustRepository -Confirm:$false -ErrorAction Stop
            } else {
                Install-Module -Name Az.Accounts -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
            Import-Module Az.Accounts -ErrorAction SilentlyContinue -Verbose:$false
            $msalReady = Initialize-MSALAssemblies
        } catch {
            Write-Host "  Failed to install Az.Accounts: $_" -ForegroundColor Red
            return $false
        }
    }

    if (-not $msalReady) {
        Write-Host "  Could not load MSAL authentication libraries." -ForegroundColor Red
        Write-Host "  Please install Az.Accounts: Install-Module Az.Accounts -Scope CurrentUser" -ForegroundColor Yellow
        return $false
    }

    # Pre-compile MSAL helper
    try {
        $null = Initialize-MSALHelper
        Write-Host "  Prerequisites OK" -ForegroundColor Green
    } catch {
        Write-Host "  Failed to initialize authentication: $_" -ForegroundColor Red
        return $false
    }

    return $true
}

# ========================= Main Application Loop =========================

function Start-LAPSApp {
    <#
    .SYNOPSIS
        Main application entry point - runs the LAPS TUI loop.
    #>

    Clear-Host
    Show-LAPSHeader
    Write-Host ""

    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        Write-Host ""
        Write-Host "Press Enter to exit" -ForegroundColor Gray
        $null = [Console]::ReadLine()
        return
    }

    Write-Host ""

    # Authenticate
    $connected = Connect-LAPSGraph
    if (-not $connected) {
        Write-Host ""
        Write-Host "Authentication failed. Press Enter to exit" -ForegroundColor Red
        $null = [Console]::ReadLine()
        return
    }

    # Main search loop
    while ($true) {
        Clear-Host
        Show-LAPSHeaderMinimal
        Write-Host ""

        $searchTerm = Read-LAPSInput -Prompt "  Search device name" -Required -ControlsText "ESC Back | Ctrl+Q Exit"

        # Clear the control bar left behind by Read-LAPSInput
        if ($script:LastControlBarLine -ge 0) {
            try {
                [Console]::SetCursorPosition(0, $script:LastControlBarLine)
                Write-Host (" " * [Console]::WindowWidth) -NoNewline
                [Console]::SetCursorPosition(0, $script:LastControlBarLine)
                $script:LastControlBarLine = -1
            } catch { }
        }

        if ($null -eq $searchTerm) {
            Invoke-LAPSExit
            return
        }

        # Empty input (just Enter) - treat as back/exit
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            Invoke-LAPSExit
            return
        }

        # Search for devices
        Write-Host "  Searching..." -ForegroundColor Gray

        $devices = Search-Device -SearchTerm $searchTerm

        if ($devices.Count -eq 0) {
            Write-Host "  No devices found matching '$searchTerm'" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Press any key to search again..." -ForegroundColor DarkGray
            $key = [Console]::ReadKey($true)
            if (Test-GlobalShortcut -Key $key) { return }
            continue
        }

        # Show device selection menu
        $selectedDevice = Show-DeviceSelectionMenu -Devices $devices
        if ($null -eq $selectedDevice) {
            continue
        }

        # Retrieve LAPS credential
        Clear-Host
        Show-LAPSHeaderMinimal
        Write-Host ""
        Write-Host "  Retrieving LAPS credential for $($selectedDevice.displayName)..." -ForegroundColor Gray

        try {
            $credential = Get-DeviceLAPSCredential -DeviceName $selectedDevice.displayName
            $action = Show-PasswordResult -Device $selectedDevice -Credential $credential
            if ($action -eq 'Quit') { Invoke-LAPSExit; return }
            # 'Search' continues the loop
        } catch {
            Write-Host ""
            Write-Host "  Failed to retrieve LAPS credential: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "  This could mean:" -ForegroundColor DarkGray
            Write-Host "    - LAPS is not configured for this device" -ForegroundColor DarkGray
            Write-Host "    - You lack DeviceLocalCredential.Read.All permission" -ForegroundColor DarkGray
            Write-Host "    - The device has not backed up credentials yet" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Press any key to search again..." -ForegroundColor DarkGray
            $key = [Console]::ReadKey($true)
            if (Test-GlobalShortcut -Key $key) { return }
        }
    }
}

# ========================= Entry Point =========================
Start-LAPSApp
# SIG # Begin signature block
# MIIsCQYJKoZIhvcNAQcCoIIr+jCCK/YCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD66XsWKm2oKiyT
# 2iUKTjhBGainJRrt3OuglgdH9DcTlqCCJRowggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# LwYJKoZIhvcNAQkEMSIEIIG/QCtFOU1GJVhozKFlyGCBT7JwsBeB8URaBTjB/qQ9
# MA0GCSqGSIb3DQEBAQUABIICABkijVThiqxilmYOZtHavfCRBCqa51RtbczUOeVS
# 7iGXkIT6bfBxfaj5D9naA2Wx/dzjTFlnNyXqLCkHR7atZV6qpVGA09SMeOzrpdhU
# BgtvluzqAC+/fLGt7w6XhCuVG0nf1a4Vi9HAdNXUbCvbnd3xQWmvp+Okq//h4xNY
# fL9bf/pwt5eMRBW7C1cqrla9E+e7OZR9sUf0pQgO6MSFVCA6dNQvUHYnh5O6iOw1
# JqVdKriX2TUG6rASxrnHF2/ilpOXAzXKPseYDuNYKfJKVU1YKJ3jFtp0UJYgiDDx
# kiiy7yTDIcI5CNi4LL7HCkOgXhiINBQ+qch/A5QLYZ3PGefw6SZtxkVJkvBAtTdj
# oFGXkqorVCb4/agcO2f3htZ4/f96ea7ZHRPwKsHE2GWxvonuRt5gRAv1X8BkcNaY
# dfFlgIcWkdXD1n9WVfqhTpEcxmvQIkyDb3ABfUWH3MZk0CbNjJ/Cs8C5RhYji0iA
# GyAMt3Zs84nhu8L1h/gEbjtzzNGpn29v8Yz/GgnsF34agX7BrW5jA1wmIdKd/Q8T
# CamjfN016UkuWfw9UrMMplIRSaLI74cdFSOKumsIyVoy+zJ5uMtEmwLofMwxdjgU
# YnZ7pET9f2Vhjk/xAGdwEv/NxYnPr46V7OSsZa1L2e6B/62YJx2tSXiI0FuzKRue
# FmnioYIDJjCCAyIGCSqGSIb3DQEJBjGCAxMwggMPAgEBMH0waTELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBU
# cnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMQIQ
# CoDvGEuN8QWC0cR2p5V0aDANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI2MDQxNDExMDEzMFowLwYJKoZI
# hvcNAQkEMSIEIJ77snq7V9Ppk9C4R3pK8iI7he/jaAog5eTPPJ3CG8AlMA0GCSqG
# SIb3DQEBAQUABIICAEHoGYQZ6P0UQ6QSmGZy0DGe3645r4juulEVUuILwokMN848
# nUIH5C9gqaLSMpoM8eJXppyHVw550/fklocS8LYWrwhxgbu5uT185Po7epJabU7L
# Z+xXS3hZ22YkazNTRKF2Mq2A7abEGUeY0lmZzjF7aOExD0x83kMADn/JDWrQfCVm
# CX+jljt7KnEyyfh5f4I4ER8hradPvVKG1wtTHK1QhYjofzihn3x7hmkJaSn9RWJC
# lrGZYjbCxn5eGvxM4O7xxEUqHK/84Afv+YKzMZcmyxf7TYUgnDgCBywdrPBZ8UYi
# jq3+nV0ER4OYiA26kNIQpcj3NufxYcTAhs/03CUtkLaDMiu/ua4rEekrQ+w+i3AP
# AYGEOlQ5/egMGRJejvFkkI+tDVoFQU6f1Ev641ABInh5D7ugNzN/z+VM4Xyn3Oy/
# XjuW3A1V2RdvkLNRIBj2N4l+qED/Eet3Iii+nWCQhVFrJRXVqw5WGlI8gNMacBgS
# +Lsrq4aud/YNRhQXaOSsmghFUB9bi61QJ9EqQxfizXQ7qy9c27rieGcu/ZjAiNBT
# 59DBak93xaHKwkPzl1YyAtfST3avDWCinHT0OtqAIekPb8XgGA6FSy+IyhZ5HnV1
# LDIlrlN7YRCBHHxS6gt589SyemoPWdCy/A+BRnlhGZEyYqaGNOWwdl8u77iI
# SIG # End signature block
