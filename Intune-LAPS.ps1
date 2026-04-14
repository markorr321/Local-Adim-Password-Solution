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
$script:Version = "1.0.0"

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
# MIIr0AYJKoZIhvcNAQcCoIIrwTCCK70CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDNegtycTnnNfRN
# Heo1Updn2bIKWVsPHLs71enIPBzJoqCCJOQwggVvMIIEV6ADAgECAhBI/JO0YFWU
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
# hvcNAQkEMSIEICaSSstCjDDoIXXWuB7Zpw1t7A6u7JjIA/6PbDfin6+PMA0GCSqG
# SIb3DQEBAQUABIICAAx5w6F0sH0QQtR2gEPHgIAL98Up8TlKw5BwApHJ3xUeYWNJ
# A5IIWEZDk5OFdf1LMyLRWAq/BWjpvEw0eDBQHnmkJWoz5w1pBib/a01lfayahYHg
# HrBUYlaz2vFr1DmicptQfezrc9ITgz+sCVp5eG5+LP/MGwA6wd7IJ3wd8bPxO4jZ
# 5GaKd6q7XNFG8FtTHZP7zXACg28trkS06cMBO0XJWEqJDlurN99Fxib10td6wweO
# uAFo9IVam6IPQ5vnLuqrtUY+Bgrl+4AlOWuxjBioRJ8dDKKgzi9rhBSesB8+rH57
# GfE1VY72MXTIEXmZOPje+8mu37PwxSAfnkYyBPnOns2E7AOEif+zIDsVKBuvZHC7
# Mg7TXLoDdDLCbN7xfqmIah1ZVQ/8ihUHXbsfV0H+59anDynBna2bs6KwVzfwZya9
# PcrKf7icSwn82nGS/vn4tHYVo2ponzb+0MaaFtjPJzg0u78mkDA1MhjVz+AgeFiQ
# +Dlm1wcDrtghP1F+i7d6A/R7RDSUW6Lb/p4u4ov2HmsfvO6CcgEpgevG7IILHf+p
# 8UR5tI+n4A8hRPYx/YUoHSQ56PXuBArazLMgiLBf5ScXjJg1KEjAIphILtH5iU8E
# ksfja0PM9wHSIkPuDNihslWqliIizfiSWbzu68mp50MoxyhylJRlGL/1120/oYID
# IzCCAx8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAW
# BgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMg
# VGltZSBTdGFtcGluZyBDQSBSMzYCEQCkKTtuHt3XpzQIh616TrckMA0GCWCGSAFl
# AwQCAgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjYwNDE0MDcyODE3WjA/BgkqhkiG9w0BCQQxMgQwoiBDWF23YJpqpS5UwnUY
# a0uhgKAmLKO/s6zhcuRE8XIeUl4aH4l8OJePdmXbnHGCMA0GCSqGSIb3DQEBAQUA
# BIICABzI82YoHKjgDD5mkJPHm2loktaBur3R3gqdXsW8n/VmXnmKuFmG23pp0avC
# me7UGnNZumOMBIuV/UoGFPK2nCvjmmf95U0Sz1ITFdIKSQQcZT9OxIFyemr3f52i
# T1pvec2Qe7ZulwisnnvwidFu4oD7sgSG9vqS2AUzJsavvT0R0kVNc/BFhxLmw3Nq
# 37nlZMwlva4/2ySXkHE2aqtJ0sCOZVaQKyWQgcDig9l8xoMkeJD4XhJ4xra7ng9m
# ktnCsWNU4vylFe1k/MFhYO/1+MQnx58njSH35uZ5GVMFgPGcfoP50u39kY8+Hdir
# 9ldRGdfaBSM3X1mSovG8lX04sRPoFX3sN/U5hglmi0mIu+GFGx9ablEZlybzxETu
# yb/tv6zwtR9RK2arcuiYEMWJGuMLuGTK4EvNt1eSBvElc8KTf9GKarmpUMQyzdQd
# Hbk1gU+d0ydiaat6FY6FkaghDRqhnVWnK8ivlEXIpIywj2quvaaNELnolmWBrq/D
# SsPaxnkxJLdyMzkr1odw49yPGBCV4VlQZPeOwHldwWIimSLGWsTVECIVgR7kl4RI
# VIReHT6giFz3ZJpgLouLOWbWmUM2/KA9Bjk6AMIhPnsUFBa02oLzCw8qShOQ/HpV
# 8BJLaGuonLIEQOi3KvTRdYU7a1PcrMgBwpu3FBaintgdFpMB
# SIG # End signature block
