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
