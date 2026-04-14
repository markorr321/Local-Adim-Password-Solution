# LAPS

PowerShell module for retrieving and managing Microsoft LAPS (Local Administrator Password Solution) credentials from Entra ID devices via Microsoft Graph API. Features an interactive TUI for device search, password retrieval, clipboard copy, and on-demand password rotation. Browser-based authentication with optional custom app registration support. Cross-platform compatible with **Windows**, **macOS**, and **Linux**. Just run `Start-LAPS` — works out of the box with no configuration, or bring your own app registration for full control.

## In Action

![LAPS Demo](Screenshots/LAPS.gif)

## Screenshots

### Launch & Authentication
![Authentication](Screenshots/Step-2.webp)

### Search for a Device
![Search](Screenshots/Step-3.webp)

### Device Selection
![Device Selection](Screenshots/Step-4.webp)

### Password Retrieved
![Password Result](Screenshots/Step-5.webp)

### Rotate Password
![Rotate Password](Screenshots/Step-7.webp)

### Copy to Clipboard
![Copy Password](Screenshots/Step-6.webp)

### Disconnect from Graph
![Disconnect](Screenshots/Step-8.webp)

## Features

- **Device Search**: Search Entra ID devices by name with real-time results
- **Password Retrieval**: Retrieve LAPS local admin credentials via Microsoft Graph `/directory/deviceLocalCredentials` endpoint
- **Clipboard Copy**: Copy passwords to clipboard with `Ctrl+C` directly from the result screen
- **On-Demand Rotation**: Trigger immediate LAPS password rotation for Intune-managed devices
- **Interactive TUI**: Arrow-key navigation, dynamic control bar, and inline prompts
- **Browser Authentication**: Secure MSAL-based browser authentication with branded success/error pages
- **Persistent Configuration**: Save custom app registration credentials as environment variables
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Auto-Dependencies**: Automatically installs required MSAL libraries on first run
- **Scope Validation**: JWT token inspection warns immediately if required permissions are missing

## Installation

### Using PowerShellGet

```powershell
Install-Module -Name LAPS -Repository PSGallery
```

### Using PSResourceGet

```powershell
Install-PSResource -Name LAPS -Repository PSGallery
```

> **Coming Soon**: This module will be published to the PowerShell Gallery shortly.

## Quick Start

```powershell
Start-LAPS
```

That's it! The tool will:
1. Check prerequisites and load MSAL libraries
2. Open your browser for authentication
3. Present a search prompt for device names
4. Display LAPS credentials with options to copy or rotate

## App Registration Setup

LAPS requires an Entra ID app registration with specific delegated permissions. The default Microsoft public client ID may not have `DeviceLocalCredential.Read.All` pre-consented in your tenant, so a custom app registration is recommended.

### Creating the App Registration

1. Go to **Microsoft Entra ID** > **App registrations** > **New registration**
2. Name it something like `LAPS-PowerShell`
3. Under **Authentication**:
   - **Platform**: Mobile and desktop applications
   - **Redirect URI**: `http://localhost`
   - **Allow public client flows**: Yes
4. Under **API permissions**, add the following **Delegated** permissions:

| Permission | Purpose |
|------------|---------|
| `Device.Read.All` | Search and read device information |
| `DeviceLocalCredential.Read.All` | Retrieve LAPS passwords |
| `DeviceManagementManagedDevices.PrivilegedOperations.All` | Rotate LAPS passwords (optional) |

5. Click **Grant admin consent** for your organization

> **Note**: `DeviceLocalCredential.Read.All` requires admin consent. Without it, authentication will succeed but password retrieval will fail. The tool validates granted scopes after login and warns if this permission is missing.

> **Note**: The rotation permission (`DeviceManagementManagedDevices.PrivilegedOperations.All`) is only required if you want to trigger on-demand password rotation. It uses the Intune beta API endpoint.

## Configuration

### Persistent Configuration (Recommended)

Configure your app registration once and use `Start-LAPS` without parameters going forward:

```powershell
# Configure once
Configure-LAPS
```

You'll be prompted to enter your Client ID and Tenant ID. These are saved as user-level environment variables that persist across PowerShell sessions.

**On Windows:** Saved to user-level environment variables automatically.

**On macOS:** You'll be offered the option to add the configuration to your PowerShell profile for persistence.

After configuration:

```powershell
Start-LAPS
```

To remove the saved configuration:

```powershell
Clear-LAPSConfig
```

### One-Time Custom App Registration

For temporary use of a custom app registration (single session only):

```powershell
Start-LAPS -ClientId "<your-client-id>" -TenantId "<your-tenant-id>"
```

## Available Commands

| Command | Description |
|---------|-------------|
| `Start-LAPS` | Launch the LAPS password retrieval TUI |
| `Configure-LAPS` | Set up persistent custom app registration configuration |
| `Clear-LAPSConfig` | Remove saved configuration and return to default auth |
| `Get-LAPSHelp` | Display comprehensive help and command reference |

## Keyboard Shortcuts

### Search Screen

| Shortcut | Action |
|----------|--------|
| Type | Enter device name |
| ENTER | Search |
| ESC | Back / Exit |
| Ctrl+Q | Disconnect and exit |

### Device Selection

| Shortcut | Action |
|----------|--------|
| ↑/↓ | Navigate device list |
| ENTER | Select device |
| ESC | Back to search |
| Ctrl+Q | Disconnect and exit |

### Password Result Screen

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy password to clipboard |
| R | Rotate password (with Y/N confirmation) |
| S | New search |
| ESC | Back to search |
| Ctrl+Q | Disconnect and exit |

## How It Works

### Authentication

LAPS uses MSAL (Microsoft Authentication Library) for browser-based interactive authentication. On launch, it:

1. Loads MSAL assemblies from the local NuGet cache or falls back to the Az.Accounts module
2. Compiles a C# helper class (`LAPSBrowserAuth`) for browser-based token acquisition
3. Opens your default browser for sign-in with a branded success/error page
4. Acquires a delegated access token with the required Graph scopes
5. Validates the JWT token to confirm all required scopes were granted

### Device Search

Searches use the Microsoft Graph `/v1.0/devices` endpoint with a `startsWith` filter on `displayName`. This is an advanced query requiring the `ConsistencyLevel: eventual` header and `$count=true` parameter.

### Password Retrieval

LAPS credentials are retrieved in two steps:
1. **Lookup**: `GET /v1.0/directory/deviceLocalCredentials?$filter=deviceName eq '{name}'` — finds the credential info object by device name
2. **Retrieve**: `GET /v1.0/directory/deviceLocalCredentials/{id}?$select=credentials` — fetches the actual password data

The password is returned as a base64-encoded string and decoded (UTF-8) for display.

> **Important**: The endpoint path must include `/directory/`. The beta endpoint and non-directory variants return errors.

### Password Rotation

On-demand rotation uses the Intune beta API:
1. **Lookup**: `GET /v1.0/deviceManagement/managedDevices?$filter=deviceName eq '{name}'` — resolves the Intune managed device ID
2. **Rotate**: `POST /beta/deviceManagement/managedDevices/{id}/rotateLocalAdminPassword` — triggers rotation

The new password is generated on the device and backed up to Entra ID on its next check-in. The old password remains valid until rotation completes.

## Requirements

- **PowerShell 7.0+**
- **MSAL Libraries** (auto-resolved from one of):
  - Local NuGet cache (`~/.nuget/packages/microsoft.identity.client`)
  - Az.Accounts module (auto-installed if needed)
- **Entra ID App Registration** with admin-consented delegated permissions
- **LAPS configured** on target devices (Windows LAPS backed up to Entra ID)
- **Intune enrollment** (only required for password rotation)

## Troubleshooting

### "DeviceLocalCredential.Read.All scope was NOT granted"

Your app registration is missing the `DeviceLocalCredential.Read.All` delegated permission, or admin consent has not been granted. Add the permission in the Azure portal and click **Grant admin consent**.

### "No LAPS credentials found for device"

- The device may not have LAPS configured
- LAPS credentials have not been backed up to Entra ID yet
- The device name doesn't match exactly

### "Device not found in Intune. Rotation requires an Intune-managed device."

Password rotation requires the device to be enrolled in Microsoft Intune. Hybrid Azure AD joined devices managed only by on-premises Active Directory cannot be rotated through Graph API.

### "Access denied (403)"

Your account does not have sufficient permissions. Ensure:
- The app registration has the correct delegated permissions
- Admin consent has been granted
- Your user account has the appropriate Entra ID role (e.g., Cloud Device Administrator)

### Search returns no results

The search uses `startsWith` matching on device display names. Ensure you're typing the beginning of the device name, not a substring from the middle.

## Tags

LAPS, Entra, Azure, Identity, Security, MicrosoftGraph, LocalAdmin, Password, CrossPlatform, TUI, Intune, PowerShell
