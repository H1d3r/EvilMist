# Invoke-SharePointEnum.ps1

## Overview

`Invoke-SharePointEnum.ps1` is a PowerShell script that searches SharePoint Online document repositories and downloads files using the SharePoint REST API with bearer token authentication. It is a PowerShell port of [ElephantPoint](https://github.com/nettitude/ElephantPoint) (Nettitude) adapted for the EvilMist toolkit.

## Purpose

This script performs authenticated enumeration of SharePoint Online to discover and exfiltrate sensitive documents during authorized penetration tests, including:

- **Full-text keyword search** - Search across all SharePoint document libraries
- **Faceted Query Language (FQL)** - Advanced search syntax for precise targeting
- **Refinement filters** - Filter results by file type, author, date, etc.
- **File download to disk** - Save discovered files locally
- **Base64-encoded download** - Output file content as Base64 for C2 exfiltration
- **Paginated retrieval** - Navigate through large result sets

## Attack Scenario Context

### Document Discovery and Exfiltration

1. Attacker obtains a valid SharePoint bearer token (via ROADtools, TokenTactics, etc.)
2. Attacker searches for sensitive keywords: "password", "credentials", "confidential"
3. Search results reveal documents containing sensitive data
4. Attacker downloads files to disk or encodes as Base64 for C2 transfer
5. Sensitive documents exfiltrated without triggering DLP if policies are weak

### Targeted File Retrieval

1. Attacker knows or discovers a server-relative file URL
2. Uses download mode to retrieve the specific file
3. Base64 encoding allows exfiltration through text-based C2 channels
4. File content transferred without touching disk (memory-only with Base64)

### Reconnaissance via Search

1. Attacker uses broad search queries ("*", "budget", "project plan")
2. Maps out document library structure and content
3. Identifies high-value targets: financial data, credentials, PII
4. Uses refinement filters to narrow scope (e.g., only .xlsx files)

### Red Team Value

- Discover sensitive documents across all SharePoint sites accessible to the compromised identity
- Download files for offline analysis or exfiltration
- Base64 output for C2-compatible data transfer
- FQL and refinement filters for surgical targeting
- Stealth features to avoid detection thresholds

### Blue Team Value

- Understand what data is discoverable via SharePoint Search API
- Test DLP policy effectiveness against REST API searches
- Validate that token-based access controls are working correctly
- Assess exposure of sensitive documents to broad search queries

## Prerequisites

- PowerShell 7.0 or later
- Azure PowerShell (`Az.Accounts` module) for auto-detection and interactive login
- No external PowerShell modules required for direct token usage

### Token Acquisition Methods

| Method | Description |
|--------|-------------|
| Auto-connect (default) | Automatically runs `Connect-AzAccount` when no session exists — browser popup with device code fallback. Token is validated against SharePoint REST API before use |
| Direct token (`-Token`) | Provide a bearer token obtained from ROADtools, TokenTactics, or other tooling. Requires `-SPOUrl` |
| Azure CLI (`-UseAzCliToken`) | Uses `az account get-access-token --resource https://{SPOUrl}` |
| Azure PowerShell (`-UseAzPowerShellToken`) | Uses `Get-AzAccessToken -ResourceUrl https://{SPOUrl}` |
| Device Code (`-UseDeviceCode`) | Forces device code authentication (useful in embedded terminals without browser popup) |

### Smart Authentication Fallback

When using auto-connect or device code mode, the script validates the obtained token against the SharePoint REST API before proceeding. If the token is rejected (401), the script automatically falls through a 3-tier fallback chain:

1. **Azure session token** — `Get-AzAccessToken` from existing `Connect-AzAccount` session. Validated with a lightweight `/_api/web` call
2. **Azure CLI token** — Falls back to `az account get-access-token` if the Azure CLI is installed
3. **MSAL device code flow** — Direct OAuth2 device code flow using the Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`), which has pre-consented SharePoint delegated permissions in most M365 tenants. Prompts the user to authenticate in a browser

This fallback is necessary because the Azure PowerShell first-party app (`1950a258-...`) may not have SharePoint API permissions consented in the target tenant, resulting in tokens with no SharePoint scopes that are rejected with 401.

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-SPOUrl` | String | Auto-detected | SharePoint Online domain (e.g., example.sharepoint.com). Auto-detected from Azure session when omitted. Required when using `-Token` |
| `-Token` | String | None | Bearer token for authentication. Requires `-SPOUrl` |
| `-UseAzCliToken` | Switch | False | Get token from Azure CLI |
| `-UseAzPowerShellToken` | Switch | False | Get token from Azure PowerShell |
| `-UseDeviceCode` | Switch | False | Force device code authentication (for terminals without browser popup) |
| `-TenantId` | String | Auto-detected | Azure AD tenant ID or domain. Auto-detected from Azure session when omitted |
| `-Query` | String | None | Search query string (keywords, filenames) |
| `-MaxRows` | Int | 50 | Maximum results to return (1-500) |
| `-EnableFQL` | Switch | False | Enable Faceted Query Language |
| `-RefinementFilter` | String | None | Refinement filter expression |
| `-StartRow` | Int | 0 | Starting row for pagination |
| `-FileUrl` | String | None | Server-relative file URL to download |
| `-SavePath` | String | None | Local path to save downloaded file |
| `-Base64` | Switch | False | Return file as Base64-encoded string |
| `-ExportPath` | String | None | Export results to CSV or JSON |
| `-Matrix` | Switch | False | Display results in table/matrix format |
| `-EnableStealth` | Switch | False | Enable stealth mode (500ms delay + 300ms jitter) |
| `-RequestDelay` | Double | 0 | Base delay in seconds between requests (0-60) |
| `-RequestJitter` | Double | 0 | Random jitter range in seconds (0-30) |
| `-MaxRetries` | Int | 3 | Maximum retries on throttling responses (1-10) |
| `-QuietStealth` | Switch | False | Suppress stealth-related status messages |

## Usage Examples

### Auto-Detection (Zero-Config)

```powershell
# Fully automatic — connects to Azure, detects SharePoint domain, searches
.\Invoke-SharePointEnum.ps1 -Query "password"

# Force device code auth (for embedded terminals without browser popup)
.\Invoke-SharePointEnum.ps1 -UseDeviceCode -Query "password"

# Auto-detect with specific tenant
.\Invoke-SharePointEnum.ps1 -TenantId "example.onmicrosoft.com" -Query "password"

# Via dispatcher — zero-config
.\Invoke-EvilMist.ps1 -Script SharePointEnum -Query "password"

# Via dispatcher with device code
.\Invoke-EvilMist.ps1 -Script SharePointEnum -UseDeviceCode -Query "password"
```

### Search Mode

```powershell
# Basic keyword search (with explicit URL)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "password"

# Search for Excel files containing 'budget'
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "filetype:xlsx budget" -MaxRows 100

# Search all documents (wildcard)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "*" -MaxRows 500

# Search with auto-detection (no -SPOUrl needed)
.\Invoke-SharePointEnum.ps1 -Query "confidential" -MaxRows 100
```

### FQL and Refinement Filters

```powershell
# FQL search with refinement filter for DOCX files
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "confidential" -EnableFQL -RefinementFilter 'filetype:equals("docx")'

# Refinement filter for files modified recently
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "*" -RefinementFilter 'write:range(2024-01-01,max)'

# FQL with auto-detection
.\Invoke-SharePointEnum.ps1 -Query "confidential" -EnableFQL -RefinementFilter 'filetype:equals("docx")'
```

### Download Mode

```powershell
# Download a specific file to disk
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -FileUrl "/sites/hr/Shared Documents/salaries.xlsx" -SavePath "C:\loot\salaries.xlsx"

# Download file as Base64 (for C2 exfiltration)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -FileUrl "/sites/hr/Shared Documents/salaries.xlsx" -Base64

# Download with auto-detection
.\Invoke-SharePointEnum.ps1 -FileUrl "/sites/hr/Shared Documents/salaries.xlsx" -SavePath "salaries.xlsx"
```

### Combined Search + Download

```powershell
# Search first, then download a known file in the same invocation
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "salary report" -FileUrl "/sites/hr/Shared Documents/salaries.xlsx" -SavePath "salaries.xlsx"
```

### Pagination

```powershell
# First page (results 1-50)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "*" -MaxRows 50 -StartRow 0

# Second page (results 51-100)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "*" -MaxRows 50 -StartRow 50
```

### Stealth Mode

```powershell
# Default stealth (500ms delay + 300ms jitter)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "credentials" -EnableStealth

# Stealth with quiet output
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "credentials" -EnableStealth -QuietStealth

# Custom timing for evasion
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "credentials" -RequestDelay 2 -RequestJitter 1

# Stealth with auto-detection
.\Invoke-SharePointEnum.ps1 -Query "credentials" -EnableStealth -QuietStealth
```

### Export Results

```powershell
# Export search results to CSV
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "budget" -ExportPath "results.csv"

# Export to JSON (includes metadata and summary)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "budget" -ExportPath "results.json"

# Matrix display with export
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "budget" -Matrix -ExportPath "results.csv"

# Export with auto-detection
.\Invoke-SharePointEnum.ps1 -Query "budget" -ExportPath "results.json" -Matrix
```

### Using Dispatcher

```powershell
# Via main EvilMist dispatcher (zero-config)
.\Invoke-EvilMist.ps1 -Script SharePointEnum -Query "password"

# Dispatcher with explicit URL and token
.\Invoke-EvilMist.ps1 -Script SharePointEnum -SPOUrl example.sharepoint.com -Token $token -Query "password"

# Dispatcher with export
.\Invoke-EvilMist.ps1 -Script SharePointEnum -Query "confidential" -ExportPath "results.json"

# Dispatcher with device code
.\Invoke-EvilMist.ps1 -Script SharePointEnum -UseDeviceCode -Query "password"

# Dispatcher with Azure CLI token
.\Invoke-EvilMist.ps1 -Script SharePointEnum -SPOUrl example.sharepoint.com -UseAzCliToken -Query "password"

# Dispatcher with Azure PowerShell token
.\Invoke-EvilMist.ps1 -Script SharePointEnum -SPOUrl example.sharepoint.com -UseAzPowerShellToken -Query "password"
```

### Azure CLI / PowerShell Token

```powershell
# Use Azure CLI cached token (requires explicit -SPOUrl)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseAzCliToken -Query "password"

# Use Azure PowerShell cached token
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseAzPowerShellToken -Query "password"

# Via dispatcher with Azure CLI token
.\Invoke-EvilMist.ps1 -Script SharePointEnum -SPOUrl example.sharepoint.com -UseAzCliToken -Query "password"

# Via dispatcher with Azure PowerShell token
.\Invoke-EvilMist.ps1 -Script SharePointEnum -SPOUrl example.sharepoint.com -UseAzPowerShellToken -Query "password"
```

### Device Code Flow

```powershell
# Device code with auto-detection
.\Invoke-SharePointEnum.ps1 -UseDeviceCode -Query "password"

# Device code with specific tenant
.\Invoke-SharePointEnum.ps1 -UseDeviceCode -TenantId "example.onmicrosoft.com" -Query "password"

# Device code with explicit URL (skips auto-detect)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseDeviceCode -Query "password"

# Via dispatcher with device code
.\Invoke-EvilMist.ps1 -Script SharePointEnum -UseDeviceCode -Query "password"
```

## Output Fields

### Search Results

| Field | Description |
|-------|-------------|
| Title | Document title |
| Path | Full URL to the document |
| Author | Document author |
| Size | File size in bytes |
| SizeFormatted | Human-readable file size (KB/MB/GB) |
| FileType | File type (docx, xlsx, pdf, etc.) |
| FileExtension | File extension |
| Created | Creation date |
| LastModifiedTime | Last modification date |
| Summary | Hit-highlighted search summary (HTML stripped) |
| SiteName | SharePoint site name |
| ServerRelativeUrl | Server-relative path (for download) |

### Download Results

| Field | Description |
|-------|-------------|
| FileUrl | Server-relative URL of the downloaded file |
| SavedTo | Local path where file was saved (disk mode) |
| Size | File size in bytes |
| SizeFormatted | Human-readable file size |
| Base64Length | Length of Base64 string (Base64 mode) |
| Base64 | Base64-encoded file content (Base64 mode) |
| Method | Download method used (Disk or Base64) |
| Success | Whether the download succeeded |

## Sample Output

### Standard Output

```
[*] Target: https://example.sharepoint.com
[*] Stealth mode: ACTIVE (delay: 0.5s, jitter: 0.3s, retries: 3)
[+] Using provided bearer token

[*] Executing SharePoint search...
[*] Query: password
[*] Max rows: 50 | Start row: 0
[+] Total results available: 12
[+] Results returned: 12

================================================================================
SEARCH RESULTS (12 items)
================================================================================

[1] Network Credentials Guide
    [+] Path: https://example.sharepoint.com/sites/IT/Shared Documents/credentials.docx
    [+] Author: John Smith
    [+] Size: 45.23 KB
    [+] Type: docx
    [+] Modified: 2024-03-15T10:30:00Z
    [+] Site: IT Department
    [+] Summary: This document contains all network password reset procedures...

[2] VPN Access Instructions
    [+] Path: https://example.sharepoint.com/sites/IT/Shared Documents/vpn-setup.pdf
    [+] Author: Jane Doe
    [+] Size: 1.24 MB
    [+] Type: pdf
    [+] Modified: 2024-02-20T14:15:00Z
    [+] Site: IT Department

================================================================================
```

### Matrix Output

```
========================================================================================================================
MATRIX VIEW - SHAREPOINT SEARCH RESULTS
========================================================================================================================

[SEARCH METADATA]
--------------------------------------------------------------------------------
  Query: password
  Total available: 12 | Returned: 12 | Start row: 0

[RESULTS]
------------------------------------------------------------------------------------------------------------------------

Title                              Type  Size       Author               Modified    Path
-----                              ----  ----       ------               --------    ----
Network Credentials Guide          docx  45.23 KB   John Smith           2024-03-15  ...IT/Shared Documents/credentials.docx
VPN Access Instructions            pdf   1.24 MB    Jane Doe             2024-02-20  ...IT/Shared Documents/vpn-setup.pdf

========================================================================================================================
```

### Base64 Download Output

```
[*] Downloading file from SharePoint...
[*] File URL: /sites/hr/Shared Documents/salaries.xlsx
[*] Downloading as Base64-encoded string...
[+] File downloaded successfully (234.56 KB)
[+] Base64 length: 312748 characters

--- BEGIN BASE64 ---
UEsDBBQAAAAIAGFiV1kAAAAAAAAAAAAAABEAHABkb2NQcm9wcy9jb3JlLnhtbFVUCQAD...
--- END BASE64 ---
```

## Stealth & Evasion

### Request Timing

| Parameter | Default | Stealth Default | Description |
|-----------|---------|-----------------|-------------|
| `-RequestDelay` | 0s | 0.5s | Base delay between API requests |
| `-RequestJitter` | 0s | 0.3s | Random jitter added/subtracted from delay |
| `-MaxRetries` | 3 | 3 | Retry count on 429/503 responses |

### Throttle Handling

The script automatically handles SharePoint API throttling:
- **429 Too Many Requests**: Respects `Retry-After` header, adds jitter, retries up to `MaxRetries`
- **503 Service Unavailable**: Exponential backoff (5s, 10s, 20s...) with retry

### OPSEC Considerations

- SharePoint REST API calls generate audit log entries in Microsoft 365
- Search queries are logged in the Unified Audit Log
- File downloads are tracked in SharePoint access logs
- Use stealth mode to reduce request velocity and avoid alerting SOC teams
- Base64 mode avoids writing files to disk on the target

## Authentication Methods

### Auto-Connect (Default)

When no authentication flag is specified, the script automatically:
1. Checks for an existing Azure PowerShell session (`Get-AzContext`)
2. If no session, opens a browser popup for interactive login
3. Falls back to device code authentication if browser popup fails
4. Auto-detects the SharePoint domain from the authenticated tenant
5. **Validates** the session token against SharePoint REST API (`/_api/web`)
6. If validation fails (401), tries Azure CLI token as fallback
7. If Azure CLI also fails, triggers **MSAL device code flow** with the Microsoft Office client ID — prompts the user to authenticate in a browser with proper SharePoint permissions

```powershell
# Just provide a query — everything else is automatic
# If the Azure session token lacks SharePoint permissions, you'll be prompted via device code
.\Invoke-SharePointEnum.ps1 -Query "password"

# Force device code for terminals without browser support
.\Invoke-SharePointEnum.ps1 -UseDeviceCode -Query "password"
```

### Bearer Token (Recommended for Red Team)

Obtain a token using your preferred method and pass it directly:

```powershell
# From ROADtools
$token = (roadrecon auth -u user@example.com -p 'password' --tokens-stdout | ConvertFrom-Json).accessToken

# From TokenTactics
$token = (Get-AzureToken -Client SharePoint).access_token

# Pass to script (requires -SPOUrl)
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "password"
```

### Azure CLI

```powershell
# Login first
az login

# Script will auto-acquire token for the SharePoint resource
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseAzCliToken -Query "password"
```

### Azure PowerShell

```powershell
# Login first
Connect-AzAccount

# Script will auto-acquire token for the SharePoint resource
.\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseAzPowerShellToken -Query "password"
```

## Related Scripts

- `Invoke-EntraSharePointCheck.ps1` - SharePoint Online sharing settings security audit
- `Invoke-EntraEnum.ps1` - Unauthenticated Azure/Entra ID enumeration and reconnaissance
- `Invoke-EntraRecon.ps1` - Authenticated Entra ID user enumeration and assessment

## References

- [SharePoint Search REST API](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api)
- [SharePoint REST API: GetFileByServerRelativeUrl](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)
- [Keyword Query Language (KQL) syntax reference](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference)
- [FAST Query Language (FQL) syntax reference](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/fast-query-language-fql-syntax-reference)
- [ElephantPoint - Nettitude (original C# implementation)](https://github.com/nettitude/ElephantPoint)

## License

This script is part of the EvilMist toolkit and is distributed under the GNU General Public License v3.0.

## Author

Logisek - https://logisek.com
