# Invoke-EntraAzureRBACCheck.ps1

Azure RBAC Role Assignment Audit & Drift Detection Tool

## Overview

This script provides comprehensive Azure RBAC role assignment auditing across all accessible tenants and subscriptions, with baseline export and drift detection capabilities. It helps identify unauthorized changes to role assignments by comparing current Azure state against a known-good baseline.

## Features

### Export Mode
- **Multi-Tenant Scanning** - Automatically scans ALL accessible tenants unless a specific tenant is specified
- **Multi-Subscription Support** - Scans all accessible subscriptions across all tenants
- **Comprehensive Coverage** - Captures role assignments at all scopes (subscription, resource group, resource level)
- **Principal Analysis** - Tracks users, groups, and service principals with role assignments
- **Role Definition Details** - Captures built-in vs custom roles, permissions, and descriptions
- **Scope Hierarchy** - Analyzes assignments across subscription, resource group, and resource scopes
- **Condition Support** - Tracks ABAC (Attribute-Based Access Control) conditions on assignments
- **Tenant Tracking** - Includes tenant information for all assignments in multi-tenant environments
- **JSON Baseline Export** - Exports all role assignments to JSON for drift detection

### DriftDetect Mode
- **Baseline Comparison** - Compares current Azure RBAC state against exported baseline
- **New Assignment Detection** - Identifies role assignments created outside of baseline
- **Removed Assignment Detection** - Detects role assignments removed since baseline
- **Modified Assignment Detection** - Identifies changes to existing role assignments (scope, role, principal, conditions)
- **Risk Assessment** - Categorizes drift by risk level (CRITICAL/HIGH/MEDIUM)
- **Detailed Recommendations** - Provides actionable recommendations for each drift issue
- **JSON Drift Report** - Exports drift findings to JSON for integration with security tools

## Requirements

- PowerShell 7+
- Az.Accounts module
- Az.Resources module
- Microsoft.Graph.Authentication module (for Graph API calls)

## Installation

```powershell
# Install required modules
Install-Module Az.Accounts -Scope CurrentUser
Install-Module Az.Resources -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `-Mode` | String | Operation mode: 'Export' or 'DriftDetect' | Export |
| `-ExportPath` | String | Path to export baseline JSON or drift report | azure-rbac-baseline.json |
| `-BaselinePath` | String | Path to baseline JSON file (required for DriftDetect) | - |
| `-SubscriptionId` | String[] | Specific subscription ID(s) to scan | All accessible |
| `-TenantId` | String | Specific Tenant ID to scan | All accessible |
| `-UseAzCliToken` | Switch | Use Azure CLI authentication | - |
| `-UseAzPowerShellToken` | Switch | Use Azure PowerShell authentication | - |
| `-UseDeviceCode` | Switch | Use device code authentication flow | - |
| `-EnableStealth` | Switch | Enable stealth mode with delays | - |
| `-RequestDelay` | Double | Base delay in seconds between requests (0-60) | 0 |
| `-RequestJitter` | Double | Random jitter range in seconds (0-30) | 0 |
| `-MaxRetries` | Int | Maximum retries on throttling (1-10) | 3 |
| `-QuietStealth` | Switch | Suppress stealth-related messages | - |
| `-IncludeInherited` | Switch | Include inherited (management group) role assignments | - |
| `-Matrix` | Switch | Display results in matrix/table format | - |
| `-SkipFailedTenants` | Switch | Continue on tenant authentication failures | - |
| `-ShowAllUsersPermissions` | Switch | Display user permissions matrix in Export mode | - |

## Usage Examples

### Export Mode

```powershell
# Export all Azure RBAC role assignments across ALL tenants and subscriptions
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export

# Export to specific file
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -ExportPath "rbac-baseline.json"

# Export with matrix view of all assignments
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -Matrix

# Export specific tenant only
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -TenantId "00000000-0000-0000-0000-000000000000"

# Export specific subscription with Azure CLI auth
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -SubscriptionId "00000000-0000-0000-0000-000000000000" -UseAzCliToken

# Export with stealth mode
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -EnableStealth -QuietStealth

# Export and show all users permissions matrix
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -ShowAllUsersPermissions
```

### DriftDetect Mode

```powershell
# Detect drift against baseline
.\Invoke-EntraAzureRBACCheck.ps1 -Mode DriftDetect -BaselinePath "rbac-baseline.json"

# Detect drift with matrix view
.\Invoke-EntraAzureRBACCheck.ps1 -Mode DriftDetect -BaselinePath "rbac-baseline.json" -Matrix

# Detect drift and export report
.\Invoke-EntraAzureRBACCheck.ps1 -Mode DriftDetect -BaselinePath "rbac-baseline.json" -ExportPath "drift-report.json"

# Detect drift with stealth mode
.\Invoke-EntraAzureRBACCheck.ps1 -Mode DriftDetect -BaselinePath "rbac-baseline.json" -EnableStealth -QuietStealth

# Detect drift for specific tenant
.\Invoke-EntraAzureRBACCheck.ps1 -Mode DriftDetect -BaselinePath "rbac-baseline.json" -TenantId "00000000-0000-0000-0000-000000000000"
```

### Using Script Dispatcher

```powershell
# Export mode (default)
.\Invoke-EvilMist.ps1 -Script EntraAzureRBACCheck -Mode Export -ExportPath "baseline.json"

# Drift detection with matrix view
.\Invoke-EvilMist.ps1 -Script EntraAzureRBACCheck -Mode DriftDetect -BaselinePath "baseline.json" -Matrix
```

## Output

### Export Mode JSON Structure

```json
{
  "ExportDate": "2025-01-03T12:00:00Z",
  "ExportVersion": "2.0",
  "Summary": {
    "TotalTenants": 2,
    "TotalSubscriptions": 5,
    "TotalAssignments": 150,
    "SkippedTenants": 0,
    "IncludeInherited": false
  },
  "ScopeStatistics": {
    "SubscriptionLevel": 50,
    "ResourceGroupLevel": 80,
    "ResourceLevel": 20,
    "ManagementGroupLevel": 0,
    "Other": 0
  },
  "PrincipalStatistics": {
    "Users": 80,
    "Groups": 50,
    "ServicePrincipals": 20,
    "Other": 0
  },
  "HighPrivilegeRoles": {
    "Owners": 10,
    "Contributors": 40,
    "UserAccessAdministrators": 5
  },
  "Tenants": [...],
  "Subscriptions": [...],
  "SkippedTenants": [...],
  "RoleAssignments": [
    {
      "AssignmentId": "/subscriptions/.../providers/Microsoft.Authorization/roleAssignments/...",
      "AssignmentName": "...",
      "TenantId": "00000000-0000-0000-0000-000000000000",
      "TenantName": "example",
      "SubscriptionId": "00000000-0000-0000-0000-000000000000",
      "SubscriptionName": "Production",
      "Scope": "/subscriptions/...",
      "ScopeType": "Subscription",
      "RoleDefinitionName": "Contributor",
      "RoleDefinitionId": "...",
      "PrincipalId": "00000000-0000-0000-0000-000000000000",
      "PrincipalType": "User",
      "PrincipalDisplayName": "John Doe",
      "PrincipalSignInName": "john.doe@example.com",
      "Condition": null,
      "ConditionVersion": null,
      "CanDelegate": false,
      "CreatedOn": "2024-01-15T10:30:00Z",
      "UpdatedOn": "2024-01-15T10:30:00Z",
      "CreatedBy": "admin@example.com",
      "UpdatedBy": "admin@example.com",
      "ExportTimestamp": "2025-01-03T12:00:00Z"
    }
  ]
}
```

### DriftDetect Mode JSON Structure

```json
{
  "ReportDate": "2025-01-03T14:00:00Z",
  "ReportVersion": "2.0",
  "BaselineFile": "rbac-baseline.json",
  "DriftDetected": true,
  "Summary": {
    "TotalDriftIssues": 5,
    "NewAssignments": 3,
    "RemovedAssignments": 1,
    "ModifiedAssignments": 1,
    "CriticalIssues": 2,
    "HighIssues": 2,
    "MediumIssues": 1
  },
  "ScanInfo": {
    "TenantsScanned": [...],
    "SubscriptionsScanned": [...],
    "TotalCurrentAssignments": 152,
    "ScanTimestamp": "2025-01-03T14:00:00Z"
  },
  "BaselineInfo": {
    "TotalBaselineAssignments": 150
  },
  "NewAssignmentsDetails": [...],
  "RemovedAssignmentsDetails": [...],
  "ModifiedAssignmentsDetails": [...]
}
```

## Risk Levels

### Drift Risk Classification

| Risk Level | Color | Description |
|------------|-------|-------------|
| **CRITICAL** | Red | New/modified Owner, Contributor, or Administrator role assignments |
| **HIGH** | Yellow | New assignments for other privileged roles, or removed high-privilege roles |
| **MEDIUM** | Green | Removed standard role assignments |

### Export Risk Classification (Matrix View)

| Risk Level | Color | Roles |
|------------|-------|-------|
| **CRITICAL** | Red | Owner, User Access Administrator |
| **HIGH** | Yellow | Contributor, *Administrator roles |
| **MEDIUM** | Green | Roles with Write/Delete/Modify actions |
| **LOW** | Gray | Reader and other read-only roles |

## Matrix View

The `-Matrix` parameter provides a formatted table view with color-coded risk levels:

### Export Mode Matrix

```
================================================================================
MATRIX VIEW - AZURE RBAC ROLE ASSIGNMENTS
================================================================================

Risk   Role                     Principal Type   Principal              Subscription              Scope Type   Tenant
----   ----                     --------------   ---------              ------------              ----------   ------
CRITICAL Owner                  User            John Doe               Production                Subscription example
HIGH     Contributor            ServicePrincipal App-Service           Development               ResourceGroup example
MEDIUM   Storage Blob Data...   User            Jane Smith             Production                Resource     example
LOW      Reader                 Group           Security-Readers       All-Subscriptions         Subscription example

================================================================================

[SUMMARY]
Total role assignments: 150
Unique principals: 75
  - CRITICAL risk: 10
  - HIGH risk: 40
  - MEDIUM risk: 30
  - LOW risk: 70

[PRINCIPAL TYPES]
  Users: 80
  Groups: 50
  Service Principals: 20

[SCOPE LEVELS]
  Subscription level: 50
  Resource Group level: 80
  Resource level: 20

[TOP ROLES]
  Contributor: 40
  Reader: 35
  Owner: 10
  ...

[!] HIGH-PRIVILEGE WARNINGS
  Owner role assignments: 10
  User Access Administrator assignments: 5
```

### DriftDetect Mode Matrix

```
================================================================================
MATRIX VIEW - DRIFT DETECTION RESULTS
================================================================================

Risk     Type              Tenant    Subscription   Role         Principal Type  Principal        Scope Type  Issue
----     ----              ------    ------------   ----         --------------  ---------        ----------  -----
CRITICAL NEW_ASSIGNMENT    example   Production     Owner        User            attacker@...     Subscription New role assignment...
HIGH     NEW_ASSIGNMENT    example   Development    Contributor  ServicePrinc... Unknown-App      ResourceGr... New role assignment...
MEDIUM   REMOVED_ASSIGN... example   Production     Reader       Group           Old-Readers      Subscription Role assignment rem...

================================================================================

[SUMMARY]
Total drift issues: 3
  - CRITICAL risk: 1
  - HIGH risk: 1
  - MEDIUM risk: 1

[DRIFT TYPES]
  New assignments: 2
  Removed assignments: 1
  Modified assignments: 0

[DRIFT BY SUBSCRIPTION]
  Production: 2
  Development: 1

[DRIFT BY ROLE]
  Owner: 1
  Contributor: 1
  Reader: 1
```

## Authentication

The script supports multiple authentication methods:

1. **Interactive Browser** (default) - OAuth login via browser popup
2. **Device Code Flow** (`-UseDeviceCode`) - Code-based authentication for embedded terminals
3. **Azure CLI Token** (`-UseAzCliToken`) - Use cached `az login` credentials
4. **Azure PowerShell Token** (`-UseAzPowerShellToken`) - Use cached `Connect-AzAccount` credentials

## Multi-Tenant Support

By default, the script scans ALL accessible tenants. This is useful for:
- Organizations with multiple Azure AD tenants
- Guest accounts with access to multiple tenants
- MSP/CSP scenarios managing multiple customer tenants

To limit to a specific tenant, use `-TenantId`:

```powershell
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -TenantId "00000000-0000-0000-0000-000000000000"
```

## Stealth Mode

Enable stealth mode to avoid detection and throttling:

```powershell
# Default stealth (500ms delay + 300ms jitter)
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -EnableStealth

# Custom stealth settings
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -RequestDelay 2 -RequestJitter 1

# Quiet stealth (no delay messages)
.\Invoke-EntraAzureRBACCheck.ps1 -Mode Export -EnableStealth -QuietStealth
```

## Use Cases

### Security Audit
1. Export baseline of all role assignments
2. Store baseline in version control
3. Periodically run drift detection to identify unauthorized changes
4. Review and investigate any detected drift

### Compliance
1. Document all role assignments across tenants and subscriptions
2. Identify high-privilege role assignments (Owner, Contributor, UAA)
3. Track who has access to what resources
4. Generate reports for compliance audits

### Incident Response
1. Compare current state against known-good baseline
2. Identify newly created role assignments
3. Detect removed or modified assignments
4. Trace unauthorized access changes

## License

GNU General Public License v3.0

---

**Part of EvilMist Toolkit** | [GitHub](https://github.com/Logisek/EvilMist)
