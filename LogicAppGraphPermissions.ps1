# Required PowerShell module:
#  - MgGraph to grant MSI permissions using the Microsoft Graph API
# Check if the required module is installed, prompt the user before installing if missing
$ModuleName = "Microsoft.Graph.Applications"
if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
    $UserResponse = Read-Host "❗ The module '$ModuleName' is not installed. Would you like to install it now? (Y/N)"
    if ($UserResponse -match "^[Yy]$") {
        Install-Module $ModuleName -Scope CurrentUser -Force
    } else {
        Write-Host "❌ Module installation declined. Script cannot proceed." -ForegroundColor Red
        exit
    }
}

$TenantId = ""
$TargetApiName = "" # Provide either the ID or name, not both.
$TargetApiAppId = "" # Provide either the ID or name, not both.
$LogicAppName = ""

# Connect to the Microsoft Graph API
Write-Host "⚙️ Connecting to the Entra tenant: $TenantId"
Connect-MgGraph -TenantId $TenantId -Scopes AppRoleAssignment.ReadWrite.All, Application.Read.All -NoWelcome

# Function to retrieve a service principal by name or ID
function Get-ServicePrincipal {
    param ([string]$AppName, [string]$AppId)

    if (-not $AppName -and -not $AppId) {
        Write-Host "❌ You must provide either an AppName or an AppId." -ForegroundColor Red
        return $null
    }

    $filter = if ($AppName) { "displayName eq '$AppName'" } else { "appId eq '$AppId'" }
    $app = Get-MgServicePrincipal -Filter $filter

    if (-not $app) {
        Write-Host "❌ Service principal '$($AppName ?? $AppId)' not found." -ForegroundColor Red
        return $null
    }

    if ($app.Count -gt 1) {
        Write-Host "❌ Multiple service principals found for '$($AppName ?? $AppId)'. Please refine your search." -ForegroundColor Red
        return $null
    }

    return $app
}

# Function to validate API parameters and retrieve necessary objects
function Get-APIObjects {
    param ([string]$MSIName, [string]$TargetApiName, [string]$TargetApiAppId, [string]$PermissionName)

    if (-not $TargetApiName -and -not $TargetApiAppId) {
        Write-Host "❌ You must provide either TargetApiName or TargetApiAppId." -ForegroundColor Red
        return $null
    }

    if ($TargetApiName -and $TargetApiAppId) {
        Write-Host "❌ Provide either TargetApiName or TargetApiAppId, not both." -ForegroundColor Red
        return $null
    }

    Write-Host "⚙️ Checking for principal: $MSIName"
    $MSI = Get-ServicePrincipal -AppName $MSIName
    if (-not $MSI) { return $null }

    Write-Host "✅ Found principal for $MSIName - ($($MSI.Id))" -ForegroundColor Green

    Write-Host "⚙️ Checking target API: $($TargetApiName ?? $TargetApiAppId)"
    $TargetApi = Get-ServicePrincipal -AppName $TargetApiName -AppId $TargetApiAppId
    if (-not $TargetApi) { return $null }

    $AppRole = $TargetApi.AppRoles | Where-Object { $_.Value -eq $PermissionName -and $_.AllowedMemberTypes -contains "Application" }
    if (-not $AppRole) {
        Write-Host "❌ Permission '$PermissionName' not found in API '$($TargetApi.DisplayName)'." -ForegroundColor Red
        return $null
    }

    return @{
        MSI       = $MSI
        TargetApi = $TargetApi
        AppRole   = $AppRole
    }
}

# Function to set or remove API permissions
function Set-APIPermissions {
    param (
        [Parameter(Mandatory=$true)][string]$MSIName,
        [Parameter(Mandatory=$false)][string]$TargetApiName,
        [Parameter(Mandatory=$false)][string]$TargetApiAppId,
        [Parameter(Mandatory=$true)][string]$PermissionName,
        [Parameter(Mandatory=$true)][ValidateSet("Grant", "Revoke")] [string]$Action
    )

    # Display input summary
    Write-Host "`n🔹 Execution Details 🔹"
    Write-Host "------------------------------------------------"
    Write-Host "📌 MSI Name: $MSIName"
    Write-Host "📌 Target API: $($TargetApiName ?? $TargetApiAppId)"
    Write-Host "📌 Permission: $PermissionName"
    Write-Host "📌 Action: $Action"
    Write-Host "------------------------------------------------"

    $apiObjects = Get-APIObjects -MSIName $MSIName -TargetApiName $TargetApiName -TargetApiAppId $TargetApiAppId -PermissionName $PermissionName
    if (-not $apiObjects) { return }

    $MSI = $apiObjects.MSI
    $TargetApi = $apiObjects.TargetApi
    $AppRole = $apiObjects.AppRole

    if ($Action -eq "Grant") {
        Write-Host "⚙️ Assigning permission '$PermissionName' to '$MSIName'"
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MSI.Id -PrincipalId $MSI.Id -ResourceId $TargetApi.Id -AppRoleId $AppRole.Id -ErrorAction Stop | Out-Null
            Write-Host "✅ Permission granted" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    elseif ($Action -eq "Revoke") {
        $AppRoleAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MSI.Id | Where-Object { $_.AppRoleId -eq $AppRole.Id }
        if (-not $AppRoleAssignment) {
            Write-Host "ℹ️ No existing assignment found for permission '$PermissionName' on '$MSIName'." -ForegroundColor Yellow
            return
        }

        Write-Host "⚙️ Removing permission '$PermissionName' from '$MSIName'"
        try {
            Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MSI.Id -AppRoleAssignmentId $AppRoleAssignment.Id -ErrorAction Stop | Out-Null
            Write-Host "✅ Permission removed successfully`n" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error: $($_.Exception.Message)`n" -ForegroundColor Red
        }
    }
}

# Execute the function with provided parameters
Set-APIPermissions -MSIName $LogicAppName -TargetApiName $TargetApiName -TargetApiAppId $TargetApiAppId -PermissionName "AdvancedQuery.Read.All" -Action "Revoke"
Set-APIPermissions -MSIName $LogicAppName -TargetApiName $TargetApiName -TargetApiAppId $TargetApiAppId -PermissionName "AdvancedQuery.Read.All" -Action "Grant"