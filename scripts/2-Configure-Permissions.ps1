#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications
<#
.SYNOPSIS
    Configures API permissions for Exchange Calendar Event Management Automation App Registration

.DESCRIPTION
    This script automates the configuration of:
    - Microsoft Graph API permissions (Application permissions)
    - Exchange Online API permissions
    - Admin consent for all permissions
    - Verification of permission grants

.PARAMETER AppId
    Application ID of the App Registration (if not provided, will try to load from config file)

.PARAMETER ConfigFilePath
    Path to the configuration file created by the first script (default: current directory)

.EXAMPLE
    .\2-Configure-Permissions.ps1

.EXAMPLE
    .\2-Configure-Permissions.ps1 -AppId "12345678-1234-1234-1234-123456789012"

.NOTES
    Prerequisites:
    - PowerShell 5.1 or higher
    - Microsoft.Graph.Authentication module
    - Microsoft.Graph.Applications module
    - Global Administrator permissions in Azure AD
    - App Registration must exist (created by script 1)

    Author: Exchange Calendar Event Management Automation
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppId,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigFilePath = (Join-Path (Get-Location).Path "app-config.json")
)

# Color coding for output
function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    switch ($Color) {
        "Red" { Write-Host $Message -ForegroundColor Red }
        "Green" { Write-Host $Message -ForegroundColor Green }
        "Yellow" { Write-Host $Message -ForegroundColor Yellow }
        "Cyan" { Write-Host $Message -ForegroundColor Cyan }
        "Magenta" { Write-Host $Message -ForegroundColor Magenta }
        default { Write-Host $Message -ForegroundColor White }
    }
}

function Test-Prerequisites {
    Write-ColorOutput "🔍 Checking prerequisites..." "Cyan"
    
    # Check required modules
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications'
    )
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-ColorOutput "❌ Missing required module: $module" "Red"
            Write-ColorOutput "Install with: Install-Module -Name $module -Scope CurrentUser" "Yellow"
            return $false
        }
    }
    
    Write-ColorOutput "✅ All required modules are installed" "Green"
    return $true
}

function Get-AppConfiguration {
    param([string]$ConfigPath, [string]$AppId)
    
    if ($AppId) {
        Write-ColorOutput "📋 Using provided App ID: $AppId" "Cyan"
        return @{ AppId = $AppId }
    }
    
    if (Test-Path $ConfigPath) {
        Write-ColorOutput "📄 Loading configuration from: $ConfigPath" "Cyan"
        try {
            $config = Get-Content $ConfigPath | ConvertFrom-Json
            Write-ColorOutput "✅ Configuration loaded successfully" "Green"
            Write-ColorOutput "   App ID: $($config.AppId)" "White"
            return $config
        }
        catch {
            Write-ColorOutput "❌ Failed to load configuration file: $($_.Exception.Message)" "Red"
            return $null
        }
    }
    else {
        Write-ColorOutput "❌ Configuration file not found: $ConfigPath" "Red"
        Write-ColorOutput "Please provide the App ID parameter or run script 1 first" "Yellow"
        return $null
    }
}

function Connect-ToMicrosoftGraph {
    Write-ColorOutput "🔐 Connecting to Microsoft Graph..." "Cyan"
    
    try {
        # Connect with admin consent scopes
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Directory.Read.All" -NoWelcome
        Write-ColorOutput "✅ Connected to Microsoft Graph" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Grant-ApiPermissions {
    param([string]$ApplicationId)
    
    Write-ColorOutput "🔑 Configuring API permissions..." "Cyan"
    
    try {
        # Get the application
        $app = Get-MgApplication -Filter "AppId eq '$ApplicationId'"
        if (-not $app) {
            Write-ColorOutput "❌ Application not found with ID: $ApplicationId" "Red"
            return $false
        }
        
        # Get service principals for Microsoft Graph and Exchange Online
        $graphSP = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
        $exchangeSP = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
        
        # Get or create service principal for our app
        $appSP = Get-MgServicePrincipal -Filter "AppId eq '$ApplicationId'"
        if (-not $appSP) {
            Write-ColorOutput "📝 Creating service principal for the application..." "Cyan"
            $appSP = New-MgServicePrincipal -AppId $ApplicationId
        }
        
        # Define required permissions with friendly names
        $permissions = @{
            # Microsoft Graph permissions
            "00000003-0000-0000-c000-000000000000" = @{
                ServicePrincipal = $graphSP
                Permissions = @(
                    @{ Id = "ef54d2bf-783f-4e0f-bca1-3210c0444d99"; Name = "Calendars.ReadWrite" },
                    @{ Id = "19dbc75e-c2e2-444c-a770-ec69d8559fc7"; Name = "Directory.Read.All" },
                    @{ Id = "df021288-bdef-4463-88db-98f22de89214"; Name = "User.Read.All" },
                    @{ Id = "5b567255-7703-4780-807c-7be8301ae99b"; Name = "Group.Read.All" },
                    @{ Id = "75359482-378d-4052-8f01-80520e7db3cd"; Name = "Files.Read" },
                    @{ Id = "8e8e4742-1d95-4f68-9d56-6ee75648c72a"; Name = "Sites.Selected" }
                )
            }
            # Exchange Online permissions
            "00000002-0000-0ff1-ce00-000000000000" = @{
                ServicePrincipal = $exchangeSP
                Permissions = @(
                    @{ Id = "dc890d15-9560-4a4c-9b7f-a736ec74ec40"; Name = "Exchange.ManageAsApp" }
                )
            }
        }
        
        $grantedPermissions = @()
        $failedPermissions = @()
        
        foreach ($resourceAppId in $permissions.Keys) {
            $resource = $permissions[$resourceAppId]
            $servicePrincipal = $resource.ServicePrincipal
            
            foreach ($permission in $resource.Permissions) {
                try {
                    Write-ColorOutput "   Granting: $($permission.Name)" "White"
                    
                    # Check if permission is already granted
                    $existingGrant = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSP.Id | 
                        Where-Object { $_.AppRoleId -eq $permission.Id -and $_.ResourceId -eq $servicePrincipal.Id }
                    
                    if ($existingGrant) {
                        Write-ColorOutput "   ✅ $($permission.Name) - Already granted" "Green"
                        $grantedPermissions += $permission.Name
                    }
                    else {
                        # Grant the permission
                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSP.Id -PrincipalId $appSP.Id -ResourceId $servicePrincipal.Id -AppRoleId $permission.Id
                        Write-ColorOutput "   ✅ $($permission.Name) - Granted successfully" "Green"
                        $grantedPermissions += $permission.Name
                    }
                }
                catch {
                    Write-ColorOutput "   ❌ $($permission.Name) - Failed: $($_.Exception.Message)" "Red"
                    $failedPermissions += $permission.Name
                }
            }
        }
        
        Write-ColorOutput "`n📊 Permission Grant Summary:" "Yellow"
        Write-ColorOutput "✅ Successfully granted: $($grantedPermissions.Count) permissions" "Green"
        if ($failedPermissions.Count -gt 0) {
            Write-ColorOutput "❌ Failed to grant: $($failedPermissions.Count) permissions" "Red"
            Write-ColorOutput "Failed permissions: $($failedPermissions -join ', ')" "Red"
        }
        
        return $failedPermissions.Count -eq 0
    }
    catch {
        Write-ColorOutput "❌ Failed to configure permissions: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Test-PermissionGrants {
    param([string]$ApplicationId)
    
    Write-ColorOutput "🔍 Verifying permission grants..." "Cyan"
    
    try {
        $appSP = Get-MgServicePrincipal -Filter "AppId eq '$ApplicationId'"
        if (-not $appSP) {
            Write-ColorOutput "❌ Service principal not found" "Red"
            return $false
        }
        
        $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSP.Id
        
        Write-ColorOutput "📋 Current permission grants:" "Yellow"
        foreach ($assignment in $assignments) {
            $resource = Get-MgServicePrincipal -ServicePrincipalId $assignment.ResourceId
            $appRole = $resource.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId }
            Write-ColorOutput "   ✅ $($resource.DisplayName): $($appRole.Value)" "Green"
        }
        
        return $assignments.Count -gt 0
    }
    catch {
        Write-ColorOutput "❌ Failed to verify permissions: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Show-AdminConsentUrl {
    param([string]$ApplicationId, [string]$TenantId)
    
    Write-ColorOutput "🌐 Admin Consent URL:" "Yellow"
    $consentUrl = "https://login.microsoftonline.com/$TenantId/v2.0/adminconsent?client_id=$ApplicationId&state=12345&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=https://graph.microsoft.com/.default"
    Write-ColorOutput $consentUrl "Cyan"
    Write-ColorOutput "`n📝 Please visit this URL to provide admin consent for the permissions." "Yellow"
    Write-ColorOutput "This step is required for the application to work properly." "Yellow"
}

# Main execution
try {
    Write-ColorOutput "🔑 Exchange Calendar Event Management Automation - API Permissions Setup" "Magenta"
    Write-ColorOutput "=" * 80 "Magenta"
    
    # Check prerequisites
    if (!(Test-Prerequisites)) {
        Write-ColorOutput "❌ Prerequisites not met. Please install required modules and try again." "Red"
        exit 1
    }
    
    # Get app configuration
    $config = Get-AppConfiguration -ConfigPath $ConfigFilePath -AppId $AppId
    if (-not $config) {
        Write-ColorOutput "❌ Could not determine App ID. Please provide it as parameter or run script 1 first." "Red"
        exit 1
    }
    
    # Connect to Microsoft Graph
    if (!(Connect-ToMicrosoftGraph)) {
        Write-ColorOutput "❌ Failed to connect to Microsoft Graph" "Red"
        exit 1
    }
    
    $currentContext = Get-MgContext
    $tenantId = if ($config.TenantId) { $config.TenantId } else { $currentContext.TenantId }
    
    # Grant API permissions
    $permissionsSuccess = Grant-ApiPermissions -ApplicationId $config.AppId
    
    # Verify permissions
    $verificationSuccess = Test-PermissionGrants -ApplicationId $config.AppId
    
    # Show admin consent URL
    Show-AdminConsentUrl -ApplicationId $config.AppId -TenantId $tenantId
    
    # Output summary
    Write-ColorOutput "`n" + "=" * 80 "Green"
    Write-ColorOutput "✅ PERMISSIONS CONFIGURATION COMPLETED!" "Green"
    Write-ColorOutput "=" * 80 "Green"
    
    Write-ColorOutput "📋 Summary:" "Yellow"
    Write-ColorOutput "$(if($permissionsSuccess){'✅'}else{'⚠️'}) API permissions: $(if($permissionsSuccess){'Configured successfully'}else{'Some permissions failed'})" "White"
    Write-ColorOutput "$(if($verificationSuccess){'✅'}else{'❌'}) Permission verification: $(if($verificationSuccess){'Passed'}else{'Failed'})" "White"
    
    Write-ColorOutput "`n📝 Next Steps:" "Yellow"
    Write-ColorOutput "1. Visit the admin consent URL above to grant admin consent" "Cyan"
    Write-ColorOutput "2. Run script 3-Configure-SharePoint-Permissions.ps1 for Sites.Selected permissions" "Cyan"
    Write-ColorOutput "3. Upload the PowerShell runbook to your Automation Account" "Cyan"
    
    if ($permissionsSuccess -and $verificationSuccess) {
        Write-ColorOutput "`n🎉 All API permissions configured successfully!" "Green"
    }
    else {
        Write-ColorOutput "`n⚠️ Some permissions may need manual configuration in the Azure Portal." "Yellow"
    }
}
catch {
    Write-ColorOutput "❌ Script execution failed: $($_.Exception.Message)" "Red"
    Write-ColorOutput $_.ScriptStackTrace "Red"
    exit 1
}
finally {
    # Disconnect from services
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnection errors
    }
}