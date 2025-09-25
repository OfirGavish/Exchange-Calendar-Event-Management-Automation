#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    Configures SharePoint Sites.Selected permissions for Exchange Calendar Event Management Automation

.DESCRIPTION
    This script automates the configuration of:
    - Sites.Selected permissions for specific SharePoint sites
    - Read and Write permissions for the calendar automation app
    - Verification of permission grants

.PARAMETER AppId
    Application ID of the App Registration (if not provided, will try to load from config file)

.PARAMETER SharePointSiteUrl
    SharePoint site URL where the Excel files are stored

.PARAMETER ConfigFilePath
    Path to the configuration file created by the first script (default: current directory)

.PARAMETER PermissionLevel
    Permission level to grant (Read, Write, or FullControl). Default: Write

.EXAMPLE
    .\3-Configure-SharePoint-Permissions.ps1 -SharePointSiteUrl "https://contoso.sharepoint.com/sites/events"

.EXAMPLE
    .\3-Configure-SharePoint-Permissions.ps1 -AppId "12345678-1234-1234-1234-123456789012" -SharePointSiteUrl "https://contoso.sharepoint.com/sites/events" -PermissionLevel "Write"

.NOTES
    Prerequisites:
    - PowerShell 5.1 or higher
    - PnP.PowerShell module
    - SharePoint Administrator or Global Administrator permissions
    - App Registration must exist with Sites.Selected permission (created by previous scripts)

    Author: Exchange Calendar Event Management Automation
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppId,
    
    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteUrl,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigFilePath = (Join-Path (Get-Location).Path "app-config.json"),
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Read", "Write", "FullControl")]
    [string]$PermissionLevel = "Write"
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
    if (!(Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-ColorOutput "❌ Missing required module: PnP.PowerShell" "Red"
        Write-ColorOutput "Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser" "Yellow"
        return $false
    }
    
    Write-ColorOutput "✅ PnP.PowerShell module is installed" "Green"
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

function Connect-ToSharePointAdmin {
    param([string]$SiteUrl)
    
    Write-ColorOutput "🔐 Connecting to SharePoint..." "Cyan"
    
    try {
        # Extract tenant from site URL
        $uri = [System.Uri]$SiteUrl
        $tenantName = $uri.Host.Split('.')[0]
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        
        Write-ColorOutput "📡 Admin URL: $adminUrl" "White"
        
        # Connect to SharePoint Admin
        Connect-PnPOnline -Url $adminUrl -Interactive -WarningAction SilentlyContinue
        Write-ColorOutput "✅ Connected to SharePoint Admin Center" "Green"
        
        return $adminUrl
    }
    catch {
        Write-ColorOutput "❌ Failed to connect to SharePoint: $($_.Exception.Message)" "Red"
        return $null
    }
}

function Grant-SharePointSitePermission {
    param([string]$ApplicationId, [string]$SiteUrl, [string]$Permission)
    
    Write-ColorOutput "🔑 Granting Sites.Selected permission to site..." "Cyan"
    Write-ColorOutput "   Site: $SiteUrl" "White"
    Write-ColorOutput "   Permission: $Permission" "White"
    
    try {
        # Grant the permission
        Grant-PnPAzureADAppSitePermission -AppId $ApplicationId -DisplayName "Exchange Calendar Event Management Automation" -Site $SiteUrl -Permissions $Permission
        
        Write-ColorOutput "✅ Sites.Selected permission granted successfully" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "❌ Failed to grant Sites.Selected permission: $($_.Exception.Message)" "Red"
        
        # Check if it's a permission already exists error
        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already granted*") {
            Write-ColorOutput "ℹ️ Permission may already exist. Checking current permissions..." "Yellow"
            return $true
        }
        
        return $false
    }
}

function Test-SharePointSitePermission {
    param([string]$ApplicationId, [string]$SiteUrl)
    
    Write-ColorOutput "🔍 Verifying Sites.Selected permissions..." "Cyan"
    
    try {
        # Get current app permissions for the site
        $permissions = Get-PnPAzureADAppSitePermission -Site $SiteUrl
        
        $appPermissions = $permissions | Where-Object { $_.AppId -eq $ApplicationId }
        
        if ($appPermissions) {
            Write-ColorOutput "✅ Sites.Selected permissions found:" "Green"
            foreach ($perm in $appPermissions) {
                Write-ColorOutput "   📋 App: $($perm.DisplayName)" "White"
                Write-ColorOutput "   🔑 Permission: $($perm.Roles -join ', ')" "White"
                Write-ColorOutput "   🆔 App ID: $($perm.AppId)" "White"
            }
            return $true
        }
        else {
            Write-ColorOutput "❌ No Sites.Selected permissions found for this app" "Red"
            return $false
        }
    }
    catch {
        Write-ColorOutput "❌ Failed to verify permissions: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Show-SharePointPermissionGuidance {
    param([string]$SiteUrl)
    
    Write-ColorOutput "`n📝 SharePoint Permission Configuration Complete!" "Yellow"
    Write-ColorOutput "=" * 60 "Yellow"
    
    Write-ColorOutput "🔍 To verify permissions manually:" "Cyan"
    Write-ColorOutput "1. Go to SharePoint Admin Center" "White"
    Write-ColorOutput "2. Navigate to Advanced > API access" "White"
    Write-ColorOutput "3. Look for 'Exchange Calendar Event Management Automation'" "White"
    Write-ColorOutput "4. Verify it has access to your site: $SiteUrl" "White"
    
    Write-ColorOutput "`n🛠️ If permissions don't work:" "Cyan"
    Write-ColorOutput "1. Ensure the app has Sites.Selected permission in Azure AD" "White"
    Write-ColorOutput "2. Check that admin consent was granted" "White"
    Write-ColorOutput "3. Verify the SharePoint site URL is correct" "White"
    Write-ColorOutput "4. Wait a few minutes for permissions to propagate" "White"
}

# Main execution
try {
    Write-ColorOutput "🌐 Exchange Calendar Event Management Automation - SharePoint Permissions Setup" "Magenta"
    Write-ColorOutput "=" * 80 "Magenta"
    
    # Check prerequisites
    if (!(Test-Prerequisites)) {
        Write-ColorOutput "❌ Prerequisites not met. Please install PnP.PowerShell module and try again." "Red"
        exit 1
    }
    
    # Get app configuration
    $config = Get-AppConfiguration -ConfigPath $ConfigFilePath -AppId $AppId
    if (-not $config) {
        Write-ColorOutput "❌ Could not determine App ID. Please provide it as parameter or run script 1 first." "Red"
        exit 1
    }
    
    # Validate SharePoint site URL
    if (-not $SharePointSiteUrl.StartsWith("https://") -or -not $SharePointSiteUrl.Contains(".sharepoint.com")) {
        Write-ColorOutput "❌ Invalid SharePoint site URL format. Expected: https://tenant.sharepoint.com/sites/sitename" "Red"
        exit 1
    }
    
    # Connect to SharePoint Admin
    $adminUrl = Connect-ToSharePointAdmin -SiteUrl $SharePointSiteUrl
    if (-not $adminUrl) {
        Write-ColorOutput "❌ Failed to connect to SharePoint Admin Center" "Red"
        exit 1
    }
    
    # Grant Sites.Selected permission
    $grantSuccess = Grant-SharePointSitePermission -ApplicationId $config.AppId -SiteUrl $SharePointSiteUrl -Permission $PermissionLevel
    
    # Test the permissions
    $testSuccess = Test-SharePointSitePermission -ApplicationId $config.AppId -SiteUrl $SharePointSiteUrl
    
    # Show guidance
    Show-SharePointPermissionGuidance -SiteUrl $SharePointSiteUrl
    
    # Output summary
    Write-ColorOutput "`n" + "=" * 80 "Green"
    Write-ColorOutput "✅ SHAREPOINT PERMISSIONS CONFIGURATION COMPLETED!" "Green"
    Write-ColorOutput "=" * 80 "Green"
    
    Write-ColorOutput "📋 Summary:" "Yellow"
    Write-ColorOutput "✅ App Registration: $($config.AppId)" "White"
    Write-ColorOutput "✅ SharePoint Site: $SharePointSiteUrl" "White"
    Write-ColorOutput "$(if($grantSuccess){'✅'}else{'❌'}) Permission Grant: $(if($grantSuccess){'Success'}else{'Failed'})" "White"
    Write-ColorOutput "$(if($testSuccess){'✅'}else{'⚠️'}) Permission Verification: $(if($testSuccess){'Verified'}else{'Needs manual check'})" "White"
    Write-ColorOutput "✅ Permission Level: $PermissionLevel" "White"
    
    Write-ColorOutput "`n📝 Next Steps:" "Yellow"
    Write-ColorOutput "1. Upload the PowerShell runbook (eventscalendarautomation.ps1) to your Automation Account" "Cyan"
    Write-ColorOutput "2. Test the automation with a small Excel file" "Cyan"
    Write-ColorOutput "3. Set up scheduled execution of the runbook" "Cyan"
    
    if ($grantSuccess -and $testSuccess) {
        Write-ColorOutput "`n🎉 All SharePoint permissions configured successfully!" "Green"
        Write-ColorOutput "Your app can now read and write Excel files in the specified SharePoint site." "Green"
    }
    else {
        Write-ColorOutput "`n⚠️ Please verify the permissions manually in SharePoint Admin Center." "Yellow"
        Write-ColorOutput "The automation may not work until Sites.Selected permissions are properly configured." "Yellow"
    }
}
catch {
    Write-ColorOutput "❌ Script execution failed: $($_.Exception.Message)" "Red"
    Write-ColorOutput $_.ScriptStackTrace "Red"
    exit 1
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnection errors
    }
}