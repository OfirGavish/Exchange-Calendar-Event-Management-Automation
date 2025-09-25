#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Az.Automation
<#
.SYNOPSIS
    Creates Azure AD App Registration and Certificate for Exchange Calendar Event Management Automation

.DESCRIPTION
    This script automates the creation of:
    - Azure AD App Registration with appropriate settings
    - Self-signed certificate for authentication
    - Uploads certificate to both App Registration and Azure Automation Account
    - Configures redirect URIs and basic app settings

.PARAMETER AppDisplayName
    Display name for the App Registration (default: "Exchange Calendar Event Management Automation")

.PARAMETER CertificatePassword
    Password for the certificate private key (if not provided, will prompt securely)

.PARAMETER AutomationAccountName
    Name of the Azure Automation Account where certificate will be uploaded

.PARAMETER AutomationResourceGroupName
    Resource Group containing the Automation Account

.PARAMETER CertificateExportPath
    Path where certificate files will be saved (default: current directory)

.EXAMPLE
    .\1-Setup-AppRegistration.ps1 -AutomationAccountName "calendar-automation-aa" -AutomationResourceGroupName "rg-calendar-automation"

.EXAMPLE
    .\1-Setup-AppRegistration.ps1 -AppDisplayName "My Calendar Automation" -CertificatePassword (ConvertTo-SecureString "MyPassword123!" -AsPlainText -Force)

.NOTES
    Prerequisites:
    - PowerShell 5.1 or higher
    - Microsoft.Graph.Authentication module
    - Microsoft.Graph.Applications module  
    - Az.Automation module
    - Global Administrator or Application Administrator permissions in Azure AD
    - Contributor permissions on the Azure Automation Account

    Author: Exchange Calendar Event Management Automation
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppDisplayName = "Exchange Calendar Event Management Automation",
    
    [Parameter(Mandatory = $false)]
    [SecureString]$CertificatePassword,
    
    [Parameter(Mandatory = $true)]
    [string]$AutomationAccountName,
    
    [Parameter(Mandatory = $true)]
    [string]$AutomationResourceGroupName,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateExportPath = (Get-Location).Path
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
        'Microsoft.Graph.Applications', 
        'Az.Automation'
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

function Connect-ToServices {
    Write-ColorOutput "🔐 Connecting to Microsoft Graph..." "Cyan"
    
    try {
        # Connect to Microsoft Graph with required scopes
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.Read.All" -NoWelcome
        Write-ColorOutput "✅ Connected to Microsoft Graph" "Green"
        
        # Connect to Azure (for Automation Account access)
        Write-ColorOutput "🔐 Connecting to Azure..." "Cyan"
        Connect-AzAccount -WarningAction SilentlyContinue
        Write-ColorOutput "✅ Connected to Azure" "Green"
        
        return $true
    }
    catch {
        Write-ColorOutput "❌ Failed to connect to services: $($_.Exception.Message)" "Red"
        return $false
    }
}

function New-SelfSignedCertificateForApp {
    param([string]$SubjectName, [SecureString]$Password)
    
    Write-ColorOutput "🔑 Creating self-signed certificate..." "Cyan"
    
    try {
        # Create certificate
        $cert = New-SelfSignedCertificate -Subject "CN=$SubjectName" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeyExportPolicy Exportable `
            -KeySpec Signature `
            -KeyLength 2048 `
            -KeyAlgorithm RSA `
            -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears(2)
        
        # Export certificate files
        $certPath = Join-Path $CertificateExportPath "CalendarAutomationCert.cer"
        $pfxPath = Join-Path $CertificateExportPath "CalendarAutomationCert.pfx"
        
        # Export public key (.cer file)
        Export-Certificate -Cert $cert -FilePath $certPath -Force | Out-Null
        
        # Export private key (.pfx file)
        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $Password -Force | Out-Null
        
        Write-ColorOutput "✅ Certificate created successfully" "Green"
        Write-ColorOutput "📄 Certificate files saved to:" "Yellow"
        Write-ColorOutput "   Public Key (.cer): $certPath" "White"
        Write-ColorOutput "   Private Key (.pfx): $pfxPath" "White"
        
        return @{
            Certificate = $cert
            CerPath = $certPath
            PfxPath = $pfxPath
        }
    }
    catch {
        Write-ColorOutput "❌ Failed to create certificate: $($_.Exception.Message)" "Red"
        return $null
    }
}

function New-AppRegistration {
    param([string]$DisplayName, $Certificate)
    
    Write-ColorOutput "📝 Creating App Registration..." "Cyan"
    
    try {
        # Prepare certificate key credential
        $keyCredential = @{
            Type = "AsymmetricX509Cert"
            Usage = "Verify"
            Key = $Certificate.Certificate.RawData
            DisplayName = "Calendar Automation Certificate"
        }
        
        # Create the app registration
        $appRegistration = New-MgApplication -DisplayName $DisplayName `
            -SignInAudience "AzureADMyOrg" `
            -KeyCredentials @($keyCredential) `
            -RequiredResourceAccess @(
                @{
                    ResourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
                    ResourceAccess = @(
                        @{ Id = "1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; Type = "Role" }, # Application.ReadWrite.All
                        @{ Id = "19dbc75e-c2e2-444c-a770-ec69d8559fc7"; Type = "Role" }, # Directory.Read.All
                        @{ Id = "ef54d2bf-783f-4e0f-bca1-3210c0444d99"; Type = "Role" }, # Calendars.ReadWrite
                        @{ Id = "df021288-bdef-4463-88db-98f22de89214"; Type = "Role" }, # User.Read.All
                        @{ Id = "5b567255-7703-4780-807c-7be8301ae99b"; Type = "Role" }, # Group.Read.All
                        @{ Id = "75359482-378d-4052-8f01-80520e7db3cd"; Type = "Role" }, # Files.Read
                        @{ Id = "8e8e4742-1d95-4f68-9d56-6ee75648c72a"; Type = "Role" }  # Sites.Selected
                    )
                },
                @{
                    ResourceAppId = "00000002-0000-0ff1-ce00-000000000000" # Exchange Online
                    ResourceAccess = @(
                        @{ Id = "dc890d15-9560-4a4c-9b7f-a736ec74ec40"; Type = "Role" } # Exchange.ManageAsApp
                    )
                }
            )
        
        Write-ColorOutput "✅ App Registration created successfully" "Green"
        Write-ColorOutput "📋 App Registration Details:" "Yellow"
        Write-ColorOutput "   Application ID: $($appRegistration.AppId)" "White"
        Write-ColorOutput "   Object ID: $($appRegistration.Id)" "White"
        Write-ColorOutput "   Display Name: $($appRegistration.DisplayName)" "White"
        
        return $appRegistration
    }
    catch {
        Write-ColorOutput "❌ Failed to create App Registration: $($_.Exception.Message)" "Red"
        return $null
    }
}

function Add-CertificateToAutomationAccount {
    param([string]$AutomationAccount, [string]$ResourceGroup, [string]$PfxPath, [SecureString]$Password)
    
    Write-ColorOutput "🔄 Uploading certificate to Automation Account..." "Cyan"
    
    try {
        # Import certificate to Automation Account
        $result = Import-AzAutomationCertificate -AutomationAccountName $AutomationAccount `
            -ResourceGroupName $ResourceGroup `
            -Name "CalendarAutomationCert" `
            -Path $PfxPath `
            -Password $Password `
            -Exportable:$false
        
        Write-ColorOutput "✅ Certificate uploaded to Automation Account successfully" "Green"
        return $true
    }
    catch {
        Write-ColorOutput "❌ Failed to upload certificate to Automation Account: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Main execution
try {
    Write-ColorOutput "🚀 Exchange Calendar Event Management Automation - App Registration Setup" "Magenta"
    Write-ColorOutput "=" * 80 "Magenta"
    
    # Check prerequisites
    if (!(Test-Prerequisites)) {
        Write-ColorOutput "❌ Prerequisites not met. Please install required modules and try again." "Red"
        exit 1
    }
    
    # Get certificate password if not provided
    if (-not $CertificatePassword) {
        Write-ColorOutput "🔐 Please enter a password for the certificate (min 8 characters):" "Yellow"
        $CertificatePassword = Read-Host "Certificate Password" -AsSecureString
        
        # Validate password strength
        $plaintextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword))
        if ($plaintextPassword.Length -lt 8) {
            Write-ColorOutput "❌ Password must be at least 8 characters long" "Red"
            exit 1
        }
    }
    
    # Connect to services
    if (!(Connect-ToServices)) {
        Write-ColorOutput "❌ Failed to connect to required services" "Red"
        exit 1
    }
    
    # Create certificate
    $certResult = New-SelfSignedCertificateForApp -SubjectName $AppDisplayName -Password $CertificatePassword
    if (-not $certResult) {
        Write-ColorOutput "❌ Certificate creation failed" "Red"
        exit 1
    }
    
    # Create App Registration
    $appReg = New-AppRegistration -DisplayName $AppDisplayName -Certificate $certResult
    if (-not $appReg) {
        Write-ColorOutput "❌ App Registration creation failed" "Red"
        exit 1
    }
    
    # Upload certificate to Automation Account
    $uploadSuccess = Add-CertificateToAutomationAccount -AutomationAccount $AutomationAccountName `
        -ResourceGroup $AutomationResourceGroupName `
        -PfxPath $certResult.PfxPath `
        -Password $CertificatePassword
    
    if (-not $uploadSuccess) {
        Write-ColorOutput "⚠️ App Registration created but certificate upload to Automation Account failed" "Yellow"
        Write-ColorOutput "You can manually upload the certificate later using the Azure Portal" "Yellow"
    }
    
    # Output summary
    Write-ColorOutput "`n" + "=" * 80 "Green"
    Write-ColorOutput "✅ SETUP COMPLETED SUCCESSFULLY!" "Green"
    Write-ColorOutput "=" * 80 "Green"
    
    Write-ColorOutput "📋 Summary:" "Yellow"
    Write-ColorOutput "✅ App Registration created: $($appReg.DisplayName)" "White"
    Write-ColorOutput "✅ Application ID: $($appReg.AppId)" "White"
    Write-ColorOutput "✅ Certificate created and uploaded to App Registration" "White"
    Write-ColorOutput "$(if($uploadSuccess){'✅'}else{'⚠️'}) Certificate upload to Automation Account: $(if($uploadSuccess){'Success'}else{'Manual upload required'})" "White"
    Write-ColorOutput "✅ Certificate files saved to: $CertificateExportPath" "White"
    
    Write-ColorOutput "`n📝 Next Steps:" "Yellow"
    Write-ColorOutput "1. Run script 2-Configure-Permissions.ps1 to configure API permissions" "Cyan"
    Write-ColorOutput "2. Run script 3-Configure-SharePoint-Permissions.ps1 for Sites.Selected permissions" "Cyan"
    Write-ColorOutput "3. Upload the PowerShell runbook to your Automation Account" "Cyan"
    
    Write-ColorOutput "`n🔑 IMPORTANT - Save these values for your Automation Account variables:" "Yellow"
    Write-ColorOutput "ClientId: $($appReg.AppId)" "White"
    Write-ColorOutput "TenantId: $((Get-MgContext).TenantId)" "White"
    
    # Save configuration to file for next scripts
    $config = @{
        AppId = $appReg.AppId
        ObjectId = $appReg.Id
        TenantId = (Get-MgContext).TenantId
        CertificateName = "CalendarAutomationCert"
        AutomationAccountName = $AutomationAccountName
        AutomationResourceGroupName = $AutomationResourceGroupName
    }
    
    $configPath = Join-Path $CertificateExportPath "app-config.json"
    $config | ConvertTo-Json | Set-Content -Path $configPath
    Write-ColorOutput "💾 Configuration saved to: $configPath" "Green"
    
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
        Disconnect-AzAccount -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnection errors
    }
}

Write-ColorOutput "`n🎉 App Registration setup completed! Run the next script to configure permissions." "Green"