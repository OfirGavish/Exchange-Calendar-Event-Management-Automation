# 🚀 Automated Setup Scripts Guide

This guide walks you through using the PowerShell automation scripts to configure your Exchange Calendar Event Management Automation solution. These scripts automate most of the manual setup process.

## 📋 Overview

The setup process is divided into three automated scripts that handle different aspects of the configuration:

1. **`1-Setup-AppRegistration.ps1`** - Creates App Registration and certificates
2. **`2-Configure-Permissions.ps1`** - Configures API permissions and admin consent
3. **`3-Configure-SharePoint-Permissions.ps1`** - Configures SharePoint Sites.Selected permissions

## ⚡ Quick Start

If you just want to get started quickly:

```powershell
# Navigate to the scripts folder
Set-Location "scripts"

# Run all three scripts in sequence
.\1-Setup-AppRegistration.ps1 -AutomationAccountName "your-automation-account" -AutomationResourceGroupName "your-resource-group"
.\2-Configure-Permissions.ps1
.\3-Configure-SharePoint-Permissions.ps1 -SharePointSiteUrl "https://your-tenant.sharepoint.com/sites/events"
```

## 📦 Prerequisites

### PowerShell Modules
Install these modules before running the scripts:

```powershell
# For App Registration and Permissions
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module -Name Microsoft.Graph.Applications -Scope CurrentUser -Force

# For Azure Automation Account
Install-Module -Name Az.Automation -Scope CurrentUser -Force

# For SharePoint permissions
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

### Azure Permissions Required
- **Global Administrator** or **Application Administrator** role in Azure AD
- **Contributor** access to the Azure Automation Account and Resource Group
- **SharePoint Administrator** or **Global Administrator** for SharePoint permissions

## 📝 Script 1: App Registration & Certificate Setup

### Purpose
Creates the Azure AD App Registration and generates certificates for authentication.

### Usage
```powershell
.\1-Setup-AppRegistration.ps1 -AutomationAccountName "calendar-automation-aa" -AutomationResourceGroupName "rg-calendar-automation"
```

### Parameters
- `AppDisplayName` (Optional): Name for the App Registration
- `CertificatePassword` (Optional): Password for certificate (will prompt if not provided)
- `AutomationAccountName` (Required): Name of your Automation Account
- `AutomationResourceGroupName` (Required): Resource Group containing the Automation Account
- `CertificateExportPath` (Optional): Where to save certificate files

### What It Does
✅ Creates Azure AD App Registration  
✅ Generates self-signed certificate (2-year validity)  
✅ Uploads certificate to App Registration  
✅ Uploads certificate to Automation Account  
✅ Configures basic app settings  
✅ Saves configuration file for next scripts  

### Output Files
- `CalendarAutomationCert.cer` - Public certificate (for App Registration)
- `CalendarAutomationCert.pfx` - Private certificate (for Automation Account)
- `app-config.json` - Configuration file for other scripts

### Example Output
```
🎉 App Registration setup completed!
📋 Application ID: 12345678-1234-1234-1234-123456789012
🔑 Certificate uploaded to both App Registration and Automation Account
💾 Configuration saved to: app-config.json
```

## 🔑 Script 2: API Permissions Configuration

### Purpose
Configures all required Microsoft Graph and Exchange Online API permissions.

### Usage
```powershell
# Uses config file from Script 1
.\2-Configure-Permissions.ps1

# Or specify App ID manually
.\2-Configure-Permissions.ps1 -AppId "12345678-1234-1234-1234-123456789012"
```

### Parameters
- `AppId` (Optional): App Registration ID (loads from config if not provided)
- `ConfigFilePath` (Optional): Path to config file from Script 1

### What It Does
✅ Configures Microsoft Graph permissions:
- `Calendars.ReadWrite` - Create calendar events
- `Directory.Read.All` - Read directory data
- `User.Read.All` - Read user profiles  
- `Group.Read.All` - Read group information
- `Files.Read` - Read SharePoint files
- `Sites.Selected` - Access specific SharePoint sites

✅ Configures Exchange Online permissions:
- `Exchange.ManageAsApp` - Manage Exchange as application

✅ Verifies permission grants  
✅ Provides admin consent URL  

### Admin Consent Required
After running this script, you **MUST** visit the provided admin consent URL to grant permissions:
```
https://login.microsoftonline.com/{tenant-id}/v2.0/adminconsent?client_id={app-id}&...
```

### Example Output
```
✅ API permissions configured successfully
🌐 Admin Consent URL: https://login.microsoftonline.com/...
📝 Please visit the URL above to grant admin consent
```

## 🌐 Script 3: SharePoint Permissions Configuration

### Purpose
Configures SharePoint Sites.Selected permissions for specific site access.

### Usage
```powershell
.\3-Configure-SharePoint-Permissions.ps1 -SharePointSiteUrl "https://contoso.sharepoint.com/sites/events"
```

### Parameters
- `AppId` (Optional): App Registration ID (loads from config if not provided)
- `SharePointSiteUrl` (Required): SharePoint site URL where Excel files are stored
- `ConfigFilePath` (Optional): Path to config file from Script 1
- `PermissionLevel` (Optional): Read, Write, or FullControl (default: Write)

### What It Does
✅ Connects to SharePoint Admin Center  
✅ Grants Sites.Selected permission to specific site  
✅ Verifies permission configuration  
✅ Provides troubleshooting guidance  

### SharePoint Site URL Format
Must be in the format: `https://tenant.sharepoint.com/sites/sitename`

### Example Output
```
✅ Sites.Selected permission granted successfully
✅ Permission verification: Verified
🎉 App can now read and write Excel files in SharePoint
```

## 🔄 Complete Setup Workflow

### Step-by-Step Process

1. **Deploy Azure Resources** (if not already done):
   ```powershell
   # Use the Deploy to Azure buttons in README.md
   # Or deploy manually via Azure Portal/CLI
   ```

2. **Run Setup Scripts**:
   ```powershell
   # Script 1: App Registration & Certificates
   .\1-Setup-AppRegistration.ps1 -AutomationAccountName "your-aa-name" -AutomationResourceGroupName "your-rg-name"
   
   # Script 2: API Permissions
   .\2-Configure-Permissions.ps1
   
   # Visit the admin consent URL provided by Script 2
   
   # Script 3: SharePoint Permissions  
   .\3-Configure-SharePoint-Permissions.ps1 -SharePointSiteUrl "https://your-tenant.sharepoint.com/sites/events"
   ```

3. **Upload Runbook**:
   - Go to Azure Portal → Automation Account → Runbooks
   - Create new PowerShell runbook
   - Upload `eventscalendarautomation.ps1`

4. **Test the Setup**:
   - Create a small test Excel file in SharePoint
   - Run the runbook manually
   - Check logs in the web dashboard

## 🛠️ Troubleshooting

### Common Issues

#### Module Installation Errors
```powershell
# If you get execution policy errors
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# If modules fail to install
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
```

#### Permission Errors
- Ensure you have **Global Administrator** or **Application Administrator** role
- Check that you're connected to the correct tenant
- Verify the Automation Account exists and you have Contributor access

#### Certificate Upload Failures
- Ensure the Automation Account name and resource group are correct
- Check that you have Contributor permissions on the Automation Account
- Verify the certificate password is correct

#### SharePoint Connection Issues
- Ensure you have SharePoint Administrator permissions
- Check that the SharePoint site URL is correct and accessible
- Verify the site exists and you have access to it

### Manual Verification Steps

#### Check App Registration
1. Azure Portal → Azure Active Directory → App registrations
2. Find your app registration
3. Verify certificates are uploaded
4. Check API permissions are configured

#### Check Automation Account
1. Azure Portal → Automation Account → Certificates
2. Verify `CalendarAutomationCert` is uploaded
3. Check variables are configured

#### Check SharePoint Permissions
1. SharePoint Admin Center → Advanced → API access
2. Look for your app registration
3. Verify it has access to your site

## 📊 Configuration Summary

After running all scripts successfully, you should have:

### ✅ App Registration
- Display Name: Exchange Calendar Event Management Automation
- Application ID: Generated unique ID
- Authentication: Certificate-based
- Permissions: Microsoft Graph + Exchange Online

### ✅ Certificates
- Self-signed certificate (2-year validity)
- Uploaded to App Registration (.cer file)
- Uploaded to Automation Account (.pfx file)
- Saved locally for backup

### ✅ API Permissions
- Microsoft Graph: Calendars, Directory, Users, Groups, Files, Sites
- Exchange Online: ManageAsApp
- Admin consent: Granted

### ✅ SharePoint Access
- Sites.Selected permission configured
- Write access to specified SharePoint site
- Ready for Excel file read/write operations

## 🔄 Re-running Scripts

The scripts are designed to be **idempotent** - you can run them multiple times safely:

- **Script 1**: Will detect existing app registration and certificates
- **Script 2**: Will skip already-granted permissions
- **Script 3**: Will detect existing SharePoint permissions

## 📞 Getting Help

If you encounter issues:

1. **Check Prerequisites**: Ensure all modules are installed and you have required permissions
2. **Review Error Messages**: Scripts provide detailed error information
3. **Manual Verification**: Use the troubleshooting steps above
4. **Check Logs**: Azure Automation Account logs can provide additional insights

## 🚀 What's Next?

After successful script execution:

1. **Upload the PowerShell runbook** to your Automation Account
2. **Test with a small Excel file** to verify everything works
3. **Set up scheduling** for automated execution
4. **Monitor the web dashboard** for ongoing operations

The automation setup is now complete and ready for production use! 🎉