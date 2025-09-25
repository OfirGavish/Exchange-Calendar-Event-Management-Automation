# 🤖 Automation Setup Scripts

This folder contains PowerShell scripts to automate the setup of Exchange Calendar Event Management Automation.

## 📝 Scripts Overview

| Script | Purpose | Prerequisites |
|--------|---------|---------------|
| `1-Setup-AppRegistration.ps1` | Creates App Registration & certificates | Global Admin + Az.Automation |
| `2-Configure-Permissions.ps1` | Configures API permissions | Global Admin |
| `3-Configure-SharePoint-Permissions.ps1` | Configures SharePoint access | SharePoint Admin + PnP.PowerShell |

## 🚀 Quick Start

```powershell
# Install required modules
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Az.Automation, PnP.PowerShell -Scope CurrentUser

# Run scripts in order
.\1-Setup-AppRegistration.ps1 -AutomationAccountName "your-aa" -AutomationResourceGroupName "your-rg"
.\2-Configure-Permissions.ps1
.\3-Configure-SharePoint-Permissions.ps1 -SharePointSiteUrl "https://tenant.sharepoint.com/sites/events"
```

## 📖 Documentation

- **[INSTALLATION-GUIDE.md](INSTALLATION-GUIDE.md)** - Complete setup documentation
- **[../README.md](../README.md)** - Main project documentation
- **[../deploy/DEPLOYMENT-GUIDE.md](../deploy/DEPLOYMENT-GUIDE.md)** - Azure resource deployment

## ⚡ What Gets Automated

These scripts automate approximately **80%** of the manual setup process:

### ✅ Fully Automated
- App Registration creation
- Certificate generation and upload
- API permissions configuration
- SharePoint Sites.Selected permissions
- Variable configuration

### 🔧 Still Manual
- PowerShell module installation in Automation Account
- Admin consent (guided with URL)
- Runbook upload
- Testing and validation

## 🎯 Benefits

- **Faster Setup**: Reduces setup time from hours to minutes
- **Consistent Configuration**: Same setup every time
- **Error Reduction**: Automated validation and error handling
- **Documentation**: Clear next steps and troubleshooting
- **Idempotent**: Safe to run multiple times

Perfect for IT administrators who need to deploy this solution across multiple environments! 🚀