# Deployment Guide

This guide walks you through deploying the Exchange Calendar Event Management Automation solution using the provided ARM templates.

## 🚀 Quick Deploy Options

### Option 1: Deploy Both Resources Together

If you prefer to deploy everything at once, deploy in this order:

1. **Deploy Storage Account first** - The Automation Account references the storage account name
2. **Deploy Automation Account second** - Requires the storage account name as parameter

### Option 2: Custom Parameters

Both templates support custom parameters. You can either:
- Use the Azure Portal UI to fill in parameters during deployment
- Download and modify the parameter files:
  - [`storage-parameters.json`](storage-parameters.json)
  - [`automation-parameters.json`](automation-parameters.json)

## 📋 Storage Account Deployment

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fofirga%2FExchange-Calendar-Event-Management-Automation%2Fmain%2Fdeploy%2Fstorage-template.json)

### Parameters Required:
- **Storage Account Name**: Must be globally unique, 3-24 characters, lowercase letters and numbers only
- **Location**: Azure region (default: same as resource group)

### What Gets Deployed:
- ✅ Azure Storage Account (Standard_LRS, StorageV2)
- ✅ Static website hosting enabled
- ✅ CORS rules configured for web dashboard access
- ✅ `$web` container created for hosting files
- ✅ `logs` container created for automation logs

### After Deployment:
1. Note the **Primary Endpoint URL** from the output
2. Upload `index.html` and `log-analyzer.html` to the `$web` container
3. The storage account is ready for use by the Automation Account

## 🤖 Automation Account Deployment

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fofirga%2FExchange-Calendar-Event-Management-Automation%2Fmain%2Fdeploy%2Fautomation-template.json)

### Parameters Required:
- **Automation Account Name**: Descriptive name for your automation account
- **Location**: Azure region (default: same as resource group)
- **Tenant ID**: Your Azure AD tenant ID
- **Tenant Domain**: Your tenant domain (e.g., `contoso.onmicrosoft.com`)
- **SharePoint Site URL**: URL to your SharePoint site
- **SharePoint Library Path**: Path to your document library
- **Excel File Name**: Name of your Excel file with event data
- **Default Organizer Email**: Email address for the default event organizer
- **Storage Account Name**: Name of the storage account (must already exist)

### What Gets Deployed:
- ✅ Azure Automation Account with managed identity
- ✅ All configuration variables pre-configured
- ✅ System-assigned managed identity enabled
- ✅ Ready for PowerShell module installation

### After Deployment:
You'll need to complete these manual steps:

1. **Install PowerShell Modules** (in order):
   - Microsoft.Graph.Authentication
   - Microsoft.Graph
   - ImportExcel
   - ExchangeOnlineManagement
   - Az.Storage
   - Az.Accounts

2. **Create App Registration**:
   - Register a new app in Azure AD
   - Generate client certificate
   - Configure API permissions

3. **Upload Certificates**:
   - Upload .cer file to App Registration
   - Upload .pfx file to Automation Account

4. **Configure Permissions**:
   - Grant admin consent for API permissions
   - Set up Sites.Selected permissions for SharePoint

5. **Upload Runbook**:
   - Create new PowerShell runbook
   - Upload the `eventscalendarautomation.ps1` script

## 🔧 Parameter Files

### Storage Parameters Template

```json
{
    "$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
        "storageAccountName": {
            "value": "yourstorageaccountname"
        },
        "location": {
            "value": "East US"
        }
    }
}
```

### Automation Parameters Template

```json
{
    "$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
        "automationAccountName": {
            "value": "your-automation-account"
        },
        "tenantId": {
            "value": "YOUR_TENANT_ID_HERE"
        },
        "tenantDomain": {
            "value": "your-tenant.onmicrosoft.com"
        },
        "sharePointSiteUrl": {
            "value": "https://your-tenant.sharepoint.com/sites/events"
        },
        "sharePointLibraryPath": {
            "value": "/sites/events/Shared%20Documents/EventData"
        },
        "defaultOrganizerEmail": {
            "value": "organizer@your-domain.com"
        },
        "storageAccountName": {
            "value": "yourstorageaccountname"
        }
    }
}
```

## 🛠️ Manual Deployment via Azure CLI

If you prefer command-line deployment:

### Deploy Storage Account

```bash
# Create resource group
az group create --name "rg-calendar-automation" --location "East US"

# Deploy storage account
az deployment group create \
  --resource-group "rg-calendar-automation" \
  --template-file "deploy/storage-template.json" \
  --parameters "deploy/storage-parameters.json"
```

### Deploy Automation Account

```bash
# Deploy automation account
az deployment group create \
  --resource-group "rg-calendar-automation" \
  --template-file "deploy/automation-template.json" \
  --parameters "deploy/automation-parameters.json"
```

## 📝 Next Steps

After successful deployment:

1. Follow the [complete setup guide](../README.md#-manual-setup-instructions) for manual configuration steps
2. Test the automation with a small Excel file
3. Monitor the web dashboard for successful operation
4. Set up scheduling for automated runs

## 🆘 Troubleshooting Deployment

### Common Issues:

1. **Storage Account Name Conflicts**: Storage account names must be globally unique
2. **Parameter Validation Errors**: Ensure all required parameters are provided
3. **Permissions Issues**: Ensure you have Contributor access to the resource group
4. **Region Limitations**: Some regions may not support all resource types

### Getting Help:

- Check the deployment logs in the Azure Portal
- Review the ARM template validation errors
- Refer to the main [README troubleshooting section](README.md#troubleshooting)

## 📋 Deployment Checklist

- [ ] Resource group created
- [ ] Storage account deployed successfully
- [ ] Automation account deployed successfully
- [ ] PowerShell modules installed
- [ ] App registration created and configured
- [ ] Certificates uploaded
- [ ] API permissions granted
- [ ] Sites.Selected permissions configured
- [ ] Runbook uploaded and tested
- [ ] Web dashboard files uploaded
- [ ] End-to-end testing completed