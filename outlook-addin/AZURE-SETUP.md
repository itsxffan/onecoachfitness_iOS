# Azure Setup Guide

This guide walks through setting up all required Azure resources for the OneCoach Email Manager add-in.

## Prerequisites

- Azure subscription
- Azure CLI installed
- Microsoft 365 tenant admin access
- Node.js 16+ installed

## Overview

Required Azure resources:
1. Azure AD App Registration
2. Hosting service (Static Web Apps, App Service, or Storage)
3. (Optional) Azure Functions for backend API
4. (Optional) Application Insights for monitoring

## Part 1: Azure AD App Registration

### Step 1: Create App Registration

```bash
# Login to Azure
az login

# Create app registration
az ad app create \
  --display-name "OneCoach Email Manager" \
  --sign-in-audience "AzureADMultipleOrgs" \
  --web-redirect-uris "https://localhost:3000/auth-callback.html"
```

Save the `appId` from the output.

### Step 2: Create Service Principal

```bash
# Replace YOUR_APP_ID with the appId from previous step
az ad sp create --id YOUR_APP_ID
```

### Step 3: Add Microsoft Graph Permissions

```bash
# Get Microsoft Graph API ID
GRAPH_API_ID="00000003-0000-0000-c000-000000000000"

# Add delegated permissions
az ad app permission add \
  --id YOUR_APP_ID \
  --api $GRAPH_API_ID \
  --api-permissions \
    e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope \  # User.Read
    024d486e-b451-40bb-833d-3e66d98c5c73=Scope \  # Mail.ReadWrite
    65e50fdc-43b7-4915-933e-e8138f11f40a=Scope \  # Sites.ReadWrite.All
    5a54b8b3-347c-476d-8f8e-42d5c7424d29=Scope    # Files.ReadWrite.All

# Grant admin consent (requires Global Admin)
az ad app permission admin-consent --id YOUR_APP_ID
```

### Step 4: Configure Authentication

```bash
# Enable implicit flow for tokens
az ad app update \
  --id YOUR_APP_ID \
  --set oauth2AllowIdTokenImplicitFlow=true \
  --set oauth2AllowImplicitFlow=true
```

### Step 5: Create Client Secret (Optional - for backend)

```bash
az ad app credential reset \
  --id YOUR_APP_ID \
  --append
```

Save the password (client secret) securely.

## Part 2: Hosting Setup

### Option A: Azure Static Web Apps

```bash
# Create resource group
az group create \
  --name onecoach-rg \
  --location eastus

# Create static web app
az staticwebapp create \
  --name onecoach-outlook-addin \
  --resource-group onecoach-rg \
  --source https://github.com/YOUR-ORG/YOUR-REPO \
  --location eastus2 \
  --branch main \
  --app-location "outlook-addin" \
  --output-location "dist"

# Get the hostname
az staticwebapp show \
  --name onecoach-outlook-addin \
  --resource-group onecoach-rg \
  --query "defaultHostname" \
  --output tsv
```

### Option B: Azure App Service

```bash
# Create App Service Plan
az appservice plan create \
  --name onecoach-plan \
  --resource-group onecoach-rg \
  --sku B1 \
  --is-linux

# Create Web App
az webapp create \
  --resource-group onecoach-rg \
  --plan onecoach-plan \
  --name onecoach-outlook-addin \
  --runtime "NODE|16-lts"

# Configure for SPA
az webapp config set \
  --resource-group onecoach-rg \
  --name onecoach-outlook-addin \
  --startup-file "pm2 serve dist --no-daemon --spa"

# Enable HTTPS only
az webapp update \
  --resource-group onecoach-rg \
  --name onecoach-outlook-addin \
  --https-only true
```

### Option C: Azure Storage + CDN

```bash
# Create storage account
az storage account create \
  --name onecoachstorage \
  --resource-group onecoach-rg \
  --location eastus \
  --sku Standard_LRS \
  --kind StorageV2

# Enable static website hosting
az storage blob service-properties update \
  --account-name onecoachstorage \
  --static-website \
  --index-document index.html \
  --404-document 404.html

# Create CDN profile
az cdn profile create \
  --resource-group onecoach-rg \
  --name onecoach-cdn \
  --sku Standard_Microsoft

# Get storage endpoint
STORAGE_ENDPOINT=$(az storage account show \
  --name onecoachstorage \
  --resource-group onecoach-rg \
  --query "primaryEndpoints.web" \
  --output tsv)

# Create CDN endpoint
az cdn endpoint create \
  --resource-group onecoach-rg \
  --profile-name onecoach-cdn \
  --name onecoach-addin \
  --origin $STORAGE_ENDPOINT \
  --origin-host-header $(echo $STORAGE_ENDPOINT | sed 's|https://||' | sed 's|/||')

# Enable HTTPS
az cdn endpoint update \
  --resource-group onecoach-rg \
  --profile-name onecoach-cdn \
  --name onecoach-addin \
  --https-only
```

## Part 3: Optional Backend API (Azure Functions)

### Create Azure Functions App

```bash
# Create storage for functions
az storage account create \
  --name onecoachfuncstorage \
  --resource-group onecoach-rg \
  --location eastus \
  --sku Standard_LRS

# Create function app
az functionapp create \
  --resource-group onecoach-rg \
  --consumption-plan-location eastus \
  --runtime node \
  --runtime-version 16 \
  --functions-version 4 \
  --name onecoach-api \
  --storage-account onecoachfuncstorage

# Configure CORS
az functionapp cors add \
  --resource-group onecoach-rg \
  --name onecoach-api \
  --allowed-origins "https://your-frontend-domain.com"

# Add app settings
az functionapp config appsettings set \
  --resource-group onecoach-rg \
  --name onecoach-api \
  --settings \
    "AZURE_AD_CLIENT_ID=YOUR_APP_ID" \
    "AZURE_AD_TENANT_ID=YOUR_TENANT_ID" \
    "AZURE_AD_CLIENT_SECRET=YOUR_CLIENT_SECRET" \
    "SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/yoursite"
```

### Create Function for SharePoint Operations

Create `azure-functions/SharePointProxy/index.js`:

```javascript
module.exports = async function (context, req) {
    // Authentication middleware
    const token = req.headers.authorization?.split('Bearer ')[1];
    if (!token) {
        context.res = {
            status: 401,
            body: { error: 'Unauthorized' }
        };
        return;
    }

    // Proxy request to SharePoint
    const { library, fileName, content } = req.body;
    
    try {
        // Upload to SharePoint using Graph API
        // Implementation here
        
        context.res = {
            status: 200,
            body: { success: true }
        };
    } catch (error) {
        context.res = {
            status: 500,
            body: { error: error.message }
        };
    }
};
```

Create `azure-functions/SharePointProxy/function.json`:

```json
{
  "bindings": [
    {
      "authLevel": "anonymous",
      "type": "httpTrigger",
      "direction": "in",
      "name": "req",
      "methods": ["post"],
      "route": "sharepoint/upload"
    },
    {
      "type": "http",
      "direction": "out",
      "name": "res"
    }
  ]
}
```

### Deploy Functions

```bash
# Install Azure Functions Core Tools
npm install -g azure-functions-core-tools@4

# Navigate to functions directory
cd azure-functions

# Install dependencies
npm install

# Deploy
func azure functionapp publish onecoach-api
```

## Part 4: Application Insights (Optional)

### Create Application Insights

```bash
# Create Application Insights
az monitor app-insights component create \
  --app onecoach-insights \
  --location eastus \
  --resource-group onecoach-rg

# Get instrumentation key
INSTRUMENTATION_KEY=$(az monitor app-insights component show \
  --app onecoach-insights \
  --resource-group onecoach-rg \
  --query "instrumentationKey" \
  --output tsv)

echo "Instrumentation Key: $INSTRUMENTATION_KEY"
```

### Connect to Your App

Add to your app configuration:

```typescript
// src/config/appInsights.ts
import { ApplicationInsights } from '@microsoft/applicationinsights-web';

const appInsights = new ApplicationInsights({
  config: {
    instrumentationKey: 'YOUR_INSTRUMENTATION_KEY'
  }
});

appInsights.loadAppInsights();
appInsights.trackPageView();

export default appInsights;
```

## Part 5: Update Configuration

### Update authService.ts

```typescript
const clientId = 'YOUR_APP_ID_HERE';
const tenantId = 'common'; // or your specific tenant ID
```

### Update manifest.xml

Replace all URLs with your production URLs:

```xml
<SourceLocation DefaultValue="https://your-production-url.azurestaticapps.net/taskpane.html"/>
```

### Update Redirect URIs

Add production redirect URIs to Azure AD:

```bash
az ad app update \
  --id YOUR_APP_ID \
  --add oauth2AllowIdTokenImplicitFlow=true \
  --web-redirect-uris \
    "https://localhost:3000/auth-callback.html" \
    "https://your-production-url.azurestaticapps.net/auth-callback.html"
```

## Part 6: Security Configuration

### Enable Azure AD Conditional Access (Optional)

```bash
# This requires additional configuration in Azure AD portal
# Navigate to: Azure AD > Security > Conditional Access
# Create policies for:
# - Require MFA for external access
# - Require compliant devices
# - Block legacy authentication
```

### Configure Key Vault for Secrets (Optional)

```bash
# Create Key Vault
az keyvault create \
  --name onecoach-vault \
  --resource-group onecoach-rg \
  --location eastus

# Store secrets
az keyvault secret set \
  --vault-name onecoach-vault \
  --name "AzureAD-ClientSecret" \
  --value "YOUR_CLIENT_SECRET"

# Grant access to your app
az keyvault set-policy \
  --name onecoach-vault \
  --spn YOUR_APP_ID \
  --secret-permissions get list
```

## Part 7: Monitoring and Alerts

### Create Alert Rules

```bash
# Alert for high error rate
az monitor metrics alert create \
  --name "High Error Rate" \
  --resource-group onecoach-rg \
  --scopes "/subscriptions/YOUR_SUBSCRIPTION_ID/resourceGroups/onecoach-rg/providers/Microsoft.Web/sites/onecoach-outlook-addin" \
  --condition "avg Percentage CPU > 80" \
  --description "Alert when CPU usage is high" \
  --evaluation-frequency 5m \
  --window-size 15m \
  --severity 2

# Alert for availability
az monitor metrics alert create \
  --name "App Down" \
  --resource-group onecoach-rg \
  --scopes "/subscriptions/YOUR_SUBSCRIPTION_ID/resourceGroups/onecoach-rg/providers/Microsoft.Web/sites/onecoach-outlook-addin" \
  --condition "avg Health check status < 1" \
  --description "Alert when app is down" \
  --evaluation-frequency 1m \
  --window-size 5m \
  --severity 0
```

## Part 8: Backup and Recovery

### Enable Backups (App Service)

```bash
# Create storage account for backups
az storage account create \
  --name onecoachbackups \
  --resource-group onecoach-rg \
  --location eastus \
  --sku Standard_LRS

# Get storage connection string
STORAGE_CONNECTION=$(az storage account show-connection-string \
  --name onecoachbackups \
  --resource-group onecoach-rg \
  --output tsv)

# Configure backup
az webapp config backup update \
  --resource-group onecoach-rg \
  --webapp-name onecoach-outlook-addin \
  --container-url "$STORAGE_CONNECTION" \
  --frequency 1d \
  --retain-one true \
  --retention 30
```

## Part 9: Cost Optimization

### Enable Auto-scaling

```bash
# Create autoscale settings
az monitor autoscale create \
  --resource-group onecoach-rg \
  --resource onecoach-outlook-addin \
  --resource-type "Microsoft.Web/sites" \
  --name "Autoscale Settings" \
  --min-count 1 \
  --max-count 3 \
  --count 1

# Add CPU-based rule
az monitor autoscale rule create \
  --resource-group onecoach-rg \
  --autoscale-name "Autoscale Settings" \
  --condition "Percentage CPU > 70 avg 5m" \
  --scale out 1
```

## Verification Checklist

- [ ] Azure AD app registered
- [ ] Permissions granted and consented
- [ ] Hosting service deployed
- [ ] SSL/HTTPS configured
- [ ] manifest.xml updated with production URLs
- [ ] authService.ts updated with client ID
- [ ] SharePoint site URL configured
- [ ] Application Insights enabled (optional)
- [ ] Monitoring alerts configured
- [ ] Backup enabled
- [ ] Documentation updated

## Useful Commands

```bash
# View all resources
az resource list --resource-group onecoach-rg --output table

# View app logs
az webapp log tail --resource-group onecoach-rg --name onecoach-outlook-addin

# Restart app
az webapp restart --resource-group onecoach-rg --name onecoach-outlook-addin

# View app metrics
az monitor metrics list \
  --resource "/subscriptions/YOUR_SUBSCRIPTION_ID/resourceGroups/onecoach-rg/providers/Microsoft.Web/sites/onecoach-outlook-addin" \
  --metric "Requests" \
  --start-time 2024-01-01T00:00:00Z \
  --end-time 2024-01-02T00:00:00Z
```

## Troubleshooting

### Issue: Permission errors in Azure AD

```bash
# Check app permissions
az ad app permission list --id YOUR_APP_ID

# Re-grant consent
az ad app permission admin-consent --id YOUR_APP_ID
```

### Issue: CORS errors

```bash
# Update CORS settings
az webapp cors add \
  --resource-group onecoach-rg \
  --name onecoach-outlook-addin \
  --allowed-origins "https://outlook.office.com" "https://outlook.live.com"
```

### Issue: SSL/HTTPS not working

```bash
# Verify HTTPS only
az webapp show \
  --resource-group onecoach-rg \
  --name onecoach-outlook-addin \
  --query httpsOnly

# Force HTTPS
az webapp update \
  --resource-group onecoach-rg \
  --name onecoach-outlook-addin \
  --https-only true
```

## Support

For Azure-specific issues:
- Azure Support Portal: https://portal.azure.com/#blade/Microsoft_Azure_Support/HelpAndSupportBlade
- Documentation: https://docs.microsoft.com/azure/

For add-in issues:
- See main README.md and DEPLOYMENT.md
