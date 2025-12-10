# Deployment Guide

This guide provides detailed instructions for deploying the OneCoach Email Manager Outlook add-in to production environments.

## Pre-Deployment Checklist

- [ ] Azure AD app registration completed
- [ ] API permissions granted and consented
- [ ] Production hosting environment set up
- [ ] SSL certificate configured
- [ ] SharePoint site and libraries created
- [ ] manifest.xml updated with production URLs
- [ ] Add-in tested in development environment
- [ ] Icons created and added to assets folder
- [ ] Environment-specific configuration reviewed

## Production Build

### 1. Update Configuration

Before building for production, ensure all configuration is updated:

**src/services/authService.ts**
```typescript
const clientId = 'YOUR_PRODUCTION_CLIENT_ID';
const tenantId = 'YOUR_TENANT_ID'; // or 'common' for multi-tenant
```

**src/services/sharepointService.ts**
```typescript
this.siteUrl = 'https://yourcompany.sharepoint.com/sites/production';
```

### 2. Build Production Assets

```bash
cd outlook-addin
npm install
npm run build
```

This creates optimized files in the `dist/` directory.

### 3. Update Manifest

Update all URLs in `manifest.xml`:

```xml
<!-- Change all instances of localhost -->
<SourceLocation DefaultValue="https://your-production-domain.com/taskpane.html"/>
<bt:Url id="Commands.Url" DefaultValue="https://your-production-domain.com/commands.html"/>
<!-- Update all icon URLs -->
<IconUrl DefaultValue="https://your-production-domain.com/assets/icon-32.png"/>
```

Update the add-in ID if needed (for new deployment):
```xml
<Id>YOUR-UNIQUE-GUID-HERE</Id>
```

## Hosting Options

### Option 1: Azure Static Web Apps (Recommended)

**Advantages:**
- Free tier available
- Automatic HTTPS
- Global CDN
- Easy GitHub integration
- Built-in CI/CD

**Steps:**

1. Install Azure CLI:
```bash
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
```

2. Create Static Web App:
```bash
az login
az staticwebapp create \
  --name onecoach-outlook-addin \
  --resource-group your-resource-group \
  --source . \
  --location "East US 2" \
  --branch main \
  --app-location "outlook-addin/dist" \
  --output-location "dist"
```

3. Get the deployment URL:
```bash
az staticwebapp show \
  --name onecoach-outlook-addin \
  --resource-group your-resource-group \
  --query "defaultHostname" \
  --output tsv
```

4. Update manifest.xml with the deployment URL

### Option 2: Azure App Service

**Steps:**

1. Create App Service:
```bash
az webapp create \
  --resource-group your-resource-group \
  --plan your-app-service-plan \
  --name onecoach-outlook-addin \
  --runtime "NODE|16-lts"
```

2. Configure deployment:
```bash
az webapp deployment source config-local-git \
  --name onecoach-outlook-addin \
  --resource-group your-resource-group
```

3. Deploy:
```bash
git remote add azure <deployment-url>
git push azure main
```

### Option 3: Azure Blob Storage with CDN

**Steps:**

1. Create storage account:
```bash
az storage account create \
  --name onecoachstorage \
  --resource-group your-resource-group \
  --location eastus \
  --sku Standard_LRS
```

2. Enable static website hosting:
```bash
az storage blob service-properties update \
  --account-name onecoachstorage \
  --static-website \
  --index-document index.html
```

3. Upload files:
```bash
az storage blob upload-batch \
  --account-name onecoachstorage \
  --source ./dist \
  --destination '$web'
```

4. Get the website URL:
```bash
az storage account show \
  --name onecoachstorage \
  --resource-group your-resource-group \
  --query "primaryEndpoints.web" \
  --output tsv
```

5. (Optional) Add Azure CDN for better performance

## Deployment Methods

### Method 1: Tenant-Wide Deployment (Enterprise)

**Use Case:** Deploy to all users in your organization

**Prerequisites:**
- Microsoft 365 admin permissions
- Completed Azure AD setup

**Steps:**

1. Access Microsoft 365 admin center:
   - Go to https://admin.microsoft.com
   - Sign in with admin credentials

2. Navigate to Integrated Apps:
   - Settings > Integrated apps
   - Click "Upload custom apps"

3. Upload manifest:
   - Select "Upload manifest file (.xml)"
   - Choose your production manifest.xml
   - Click "Next"

4. Configure deployment:
   - **Assign users**: Choose from:
     - Entire organization
     - Specific users/groups
     - Just me (for testing)
   - **Deployment type**: 
     - Required (auto-installed)
     - Optional (users can install)
   - Click "Deploy"

5. Monitor deployment:
   - Check deployment status in admin center
   - Users will see the add-in within 24 hours
   - Send communication to users about the new add-in

### Method 2: AppSource Publication (Public)

**Use Case:** Make add-in available to all Microsoft 365 users

**Steps:**

1. Prepare for submission:
   - Complete Partner Center registration
   - Prepare marketing materials
   - Create demo video
   - Write comprehensive documentation

2. Create Partner Center offer:
   - Go to https://partner.microsoft.com/dashboard
   - Create new Office Add-in offer
   - Fill in product information

3. Technical configuration:
   - Upload manifest.xml
   - Provide test accounts
   - Add support contact

4. Submit for validation:
   - Microsoft reviews (typically 1-2 weeks)
   - Address any feedback
   - Resubmit if needed

5. Go live:
   - Once approved, set go-live date
   - Add-in appears in AppSource
   - Monitor reviews and analytics

### Method 3: Individual User Installation

**Use Case:** Users install add-in themselves

**Steps:**

1. Share manifest file:
   - Host manifest.xml on accessible URL
   - Or provide file directly to users

2. User installation:
   - User opens Outlook
   - Clicks "Get Add-ins"
   - Selects "My add-ins"
   - Clicks "+ Add a custom add-in"
   - Chooses "Add from file" or "Add from URL"
   - Provides manifest location

3. Verify installation:
   - Add-in appears in ribbon
   - User can open and use features

### Method 4: SharePoint App Catalog

**Use Case:** Organization-wide deployment via SharePoint

**Steps:**

1. Create App Catalog (if not exists):
   - SharePoint admin center
   - More features > Apps > Open
   - Create App Catalog site

2. Upload add-in:
   - Go to App Catalog
   - Upload manifest.xml
   - Check "Make available to all sites"
   - Click "OK"

3. Deploy to users:
   - Apps > Manage App Licenses
   - Select the add-in
   - Add users/groups
   - Click "OK"

4. Users access:
   - Add-in automatically appears
   - Or users can install from organization store

## Post-Deployment

### Monitoring

1. **Usage Analytics**:
   - Set up Application Insights
   - Track user adoption
   - Monitor error rates
   - Analyze performance metrics

2. **User Feedback**:
   - Collect user feedback
   - Monitor support tickets
   - Track feature requests
   - Regular user surveys

3. **Performance Monitoring**:
   - API response times
   - SharePoint upload success rate
   - Authentication success rate
   - Client-side errors

### Maintenance

1. **Regular Updates**:
   - Security patches
   - Feature enhancements
   - Bug fixes
   - Dependency updates

2. **Version Management**:
   - Update version in manifest.xml
   - Document changes in CHANGELOG.md
   - Test thoroughly before deployment
   - Communicate changes to users

3. **Certificate Management**:
   - Monitor SSL certificate expiration
   - Renew before expiry
   - Update if domain changes

### Rollback Procedure

If issues occur post-deployment:

1. **Immediate action**:
   - Access Microsoft 365 admin center
   - Navigate to deployed add-in
   - Click "Remove" or "Disable"

2. **Communicate**:
   - Notify users of issue
   - Provide timeline for fix
   - Offer workaround if available

3. **Fix and redeploy**:
   - Identify and fix issue
   - Test thoroughly
   - Deploy updated version
   - Monitor closely

## Security Best Practices

1. **Authentication**:
   - Use Azure AD SSO
   - Implement token refresh
   - Never store secrets client-side
   - Use secure token storage

2. **Data Protection**:
   - Encrypt data in transit (HTTPS)
   - Minimize data collection
   - Follow GDPR guidelines
   - Regular security audits

3. **Access Control**:
   - Principle of least privilege
   - Regular permission reviews
   - Monitor access logs
   - Implement MFA where possible

4. **Code Security**:
   - Regular dependency updates
   - Security scanning (CodeQL)
   - Input validation
   - Output encoding

## Compliance

### GDPR Compliance

- Document data processing
- Implement data retention policies
- Provide data export functionality
- Honor deletion requests
- Maintain privacy policy

### Microsoft 365 Compliance

- Follow Office Store policies
- Implement required disclosures
- Handle PII appropriately
- Regular compliance reviews

## Troubleshooting Deployment Issues

### Issue: Manifest validation fails

**Solution:**
- Run `npm run validate` locally
- Check XML syntax
- Verify all URLs are accessible
- Ensure HTTPS is used
- Check icon sizes and formats

### Issue: Add-in doesn't appear after deployment

**Solution:**
- Wait 24 hours for propagation
- Check deployment assignment
- Clear Office cache
- Restart Outlook
- Verify user has correct license

### Issue: Authentication errors in production

**Solution:**
- Verify redirect URIs in Azure AD
- Check client ID matches
- Ensure permissions are granted
- Review consent requirements
- Check token expiration handling

### Issue: SharePoint integration fails

**Solution:**
- Verify site URL is correct
- Check user permissions
- Confirm library exists
- Review API permissions
- Test Graph API access directly

## Support Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Microsoft Graph Documentation](https://docs.microsoft.com/graph/)
- [Azure AD Documentation](https://docs.microsoft.com/azure/active-directory/)
- [Microsoft 365 Admin Help](https://docs.microsoft.com/microsoft-365/admin/)

## Contact

For deployment assistance:
- Technical Support: tech@onecoachfitness.com
- Emergency Contact: +1-XXX-XXX-XXXX
- Documentation: https://docs.onecoachfitness.com
