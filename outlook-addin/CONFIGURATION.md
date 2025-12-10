# Configuration Guide

This guide explains all configuration options for the OneCoach Email Manager add-in.

## Required Configuration

### 1. Manifest Configuration

**File**: `manifest.xml`

#### Generate Unique ID

⚠️ **IMPORTANT**: Replace the placeholder GUID with a unique identifier before deployment.

```xml
<!-- Line 9 in manifest.xml -->
<Id>a1b2c3d4-e5f6-4a5b-8c9d-0e1f2a3b4c5d</Id>
```

**How to generate a new GUID:**

**Using PowerShell (Windows):**
```powershell
[guid]::NewGuid().ToString()
```

**Using Terminal (Mac/Linux):**
```bash
uuidgen
```

**Using Node.js:**
```bash
node -e "console.log(require('crypto').randomUUID())"
```

**Using Online Tool:**
- Visit: https://www.uuidgenerator.net/
- Copy the generated GUID
- Replace in manifest.xml

#### Update URLs for Production

Replace all `localhost:3000` URLs with your production domain:

```xml
<!-- Icons -->
<IconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://YOUR-DOMAIN.com/assets/icon-64.png"/>

<!-- Source locations -->
<SourceLocation DefaultValue="https://YOUR-DOMAIN.com/taskpane.html"/>

<!-- In VersionOverrides section -->
<bt:Image id="Icon.16x16" DefaultValue="https://YOUR-DOMAIN.com/assets/icon-16.png"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://YOUR-DOMAIN.com/taskpane.html"/>
```

### 2. Environment Variables

**File**: `.env` (create from `.env.example`)

```bash
# Copy example file
cp .env.example .env

# Edit with your values
nano .env  # or use your preferred editor
```

**Required variables:**

```env
# Azure AD Configuration
AZURE_AD_CLIENT_ID=your-client-id-here
AZURE_AD_TENANT_ID=common
# Optional: AZURE_AD_CLIENT_SECRET=your-secret-here

# SharePoint Configuration
SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/yoursite

# Application Configuration
APP_URL=https://your-production-domain.com
NODE_ENV=production
```

### 3. Azure AD App Registration

Before configuring environment variables, you need to create an Azure AD app registration.

**Step-by-step:**

1. Go to https://portal.azure.com
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in:
   - Name: `OneCoach Email Manager`
   - Account types: `Accounts in any organizational directory (multitenant)`
   - Redirect URI: 
     - Platform: `Web`
     - URI: `https://your-domain.com/auth-callback.html`
5. Click **Register**
6. Copy the **Application (client) ID** → Use for `AZURE_AD_CLIENT_ID`
7. Copy the **Directory (tenant) ID** → Use for `AZURE_AD_TENANT_ID` (or use "common" for multitenant)

**Add API Permissions:**

1. Click **API permissions** in left menu
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add these permissions:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
6. Click **Add permissions**
7. Click **Grant admin consent for [Your Organization]**

**Configure Authentication:**

1. Click **Authentication** in left menu
2. Under **Implicit grant and hybrid flows**, enable:
   - ✅ Access tokens (used for implicit flows)
   - ✅ ID tokens (used for implicit and hybrid flows)
3. Add additional redirect URIs for different environments:
   - Development: `https://localhost:3000/auth-callback.html`
   - Staging: `https://staging.your-domain.com/auth-callback.html`
   - Production: `https://your-domain.com/auth-callback.html`
4. Click **Save**

### 4. SharePoint Site Configuration

You need a SharePoint site with document libraries for storing email content.

**Setup:**

1. Go to your SharePoint site
2. Create document libraries (or use existing):
   - **Client Documents** - For client-related emails
   - **Training Materials** - For training content
   - **General Archive** - For general emails
3. Grant permissions:
   - Users need at least **Contribute** permission to upload files
   - Check under: Site Settings > Site Permissions
4. Copy the site URL:
   - Format: `https://yourtenant.sharepoint.com/sites/yoursite`
   - Use for `SHAREPOINT_SITE_URL`

## Optional Configuration

### 5. Application Configuration

**File**: `src/config/config.ts`

#### Categories Configuration

Customize email categories:

```typescript
categories: [
  { id: 'Important', name: 'Important', color: 'red' },
  { id: 'Follow Up', name: 'Follow Up', color: 'yellow' },
  { id: 'Client', name: 'Client', color: 'blue' },
  { id: 'Training', name: 'Training', color: 'green' },
  // Add your custom categories here
  { id: 'Project', name: 'Project', color: 'purple' },
  { id: 'Urgent', name: 'Urgent', color: 'orange' }
]
```

#### SharePoint Libraries

Customize available document libraries:

```typescript
sharepoint: {
  siteUrl: process.env.SHAREPOINT_SITE_URL || '',
  defaultLibraries: [
    { id: 'ClientDocuments', name: 'Client Documents' },
    { id: 'TrainingMaterials', name: 'Training Materials' },
    { id: 'GeneralArchive', name: 'General Archive' },
    // Add your custom libraries here
    { id: 'Projects', name: 'Project Files' },
    { id: 'HR', name: 'HR Documents' }
  ]
}
```

#### Feature Flags

Enable/disable features:

```typescript
features: {
  enableSharePointIntegration: true,  // Enable/disable SharePoint
  enableCustomTags: true,             // Enable/disable custom tags
  enableArchive: true,                // Enable/disable archive
  enableCategorySync: true            // Enable/disable category sync
}
```

#### Logging Configuration

Configure logging behavior:

```typescript
logging: {
  enableConsoleLogging: true,  // Console logs (disable for production)
  enableErrorTracking: true,   // Error tracking
  logLevel: 'info'             // 'debug', 'info', 'warn', 'error'
}
```

### 6. Application Insights (Optional)

For monitoring and analytics:

1. Create Application Insights resource in Azure
2. Get the Instrumentation Key
3. Add to `.env`:
   ```env
   APPINSIGHTS_INSTRUMENTATIONKEY=your-key-here
   ```

### 7. Custom Backend API (Optional)

If using Azure Functions or custom API:

1. Deploy your backend API
2. Add to `.env`:
   ```env
   API_BASE_URL=https://your-api.azurewebsites.net/api
   ```
3. Configure CORS in your API to allow your frontend domain

## Build Configuration

### Webpack Configuration

**File**: `webpack.config.js`

#### CORS Configuration (Development)

For development, specify allowed origins:

```javascript
headers: {
  'Access-Control-Allow-Origin': 'https://outlook.office.com, https://outlook.live.com'
}
```

For production, this is handled by your hosting service.

#### Port Configuration

Change the development server port:

```javascript
devServer: {
  port: 3000,  // Change to your preferred port
  // ...
}
```

### TypeScript Configuration

**File**: `tsconfig.json`

Default configuration is optimized for Office Add-ins. Only modify if you know what you're doing.

**Common changes:**

```json
{
  "compilerOptions": {
    "strict": true,        // Strict type checking (recommended)
    "target": "ES2020",    // JavaScript version target
    "lib": ["ES2020", "DOM"]  // Available APIs
  }
}
```

### ESLint Configuration

**File**: `.eslintrc.json`

Customize linting rules:

```json
{
  "rules": {
    "@typescript-eslint/no-explicit-any": "warn",  // Change to "error" for stricter checking
    "no-console": "off",  // Allow console.log (change to "warn" for production)
    // Add your custom rules
  }
}
```

## Production Checklist

Before deploying to production:

- [ ] Generate unique GUID for manifest.xml
- [ ] Update all URLs from localhost to production domain
- [ ] Configure Azure AD app registration
- [ ] Set environment variables (.env file)
- [ ] Configure SharePoint site and libraries
- [ ] Test Azure AD authentication
- [ ] Verify SharePoint connectivity
- [ ] Create and upload icons to assets/
- [ ] Build production assets (`npm run build`)
- [ ] Validate manifest (`npm run validate`)
- [ ] Test in all Outlook platforms
- [ ] Review and adjust logging levels
- [ ] Set up monitoring (Application Insights)
- [ ] Configure backup procedures
- [ ] Document your specific configuration

## Environment-Specific Configuration

### Development Environment

```env
NODE_ENV=development
APP_URL=https://localhost:3000
AZURE_AD_CLIENT_ID=dev-client-id
SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/dev
```

### Staging Environment

```env
NODE_ENV=staging
APP_URL=https://staging.your-domain.com
AZURE_AD_CLIENT_ID=staging-client-id
SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/staging
```

### Production Environment

```env
NODE_ENV=production
APP_URL=https://your-domain.com
AZURE_AD_CLIENT_ID=prod-client-id
SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/production
```

## Security Best Practices

1. **Never commit secrets to source control**
   - Use `.env` file (in .gitignore)
   - Use Azure Key Vault for production secrets
   - Rotate secrets regularly

2. **Use separate Azure AD apps for each environment**
   - Development app
   - Staging app
   - Production app

3. **Restrict redirect URIs**
   - Only add trusted domains
   - Use HTTPS only
   - Match exactly (including trailing slashes)

4. **Minimum permissions**
   - Only grant required API permissions
   - Review regularly
   - Remove unused permissions

5. **Monitor and audit**
   - Enable Azure AD sign-in logs
   - Set up alerts for suspicious activity
   - Review access regularly

## Troubleshooting Configuration Issues

### Issue: "Client ID not configured"

**Solution**: Set `AZURE_AD_CLIENT_ID` in `.env` file

### Issue: "SharePoint site URL not configured"

**Solution**: Set `SHAREPOINT_SITE_URL` in `.env` file

### Issue: Authentication fails

**Solution**: 
- Verify client ID is correct
- Check redirect URIs match exactly
- Ensure permissions are granted
- Check tenant ID (use "common" for multitenant)

### Issue: SharePoint upload fails

**Solution**:
- Verify site URL is correct (no trailing slash)
- Check user has permissions to the site
- Ensure library names match exactly
- Verify SharePoint is accessible

## Getting Help

If you need help with configuration:

- **Documentation**: See README.md, AZURE-SETUP.md
- **Support**: support@onecoachfitness.com
- **GitHub Issues**: Report configuration problems

---

**Note**: Always test configuration in a development/staging environment before deploying to production.
