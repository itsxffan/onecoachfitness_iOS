# OneCoach Email Manager - Outlook Add-in

A modern Office JavaScript add-in for smarter email management with SharePoint integration and Azure AD authentication.

## Features

- **Email Management**: Tag, categorize, and archive messages directly from the reading pane
- **SharePoint Integration**: Push email content and attachments to SharePoint document libraries
- **Azure AD Authentication**: Seamless single sign-on for Microsoft 365 tenants
- **Cross-Platform**: Works on Outlook Desktop, Web, and Mac
- **Fluent UI**: Modern interface aligned with Microsoft's design standards

## Prerequisites

- Node.js 16.x or higher
- npm or yarn package manager
- Microsoft 365 account with Outlook
- Azure AD tenant (for authentication setup)
- SharePoint site (for document library integration)

## Installation

### 1. Install Dependencies

```bash
cd outlook-addin
npm install
```

### 2. Azure AD App Registration

Before you can use the add-in, you need to register an application in Azure AD:

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Configure the application:
   - **Name**: OneCoach Email Manager
   - **Supported account types**: Accounts in any organizational directory (multitenant)
   - **Redirect URI**: 
     - Type: Web
     - URI: `https://localhost:3000/auth-callback.html`
4. Click **Register**
5. Note the **Application (client) ID** - you'll need this later
6. Under **Certificates & secrets**, create a new client secret and save it securely

#### Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Add the following **Delegated permissions**:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
5. Click **Grant admin consent** for your organization

#### Configure Authentication

1. Go to **Authentication** in your app registration
2. Under **Implicit grant and hybrid flows**, enable:
   - Access tokens
   - ID tokens
3. Add additional redirect URIs as needed for production deployment
4. Save changes

### 3. Configure the Add-in

Update the Azure AD configuration in `src/services/authService.ts`:

```typescript
// Replace these values with your Azure AD app registration details
const clientId = 'YOUR_AZURE_AD_CLIENT_ID';
const tenantId = 'common'; // or your specific tenant ID
```

Update the SharePoint site URL in `src/services/sharepointService.ts`:

```typescript
this.siteUrl = 'https://yourtenant.sharepoint.com/sites/yoursite';
```

### 4. Update Manifest URLs

For production deployment, update all URLs in `manifest.xml` from `https://localhost:3000` to your production hosting URL.

### 5. Generate Development Certificates

For local development with HTTPS:

```bash
npx office-addin-dev-certs install
```

## Development

### Start Development Server

```bash
npm run dev-server
```

This will start a webpack dev server at `https://localhost:3000` with hot reload enabled.

### Build for Production

```bash
npm run build
```

The production-ready files will be in the `dist/` directory.

### Validate Manifest

```bash
npm run validate
```

### Sideload Add-in for Testing

#### Outlook Desktop (Windows)

1. Run: `npm start`
2. This will automatically sideload the add-in in Outlook Desktop
3. To stop: `npm stop`

#### Outlook Web

1. Go to Outlook on the web
2. Open any email message
3. Click the **More actions** (...) menu
4. Select **Get Add-ins**
5. Click **My add-ins** tab
6. Under **Custom add-ins**, click **+ Add a custom add-in**
7. Select **Add from file**
8. Upload the `manifest.xml` file
9. Click **Install**

#### Outlook Mac

1. Open Outlook for Mac
2. Click **Get Add-ins** button in the ribbon
3. Click **My add-ins** tab
4. Under **Custom add-ins**, click **+ Add a custom add-in**
5. Select **Add from file**
6. Browse and select the `manifest.xml` file
7. Click **Install**

## Deployment

### Option 1: Tenant-Wide Deployment (Centralized)

For deploying to all users in your Microsoft 365 tenant:

1. Go to the [Microsoft 365 admin center](https://admin.microsoft.com)
2. Navigate to **Settings** > **Integrated apps**
3. Click **Upload custom apps**
4. Select **Upload manifest file (.xml)**
5. Choose the `manifest.xml` file
6. Follow the deployment wizard:
   - Assign to users/groups or make available to entire organization
   - Configure deployment settings
7. Click **Deploy**

The add-in will be automatically available to assigned users within 24 hours.

### Option 2: Individual Deployment

Users can install the add-in themselves:

1. Follow the **Sideload Add-in for Testing** instructions above
2. The add-in will be available only to that specific user
3. User must re-install if manifest is updated

### Option 3: SharePoint App Catalog

1. Go to your SharePoint admin center
2. Navigate to **More features** > **Apps** > **Open**
3. Click **App Catalog**
4. Upload the `manifest.xml` to the App Catalog
5. Check **Make this solution available to all sites**
6. Users can then add the app from the App Catalog

### Production Hosting

For production, you need to host the add-in files on a publicly accessible HTTPS server:

1. Build the production files: `npm run build`
2. Upload the contents of `dist/` to your web server
3. Update all URLs in `manifest.xml` to point to your production URL
4. Ensure your server supports HTTPS with a valid SSL certificate
5. Deploy the updated manifest using one of the methods above

**Recommended hosting options:**
- Azure Static Web Apps
- Azure App Service
- GitHub Pages (with custom domain for HTTPS)
- Any CDN with HTTPS support

## Usage

### Opening the Add-in

1. Open Outlook (Desktop, Web, or Mac)
2. Select any email message
3. Look for the **OneCoach Email** tab in the ribbon
4. Click **Manage Email** to open the task pane

### Signing In

1. When you first open the add-in, click **Sign In with Microsoft**
2. Authenticate with your Microsoft 365 credentials
3. Grant the requested permissions
4. You'll be redirected back to the add-in

### Managing Emails

#### Categorize Messages

- Click any of the predefined category buttons (Important, Follow Up, Client, Training)
- Categories are applied to the message immediately
- Active categories are highlighted

#### Add Custom Tags

1. Type a tag name in the tag input field
2. Click **Add Tag**
3. The tag appears below the input
4. Click the × to remove a tag

#### Archive Messages

- Click **Archive Message** to move the message to your archive
- A confirmation message will appear

#### Push to SharePoint

1. Select a document library from the dropdown
2. Check/uncheck **Include attachments** as needed
3. Click **Push to SharePoint**
4. The email content and attachments will be uploaded to the selected library

### Quick Archive from Ribbon

You can also archive messages directly from the ribbon without opening the task pane:

1. Select an email message
2. Click the **Quick Archive** button in the OneCoach Email ribbon tab

## Architecture

### Technology Stack

- **Frontend**: TypeScript, HTML5, CSS3
- **UI Framework**: Fluent UI principles
- **Office API**: Office.js 1.13+
- **Authentication**: MSAL (Microsoft Authentication Library)
- **API Integration**: Microsoft Graph API
- **Build Tool**: Webpack 5
- **Package Manager**: npm

### Project Structure

```
outlook-addin/
├── src/
│   ├── taskpane/          # Main task pane UI
│   │   ├── taskpane.html
│   │   ├── taskpane.ts
│   │   └── taskpane.css
│   ├── commands/          # Ribbon commands
│   │   ├── commands.html
│   │   └── commands.ts
│   └── services/          # Business logic
│       ├── authService.ts       # Azure AD authentication
│       ├── emailService.ts      # Email operations
│       └── sharepointService.ts # SharePoint integration
├── assets/                # Icons and images
├── dist/                  # Build output
├── manifest.xml           # Office Add-in manifest
├── package.json
├── tsconfig.json
└── webpack.config.js
```

### Services

#### AuthService
Handles Azure AD authentication and token management:
- SSO authentication flow
- Token storage and retrieval
- Microsoft Graph client initialization
- Fallback to dialog-based authentication

#### EmailService
Manages email operations:
- Message categorization
- Custom tagging
- Archive functionality
- Content extraction

#### SharePointService
Integrates with SharePoint:
- Document library access
- File upload
- Attachment handling
- Drive ID resolution

## Troubleshooting

### Add-in Doesn't Load

1. Check that HTTPS is enabled on your server
2. Verify the manifest URLs are correct
3. Clear Office cache:
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: `~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef`
4. Restart Outlook

### Authentication Fails

1. Verify Azure AD app registration is correct
2. Check that API permissions are granted
3. Ensure redirect URIs match exactly
4. Clear browser cookies and cache
5. Try signing out and back in to Microsoft 365

### SharePoint Upload Fails

1. Check that the user has write permissions to the SharePoint library
2. Verify the SharePoint site URL is correct
3. Ensure the document library name matches exactly
4. Check network connectivity to SharePoint

### CORS Errors

1. Verify your server has proper CORS headers configured
2. Check that Azure AD app allows your domain
3. Ensure redirect URIs include your domain

## Security Considerations

- Never commit secrets or client IDs to source control
- Use environment variables for sensitive configuration
- Implement proper token refresh logic
- Follow Microsoft's security best practices for Office Add-ins
- Regular security audits and dependency updates
- Validate and sanitize all user inputs
- Use HTTPS everywhere in production

## Performance Optimization

- Lazy load services only when needed
- Cache API responses appropriately
- Minimize bundle size with tree shaking
- Use async/await for non-blocking operations
- Implement proper error handling and retry logic

## Browser Compatibility

- Microsoft Edge (Chromium)
- Google Chrome (latest)
- Safari (latest)
- Firefox (latest)

Note: Internet Explorer 11 is not supported.

## Support

For issues, questions, or feature requests:
- Email: support@onecoachfitness.com
- Documentation: https://docs.onecoachfitness.com/outlook-addin

## License

MIT License - see LICENSE file for details

## Credits

Developed by OneCoach Fitness team
Powered by Microsoft Office JavaScript APIs and Microsoft Graph
