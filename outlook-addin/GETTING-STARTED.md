# Getting Started with OneCoach Email Manager

This quick start guide will get you up and running with the OneCoach Email Manager Outlook add-in in minutes.

## 🚀 Quick Start (5 minutes)

### Step 1: Prerequisites

Make sure you have:
- ✅ Node.js 16.x or higher installed
- ✅ npm or yarn package manager
- ✅ A Microsoft 365 account
- ✅ Access to Azure Portal (for Azure AD setup)

### Step 2: Clone and Install

```bash
# Navigate to the add-in directory
cd outlook-addin

# Install dependencies
npm install
```

This will install all required packages including Office.js, Microsoft Graph Client, Fluent UI components, and development tools.

### Step 3: Azure AD Setup

Before you can run the add-in, you need an Azure AD app registration:

**Quick setup (5 minutes):**

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Fill in:
   - Name: `OneCoach Email Manager`
   - Supported account types: `Multitenant`
   - Redirect URI: `Web` - `https://localhost:3000/auth-callback.html`
4. Click **Register**
5. Copy the **Application (client) ID**

**Add API permissions:**

1. Go to **API permissions** > **Add a permission** > **Microsoft Graph**
2. Select **Delegated permissions**
3. Add these permissions:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
4. Click **Grant admin consent**

**Enable authentication:**

1. Go to **Authentication**
2. Under **Implicit grant**, enable:
   - ✅ Access tokens
   - ✅ ID tokens
3. Click **Save**

### Step 4: Configure the Add-in

Update the client ID in `src/services/authService.ts`:

```typescript
const clientId = 'YOUR_CLIENT_ID_HERE'; // Replace with your actual client ID
```

Update SharePoint URL in `src/services/sharepointService.ts`:

```typescript
this.siteUrl = 'https://yourtenant.sharepoint.com/sites/yoursite';
```

### Step 5: Generate SSL Certificate (for HTTPS)

```bash
# Generate self-signed certificate for local development
npx office-addin-dev-certs install
```

When prompted, trust the certificate.

### Step 6: Start Development Server

```bash
# Start the webpack dev server
npm run dev-server
```

The add-in will be available at `https://localhost:3000`

### Step 7: Sideload in Outlook

**For Outlook Web:**

1. Go to https://outlook.office.com
2. Open any email
3. Click **...** (More actions) > **Get Add-ins**
4. Click **My add-ins** > **+ Add a custom add-in** > **Add from file**
5. Upload `manifest.xml` from the `outlook-addin` folder
6. Click **Install**

**For Outlook Desktop (Windows):**

```bash
# This will automatically sideload the add-in
npm start
```

**For Outlook Mac:**

1. Open Outlook
2. Click **Get Add-ins** in ribbon
3. Go to **My add-ins** > **+ Add a custom add-in** > **Add from file**
4. Browse and select `manifest.xml`
5. Click **Install**

### Step 8: Test the Add-in

1. Open any email in Outlook
2. Look for the **OneCoach Email** tab in the ribbon
3. Click **Manage Email**
4. The task pane opens on the right
5. Click **Sign In with Microsoft**
6. Grant permissions when prompted
7. Start managing your emails! 🎉

## 🔧 Development Workflow

### Making Changes

1. Edit files in `src/` directory
2. Webpack will automatically rebuild (in dev mode)
3. Refresh the add-in in Outlook to see changes
4. Check browser console for errors (F12)

### Build for Production

```bash
# Create optimized production build
npm run build

# Output will be in dist/ folder
```

### Validate Manifest

```bash
# Check manifest.xml for errors
npm run validate
```

### Lint Code

```bash
# Check code quality
npm run lint

# Auto-fix issues
npm run lint:fix
```

## 📁 Project Structure

```
outlook-addin/
├── src/                      # Source code
│   ├── taskpane/            # Task pane UI
│   │   ├── taskpane.html   # UI structure
│   │   ├── taskpane.ts     # Logic and event handlers
│   │   └── taskpane.css    # Fluent UI styling
│   ├── commands/            # Ribbon commands
│   ├── services/            # Business logic
│   │   ├── authService.ts  # Azure AD authentication
│   │   ├── emailService.ts # Email operations
│   │   └── sharepointService.ts # SharePoint integration
│   └── config/              # Configuration
├── assets/                  # Icons and images
├── dist/                    # Build output (generated)
├── manifest.xml             # Add-in manifest
├── package.json             # Dependencies
├── tsconfig.json            # TypeScript config
└── webpack.config.js        # Build config
```

## 🎯 Key Features to Try

### 1. Categorize an Email

1. Select an email
2. Open the add-in
3. Click any category button (Important, Follow Up, etc.)
4. The email is immediately categorized

### 2. Add Custom Tags

1. Type a tag name (e.g., "Project Alpha")
2. Click **Add Tag**
3. Tag appears below, linked to the email

### 3. Archive Message

1. Click **Archive Message** in the task pane
2. Or use **Quick Archive** button in ribbon
3. Message moves to archive folder

### 4. Push to SharePoint

1. Select a document library from dropdown
2. Check/uncheck "Include attachments"
3. Click **Push to SharePoint**
4. Email content and attachments are uploaded

## 🐛 Common Issues

### Issue: Add-in doesn't load

**Solution:**
```bash
# Clear Office cache
# Windows: Delete this folder
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\

# Mac: Delete this folder
~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef

# Restart Outlook
```

### Issue: SSL certificate error

**Solution:**
```bash
# Reinstall certificate
npx office-addin-dev-certs install --force

# Trust the certificate when prompted
```

### Issue: Authentication fails

**Solution:**
1. Verify client ID is correct
2. Check redirect URI matches exactly
3. Ensure permissions are granted
4. Clear browser cookies and retry

### Issue: CORS errors

**Solution:**
1. Make sure dev server is running on HTTPS
2. Check browser console for specific CORS error
3. Verify Azure AD app allows your domain

## 📚 Next Steps

Now that you have the add-in running:

1. **Customize Categories**: Edit `src/config/config.ts` to add your own categories
2. **Update Styling**: Modify `src/taskpane/taskpane.css` for custom branding
3. **Add Features**: Extend functionality in `src/services/`
4. **Read Documentation**: 
   - [README.md](./README.md) - Complete documentation
   - [USER-GUIDE.md](./USER-GUIDE.md) - End-user instructions
   - [DEPLOYMENT.md](./DEPLOYMENT.md) - Production deployment
   - [AZURE-SETUP.md](./AZURE-SETUP.md) - Detailed Azure configuration

## 🤝 Getting Help

### Documentation
- Full README: [README.md](./README.md)
- User Guide: [USER-GUIDE.md](./USER-GUIDE.md)
- Deployment: [DEPLOYMENT.md](./DEPLOYMENT.md)

### Support
- Email: support@onecoachfitness.com
- Issues: GitHub Issues tab
- Microsoft Docs: https://docs.microsoft.com/office/dev/add-ins/

### Useful Links
- Office Add-ins Documentation: https://docs.microsoft.com/office/dev/add-ins/
- Microsoft Graph: https://docs.microsoft.com/graph/
- Fluent UI: https://developer.microsoft.com/fluentui
- Azure AD: https://docs.microsoft.com/azure/active-directory/

## ✅ Checklist

Before moving to production:

- [ ] Azure AD app registered with correct settings
- [ ] All permissions granted and consented
- [ ] Client ID updated in authService.ts
- [ ] SharePoint URL configured in sharepointService.ts
- [ ] manifest.xml validated successfully
- [ ] Add-in tested in Outlook Web
- [ ] Add-in tested in Outlook Desktop
- [ ] Add-in tested in Outlook Mac
- [ ] Icons created and added to assets/
- [ ] Production hosting environment set up
- [ ] manifest.xml updated with production URLs
- [ ] SSL certificate configured for production
- [ ] Documentation reviewed and updated
- [ ] Code linted and cleaned
- [ ] Production build created and tested

## 🎉 Success!

You now have a working Outlook add-in! Start customizing it for your needs and deploy it to your organization.

Happy coding! 💻

---

**Need help?** Don't hesitate to reach out to support@onecoachfitness.com
