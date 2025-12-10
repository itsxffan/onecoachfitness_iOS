# Troubleshooting Guide

Common issues and solutions for the OneCoach Email Manager Outlook add-in.

## Table of Contents

1. [Installation Issues](#installation-issues)
2. [Authentication Problems](#authentication-problems)
3. [Add-in Not Loading](#add-in-not-loading)
4. [SharePoint Integration Issues](#sharepoint-integration-issues)
5. [Performance Problems](#performance-problems)
6. [Development Issues](#development-issues)
7. [Platform-Specific Issues](#platform-specific-issues)
8. [Error Messages](#error-messages)

---

## Installation Issues

### Issue: npm install fails

**Symptoms:**
- Error messages during `npm install`
- Dependencies not installing correctly

**Solutions:**

1. **Clear npm cache:**
```bash
npm cache clean --force
rm -rf node_modules
rm package-lock.json
npm install
```

2. **Check Node.js version:**
```bash
node --version  # Should be 16.x or higher
npm --version   # Should be 8.x or higher
```

3. **Use different registry (if corporate firewall):**
```bash
npm config set registry https://registry.npmjs.org/
npm install
```

4. **Try with --legacy-peer-deps:**
```bash
npm install --legacy-peer-deps
```

### Issue: Manifest validation fails

**Symptoms:**
- `npm run validate` shows errors
- Manifest upload rejected in Outlook

**Solutions:**

1. **Check XML syntax:**
```bash
# Use an XML validator
xmllint --noout manifest.xml
```

2. **Verify all URLs are accessible:**
- Check that all URLs in manifest use HTTPS
- Verify URLs are publicly accessible
- Test URLs in browser

3. **Check ID uniqueness:**
```xml
<!-- Ensure this is a unique GUID -->
<Id>your-unique-guid-here</Id>
```

4. **Validate against schema:**
```bash
npm run validate
# Fix any errors shown
```

---

## Authentication Problems

### Issue: Sign-in button doesn't work

**Symptoms:**
- Clicking "Sign In with Microsoft" does nothing
- No authentication dialog appears

**Solutions:**

1. **Check popup blockers:**
- Disable popup blockers for your domain
- Allow popups in browser settings
- Check Office add-in settings

2. **Verify Azure AD configuration:**
```typescript
// Check src/services/authService.ts
const clientId = 'YOUR_CLIENT_ID'; // Should be valid GUID
const tenantId = 'common'; // or your tenant ID
```

3. **Check redirect URI:**
- Must match exactly in Azure AD
- Must use HTTPS
- Should be accessible

4. **Clear browser cache:**
- Clear cookies and cache
- Try incognito/private mode
- Restart browser

### Issue: "Permission denied" error

**Symptoms:**
- Authentication succeeds but operations fail
- "Access denied" or "Permission denied" errors

**Solutions:**

1. **Grant admin consent:**
- Go to Azure Portal
- Navigate to your app registration
- API permissions > Grant admin consent

2. **Check required permissions:**
```
Required permissions:
- User.Read
- Mail.ReadWrite
- Sites.ReadWrite.All
- Files.ReadWrite.All
```

3. **Verify permission type:**
- Should be "Delegated permissions"
- Not "Application permissions"

4. **Wait for propagation:**
- Permissions can take up to 30 minutes to propagate
- Try again after waiting

### Issue: Token expired error

**Symptoms:**
- Add-in works initially but stops after some time
- "Token expired" or "Invalid token" errors

**Solutions:**

1. **Sign out and back in:**
- Click sign out (if available)
- Clear browser cache
- Sign in again

2. **Check token refresh logic:**
- Should auto-refresh in background
- Check console for errors

3. **Implement token refresh:**
```typescript
// In authService.ts
// Token refresh should happen automatically
// If not, sign out and sign back in
```

---

## Add-in Not Loading

### Issue: Add-in doesn't appear in Outlook

**Symptoms:**
- Ribbon tab missing
- Add-in not in list
- No way to open add-in

**Solutions:**

1. **Verify installation:**
- Check "Get Add-ins" > "My add-ins"
- Look for "OneCoach Email Manager"
- Reinstall if missing

2. **Clear Office cache:**

**Windows:**
```
1. Close all Office apps
2. Delete folder: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
3. Restart Outlook
```

**Mac:**
```
1. Close Outlook
2. Delete: ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
3. Restart Outlook
```

3. **Check manifest deployment:**
- For centralized: Wait 24 hours for propagation
- For sideload: Try removing and re-adding
- Verify manifest.xml is valid

4. **Restart Outlook:**
- Completely quit Outlook (File > Exit)
- Wait 30 seconds
- Restart Outlook

### Issue: Task pane blank or shows error

**Symptoms:**
- Task pane opens but shows nothing
- White screen
- Error message in pane

**Solutions:**

1. **Check browser console:**
- Press F12 to open dev tools
- Look for JavaScript errors
- Check network tab for failed requests

2. **Verify HTTPS:**
```bash
# Must use HTTPS, not HTTP
https://localhost:3000  ✓
http://localhost:3000   ✗
```

3. **Check dev server status:**
```bash
# Ensure dev server is running
npm run dev-server

# Should see: "webpack compiled successfully"
```

4. **Trust SSL certificate:**
```bash
npx office-addin-dev-certs install
# Follow prompts to trust certificate
```

---

## SharePoint Integration Issues

### Issue: "Failed to upload to SharePoint"

**Symptoms:**
- Push to SharePoint button fails
- Error message shown
- Files don't appear in SharePoint

**Solutions:**

1. **Verify SharePoint URL:**
```typescript
// In src/services/sharepointService.ts
this.siteUrl = 'https://yourtenant.sharepoint.com/sites/yoursite';
// Must be exact, no trailing slash
```

2. **Check user permissions:**
- User must have "Contribute" or "Edit" permissions
- Verify in SharePoint site permissions
- Contact SharePoint admin if needed

3. **Verify library exists:**
- Library name must match exactly
- Case sensitive
- Check dropdown options match actual libraries

4. **Check file name length:**
```typescript
// File names must be < 128 characters
// Special characters not allowed: ~ # % & * { } \ : < > ? / | "
```

5. **Test Graph API access:**
```bash
# Use Graph Explorer to test
# https://developer.microsoft.com/graph/graph-explorer
GET /sites/{site-id}/drives
```

### Issue: Attachments not uploading

**Symptoms:**
- Email content uploads but attachments don't
- Partial uploads
- Attachment errors

**Solutions:**

1. **Check attachment size:**
```
Maximum size: 150MB (Graph API limit)
If larger, split into chunks or skip large files
```

2. **Check file types:**
- All file types should be supported
- Some may be blocked by SharePoint admin
- Check SharePoint blocked file types list

3. **Verify "Include attachments" checkbox:**
- Must be checked to upload attachments
- Located in SharePoint Integration section

4. **Check network connection:**
- Large attachments take time
- Don't navigate away during upload
- Check network speed/stability

---

## Performance Problems

### Issue: Add-in is slow

**Symptoms:**
- Operations take long time
- UI feels sluggish
- High memory usage

**Solutions:**

1. **Clear Office cache:**
- See "Add-in Not Loading" section above
- Clears old cached data

2. **Check memory usage:**
```
Task Manager (Windows) or Activity Monitor (Mac)
- Outlook should be < 1GB
- Add-in should be < 50MB
```

3. **Reduce concurrent operations:**
- Don't operate on too many emails at once
- Process in batches of 10-20

4. **Check network speed:**
- SharePoint operations require good connection
- Test at different times/locations

5. **Optimize build:**
```bash
# Use production build
npm run build

# Should be smaller and faster than dev build
```

### Issue: High CPU usage

**Symptoms:**
- Computer fan spinning
- Outlook slow
- Other apps affected

**Solutions:**

1. **Disable hot reload (dev mode):**
```bash
# Use production mode
npm run build
# Then use the dist/ files
```

2. **Limit console logging:**
```typescript
// In src/config/config.ts
logging: {
  enableConsoleLogging: false, // Disable for production
}
```

3. **Close unused Outlook items:**
- Close old emails
- Minimize open folders
- Clear deleted items

---

## Development Issues

### Issue: Hot reload not working

**Symptoms:**
- Changes not reflected
- Must restart server manually
- Stale code running

**Solutions:**

1. **Check webpack dev server:**
```bash
# Look for "webpack compiled successfully"
# If errors, fix them first
```

2. **Hard refresh browser:**
```
Windows/Linux: Ctrl + Shift + R
Mac: Cmd + Shift + R
```

3. **Clear browser cache:**
- Settings > Privacy > Clear browsing data
- Select "Cached images and files"

4. **Restart dev server:**
```bash
# Stop server (Ctrl+C)
# Start again
npm run dev-server
```

### Issue: TypeScript errors

**Symptoms:**
- Red squiggly lines in editor
- Compilation fails
- Type errors

**Solutions:**

1. **Check TypeScript version:**
```bash
npx tsc --version  # Should be 5.3+
```

2. **Rebuild:**
```bash
rm -rf dist node_modules
npm install
npm run build
```

3. **Check tsconfig.json:**
```json
{
  "compilerOptions": {
    "strict": true,  // For strict type checking
    "skipLibCheck": true  // Skip checking .d.ts files
  }
}
```

4. **Install type definitions:**
```bash
npm install --save-dev @types/office-js
```

### Issue: ESLint errors

**Symptoms:**
- Linting fails
- Can't commit code
- Red errors in editor

**Solutions:**

1. **Auto-fix:**
```bash
npm run lint:fix
```

2. **Check configuration:**
```bash
# Verify .eslintrc.json exists
cat .eslintrc.json
```

3. **Disable rule temporarily:**
```typescript
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const data: any = response;
```

4. **Update ESLint:**
```bash
npm update eslint @typescript-eslint/eslint-plugin
```

---

## Platform-Specific Issues

### Windows-Specific

**Issue: Outlook crashes when loading add-in**

Solutions:
1. Update Outlook to latest version
2. Repair Office installation
3. Disable hardware acceleration in Outlook
4. Check Windows Event Viewer for errors

**Issue: Can't install certificate**

Solutions:
1. Run PowerShell as administrator
2. Execute: `npx office-addin-dev-certs install`
3. If fails, manually import certificate from `.office-addin-dev-certs` folder

### Mac-Specific

**Issue: Add-in doesn't load in Outlook Mac**

Solutions:
1. Update to latest Outlook for Mac
2. Check macOS version (10.14+)
3. Clear cache: `~/Library/Containers/com.microsoft.Outlook/`
4. Reinstall Outlook if persistent

**Issue: Certificate trust issues**

Solutions:
1. Open Keychain Access
2. Find certificate
3. Double click > Trust > Always Trust
4. Restart Outlook

### Outlook Web-Specific

**Issue: Add-in works in Desktop but not Web**

Solutions:
1. Check browser compatibility (Chrome, Edge, Safari, Firefox latest)
2. Disable browser extensions temporarily
3. Try different browser
4. Check console for CORS errors

**Issue: Popups blocked**

Solutions:
1. Allow popups for outlook.office.com
2. Check browser popup settings
3. Use authentication fallback method

---

## Error Messages

### "Network request failed"

**Cause:** Can't reach server/API

**Solutions:**
1. Check internet connection
2. Verify server is running (for dev)
3. Check firewall/proxy settings
4. Test URL in browser

### "Invalid client"

**Cause:** Azure AD configuration issue

**Solutions:**
1. Verify client ID is correct
2. Check redirect URI matches
3. Ensure app is not deleted/expired
4. Recreate app registration if needed

### "Forbidden" or "403"

**Cause:** Permission denied

**Solutions:**
1. Check API permissions granted
2. Verify user has necessary SharePoint permissions
3. Grant admin consent
4. Check conditional access policies

### "Not found" or "404"

**Cause:** Resource doesn't exist

**Solutions:**
1. Check SharePoint site URL
2. Verify library name
3. Check file path
4. Ensure site/library not deleted

### "CORS error"

**Cause:** Cross-origin request blocked

**Solutions:**
1. Ensure using HTTPS
2. Check Azure AD app allows domain
3. Verify redirect URIs configured
4. Check server CORS headers

---

## Getting More Help

### Self-Service Resources
1. Check browser console (F12) for detailed errors
2. Review logs in Application Insights (if configured)
3. Search GitHub Issues
4. Review documentation

### Contact Support
- **Email:** support@onecoachfitness.com
- **Include:**
  - Detailed error description
  - Steps to reproduce
  - Screenshots/error messages
  - Platform (Windows/Mac/Web)
  - Outlook version
  - Browser (for Web)

### Diagnostic Information

When reporting issues, include:

```
Platform: [Windows/Mac/Web]
Outlook Version: [e.g., 16.0.15225.20288]
Browser (Web): [e.g., Chrome 120]
Add-in Version: 1.0.0
Node.js Version: [from npm --version]
Error Message: [exact error text]
Steps to Reproduce: [numbered list]
Expected Behavior: [what should happen]
Actual Behavior: [what actually happens]
```

---

**Still having issues?** Don't hesitate to contact support@onecoachfitness.com
