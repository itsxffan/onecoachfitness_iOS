#!/bin/bash

# OneCoach Email Manager - Setup Script
# This script helps with initial configuration

set -e

echo "================================================"
echo "OneCoach Email Manager - Setup"
echo "================================================"
echo ""

# Check Node.js version
echo "Checking Node.js version..."
NODE_VERSION=$(node -v | cut -d'v' -f2 | cut -d'.' -f1)
if [ "$NODE_VERSION" -lt 16 ]; then
    echo "❌ Error: Node.js 16.x or higher is required"
    echo "Current version: $(node -v)"
    exit 1
fi
echo "✅ Node.js version: $(node -v)"
echo ""

# Install dependencies
echo "Installing dependencies..."
if npm install; then
    echo "✅ Dependencies installed successfully"
else
    echo "❌ Failed to install dependencies"
    exit 1
fi
echo ""

# Check if .env exists
if [ ! -f .env ]; then
    echo "Creating .env file from template..."
    cp .env.example .env
    echo "✅ .env file created"
    echo ""
    echo "⚠️  IMPORTANT: You need to configure .env with your values:"
    echo "   - AZURE_AD_CLIENT_ID"
    echo "   - AZURE_AD_TENANT_ID"
    echo "   - SHAREPOINT_SITE_URL"
    echo ""
    echo "   Edit .env file and add your configuration."
    echo ""
else
    echo "✅ .env file already exists"
fi

# Generate a new GUID for manifest
echo ""
echo "================================================"
echo "Manifest Configuration"
echo "================================================"
echo ""
echo "Current manifest ID:"
CURRENT_ID=$(grep -A 1 "<Id>" manifest.xml | tail -n 1 | sed 's/.*<Id>\(.*\)<\/Id>/\1/' | tr -d ' ')
echo "  $CURRENT_ID"
echo ""

# Check if it's the default GUID
if [ "$CURRENT_ID" = "a1b2c3d4-e5f6-4a5b-8c9d-0e1f2a3b4c5d" ]; then
    echo "⚠️  Warning: You're using the default GUID"
    echo ""
    read -p "Would you like to generate a new unique GUID? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        # Generate new GUID
        if command -v uuidgen &> /dev/null; then
            NEW_GUID=$(uuidgen | tr '[:upper:]' '[:lower:]')
        else
            NEW_GUID=$(node -e "console.log(require('crypto').randomUUID())")
        fi
        
        echo "Generated new GUID: $NEW_GUID"
        
        # Update manifest.xml
        if [[ "$OSTYPE" == "darwin"* ]]; then
            # macOS
            sed -i '' "s/$CURRENT_ID/$NEW_GUID/" manifest.xml
        else
            # Linux
            sed -i "s/$CURRENT_ID/$NEW_GUID/" manifest.xml
        fi
        
        echo "✅ Manifest ID updated"
    else
        echo "⚠️  Remember to update the manifest ID before deploying to production"
    fi
else
    echo "✅ Manifest has a custom GUID"
fi

echo ""
echo "================================================"
echo "Development Certificates"
echo "================================================"
echo ""
echo "Generating development SSL certificates..."
if npx office-addin-dev-certs install; then
    echo "✅ SSL certificates installed"
else
    echo "⚠️  Warning: Failed to install SSL certificates"
    echo "   You may need to do this manually"
fi

echo ""
echo "================================================"
echo "Validation"
echo "================================================"
echo ""
echo "Validating manifest.xml..."
if npm run validate; then
    echo "✅ Manifest is valid"
else
    echo "⚠️  Warning: Manifest validation issues detected"
    echo "   Review and fix any errors shown above"
fi

echo ""
echo "================================================"
echo "Setup Complete!"
echo "================================================"
echo ""
echo "Next steps:"
echo ""
echo "1. Configure Azure AD:"
echo "   - Go to https://portal.azure.com"
echo "   - Create an App Registration"
echo "   - Copy the Client ID to .env file"
echo "   - Add API permissions (see AZURE-SETUP.md)"
echo ""
echo "2. Configure SharePoint:"
echo "   - Update SHAREPOINT_SITE_URL in .env"
echo "   - Ensure you have access to the SharePoint site"
echo ""
echo "3. Update manifest.xml URLs:"
echo "   - For production, replace localhost:3000 with your domain"
echo "   - Update icon URLs"
echo "   - Update support URL"
echo ""
echo "4. Start development server:"
echo "   npm run dev-server"
echo ""
echo "5. Test in Outlook:"
echo "   - Outlook Web: Get Add-ins > Upload manifest"
echo "   - Outlook Desktop: npm start"
echo ""
echo "For detailed instructions, see:"
echo "  - README.md"
echo "  - GETTING-STARTED.md"
echo "  - AZURE-SETUP.md"
echo "  - CONFIGURATION.md"
echo ""
echo "Need help? support@onecoachfitness.com"
echo ""
