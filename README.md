# OneCoach Fitness

A comprehensive fitness coaching platform with both iOS mobile app and Outlook email management add-in.

## Repository Contents

This repository contains two main components:

### 1. iOS Fitness App (`onecoachfitness/`)

A native iOS application for fitness coaching and gym location services.

**Features:**
- Gym location finder with map integration
- Exercise tracking and management
- User profile and settings
- Google Maps and Places integration

**Technology:**
- Swift/SwiftUI
- Google Maps SDK
- Google Places SDK
- CocoaPods dependency management

**Getting Started with iOS App:**
```bash
# Install dependencies
pod install

# Open workspace
open onecoachfitness.xcworkspace
```

### 2. Outlook Email Manager Add-in (`outlook-addin/`)

A modern Office JavaScript add-in for smarter email management with SharePoint integration.

**Features:**
- Email categorization with predefined categories
- Custom tagging system
- One-click message archiving
- SharePoint document library integration
- Azure AD authentication with SSO
- Cross-platform support (Desktop, Web, Mac)

**Technology:**
- TypeScript
- Office.js API
- Microsoft Graph API
- Azure AD OAuth 2.0
- Webpack build system

**Getting Started with Outlook Add-in:**
```bash
cd outlook-addin
npm install
npm run dev-server
```

**Quick Setup:**
```bash
cd outlook-addin
./setup.sh  # Automated setup script
```

## Documentation

### iOS App Documentation
- Podfile - Dependency configuration
- APIKeyDetails/ - API key management

### Outlook Add-in Documentation

**Essential Guides:**
- [📖 README.md](./outlook-addin/README.md) - Complete documentation (10,000+ words)
- [🚀 GETTING-STARTED.md](./outlook-addin/GETTING-STARTED.md) - Quick start guide (8,000+ words)
- [👤 USER-GUIDE.md](./outlook-addin/USER-GUIDE.md) - End-user instructions (10,000+ words)

**Administration:**
- [🚀 DEPLOYMENT.md](./outlook-addin/DEPLOYMENT.md) - Deployment guide (10,000+ words)
- [☁️ AZURE-SETUP.md](./outlook-addin/AZURE-SETUP.md) - Azure configuration (13,000+ words)
- [⚙️ CONFIGURATION.md](./outlook-addin/CONFIGURATION.md) - Configuration guide (10,000+ words)

**Reference:**
- [✨ FEATURES.md](./outlook-addin/FEATURES.md) - Feature overview (10,000+ words)
- [🔧 TROUBLESHOOTING.md](./outlook-addin/TROUBLESHOOTING.md) - Problem solving (12,000+ words)
- [📝 CHANGELOG.md](./outlook-addin/CHANGELOG.md) - Version history
- [📋 IMPLEMENTATION-SUMMARY.md](./outlook-addin/IMPLEMENTATION-SUMMARY.md) - Technical summary

## Project Structure

```
onecoachfitness_iOS/
├── onecoachfitness/           # iOS fitness app
│   ├── AppDelegate.swift
│   ├── Model/
│   ├── View/
│   └── SceneDelegate.swift
├── onecoachfitness.xcodeproj/ # Xcode project
├── onecoachfitness.xcworkspace/
├── Pods/                      # iOS dependencies
├── Podfile                    # iOS dependency config
├── APIKeyDetails/             # API keys (not in git)
│
└── outlook-addin/             # Outlook email manager
    ├── src/
    │   ├── taskpane/         # Main UI
    │   ├── commands/         # Ribbon commands
    │   ├── services/         # Business logic
    │   └── config/           # Configuration
    ├── assets/               # Icons and images
    ├── dist/                 # Build output (generated)
    ├── manifest.xml          # Add-in manifest
    ├── package.json          # Dependencies
    ├── tsconfig.json         # TypeScript config
    ├── webpack.config.js     # Build config
    ├── setup.sh              # Automated setup script
    └── [Documentation]/      # Comprehensive guides
```

## Quick Start

### For iOS Development

```bash
# Clone repository
git clone https://github.com/itsxffan/onecoachfitness_iOS.git
cd onecoachfitness_iOS

# Install CocoaPods dependencies
pod install

# Open in Xcode
open onecoachfitness.xcworkspace
```

### For Outlook Add-in Development

```bash
# Navigate to add-in directory
cd outlook-addin

# Run automated setup
./setup.sh

# Or manually:
npm install
npm run dev-server

# Open in Outlook and sideload manifest.xml
```

## Technologies Used

### iOS App
- **Language**: Swift
- **UI**: SwiftUI
- **Maps**: Google Maps SDK
- **Location**: Google Places SDK
- **Dependency Manager**: CocoaPods

### Outlook Add-in
- **Language**: TypeScript
- **Framework**: Office.js
- **API**: Microsoft Graph
- **Authentication**: Azure AD OAuth 2.0
- **UI**: Fluent UI principles
- **Build**: Webpack 5
- **Package Manager**: npm

## Development Requirements

### iOS App Requirements
- macOS 10.14+
- Xcode 12.0+
- Swift 5+
- CocoaPods
- Google Maps API key
- Google Places API key

### Outlook Add-in Requirements
- Node.js 16.x or higher
- npm or yarn
- Microsoft 365 account
- Azure subscription (for Azure AD)
- SharePoint site access

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Code Style

### iOS (Swift)
- Follow Swift API Design Guidelines
- Use SwiftLint for code consistency
- Comprehensive documentation comments

### Outlook Add-in (TypeScript)
- ESLint configured with TypeScript rules
- Strict TypeScript mode enabled
- Comprehensive JSDoc comments

## Testing

### iOS App
```bash
# Run tests in Xcode
# Product > Test (Cmd+U)
```

### Outlook Add-in
```bash
cd outlook-addin

# Validate manifest
npm run validate

# Lint code
npm run lint

# Build
npm run build
```

## Deployment

### iOS App Deployment
- Follow Apple's App Store guidelines
- Use Xcode Archive and Distribution
- See Apple Developer documentation

### Outlook Add-in Deployment

**For Organizations (Tenant-wide):**
1. Configure Azure AD app registration
2. Build production assets
3. Deploy to hosting service
4. Update manifest with production URLs
5. Deploy via Microsoft 365 admin center

**For Individual Users:**
1. Sideload via "Get Add-ins" in Outlook
2. Upload manifest.xml file

See [DEPLOYMENT.md](./outlook-addin/DEPLOYMENT.md) for detailed instructions.

## License

MIT License - see [LICENSE](./outlook-addin/LICENSE) file for details.

## Support

### iOS App Support
- Issues: Use GitHub Issues tab
- Email: support@onecoachfitness.com

### Outlook Add-in Support
- Documentation: See `outlook-addin/` folder
- Email: support@onecoachfitness.com
- Issues: Use GitHub Issues tab with "outlook-addin" label

## Acknowledgments

### iOS App
- Google Maps Platform
- Google Places API
- CocoaPods community

### Outlook Add-in
- Microsoft Office Add-ins team
- Microsoft Graph API
- Fluent UI design system
- TypeScript and Webpack communities

## Roadmap

### iOS App
- [ ] Social features for fitness tracking
- [ ] Workout plan creation
- [ ] Integration with health apps
- [ ] Personal trainer matching

### Outlook Add-in
- [ ] Batch email operations
- [ ] Advanced search and filtering
- [ ] Mobile platform support
- [ ] AI-powered email classification
- [ ] Teams and OneNote integration
- [ ] Dark mode support
- [ ] Localization (multiple languages)

## Version History

### iOS App
- Current version: See Xcode project version

### Outlook Add-in
- **v1.0.0** (December 2024) - Initial release
  - Email categorization
  - Custom tagging
  - Archive functionality
  - SharePoint integration
  - Azure AD authentication
  - Cross-platform support

See [CHANGELOG.md](./outlook-addin/CHANGELOG.md) for detailed version history.

## Security

### Reporting Security Issues
If you discover a security vulnerability:
- **DO NOT** create a public GitHub issue
- Email: security@onecoachfitness.com
- Include detailed description and reproduction steps

### Security Features

**iOS App:**
- API keys not stored in repository
- Secure communication with backend services

**Outlook Add-in:**
- Azure AD OAuth 2.0 authentication
- HTTPS-only communication
- No sensitive data in code
- Regular dependency updates
- CodeQL security scanning

## Additional Resources

### iOS Development
- [Apple Developer Documentation](https://developer.apple.com/documentation/)
- [Swift Programming Language](https://swift.org/documentation/)
- [SwiftUI Tutorials](https://developer.apple.com/tutorials/swiftui)

### Outlook Add-in Development
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Microsoft Graph Documentation](https://docs.microsoft.com/graph/)
- [Azure AD Documentation](https://docs.microsoft.com/azure/active-directory/)
- [Fluent UI Components](https://developer.microsoft.com/fluentui)

## Contact

- **Website**: https://www.onecoachfitness.com
- **Email**: support@onecoachfitness.com
- **GitHub**: https://github.com/itsxffan/onecoachfitness_iOS

---

**OneCoach Fitness** - Empowering fitness coaching through technology 💪

Made with ❤️ for coaches and fitness enthusiasts
