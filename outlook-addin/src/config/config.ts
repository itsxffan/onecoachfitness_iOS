/**
 * Configuration file for the add-in
 * Update these values for your environment
 */

export const config = {
  // Azure AD Configuration
  azureAd: {
    clientId: 'YOUR_AZURE_AD_CLIENT_ID',
    tenantId: 'common', // or your specific tenant ID
    redirectUri: window.location.origin + '/auth-callback.html',
    scopes: [
      'User.Read',
      'Mail.ReadWrite',
      'Sites.ReadWrite.All',
      'Files.ReadWrite.All'
    ]
  },
  
  // SharePoint Configuration
  sharepoint: {
    siteUrl: 'https://yourtenant.sharepoint.com/sites/yoursite',
    defaultLibraries: [
      { id: 'ClientDocuments', name: 'Client Documents' },
      { id: 'TrainingMaterials', name: 'Training Materials' },
      { id: 'GeneralArchive', name: 'General Archive' }
    ]
  },
  
  // Email Categories
  categories: [
    { id: 'Important', name: 'Important', color: 'red' },
    { id: 'Follow Up', name: 'Follow Up', color: 'yellow' },
    { id: 'Client', name: 'Client', color: 'blue' },
    { id: 'Training', name: 'Training', color: 'green' }
  ],
  
  // Application Settings
  app: {
    name: 'OneCoach Email Manager',
    version: '1.0.0',
    supportEmail: 'support@onecoachfitness.com',
    documentationUrl: 'https://docs.onecoachfitness.com/outlook-addin'
  },
  
  // Feature Flags
  features: {
    enableSharePointIntegration: true,
    enableCustomTags: true,
    enableArchive: true,
    enableCategorySync: true
  },
  
  // API Configuration (if using backend)
  api: {
    baseUrl: process.env.API_BASE_URL || '',
    timeout: 30000, // 30 seconds
    retryAttempts: 3
  },
  
  // Logging Configuration
  logging: {
    enableConsoleLogging: true,
    enableErrorTracking: true,
    logLevel: 'info' // 'debug', 'info', 'warn', 'error'
  }
};

export default config;
