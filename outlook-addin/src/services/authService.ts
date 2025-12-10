/**
 * Authentication Service
 * Handles Azure AD authentication and Microsoft Graph API token management
 */

import { Client } from '@microsoft/microsoft-graph-client';

export class AuthService {
  private accessToken: string | null = null;
  private graphClient: Client | null = null;
  
  constructor() {
    // Initialize with stored token if available
    this.accessToken = this.getStoredToken();
  }
  
  /**
   * Check if user is authenticated
   */
  async isAuthenticated(): Promise<boolean> {
    if (this.accessToken) {
      // Verify token is still valid
      try {
        await this.getGraphClient();
        return true;
      } catch (error) {
        console.error('Token validation failed:', error);
        this.clearToken();
        return false;
      }
    }
    return false;
  }
  
  /**
   * Perform SSO login using Office dialog API
   */
  async login(): Promise<void> {
    return new Promise((resolve, reject) => {
      // Use Office SSO API for authentication
      Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          this.accessToken = result.value;
          this.storeToken(result.value);
          this.initializeGraphClient();
          resolve();
        } else {
          // Fall back to dialog-based authentication
          this.loginWithDialog()
            .then(resolve)
            .catch(reject);
        }
      });
    });
  }
  
  /**
   * Login using dialog-based authentication (fallback)
   */
  private async loginWithDialog(): Promise<void> {
    return new Promise((resolve, reject) => {
      const dialogUrl = this.buildAuthUrl();
      
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 30 },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const dialog = result.value;
            
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
              dialog.close();
              
              if (arg.message) {
                try {
                  const response = JSON.parse(arg.message);
                  if (response.access_token) {
                    this.accessToken = response.access_token;
                    this.storeToken(response.access_token);
                    this.initializeGraphClient();
                    resolve();
                  } else {
                    reject(new Error('No access token received'));
                  }
                } catch (error) {
                  reject(error);
                }
              }
            });
          } else {
            reject(new Error('Failed to open authentication dialog'));
          }
        }
      );
    });
  }
  
  /**
   * Build Azure AD authentication URL
   */
  private buildAuthUrl(): string {
    // Load from environment or config - MUST be configured before deployment
    const clientId = process.env.AZURE_AD_CLIENT_ID;
    const tenantId = process.env.AZURE_AD_TENANT_ID || 'common';
    
    if (!clientId || clientId === 'CONFIGURE_BEFORE_DEPLOYMENT') {
      throw new Error('Azure AD Client ID not configured. Please set AZURE_AD_CLIENT_ID environment variable.');
    }
    
    const redirectUri = encodeURIComponent(window.location.origin + '/auth-callback.html');
    const scope = encodeURIComponent('https://graph.microsoft.com/.default');
    
    return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
           `client_id=${clientId}&` +
           `response_type=token&` +
           `redirect_uri=${redirectUri}&` +
           `scope=${scope}&` +
           `response_mode=fragment`;
  }
  
  /**
   * Logout and clear tokens
   */
  logout(): void {
    this.accessToken = null;
    this.graphClient = null;
    this.clearToken();
  }
  
  /**
   * Get authenticated Microsoft Graph client
   */
  async getGraphClient(): Promise<Client> {
    if (!this.graphClient) {
      this.initializeGraphClient();
    }
    
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }
    
    return this.graphClient;
  }
  
  /**
   * Initialize Microsoft Graph client
   */
  private initializeGraphClient(): void {
    if (!this.accessToken) {
      throw new Error('No access token available');
    }
    
    this.graphClient = Client.init({
      authProvider: (done) => {
        done(null, this.accessToken!);
      }
    });
  }
  
  /**
   * Store token in session storage
   */
  private storeToken(token: string): void {
    try {
      sessionStorage.setItem('ms_graph_token', token);
    } catch (error) {
      console.error('Error storing token:', error);
    }
  }
  
  /**
   * Get stored token from session storage
   */
  private getStoredToken(): string | null {
    try {
      return sessionStorage.getItem('ms_graph_token');
    } catch (error) {
      console.error('Error retrieving token:', error);
      return null;
    }
  }
  
  /**
   * Clear stored token
   */
  private clearToken(): void {
    try {
      sessionStorage.removeItem('ms_graph_token');
    } catch (error) {
      console.error('Error clearing token:', error);
    }
  }
  
  /**
   * Get current access token
   */
  getAccessToken(): string | null {
    return this.accessToken;
  }
}
