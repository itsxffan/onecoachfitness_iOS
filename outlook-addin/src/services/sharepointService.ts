/**
 * SharePoint Service
 * Handles integration with SharePoint document libraries
 */

import { AuthService } from './authService';

export interface MessageContent {
  subject: string;
  from: string;
  date: string;
  body: string;
}

export interface Attachment {
  name: string;
  id: string;
  contentType: string;
  size: number;
}

export class SharePointService {
  private authService: AuthService;
  private siteUrl: string;
  
  constructor() {
    this.authService = new AuthService();
    // This should be configured in environment or config file
    this.siteUrl = 'https://yourtenant.sharepoint.com/sites/yoursite';
  }
  
  /**
   * Push message content and attachments to SharePoint library
   */
  async pushToLibrary(
    libraryName: string,
    content: MessageContent,
    attachments: Attachment[]
  ): Promise<void> {
    try {
      const graphClient = await this.authService.getGraphClient();
      
      // Create document in SharePoint
      const fileName = this.sanitizeFileName(`${content.subject}_${new Date().getTime()}.txt`);
      const fileContent = this.formatMessageContent(content);
      
      // Upload main content as text file
      await this.uploadFile(graphClient, libraryName, fileName, fileContent);
      
      // Upload attachments if any
      if (attachments.length > 0) {
        await this.uploadAttachments(graphClient, libraryName, attachments);
      }
      
      console.log('Successfully pushed to SharePoint');
    } catch (error) {
      console.error('Error pushing to SharePoint:', error);
      throw new Error('Failed to push to SharePoint: ' + (error as Error).message);
    }
  }
  
  /**
   * Upload a file to SharePoint
   */
  private async uploadFile(
    graphClient: any,
    libraryName: string,
    fileName: string,
    content: string
  ): Promise<void> {
    try {
      const driveId = await this.getDriveId(graphClient, libraryName);
      
      // Upload file content
      await graphClient
        .api(`/drives/${driveId}/root:/${fileName}:/content`)
        .put(content);
      
      console.log(`File ${fileName} uploaded successfully`);
    } catch (error) {
      console.error('Error uploading file:', error);
      throw error;
    }
  }
  
  /**
   * Upload attachments to SharePoint
   */
  private async uploadAttachments(
    graphClient: any,
    libraryName: string,
    attachments: Attachment[]
  ): Promise<void> {
    try {
      const driveId = await this.getDriveId(graphClient, libraryName);
      
      for (const attachment of attachments) {
        // Get attachment content from Outlook
        const attachmentContent = await this.getAttachmentContent(attachment.id);
        
        // Upload to SharePoint
        const fileName = this.sanitizeFileName(attachment.name);
        await graphClient
          .api(`/drives/${driveId}/root:/Attachments/${fileName}:/content`)
          .put(attachmentContent);
        
        console.log(`Attachment ${fileName} uploaded successfully`);
      }
    } catch (error) {
      console.error('Error uploading attachments:', error);
      throw error;
    }
  }
  
  /**
   * Get attachment content from Outlook
   */
  private async getAttachmentContent(attachmentId: string): Promise<any> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error('No message item available'));
        return;
      }
      
      item.getAttachmentContentAsync(attachmentId, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.content);
        } else {
          reject(new Error('Failed to get attachment content'));
        }
      });
    });
  }
  
  /**
   * Get SharePoint drive ID for a library
   */
  private async getDriveId(graphClient: any, libraryName: string): Promise<string> {
    try {
      // Parse site URL to get site path
      const sitePathMatch = this.siteUrl.match(/\/sites\/([^\/]+)/);
      if (!sitePathMatch) {
        throw new Error('Invalid site URL');
      }
      const sitePath = sitePathMatch[1];
      
      // Get site information
      const site = await graphClient
        .api(`/sites/root:/sites/${sitePath}`)
        .get();
      
      // Get document library by name
      const drives = await graphClient
        .api(`/sites/${site.id}/drives`)
        .get();
      
      const targetDrive = drives.value.find((drive: any) => 
        drive.name === libraryName || drive.name === `${libraryName}`
      );
      
      if (!targetDrive) {
        // If not found by name, use default documents library
        return drives.value[0].id;
      }
      
      return targetDrive.id;
    } catch (error) {
      console.error('Error getting drive ID:', error);
      throw new Error('Failed to locate SharePoint library');
    }
  }
  
  /**
   * Format message content as readable text
   */
  private formatMessageContent(content: MessageContent): string {
    return `
Subject: ${content.subject}
From: ${content.from}
Date: ${content.date}

---

${content.body}
    `.trim();
  }
  
  /**
   * Sanitize file name to be SharePoint-compatible
   */
  private sanitizeFileName(fileName: string): string {
    // Remove invalid characters for SharePoint
    return fileName
      .replace(/[~#%&*{}\\:<>?/|"]/g, '_')
      .replace(/\s+/g, '_')
      .substring(0, 128); // Limit length
  }
  
  /**
   * Configure SharePoint site URL
   */
  setSiteUrl(url: string): void {
    this.siteUrl = url;
  }
  
  /**
   * Get list of available document libraries
   */
  async getAvailableLibraries(): Promise<string[]> {
    try {
      const graphClient = await this.authService.getGraphClient();
      
      const sitePathMatch = this.siteUrl.match(/\/sites\/([^\/]+)/);
      if (!sitePathMatch) {
        throw new Error('Invalid site URL');
      }
      const sitePath = sitePathMatch[1];
      
      const site = await graphClient
        .api(`/sites/root:/sites/${sitePath}`)
        .get();
      
      const drives = await graphClient
        .api(`/sites/${site.id}/drives`)
        .get();
      
      return drives.value.map((drive: any) => drive.name);
    } catch (error) {
      console.error('Error getting libraries:', error);
      return ['ClientDocuments', 'TrainingMaterials', 'GeneralArchive'];
    }
  }
}
