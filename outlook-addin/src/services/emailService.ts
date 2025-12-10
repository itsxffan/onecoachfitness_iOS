/**
 * Email Service
 * Handles email operations like tagging, archiving, and categorization
 */

export class EmailService {
  private customTags: Map<string, string[]>;
  
  constructor() {
    this.customTags = new Map();
    this.loadCustomTags();
  }
  
  /**
   * Archive the current message
   */
  async archiveMessage(): Promise<void> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error('No message selected'));
        return;
      }
      
      // Move message to Archive folder
      const archiveFolderId = this.getArchiveFolderId();
      
      if (archiveFolderId) {
        item.move(archiveFolderId, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('Message archived successfully');
            resolve();
          } else {
            reject(new Error('Failed to archive message'));
          }
        });
      } else {
        // Fallback: Add "Archived" category
        this.addCategory('Archived')
          .then(resolve)
          .catch(reject);
      }
    });
  }
  
  /**
   * Add a category to the current message
   */
  async addCategory(category: string): Promise<void> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item || !item.categories) {
        reject(new Error('Categories not available'));
        return;
      }
      
      item.categories.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const currentCategories = result.value;
          if (!currentCategories.includes(category)) {
            const newCategories = [...currentCategories, category];
            item.categories.setAsync(newCategories, (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
              } else {
                reject(new Error('Failed to set category'));
              }
            });
          } else {
            resolve(); // Category already exists
          }
        } else {
          reject(new Error('Failed to get categories'));
        }
      });
    });
  }
  
  /**
   * Remove a category from the current message
   */
  async removeCategory(category: string): Promise<void> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item || !item.categories) {
        reject(new Error('Categories not available'));
        return;
      }
      
      item.categories.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const currentCategories = result.value;
          const newCategories = currentCategories.filter(c => c !== category);
          
          item.categories.setAsync(newCategories, (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(new Error('Failed to remove category'));
            }
          });
        } else {
          reject(new Error('Failed to get categories'));
        }
      });
    });
  }
  
  /**
   * Add a custom tag to the current message
   */
  async addCustomTag(tag: string): Promise<void> {
    const item = Office.context.mailbox.item;
    if (!item) {
      throw new Error('No message selected');
    }
    
    const messageId = item.itemId;
    if (!messageId) {
      throw new Error('Message ID not available');
    }
    
    // Store custom tags in local storage
    let tags = this.customTags.get(messageId) || [];
    if (!tags.includes(tag)) {
      tags.push(tag);
      this.customTags.set(messageId, tags);
      this.saveCustomTags();
    }
    
    // Also add as a category for visibility
    return this.addCategory(`Tag: ${tag}`);
  }
  
  /**
   * Remove a custom tag from the current message
   */
  async removeCustomTag(tag: string): Promise<void> {
    const item = Office.context.mailbox.item;
    if (!item) {
      throw new Error('No message selected');
    }
    
    const messageId = item.itemId;
    if (!messageId) {
      throw new Error('Message ID not available');
    }
    
    // Remove from custom tags storage
    let tags = this.customTags.get(messageId) || [];
    tags = tags.filter(t => t !== tag);
    this.customTags.set(messageId, tags);
    this.saveCustomTags();
    
    // Remove category
    return this.removeCategory(`Tag: ${tag}`);
  }
  
  /**
   * Get custom tags for the current message
   */
  getCustomTags(): string[] {
    const item = Office.context.mailbox.item;
    if (!item || !item.itemId) {
      return [];
    }
    
    return this.customTags.get(item.itemId) || [];
  }
  
  /**
   * Get message content
   */
  async getMessageContent(): Promise<string> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error('No message selected'));
        return;
      }
      
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error('Failed to get message content'));
        }
      });
    });
  }
  
  /**
   * Get message HTML content
   */
  async getMessageHtmlContent(): Promise<string> {
    return new Promise((resolve, reject) => {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error('No message selected'));
        return;
      }
      
      item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error('Failed to get message HTML content'));
        }
      });
    });
  }
  
  /**
   * Get Archive folder ID (if available)
   */
  private getArchiveFolderId(): string | null {
    try {
      // Try to get the Archive folder
      const mailbox = Office.context.mailbox;
      
      // This is a simplified version - in production, you would need to
      // use EWS or Graph API to find the Archive folder
      return null; // Will use category fallback
    } catch (error) {
      console.error('Error getting archive folder:', error);
      return null;
    }
  }
  
  /**
   * Save custom tags to local storage
   */
  private saveCustomTags(): void {
    try {
      const tagsObject: { [key: string]: string[] } = {};
      this.customTags.forEach((value, key) => {
        tagsObject[key] = value;
      });
      localStorage.setItem('custom_email_tags', JSON.stringify(tagsObject));
    } catch (error) {
      console.error('Error saving custom tags:', error);
    }
  }
  
  /**
   * Load custom tags from local storage
   */
  private loadCustomTags(): void {
    try {
      const stored = localStorage.getItem('custom_email_tags');
      if (stored) {
        const tagsObject = JSON.parse(stored);
        this.customTags = new Map(Object.entries(tagsObject));
      }
    } catch (error) {
      console.error('Error loading custom tags:', error);
    }
  }
}
