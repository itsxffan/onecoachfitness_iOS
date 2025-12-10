import './taskpane.css';
import { AuthService } from '../services/authService';
import { SharePointService } from '../services/sharepointService';
import { EmailService } from '../services/emailService';

let authService: AuthService;
let sharepointService: SharePointService;
let emailService: EmailService;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('Office Add-in initialized');
    
    // Initialize services
    authService = new AuthService();
    sharepointService = new SharePointService();
    emailService = new EmailService();
    
    // Set up event listeners
    document.getElementById('loginButton')?.addEventListener('click', handleLogin);
    document.getElementById('archiveButton')?.addEventListener('click', handleArchive);
    document.getElementById('addTagButton')?.addEventListener('click', handleAddTag);
    document.getElementById('pushToSharePointButton')?.addEventListener('click', handlePushToSharePoint);
    
    // Category buttons
    document.querySelectorAll('.category-btn').forEach(btn => {
      btn.addEventListener('click', (e) => handleCategoryClick(e.target as HTMLElement));
    });
    
    // Check if user is already logged in
    checkAuthStatus();
  }
});

async function checkAuthStatus() {
  try {
    const isAuthenticated = await authService.isAuthenticated();
    if (isAuthenticated) {
      showMainContent();
      await loadMessageDetails();
    } else {
      showAuthSection();
    }
  } catch (error) {
    console.error('Error checking auth status:', error);
    showAuthSection();
  }
}

async function handleLogin() {
  try {
    showStatus('Signing in...', 'info');
    await authService.login();
    showMainContent();
    await loadMessageDetails();
    showStatus('Successfully signed in!', 'success');
  } catch (error) {
    console.error('Login error:', error);
    showStatus('Failed to sign in. Please try again.', 'error');
  }
}

async function loadMessageDetails() {
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      showStatus('No message selected', 'error');
      return;
    }
    
    const messageDetails = document.getElementById('messageDetails');
    if (messageDetails) {
      messageDetails.innerHTML = `
        <p><strong>Subject:</strong> ${item.subject || 'No subject'}</p>
        <p><strong>From:</strong> ${item.from?.displayName || 'Unknown'}</p>
        <p><strong>Received:</strong> ${item.dateTimeCreated?.toLocaleString() || 'Unknown'}</p>
      `;
    }
    
    // Load existing categories
    if (item.categories) {
      item.categories.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCategoryButtons(result.value);
        }
      });
    }
  } catch (error) {
    console.error('Error loading message details:', error);
    showStatus('Error loading message details', 'error');
  }
}

function updateCategoryButtons(categories: string[]) {
  document.querySelectorAll('.category-btn').forEach(btn => {
    const categoryName = (btn as HTMLElement).dataset.category;
    if (categories.includes(categoryName || '')) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });
}

async function handleCategoryClick(button: HTMLElement) {
  const category = button.dataset.category;
  if (!category) return;
  
  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.categories) return;
    
    item.categories.getAsync(async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const currentCategories = result.value;
        let newCategories: string[];
        
        if (currentCategories.includes(category)) {
          // Remove category
          newCategories = currentCategories.filter(c => c !== category);
          button.classList.remove('active');
        } else {
          // Add category
          newCategories = [...currentCategories, category];
          button.classList.add('active');
        }
        
        item.categories.setAsync(newCategories, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            showStatus(`Category "${category}" updated`, 'success');
          } else {
            showStatus(`Failed to update category`, 'error');
          }
        });
      }
    });
  } catch (error) {
    console.error('Error handling category:', error);
    showStatus('Error updating category', 'error');
  }
}

async function handleAddTag() {
  const tagInput = document.getElementById('tagInput') as HTMLInputElement;
  const tagValue = tagInput?.value.trim();
  
  if (!tagValue) {
    showStatus('Please enter a tag', 'error');
    return;
  }
  
  try {
    await emailService.addCustomTag(tagValue);
    addTagToList(tagValue);
    tagInput.value = '';
    showStatus('Tag added successfully', 'success');
  } catch (error) {
    console.error('Error adding tag:', error);
    showStatus('Error adding tag', 'error');
  }
}

function addTagToList(tag: string) {
  const tagsList = document.getElementById('tagsList');
  if (!tagsList) return;
  
  const tagElement = document.createElement('div');
  tagElement.className = 'tag-item';
  tagElement.innerHTML = `
    ${tag}
    <span class="tag-remove" data-tag="${tag}">&times;</span>
  `;
  
  tagElement.querySelector('.tag-remove')?.addEventListener('click', (e) => {
    const tagToRemove = (e.target as HTMLElement).dataset.tag;
    if (tagToRemove) {
      emailService.removeCustomTag(tagToRemove);
      tagElement.remove();
      showStatus('Tag removed', 'success');
    }
  });
  
  tagsList.appendChild(tagElement);
}

async function handleArchive() {
  try {
    showStatus('Archiving message...', 'info');
    await emailService.archiveMessage();
    showStatus('Message archived successfully!', 'success');
  } catch (error) {
    console.error('Error archiving message:', error);
    showStatus('Error archiving message', 'error');
  }
}

async function handlePushToSharePoint() {
  const librarySelect = document.getElementById('librarySelect') as HTMLSelectElement;
  const includeAttachments = (document.getElementById('includeAttachments') as HTMLInputElement).checked;
  const selectedLibrary = librarySelect?.value;
  
  if (!selectedLibrary) {
    showStatus('Please select a document library', 'error');
    return;
  }
  
  try {
    showStatus('Pushing to SharePoint...', 'info');
    
    const item = Office.context.mailbox.item;
    if (!item) {
      showStatus('No message selected', 'error');
      return;
    }
    
    // Get message body
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const messageContent = {
          subject: item.subject || 'No Subject',
          from: item.from?.displayName || 'Unknown',
          date: item.dateTimeCreated?.toISOString() || new Date().toISOString(),
          body: result.value
        };
        
        // Push to SharePoint
        await sharepointService.pushToLibrary(
          selectedLibrary,
          messageContent,
          includeAttachments ? await getAttachments() : []
        );
        
        showStatus('Successfully pushed to SharePoint!', 'success');
      } else {
        showStatus('Error reading message content', 'error');
      }
    });
  } catch (error) {
    console.error('Error pushing to SharePoint:', error);
    showStatus('Error pushing to SharePoint', 'error');
  }
}

async function getAttachments(): Promise<any[]> {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.attachments || item.attachments.length === 0) {
      resolve([]);
      return;
    }
    
    const attachments = item.attachments.map(att => ({
      name: att.name,
      id: att.id,
      contentType: att.contentType,
      size: att.size
    }));
    
    resolve(attachments);
  });
}

function showAuthSection() {
  const authSection = document.getElementById('authSection');
  const mainContent = document.getElementById('mainContent');
  if (authSection) authSection.style.display = 'block';
  if (mainContent) mainContent.style.display = 'none';
}

function showMainContent() {
  const authSection = document.getElementById('authSection');
  const mainContent = document.getElementById('mainContent');
  if (authSection) authSection.style.display = 'none';
  if (mainContent) mainContent.style.display = 'block';
}

function showStatus(message: string, type: 'success' | 'error' | 'info') {
  const statusSection = document.getElementById('statusSection');
  if (!statusSection) return;
  
  const statusDiv = document.createElement('div');
  statusDiv.className = `status-message status-${type}`;
  statusDiv.textContent = message;
  
  statusSection.innerHTML = '';
  statusSection.appendChild(statusDiv);
  
  // Auto-remove after 5 seconds
  setTimeout(() => {
    statusDiv.remove();
  }, 5000);
}
