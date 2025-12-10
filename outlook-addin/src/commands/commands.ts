/**
 * Commands
 * Function commands that can be triggered from the Outlook ribbon
 */

import { EmailService } from '../services/emailService';

Office.onReady(() => {
  console.log('Commands initialized');
});

/**
 * Archive message function (called from ribbon button)
 */
function archiveMessage(event: Office.AddinCommands.Event) {
  const emailService = new EmailService();
  
  emailService.archiveMessage()
    .then(() => {
      // Show notification
      Office.context.mailbox.item?.notificationMessages.addAsync(
        'archiveSuccess',
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: 'Message archived successfully',
          icon: 'icon-16',
          persistent: false
        }
      );
      
      event.completed();
    })
    .catch((error) => {
      console.error('Error archiving message:', error);
      
      // Show error notification
      Office.context.mailbox.item?.notificationMessages.addAsync(
        'archiveError',
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: 'Failed to archive message',
          icon: 'icon-16',
          persistent: false
        }
      );
      
      event.completed();
    });
}

// Register function commands
(Office as any).actions.associate('archiveMessage', archiveMessage);
