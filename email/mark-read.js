/**
 * Marks emails as read or unread
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Marks emails as read or unread
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMarkRead(args) {
  const { emailIds, isRead = true, sharedMailbox } = args;
  
  if (!emailIds) {
    return {
      content: [{ 
        type: "text", 
        text: "Email IDs are required. Provide a comma-separated list of email IDs."
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    const emailIdList = emailIds.split(',').map(id => id.trim()).filter(id => id);
    
    if (emailIdList.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No valid email IDs provided."
        }]
      };
    }
    
    const results = [];
    
    // Process each email ID
    for (const emailId of emailIdList) {
      try {
        // Build the endpoint - handle shared mailbox
        let endpoint = `me/messages/${emailId}`;
        if (sharedMailbox) {
          endpoint = `users/${sharedMailbox}/messages/${emailId}`;
        }
        
        // Update the email's read status
        await callGraphAPI(
          accessToken,
          'PATCH',
          endpoint,
          {
            isRead: isRead
          }
        );
        
        results.push(`✓ ${emailId}: marked as ${isRead ? 'read' : 'unread'}`);
      } catch (error) {
        results.push(`✗ ${emailId}: ${error.message}`);
      }
    }
    
    const successCount = results.filter(r => r.startsWith('✓')).length;
    const failureCount = results.filter(r => r.startsWith('✗')).length;
    
    return {
      content: [{ 
        type: "text", 
        text: `Email read status update completed!\n\nSuccessful: ${successCount}\nFailed: ${failureCount}\n\n${results.join('\n')}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error updating email read status: ${error.message}`
      }]
    };
  }
}

module.exports = handleMarkRead;