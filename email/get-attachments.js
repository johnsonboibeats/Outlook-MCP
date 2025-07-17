/**
 * Gets information about email attachments
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Gets attachment information from an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetAttachments(args) {
  const { emailId, sharedMailbox } = args;
  
  if (!emailId) {
    return {
      content: [{ 
        type: "text", 
        text: "Email ID is required to get attachments."
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    // Build the endpoint - handle shared mailbox
    let endpoint = `me/messages/${emailId}/attachments`;
    if (sharedMailbox) {
      endpoint = `users/${sharedMailbox}/messages/${emailId}/attachments`;
    }
    
    // Get attachments from the email
    const response = await callGraphAPI(accessToken, 'GET', endpoint);
    const attachments = response.value || [];
    
    if (attachments.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No attachments found in this email."
        }]
      };
    }
    
    // Format attachment information
    const attachmentInfo = attachments.map(attachment => {
      const sizeKB = Math.round(attachment.size / 1024);
      return `📎 ${attachment.name}\n   Type: ${attachment.contentType}\n   Size: ${sizeKB} KB\n   ID: ${attachment.id}`;
    });
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${attachments.length} attachment(s):\n\n${attachmentInfo.join('\n\n')}\n\nNote: This tool shows attachment metadata. To download attachments, you would need additional implementation with file system access.`
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
        text: `Error getting attachments: ${error.message}`
      }]
    };
  }
}

module.exports = handleGetAttachments;